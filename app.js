// ============================================================
//  Job Card Manager — app.js
// ============================================================

// ---------- The Apps Script code that will be shown ----------
const APPS_SCRIPT_CODE = `// =====================================================
//  Job Card Manager — Google Apps Script (Code.gs)
//  Paste this into Extensions > Apps Script in your
//  Google Sheet, then deploy as a Web App.
// =====================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    if (action === "insert") return insertRow(data);
    if (action === "test")   return jsonResponse({ success: true, message: "Connection successful!" });
    return jsonResponse({ success: false, message: "Unknown action" });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function doGet(e) {
  if (e.parameter.action === "test")
    return jsonResponse({ success: true, message: "Connection successful!" });
  return jsonResponse({ success: false, message: "Use POST for data" });
}

function insertRow(data) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = data.sheetName || "Sheet1";
  var sheet     = ss.getSheetByName(sheetName);

  if (!sheet)
    return jsonResponse({ success: false, message: "Sheet '" + sheetName + "' not found." });

  // --- Columns: A=SL.NO  B=JOB CARD NO.  C=JC NO.  D=DATE  E=DESCRIPTION  F=AMOUNT ---
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    sheet.appendRow(["SL.NO", "JOB CARD NO.", "JC NO.", "DATE", "DESCRIPTION", "AMOUNT"]);
    formatHeader(sheet);
    lastRow = 1;
  }

  var fullJobCardNo = data.prefix + data.jobNumber;
  var numericPart   = parseInt(data.jobNumber, 10);

  // --- Duplicate check: scan columns B (Prefix) and C (JC NO.) ---
  if (lastRow >= 2) {
    var existing = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
    for (var r = 0; r < existing.length; r++) {
      if (String(existing[r][0]).trim() === data.prefix.trim() && existing[r][1] == numericPart) {
        return jsonResponse({
          success: false, duplicate: true,
          message: "Duplicate! '" + fullJobCardNo + "' already exists in the sheet."
        });
      }
    }
  }

  // --- Append new row ---
  sheet.appendRow([
    0,                       // A: SL.NO (placeholder)
    data.prefix,             // B: JOB CARD NO. (prefix only)
    numericPart,             // C: JC NO. (numeric part only — used for sorting)
    data.date,               // D: DATE
    data.description,        // E: DESCRIPTION
    parseFloat(data.amount)  // F: AMOUNT
  ]);

  // --- Sort by column C (JC NO. numeric) ascending — one fast call ---
  var total = sheet.getLastRow();
  if (total > 2) {
    sheet.getRange(2, 1, total - 1, 6).sort({ column: 3, ascending: true });
  }

  // --- Batch-update SL.NO (single write = much faster than row-by-row) ---
  var finalLastRow = sheet.getLastRow();
  var slValues = [];
  for (var i = 2; i <= finalLastRow; i++) slValues.push([i - 1]);
  if (slValues.length > 0) {
    sheet.getRange(2, 1, slValues.length, 1).setValues(slValues);
  }

  // --- Apply cell formatting in bulk ---
  applyFormatting(sheet, finalLastRow);

  return jsonResponse({
    success: true,
    message: "Inserted and sorted!",
    jobCardNo: fullJobCardNo,
    slNo: finalLastRow - 1
  });
}

// ---- Format the header row ----
function formatHeader(sheet) {
  var h = sheet.getRange(1, 1, 1, 6);
  h.setFontWeight("bold");
  h.setHorizontalAlignment("center");
  h.setVerticalAlignment("middle");
  h.setBackground("#3c3f8f");
  h.setFontColor("#ffffff");
  h.setFontSize(11);
}

// ---- Apply bulk formatting to all data rows ----
function applyFormatting(sheet, lastRow) {
  if (lastRow < 2) return;
  var rows = lastRow - 1;

  // Center + Middle: SL.NO(1), JOB CARD NO.(2), JC NO.(3), DATE(4), AMOUNT(6)
  [1, 2, 3, 4, 6].forEach(function(col) {
    var rng = sheet.getRange(2, col, rows, 1);
    rng.setHorizontalAlignment("center");
    rng.setVerticalAlignment("middle");
  });

  // Left + Middle + Wrap text: DESCRIPTION (E=5)
  var desc = sheet.getRange(2, 5, rows, 1);
  desc.setHorizontalAlignment("left");
  desc.setVerticalAlignment("middle");
  desc.setWrap(true);

  // AMOUNT (F=6): always 2 decimal places with thousand separator e.g. 1,500.00
  sheet.getRange(2, 6, rows, 1).setNumberFormat("#,##0.00");

  // DATE format (D=4)
  sheet.getRange(2, 4, rows, 1).setNumberFormat("dd-mm-yyyy");

  // JC NO. (C=3): plain integer, no decimals
  sheet.getRange(2, 3, rows, 1).setNumberFormat("0");
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}`;

// ============================================================
//  SETTINGS (localStorage)
// ============================================================
const STORAGE_KEY  = 'jobcard_settings';
const LOG_KEY      = 'jobcard_log';

// --- Your Google Apps Script URL (pre-configured) ---
const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxsS32uZ_dPRP3CGlBtXopXX-mZjsFy1SIWQ1-_rUvRridAOE1lstqgfsgMd5vqn3iF/exec';

function loadSettings() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {};
  } catch { return {}; }
}

function saveSettings(settings) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

function loadLog() {
  try {
    return JSON.parse(localStorage.getItem(LOG_KEY)) || [];
  } catch { return []; }
}

function saveLog(log) {
  // Keep last 50 entries
  localStorage.setItem(LOG_KEY, JSON.stringify(log.slice(-50)));
}

// ============================================================
//  TOAST
// ============================================================
function showToast(message, type = 'info', duration = 3500) {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;

  const icons = {
    success: `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>`,
    error:   `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`,
    info:    `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>`,
  };

  toast.innerHTML = `${icons[type] || icons.info}<span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.animation = 'toastOut 0.3s ease forwards';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// ============================================================
//  CONNECTION STATUS
// ============================================================
function setConnectionStatus(state) {
  const badge  = document.getElementById('connectionStatus');
  const textEl = document.getElementById('statusText');
  badge.className = `status-badge status-${state}`;
  const labels = { disconnected: 'Not Connected', connected: 'Connected', testing: 'Testing...' };
  textEl.textContent = labels[state] || state;
}

// ============================================================
//  UPDATE PREVIEW
// ============================================================
function updatePreview() {
  const prefix = document.getElementById('jobPrefix').value || '';
  const numEl  = document.getElementById('jobNumber');
  const num    = numEl ? (numEl.value || '12345') : '12345';

  document.getElementById('previewPrefix').textContent = prefix || '—';
  document.getElementById('previewNumber').textContent = num || '12345';
  document.getElementById('previewFull').textContent   = prefix + (num || '12345');
  document.getElementById('prefixDisplay').textContent = prefix || '—';
}

// ============================================================
//  RENDER LOG
// ============================================================
function renderLog() {
  const list = document.getElementById('recentList');
  const log  = loadLog();

  if (log.length === 0) {
    list.innerHTML = `
      <div class="empty-state">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        <p>No entries yet. Insert your first job card above!</p>
      </div>`;
    return;
  }

  list.innerHTML = '';
  [...log].reverse().forEach(entry => {
    const row = document.createElement('div');
    row.className = 'entry-row';
    const statusClass = entry.success ? 'success' : 'error';
    const amountStr   = entry.amount
      ? '₹' + parseFloat(entry.amount).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
      : '—';
    row.innerHTML = `
      <div class="entry-row-status ${statusClass}"></div>
      <div class="entry-row-jc">${escHtml(entry.jobCardNo || '—')}</div>
      <div class="entry-row-desc">${escHtml(entry.description || '—')}</div>
      <div class="entry-row-amount">${amountStr}</div>
      <div class="entry-row-time">${entry.time || ''}</div>`;
    list.appendChild(row);
  });
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ============================================================
//  SEND TO APPS SCRIPT
// ============================================================
async function sendToSheet(payload) {
  const settings = loadSettings();
  const url = settings.scriptUrl || DEFAULT_SCRIPT_URL;

  if (!url) throw new Error('Script URL not configured. Please go to Settings.');

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(payload),
  });

  if (!response.ok) throw new Error(`HTTP error: ${response.status}`);
  return await response.json();
}

// ============================================================
//  TEST CONNECTION  (manual button click)
// ============================================================
async function testConnection() {
  const url = document.getElementById('scriptUrl').value.trim();
  if (!url) { showToast('Please enter the Script URL first.', 'error'); return; }

  setConnectionStatus('testing');
  document.getElementById('testConnectionBtn').textContent = 'Testing...';

  try {
    const result = await pingScript(url);
    if (result && result.success) {
      setConnectionStatus('connected');
      showToast('✅ Connection successful! Your Google Sheet is linked.', 'success', 4000);
    } else {
      throw new Error(result?.message || 'Unknown error');
    }
  } catch (err) {
    setConnectionStatus('disconnected');
    showToast('Connection failed: ' + err.message, 'error', 5000);
  } finally {
    document.getElementById('testConnectionBtn').textContent = 'Test';
  }
}

// ============================================================
//  PING SCRIPT  (shared by auto-connect + manual test)
// ============================================================
async function pingScript(url) {
  try {
    const resp = await fetch(url + '?action=test', { method: 'GET' });
    return await resp.json();
  } catch {
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ action: 'test' }),
    });
    return await resp.json();
  }
}

// ============================================================
//  AUTO-CONNECT  (silent, runs on page load)
// ============================================================
async function autoConnect(url) {
  setConnectionStatus('testing');
  try {
    const result = await pingScript(url);
    if (result && result.success) {
      setConnectionStatus('connected');
    } else {
      setConnectionStatus('disconnected');
    }
  } catch {
    setConnectionStatus('disconnected');
  }
}

// ============================================================
//  INSERT ENTRY
// ============================================================
async function insertEntry(e) {
  e.preventDefault();

  const settings    = loadSettings();
  const prefix      = settings.jobPrefix || '';
  const jobNumber   = document.getElementById('jobNumber').value.trim();
  const date        = document.getElementById('entryDate').value;
  const description = document.getElementById('description').value.trim();
  const amount      = document.getElementById('amount').value;
  const sheetName   = settings.sheetName || 'Sheet1';

  if (!jobNumber || !/^\d+$/.test(jobNumber)) {
    showToast('Please enter a valid numeric job card number.', 'error');
    return;
  }

  const fullJobCardNo = prefix + jobNumber;

  // --------------------------------------------------------
  // FRONT-END DUPLICATE CHECK against local session log
  // --------------------------------------------------------
  const existingLog  = loadLog();
  const alreadyInLog = existingLog.some(
    entry => entry.success && entry.jobCardNo === fullJobCardNo
  );
  if (alreadyInLog) {
    showToast(
      `⚠️ Duplicate! "${fullJobCardNo}" was already inserted in this session.`,
      'error', 5000
    );
    highlightJobNumberField();
    return;
  }

  // Set loading state
  const btn     = document.getElementById('insertBtn');
  const btnText = document.getElementById('insertBtnText');
  const spinner = document.getElementById('insertSpinner');
  btn.disabled  = true;
  btnText.textContent = 'Inserting...';
  spinner.classList.remove('hidden');

  try {
    const result = await sendToSheet({
      action: 'insert',
      prefix,
      jobNumber,
      date,
      description,
      amount,
      sheetName,
    });

    if (result && result.success) {
      const log = loadLog();
      log.push({
        jobCardNo: fullJobCardNo,
        description,
        amount,
        date,
        time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }),
        success: true,
      });
      saveLog(log);
      renderLog();

      showToast(`✅ Inserted: ${fullJobCardNo} — Sheet sorted!`, 'success', 4000);
      setConnectionStatus('connected');

      // Reset entry fields (keep date & prefix settings)
      document.getElementById('jobNumber').value   = '';
      document.getElementById('description').value = '';
      document.getElementById('amount').value      = '';
      document.getElementById('jobNumber').focus();
      updatePreview();

    } else if (result && result.duplicate) {
      showToast(
        `⚠️ Duplicate! "${fullJobCardNo}" already exists in the Google Sheet.`,
        'error', 6000
      );
      highlightJobNumberField();

    } else {
      throw new Error(result?.message || 'Insertion failed');
    }
  } catch (err) {
    const log = loadLog();
    log.push({
      jobCardNo: fullJobCardNo,
      description,
      amount,
      date,
      time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }),
      success: false,
    });
    saveLog(log);
    renderLog();
    showToast('❌ Error: ' + err.message, 'error', 6000);
  } finally {
    btn.disabled = false;
    btnText.textContent = 'Insert to Sheet';
    spinner.classList.add('hidden');
  }
}

// Highlight job number field red on duplicate error
function highlightJobNumberField() {
  const el = document.getElementById('jobNumber');
  el.style.borderColor = 'var(--clr-red)';
  el.style.boxShadow   = '0 0 0 3px rgba(239,68,68,0.25)';
  el.focus();
  el.select();
  setTimeout(() => {
    el.style.borderColor = '';
    el.style.boxShadow   = '';
  }, 2500);
}

// ============================================================
//  INIT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {

  // Load Apps Script code into the pre element
  document.getElementById('appsScriptCode').textContent = APPS_SCRIPT_CODE;

  // Restore saved settings (merge with defaults)
  const saved = loadSettings();

  // If no URL saved yet, inject the pre-configured default and save it
  if (!saved.scriptUrl) {
    saved.scriptUrl = DEFAULT_SCRIPT_URL;
    saveSettings(saved);
  }

  const settings = saved;

  document.getElementById('scriptUrl').value    = settings.scriptUrl    || DEFAULT_SCRIPT_URL;
  if (settings.jobPrefix)    document.getElementById('jobPrefix').value    = settings.jobPrefix;
  if (settings.startingDate) document.getElementById('startingDate').value = settings.startingDate;
  if (settings.sheetName)    document.getElementById('sheetName').value    = settings.sheetName;

  // Auto-fill entry date from default date
  const entryDate = document.getElementById('entryDate');
  if (settings.startingDate) {
    entryDate.value = settings.startingDate;
  } else {
    entryDate.value = new Date().toISOString().split('T')[0];
  }

  // Auto-connect silently on every page load
  autoConnect(settings.scriptUrl || DEFAULT_SCRIPT_URL);

  // Update preview
  updatePreview();

  // Settings toggle
  document.getElementById('settingsToggleBtn').addEventListener('click', () => {
    document.getElementById('settingsPanel').classList.toggle('collapsed');
  });

  // Save settings
  document.getElementById('saveSettingsBtn').addEventListener('click', () => {
    const newSettings = {
      scriptUrl:    document.getElementById('scriptUrl').value.trim(),
      jobPrefix:    document.getElementById('jobPrefix').value,
      startingDate: document.getElementById('startingDate').value,
      sheetName:    document.getElementById('sheetName').value.trim() || 'Sheet1',
    };
    saveSettings(newSettings);

    if (newSettings.startingDate) {
      document.getElementById('entryDate').value = newSettings.startingDate;
    }

    document.getElementById('prefixDisplay').textContent = newSettings.jobPrefix || '—';
    updatePreview();

    if (newSettings.scriptUrl) setConnectionStatus('connected');
    showToast('✅ Settings saved successfully!', 'success');
    document.getElementById('settingsPanel').classList.add('collapsed');
  });

  // Test connection
  document.getElementById('testConnectionBtn').addEventListener('click', testConnection);

  // Prefix live preview
  document.getElementById('jobPrefix').addEventListener('input', updatePreview);
  document.getElementById('jobNumber').addEventListener('input', updatePreview);

  // Entry form submit
  document.getElementById('entryForm').addEventListener('submit', insertEntry);

  // Clear form
  document.getElementById('clearBtn').addEventListener('click', () => {
    document.getElementById('jobNumber').value   = '';
    document.getElementById('description').value = '';
    document.getElementById('amount').value      = '';
    const s = loadSettings();
    if (s.startingDate) document.getElementById('entryDate').value = s.startingDate;
    updatePreview();
    showToast('Form cleared.', 'info', 2000);
  });

  // Clear log
  document.getElementById('clearLogBtn').addEventListener('click', () => {
    localStorage.removeItem(LOG_KEY);
    renderLog();
    showToast('Log cleared.', 'info', 2000);
  });

  // Copy Apps Script
  document.getElementById('copyScriptBtn').addEventListener('click', () => {
    navigator.clipboard.writeText(APPS_SCRIPT_CODE).then(() => {
      const btn = document.getElementById('copyScriptBtn');
      btn.classList.add('copied');
      btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg> Copied!`;
      setTimeout(() => {
        btn.classList.remove('copied');
        btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg> Copy Script`;
      }, 2500);
      showToast('Script copied to clipboard!', 'success', 2500);
    }).catch(() => showToast('Copy failed — please select and copy manually.', 'error'));
  });

  // How-to link smooth scroll
  document.getElementById('howtoLink').addEventListener('click', (e) => {
    e.preventDefault();
    document.getElementById('howto').scrollIntoView({ behavior: 'smooth' });
  });

  // Render existing log
  renderLog();
});
