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

    if (action === "insert") {
      return insertRow(data);
    } else if (action === "test") {
      return jsonResponse({ success: true, message: "Connection successful!" });
    }

    return jsonResponse({ success: false, message: "Unknown action" });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function doGet(e) {
  var action = e.parameter.action;
  if (action === "test") {
    return jsonResponse({ success: true, message: "Connection successful!" });
  }
  return jsonResponse({ success: false, message: "Use POST requests for data" });
}

function insertRow(data) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = data.sheetName || "Sheet1";
  var sheet     = ss.getSheetByName(sheetName);

  if (!sheet) {
    return jsonResponse({ success: false, message: "Sheet '" + sheetName + "' not found." });
  }

  // --- Ensure header row exists ---
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    sheet.appendRow(["SL.NO", "JOB CARD NO.", "DATE", "DESCRIPTION", "AMOUNT"]);
    lastRow = 1;
  }

  // --- Build full job card number ---
  var fullJobCardNo = data.prefix + data.jobNumber;

  // -------------------------------------------------------
  // DUPLICATE CHECK: scan column B for matching job card no.
  // -------------------------------------------------------
  if (lastRow >= 2) {
    var existingCards = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var r = 0; r < existingCards.length; r++) {
      if (String(existingCards[r][0]).trim() === fullJobCardNo.trim()) {
        return jsonResponse({
          success: false,
          duplicate: true,
          message: "Duplicate entry! Job Card '" + fullJobCardNo + "' already exists in the sheet."
        });
      }
    }
  }

  // --- Append the new row ---
  sheet.appendRow([
    0,                        // SL.NO placeholder
    fullJobCardNo,
    data.date,
    data.description,
    parseFloat(data.amount)
  ]);

  // --- Sort data rows by Job Card Number (ascending) ---
  var totalRows = sheet.getLastRow();
  if (totalRows > 2) {
    sheet.getRange(2, 1, totalRows - 1, 5).sort({ column: 2, ascending: true });
  }

  // --- Recalculate SL.NO sequentially ---
  var finalLastRow = sheet.getLastRow();
  for (var i = 2; i <= finalLastRow; i++) {
    sheet.getRange(i, 1).setValue(i - 1);
  }

  return jsonResponse({
    success: true,
    message: "Inserted and sorted successfully!",
    jobCardNo: fullJobCardNo,
    slNo: finalLastRow - 1
  });
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
  const badge    = document.getElementById('connectionStatus');
  const textEl   = document.getElementById('statusText');
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
  // Show newest first
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
  const url = settings.scriptUrl;

  if (!url) throw new Error('Script URL not configured. Please go to Settings.');

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(payload),
  });

  if (!response.ok) throw new Error(`HTTP error: ${response.status}`);
  const result = await response.json();
  return result;
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
  // Try GET first (no CORS preflight), fall back to POST
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
    // Network issue — keep showing disconnected but don't alert
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
  const existingLog = loadLog();
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
  const btn      = document.getElementById('insertBtn');
  const btnText  = document.getElementById('insertBtnText');
  const spinner  = document.getElementById('insertSpinner');
  btn.disabled   = true;
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
      // Save to local log
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
      document.getElementById('jobNumber').value  = '';
      document.getElementById('description').value = '';
      document.getElementById('amount').value      = '';
      document.getElementById('jobNumber').focus();
      updatePreview();

    } else if (result && result.duplicate) {
      // Sheet-side duplicate detected
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

// Briefly highlight the job number field on duplicate error
function highlightJobNumberField() {
  const el = document.getElementById('jobNumber');
  el.style.borderColor = 'var(--clr-red)';
  el.style.boxShadow  = '0 0 0 3px rgba(239,68,68,0.25)';
  el.focus();
  el.select();
  setTimeout(() => {
    el.style.borderColor = '';
    el.style.boxShadow  = '';
  }, 2500);
}

// ============================================================
//  INIT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {

  // --- Load Apps Script code into the pre element ---
  document.getElementById('appsScriptCode').textContent = APPS_SCRIPT_CODE;

  // --- Restore saved settings (merge with defaults) ---
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
  const urlToConnect = settings.scriptUrl || DEFAULT_SCRIPT_URL;
  autoConnect(urlToConnect);

  // --- Update preview ---
  updatePreview();

  // --- Settings toggle ---
  document.getElementById('settingsToggleBtn').addEventListener('click', () => {
    const panel = document.getElementById('settingsPanel');
    panel.classList.toggle('collapsed');
  });

  // --- Save settings ---
  document.getElementById('saveSettingsBtn').addEventListener('click', () => {
    const newSettings = {
      scriptUrl:    document.getElementById('scriptUrl').value.trim(),
      jobPrefix:    document.getElementById('jobPrefix').value,
      startingDate: document.getElementById('startingDate').value,
      sheetName:    document.getElementById('sheetName').value.trim() || 'Sheet1',
    };
    saveSettings(newSettings);

    // Update entry date from default date
    if (newSettings.startingDate) {
      document.getElementById('entryDate').value = newSettings.startingDate;
    }

    // Update prefix display
    document.getElementById('prefixDisplay').textContent = newSettings.jobPrefix || '—';
    updatePreview();

    if (newSettings.scriptUrl) setConnectionStatus('connected');
    showToast('✅ Settings saved successfully!', 'success');

    // Collapse after save
    document.getElementById('settingsPanel').classList.add('collapsed');
  });

  // --- Test connection ---
  document.getElementById('testConnectionBtn').addEventListener('click', testConnection);

  // --- Prefix live preview ---
  document.getElementById('jobPrefix').addEventListener('input', updatePreview);
  document.getElementById('jobNumber').addEventListener('input', updatePreview);

  // --- Entry form submit ---
  document.getElementById('entryForm').addEventListener('submit', insertEntry);

  // --- Clear form ---
  document.getElementById('clearBtn').addEventListener('click', () => {
    document.getElementById('jobNumber').value  = '';
    document.getElementById('description').value = '';
    document.getElementById('amount').value     = '';
    const settings = loadSettings();
    if (settings.startingDate) {
      document.getElementById('entryDate').value = settings.startingDate;
    }
    updatePreview();
    showToast('Form cleared.', 'info', 2000);
  });

  // --- Clear log ---
  document.getElementById('clearLogBtn').addEventListener('click', () => {
    localStorage.removeItem(LOG_KEY);
    renderLog();
    showToast('Log cleared.', 'info', 2000);
  });

  // --- Copy Apps Script ---
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

  // --- How-to link smooth scroll ---
  document.getElementById('howtoLink').addEventListener('click', (e) => {
    e.preventDefault();
    document.getElementById('howto').scrollIntoView({ behavior: 'smooth' });
  });

  // --- Render existing log ---
  renderLog();
});
