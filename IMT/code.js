// ============================================================
//  IMT – Invoice Management Tool
//  Google Apps Script backend
// ============================================================

// ── Sheet / tab names ────────────────────────────────────────
const SHEET_SYSTEM_CONFIG    = 'System Config';
const SHEET_GL_CONFIG        = 'GL Config';
const SHEET_RDC_ALIASES      = 'RDC Aliases';
const SHEET_EMAIL_TEMPLATE   = 'Email Template';
const SHEET_HEADER_CONFIG    = 'Header Config';
const SHEET_CARRIER_CREDITS  = 'Carrier Credits';
const SHEET_HAULIER_CALLBACKS = 'Haulier Callbacks';
const SHEET_INVOICES         = 'Invoices';

// ── Web-app entry point ──────────────────────────────────────
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'dashboard';
  if (page === 'config') {
    return HtmlService.createTemplateFromFile('Config')
      .evaluate()
      .setTitle('IMT – Configuration')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('IMT – Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Spreadsheet helper ───────────────────────────────────────
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ── One-time setup: initialise all config tabs ───────────────
function setupConfigSheet() {
  // System Config  (key → value)
  const sc = getOrCreateSheet(SHEET_SYSTEM_CONFIG, ['Key', 'Value']);
  if (sc.getLastRow() < 2) {
    sc.appendRow(['companyName',  'My Company']);
    sc.appendRow(['invoicePrefix', 'INV-']);
    sc.appendRow(['defaultCurrency', 'GBP']);
    sc.appendRow(['vatRate', '20']);
  }

  // GL Config
  const gl = getOrCreateSheet(SHEET_GL_CONFIG, ['Account Code', 'Description', 'Type']);
  if (gl.getLastRow() < 2) {
    gl.appendRow(['4000', 'Haulage Revenue', 'Income']);
    gl.appendRow(['5000', 'Transport Costs', 'Expense']);
  }

  // RDC Aliases
  const rdc = getOrCreateSheet(SHEET_RDC_ALIASES, ['RDC Code', 'Alias']);
  if (rdc.getLastRow() < 2) {
    rdc.appendRow(['RDC01', 'Northern Hub']);
    rdc.appendRow(['RDC02', 'Southern Hub']);
  }

  // Email Template
  const et = getOrCreateSheet(SHEET_EMAIL_TEMPLATE, ['Field', 'Value']);
  if (et.getLastRow() < 2) {
    et.appendRow(['subject',  'Invoice {{invoiceNumber}} – {{period}} – {{carrierName}}']);
    et.appendRow(['bodyHtml', '<p>Dear {{carrierName}},</p><p>Please find attached invoice {{invoiceNumber}} for period {{period}}.</p>']);
    et.appendRow(['signature', 'Kind regards,\n{{companyName}}']);
  }

  // Header Config
  const hc = getOrCreateSheet(SHEET_HEADER_CONFIG, ['Sheet Name', 'Column', 'Header Label']);
  if (hc.getLastRow() < 2) {
    hc.appendRow(['Haulier Sheet', 'A', 'Carrier Name']);
    hc.appendRow(['Haulier Sheet', 'B', 'Route']);
    hc.appendRow(['Haulier Sheet', 'C', 'Base Rate']);
  }

  // Carrier Credits
  const cc = getOrCreateSheet(SHEET_CARRIER_CREDITS, [
    'Carrier Name', 'Credit Type', 'Credit Rate', 'Min Loads', 'Max Loads', 'Active', 'Notes'
  ]);
  if (cc.getLastRow() < 2) {
    cc.appendRow(['Example Carrier', 'Backhaul', '0.00', '1', '', 'TRUE', 'EXAMPLE: Remove this row and add your carrier credit configurations']);
  }

  // Haulier Callbacks
  const cb = getOrCreateSheet(SHEET_HAULIER_CALLBACKS, [
    'Callback Name', 'Target Sheet', 'Target Column', 'Applies To Carrier', 'Cost Formula', 'Active', 'Description'
  ]);
  if (cb.getLastRow() < 2) {
    cb.appendRow([
      'Fuel Surcharge',
      'Haulier Sheet',
      'D',
      'ALL',
      '=C{row}*0.05',
      'TRUE',
      'Adds 5% fuel surcharge to base rate'
    ]);
  }

  // Invoices log
  getOrCreateSheet(SHEET_INVOICES, [
    'Invoice ID', 'Date Generated', 'Period', 'Carrier', 'Amount', 'Status', 'File URL', 'Notes'
  ]);

  return { success: true, message: 'Configuration sheets initialised.' };
}

// ── System Config helpers ────────────────────────────────────
/**
 * Returns a single config value by key.
 */
function getConfig(key) {
  const sheet = getSpreadsheet().getSheetByName(SHEET_SYSTEM_CONFIG);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return String(data[i][1]);
  }
  return null;
}

/**
 * Returns all System Config rows as { key, value } objects.
 */
function getAllSystemConfig() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_SYSTEM_CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data
    .filter(r => String(r[0]).trim())
    .map(r => ({ key: String(r[0]).trim(), value: String(r[1]) }));
}

/**
 * Saves (overwrites) System Config rows.
 * @param {Array<{key:string,value:string}>} rows
 */
function saveSystemConfig(rows) {
  const sheet = getOrCreateSheet(SHEET_SYSTEM_CONFIG, ['Key', 'Value']);
  // Clear existing data rows
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }
  if (rows && rows.length) {
    const values = rows.map(r => [String(r.key).trim(), String(r.value)]);
    sheet.getRange(2, 1, values.length, 2).setValues(values);
  }
  return { success: true };
}

// ── GL Config helpers ────────────────────────────────────────
function getAllGlConfig() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_GL_CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data
    .filter(r => String(r[0]).trim())
    .map(r => ({ accountCode: String(r[0]), description: String(r[1]), type: String(r[2]) }));
}

function saveGlConfig(rows) {
  const sheet = getOrCreateSheet(SHEET_GL_CONFIG, ['Account Code', 'Description', 'Type']);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
  if (rows && rows.length) {
    const values = rows.map(r => [r.accountCode, r.description, r.type]);
    sheet.getRange(2, 1, values.length, 3).setValues(values);
  }
  return { success: true };
}

// ── RDC Aliases helpers ──────────────────────────────────────
function getAllRdcAliases() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_RDC_ALIASES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data
    .filter(r => String(r[0]).trim())
    .map(r => ({ code: String(r[0]), alias: String(r[1]) }));
}

function saveRdcAliases(rows) {
  const sheet = getOrCreateSheet(SHEET_RDC_ALIASES, ['RDC Code', 'Alias']);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  if (rows && rows.length) {
    const values = rows.map(r => [r.code, r.alias]);
    sheet.getRange(2, 1, values.length, 2).setValues(values);
  }
  return { success: true };
}

// ── Email Template helpers ───────────────────────────────────
function getAllEmailTemplate() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_EMAIL_TEMPLATE);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data
    .filter(r => String(r[0]).trim())
    .map(r => ({ field: String(r[0]), value: String(r[1]) }));
}

function saveEmailTemplate(rows) {
  const sheet = getOrCreateSheet(SHEET_EMAIL_TEMPLATE, ['Field', 'Value']);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  if (rows && rows.length) {
    const values = rows.map(r => [r.field, r.value]);
    sheet.getRange(2, 1, values.length, 2).setValues(values);
  }
  return { success: true };
}

// ── Header Config helpers ────────────────────────────────────
function getAllHeaderConfig() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_HEADER_CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data
    .filter(r => String(r[0]).trim())
    .map(r => ({ sheetName: String(r[0]), column: String(r[1]), label: String(r[2]) }));
}

function saveHeaderConfig(rows) {
  const sheet = getOrCreateSheet(SHEET_HEADER_CONFIG, ['Sheet Name', 'Column', 'Header Label']);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
  if (rows && rows.length) {
    const values = rows.map(r => [r.sheetName, r.column, r.label]);
    sheet.getRange(2, 1, values.length, 3).setValues(values);
  }
  return { success: true };
}

// ── Carrier Credits helpers ──────────────────────────────────
/**
 * Returns all carrier credit configurations.
 * Each row: { carrier, creditType, creditRate, minLoads, maxLoads, active, notes }
 */
function getAllCarrierCredits() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_CARRIER_CREDITS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return data
    .filter(r => String(r[0]).trim())
    .map(r => ({
      carrier:    String(r[0]).trim(),
      creditType: String(r[1]).trim(),
      creditRate: String(r[2]),
      minLoads:   String(r[3]),
      maxLoads:   String(r[4]),
      active:     String(r[5]).toLowerCase() === 'true',
      notes:      String(r[6])
    }));
}

/**
 * Saves carrier credit configurations.
 * @param {Array} rows
 */
function saveCarrierCredits(rows) {
  const sheet = getOrCreateSheet(SHEET_CARRIER_CREDITS, [
    'Carrier Name', 'Credit Type', 'Credit Rate', 'Min Loads', 'Max Loads', 'Active', 'Notes'
  ]);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
  if (rows && rows.length) {
    const values = rows.map(r => [
      r.carrier, r.creditType, r.creditRate,
      r.minLoads, r.maxLoads,
      r.active ? 'TRUE' : 'FALSE',
      r.notes || ''
    ]);
    sheet.getRange(2, 1, values.length, 7).setValues(values);
  }
  return { success: true };
}

// ── Haulier Callbacks helpers ────────────────────────────────
/**
 * Returns all haulier sheet callback configurations.
 * Each callback describes an additional cost column to inject.
 */
function getAllHaulierCallbacks() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_HAULIER_CALLBACKS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return data
    .filter(r => String(r[0]).trim())
    .map(r => ({
      name:        String(r[0]).trim(),
      targetSheet: String(r[1]).trim(),
      targetColumn: String(r[2]).trim(),
      appliesTo:   String(r[3]).trim(),
      formula:     String(r[4]),
      active:      String(r[5]).toLowerCase() === 'true',
      description: String(r[6])
    }));
}

/**
 * Saves haulier callback configurations.
 * @param {Array} rows
 */
function saveHaulierCallbacks(rows) {
  const sheet = getOrCreateSheet(SHEET_HAULIER_CALLBACKS, [
    'Callback Name', 'Target Sheet', 'Target Column', 'Applies To Carrier', 'Cost Formula', 'Active', 'Description'
  ]);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
  if (rows && rows.length) {
    const values = rows.map(r => [
      r.name, r.targetSheet, r.targetColumn,
      r.appliesTo, r.formula,
      r.active ? 'TRUE' : 'FALSE',
      r.description || ''
    ]);
    sheet.getRange(2, 1, values.length, 7).setValues(values);
  }
  return { success: true };
}

// ── Invoice helpers ──────────────────────────────────────────
/**
 * Returns all recorded invoices, newest first.
 */
function getAllInvoices() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_INVOICES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return data
    .filter(r => String(r[0]).trim())
    .map(r => ({
      id:        String(r[0]).trim(),
      date:      r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      period:    String(r[2]),
      carrier:   String(r[3]),
      amount:    String(r[4]),
      status:    String(r[5]),
      fileUrl:   String(r[6]),
      notes:     String(r[7])
    }))
    .reverse();
}

/**
 * Returns aggregated summary metrics for the dashboard.
 */
function getInvoiceSummary() {
  const invoices = getAllInvoices();
  const total = invoices.length;
  const byStatus = {};
  let totalAmount = 0;
  const byCarrier = {};

  invoices.forEach(inv => {
    const status = inv.status || 'Unknown';
    byStatus[status] = (byStatus[status] || 0) + 1;

    const amt = parseFloat(String(inv.amount).replace(/[^0-9.-]/g, '')) || 0;
    totalAmount += amt;

    const c = inv.carrier || 'Unknown';
    if (!byCarrier[c]) byCarrier[c] = { count: 0, total: 0 };
    byCarrier[c].count += 1;
    byCarrier[c].total += amt;
  });

  // Recent (last 5)
  const recent = invoices.slice(0, 5);

  return { total, byStatus, totalAmount, byCarrier, recent };
}

/**
 * Records a new invoice entry.
 * @param {{id,period,carrier,amount,status,fileUrl,notes}} inv
 */
function recordInvoice(inv) {
  const sheet = getOrCreateSheet(SHEET_INVOICES, [
    'Invoice ID', 'Date Generated', 'Period', 'Carrier', 'Amount', 'Status', 'File URL', 'Notes'
  ]);
  sheet.appendRow([
    inv.id || '',
    new Date(),
    inv.period || '',
    inv.carrier || '',
    inv.amount || 0,
    inv.status || 'Draft',
    inv.fileUrl || '',
    inv.notes || ''
  ]);
  return { success: true };
}

/**
 * Updates the status of an existing invoice.
 * @param {string} invoiceId
 * @param {string} newStatus
 */
function updateInvoiceStatus(invoiceId, newStatus) {
  const sheet = getSpreadsheet().getSheetByName(SHEET_INVOICES);
  if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Invoice not found.' };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(invoiceId).trim()) {
      sheet.getRange(i + 2, 6).setValue(newStatus);
      return { success: true };
    }
  }
  return { success: false, message: 'Invoice not found.' };
}

// ── Unified config getter for the Config page ─────────────────
/**
 * Returns all config sections in one call to minimise round-trips.
 */
function getAllConfig() {
  return {
    systemConfig:      getAllSystemConfig(),
    glConfig:          getAllGlConfig(),
    rdcAliases:        getAllRdcAliases(),
    emailTemplate:     getAllEmailTemplate(),
    headerConfig:      getAllHeaderConfig(),
    carrierCredits:    getAllCarrierCredits(),
    haulierCallbacks:  getAllHaulierCallbacks()
  };
}
