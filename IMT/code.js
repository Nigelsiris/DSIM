/**
 * INVOICE MASTER TRACKER AUTOMATION
 * * Features:
 * - Auto-creates custom menu, system config, GL config, RDC aliases, Email template, and Header config tabs.
 * - Dynamically loads any carrier ending in _ROOT_FOLDER from config.
 * - Extracts Shift Data and Additional Costs dynamically based on custom Header Config aliases.
 * - Enforces GL Cost Tolerances and Ignore Rules to prevent $0.00 or irrelevant items from being coded.
 * - Cross-references with Haulier Reports AND writes the Invoice # & Amount back to the Haulier sheet.
 * - Aggregates invoice totals and applies customizable GL coding.
 * - Auto-generates stamped PDFs for the accounting team.
 * - Processes from root folders and automatically sorts files into RDC sub-folders.
 * - Automatically archives old data and clears sheets before new imports.
 * - Generates Discrepancy Email drafts and suspends GL coding on problematic invoices.
 */

// --- MENU & UI ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Invoice Automation')
    .addItem('1. Initialize/Setup Configuration', 'setupConfigTab')
    .addSeparator()
    .addItem('2. Process All Invoices', 'runMainProcess')
    .addItem('3. Clear Tracker Data', 'clearTrackerData')
    .addToUi();
}

// --- CONFIGURATION MANAGEMENT ---
function setupConfigTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. System Config
  let configSheet = ss.getSheetByName('System Config');
  if (!configSheet) {
    configSheet = ss.insertSheet('System Config');
    const headers = [['Setting Name', 'Value / ID', 'Description']];
    configSheet.getRange(1, 1, 1, 3).setValues(headers).setFontWeight('bold').setBackground('#d9ead3');
    
    const defaultSettings = [
      ['FRG_HAULIER_ID', '1VyvVZ3BB2-tArldhuNKKBzfjU7FATVMJbtzhZP3P8pA', 'FRG Weekly Haulier Sheet ID'],
      ['GRM_HAULIER_ID', '13lvErboGKp-cd06XV1PvhB7zSrCEXXVm0POZlWPBZV4', 'GRM Weekly Haulier Sheet ID'],
      ['PYE_HAULIER_ID', '1AP_qyToLfhyuGihvfV3sUbLpj4WYx-wgAUVjtluHCoY', 'PYE Weekly Haulier Sheet ID'],
      ['CRE_ROOT_FOLDER', '1BCbMyWUu0npRGIQvN1lRtrL5oBh1PdAi', 'Root Folder ID for CRE Invoices'],
      ['HB_ROOT_FOLDER', '1awkIvApl6iyExTnrNTmD1Pt1oM1d-2GO', 'Root Folder ID for HB Invoices'],
      ['SCH_ROOT_FOLDER', '', 'Root Folder ID for SCH Invoices (Optional)'],
      ['WERNER_ROOT_FOLDER', '', 'Root Folder ID for Werner Invoices (Optional)'],
      ['TEMP_PROCESSING_FOLDER', '', 'Optional: Folder ID to store temp converted sheets (leave blank to use root)']
    ];
    configSheet.getRange(2, 1, defaultSettings.length, 3).setValues(defaultSettings);
    configSheet.setColumnWidth(1, 250); configSheet.setColumnWidth(2, 350); configSheet.setColumnWidth(3, 400);
  }

  // 2. GL Config (Now with Ignore Rules & Tolerance)
  let glSheet = ss.getSheetByName('GL Config');
  if (!glSheet) {
    glSheet = ss.insertSheet('GL Config');
    const glHeaders = [['Rule Type', 'Keyword / RDC', 'GL Account / Center Code / Value', 'Description']];
    glSheet.getRange(1, 1, 1, 4).setValues(glHeaders).setFontWeight('bold').setBackground('#fff2cc');
    
    const defaultGL = [
      ['Cost Center', 'FRG', '50001', 'Fredericksburg Cost Center Code'],
      ['Cost Center', 'GRM', '60001', 'Graham Cost Center Code'],
      ['Cost Center', 'PYE', '70001', 'Perryville Cost Center Code'],
      ['Category Mapping', 'toll', '471000, 47100002', 'Catches any line item containing "toll"'],
      ['Category Mapping', 'fuel', '471000, 47100099', 'Catches any line item containing "fuel"'],
      ['Default Category', 'BASE', '471000, 47100001', 'Base freight applied to all other costs'],
      ['Ignore Rule', 'discount', '', 'Completely ignores any line item containing this word'],
      ['Ignore Rule', 'rebate', '', 'Completely ignores any line item containing this word'],
      ['Tolerance', 'MIN_COST', '0.01', 'Ignores any line item with an amount smaller than this value']
    ];
    glSheet.getRange(2, 1, defaultGL.length, 4).setValues(defaultGL);
    glSheet.setColumnWidth(1, 150); glSheet.setColumnWidth(2, 150); glSheet.setColumnWidth(3, 220); glSheet.setColumnWidth(4, 350);
  }
  
  // 3. RDC Aliases Config
  let rdcSheet = ss.getSheetByName('RDC Aliases');
  if (!rdcSheet) {
    rdcSheet = ss.insertSheet('RDC Aliases');
    rdcSheet.getRange(1, 1, 1, 2).setValues([['RDC Code', 'Aliases (Comma Separated)']]).setFontWeight('bold').setBackground('#cfe2f3');
    const defaultAliases = [
      ['FRG', 'FREDERICKSBURG, FRG, VA, LIDL VA, _VA'],
      ['GRM', 'GRAHAM, GRM, NC, LIDL NC, _NC, MEBANE'],
      ['PYE', 'PERRYVILLE, PYE, MD, LIDL MD, _MD, PER']
    ];
    rdcSheet.getRange(2, 1, defaultAliases.length, 2).setValues(defaultAliases);
    rdcSheet.setColumnWidth(1, 150); rdcSheet.setColumnWidth(2, 500);
  }
  
  // 4. Email Template Config
  let emailSheet = ss.getSheetByName('Email Template');
  if (!emailSheet) {
    emailSheet = ss.insertSheet('Email Template');
    emailSheet.getRange(1, 1, 1, 3).setValues([['Setting', 'Template Text', 'Available Variables']]).setFontWeight('bold').setBackground('#fce5cd');
    const defaultEmail = [
      ['Subject', 'Discrepancy Notice: Invoice {FileName}', '{FileName}, {CarrierName}'],
      ['Greeting', 'Hello {CarrierName} Team,\n\nThe following items are showing as discrepancies in our system. Please advise:', '{CarrierName}'],
      ['Outro', 'Thank you.', '']
    ];
    emailSheet.getRange(2, 1, defaultEmail.length, 3).setValues(defaultEmail);
    emailSheet.setColumnWidth(1, 150); emailSheet.setColumnWidth(2, 500); emailSheet.setColumnWidth(3, 200);
  }

  // 5. Header Aliases Config (Dynamic Column Mapping)
  let headerSheet = ss.getSheetByName('Header Config');
  if (!headerSheet) {
    headerSheet = ss.insertSheet('Header Config');
    headerSheet.getRange(1, 1, 1, 2).setValues([['Target Field', 'Column Name Aliases (Comma Separated)']]).setFontWeight('bold').setBackground('#e6b8af');
    const defaultHeaders = [
      ['Date', 'date, pickup dt, delivery date'],
      ['TU', 'tu, mb number'],
      ['Store', 'store, dest location, dest city, destination name'],
      ['Tour', 'tour, route'],
      ['Miles', 'miles, total miles'],
      ['NY Pay', 'ny pay, ny, new york, borough fee, dhu $ amt, stp $ amt'],
      ['Tolls', 'toll, tol $ amt'],
      ['Total Cost', 'total $ amt, total cost, cost, total'],
      ['Shift', 'shift, shifts'],
      ['Type', 'delivery type, tour type, shift type'],
      ['Haulier Invoice', 'invoice, invoice #, invoice number'],
      ['Haulier Amount', 'amount, actual amount, invoice amount, total cost']
    ];
    headerSheet.getRange(2, 1, defaultHeaders.length, 2).setValues(defaultHeaders);
    headerSheet.setColumnWidth(1, 150); headerSheet.setColumnWidth(2, 500);
  }
  
  SpreadsheetApp.getUi().alert('Configuration setup complete. Please review the 5 Configuration tabs at the bottom of your sheet.');
}

// --- DATA FETCHERS ---

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('System Config');
  if (!configSheet) throw new Error('System Config tab not found. Please run Initialization first.');
  const data = configSheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) config[data[i][0]] = data[i][1];
  return config;
}

function getGlConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const glSheet = ss.getSheetByName('GL Config');
  if (!glSheet) return null;
  const data = glSheet.getDataRange().getValues();
  
  const config = { costCenters: {}, categories: [], ignoreKeywords: [], tolerance: 0.01 };
  
  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][0]).trim();
    const key = String(data[i][1]).trim().toLowerCase();
    const value = String(data[i][2]).trim();
    
    if (!type || (!key && type !== 'Tolerance')) continue;
    
    if (type === 'Cost Center') config.costCenters[key.toUpperCase()] = value;
    else if (type === 'Category Mapping') config.categories.push({ keyword: key, glPrefix: value });
    else if (type === 'Default Category') config.defaultCategory = value;
    else if (type === 'Ignore Rule') config.ignoreKeywords.push(key);
    else if (type === 'Tolerance') config.tolerance = parseFloat(value) || 0.01;
  }
  return config;
}

function getRdcAliases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('RDC Aliases');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const aliases = {};
  for (let i = 1; i < data.length; i++) {
    const rdc = String(data[i][0]).trim();
    if (!rdc) continue;
    aliases[rdc] = String(data[i][1]).toUpperCase().split(',').map(s => s.trim());
  }
  return aliases;
}

function getEmailTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Email Template');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const template = {};
  for (let i = 1; i < data.length; i++) {
    template[String(data[i][0]).trim()] = String(data[i][1]);
  }
  return template;
}

function getHeaderAliases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Header Config');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const aliases = {};
  for (let i = 1; i < data.length; i++) {
    const field = String(data[i][0]).trim();
    if (!field) continue;
    aliases[field] = String(data[i][1]).toLowerCase().split(',').map(s => s.trim());
  }
  return aliases;
}

function matchHeader(headerStr, fieldName, aliasesMap, exactOnly = false) {
  const keywords = aliasesMap[fieldName] || [];
  
  // Pass 1: Prioritize exact matches (prevents partial string hijacking)
  for (let kw of keywords) {
    if (headerStr === kw) return true;
  }
  
  // If we are strictly running an exact match pass, abort here
  if (exactOnly) return false;
  
  // Pass 2: Permissive partial matches
  for (let kw of keywords) {
    // Strict exclusion to prevent partial word matches for very short acronyms
    if (kw === 'tu' && (headerStr.includes('status') || headerStr.includes('return'))) continue;
    if (kw === 'ny' && headerStr.includes('company')) continue;
    
    if (headerStr.includes(kw)) {
       // Enforce word boundaries for short acronyms to avoid substring false positives
       if (kw === 'tu' || kw === 'ny') {
          if (new RegExp('\\b' + kw + '\\b', 'i').test(headerStr)) return true;
       } else {
          return true;
       }
    }
  }
  return false;
}

function rowHasHeader(rowArray, fieldName, aliasesMap) {
  for (let i = 0; i < rowArray.length; i++) {
    // Flatten multiple spaces into a single space for consistent matching
    const h = String(rowArray[i]).toLowerCase().trim().replace(/\s+/g, ' ');
    if (matchHeader(h, fieldName, aliasesMap, false)) return true;
  }
  return false;
}

// --- MAIN PROCESS ---

function runMainProcess(isWebApp = false) {
  let ui = null;
  if (!isWebApp) {
    ui = SpreadsheetApp.getUi();
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    if (ui) ui.alert('Processing Started. This may take a few minutes. Please wait...');
    Logger.log('--- STARTING INVOICE PROCESSING ---');
    
    const config = getConfig();
    const glConfig = getGlConfig();
    const rdcAliases = getRdcAliases();
    const emailTemplate = getEmailTemplate();
    const headerAliases = getHeaderAliases();
    
    Logger.log('Fetching Haulier Data...');
    const haulierData = {
      'FRG': fetchHaulierData(config.FRG_HAULIER_ID, headerAliases, 'FRG'),
      'GRM': fetchHaulierData(config.GRM_HAULIER_ID, headerAliases, 'GRM'),
      'PYE': fetchHaulierData(config.PYE_HAULIER_ID, headerAliases, 'PYE')
    };
    
    const masterData = [];
    const additionalCostsData = [];
    const discrepancyData = [];
    const tmstData = [];
    const haulierUpdates = { 'FRG': [], 'GRM': [], 'PYE': [] };
    
    // Dynamically load ALL Carriers based on System Config (_ROOT_FOLDER suffix)
    const carriers = [];
    for (let key in config) {
      if (key.endsWith('_ROOT_FOLDER') && config[key]) {
        carriers.push({ name: key.replace('_ROOT_FOLDER', ''), rootId: config[key] });
      }
    }
    
    carriers.forEach(carrier => {
      Logger.log(`\n--- Processing Carrier: ${carrier.name} ---`);
      let rootFolder;
      try { rootFolder = DriveApp.getFolderById(carrier.rootId); } catch (e) { return; }
      
      searchAndSortInvoices(rootFolder, carrier.name, haulierData, masterData, additionalCostsData, tmstData, discrepancyData, glConfig, rdcAliases, emailTemplate, headerAliases, haulierUpdates);
    });
    
    Logger.log('--- EXECUTING HAULIER REPORT WRITE-BACKS ---');
    for (let rdc in haulierUpdates) {
      if (haulierUpdates[rdc].length > 0 && haulierData[rdc] && haulierData[rdc].spreadsheetId) {
         applyUpdatesToHaulier(haulierData[rdc], haulierUpdates[rdc]);
      }
    }

    archiveAndClearSheets(ss);
    
    writeDataToSheet(ss, 'Master Input', masterData, 9); 
    writeDataToSheet(ss, 'TMST', tmstData, 15); // Expanded to 15 columns for independent Store vs Tour Matches
    writeDataToSheet(ss, 'Additonal Costs', additionalCostsData, 5); 
    writeDataToSheet(ss, 'Discrepancy Tracker', discrepancyData, 6); 
    
    if (masterData.length === 0 && additionalCostsData.length === 0) {
      if (ui) ui.alert('⚠️ Finished running, but NO DATA was extracted. Files may have already been processed.');
      return 'WARNING: Finished running, but NO DATA was extracted. Files may have already been processed.';
    } else {
      if (ui) ui.alert('✅ Processing Complete! Data added, GL Codes applied (if valid), drafts created for discrepancies, and files sorted.');
      return 'SUCCESS: Processing Complete! Data added, GL Codes applied, drafts created, and files sorted.';
    }
    
  } catch (error) {
    Logger.log(`CRITICAL ERROR: ${error.message}`);
    if (ui) ui.alert('❌ Error during processing: ' + error.message);
    return 'ERROR: ' + error.message;
  }
}

// --- HAULIER WRITE-BACK LOGIC ---

function applyUpdatesToHaulier(haulierInfo, updates) {
  try {
    const ss = SpreadsheetApp.openById(haulierInfo.spreadsheetId);
    const sheet = ss.getSheetByName(haulierInfo.sheetName);
    if (!sheet) return;

    let invCol = haulierInfo.invoiceColIdx;
    let amtCol = haulierInfo.amountColIdx;

    // Aggregate by row to sum split-shifts correctly and prevent overwriting
    const rowUpdates = {};
    updates.forEach(u => {
      if (!rowUpdates[u.rowNumber]) {
         rowUpdates[u.rowNumber] = { invoiceNumber: u.invoiceNumber, amount: 0, isValid: true };
      }
      // If ANY part of the row's match sequence failed, we void the amount update
      if (u.amount === "") {
         rowUpdates[u.rowNumber].isValid = false;
      } else {
         rowUpdates[u.rowNumber].amount += (parseFloat(u.amount) || 0);
      }
    });
    
    // Write the aggregated data back to the sheet
    let updateCount = 0;
    for (let row in rowUpdates) {
       const r = parseInt(row);
       sheet.getRange(r, invCol + 1).setValue(rowUpdates[row].invoiceNumber);
       // ONLY update amount if shift, store, and tour perfectly matched
       if (rowUpdates[row].isValid) {
           sheet.getRange(r, amtCol + 1).setValue(rowUpdates[row].amount);
       }
       updateCount++;
    }
    
    Logger.log(`      -> Successfully wrote ${updateCount} aggregated records back to Haulier sheet: ${haulierInfo.sheetName}`);
  } catch (e) {
    Logger.log(`[ERROR] Failed to write back to Haulier sheet: ${e.message}`);
  }
}

// --- FOLDER & FILE SEARCHING LOGIC ---

function searchAndSortInvoices(rootFolder, carrierName, haulierData, masterData, additionalCostsData, tmstData, discrepancyData, glConfig, rdcAliases, emailTemplate, headerAliases, haulierUpdates) {
  let count = 0;
  const files = rootFolder.getFiles();
  
  while (files.hasNext()) {
    count++;
    let originalFile = files.next(); 
    let file = originalFile; // Retain original item (e.g. the shortcut) for moving
    let fileName = file.getName();
    let mimeType = file.getMimeType();
    
    // If it's a shortcut, resolve it to read data, but keep 'originalFile' pointing to the shortcut!
    if (mimeType === 'application/vnd.google-apps.shortcut') {
      try {
        file = DriveApp.getFileById(file.getTargetId());
        fileName = file.getName();
        mimeType = file.getMimeType();
      } catch(e) { continue; }
    }
    
    const isExcelOrSheet = fileName.toLowerCase().indexOf('.xlsx') > -1 || mimeType === MimeType.GOOGLE_SHEETS || mimeType === MimeType.MICROSOFT_EXCEL || mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    if (isExcelOrSheet) {
      let finalRdc = parseRDCName(fileName, rdcAliases);
      
      if (finalRdc === 'UNKNOWN') {
         processInvoiceRouter(file, finalRdc, carrierName, {}, masterData, additionalCostsData, tmstData, discrepancyData, glConfig, rootFolder, rdcAliases, emailTemplate, headerAliases, haulierUpdates);
      } else {
         let targetFolder = getOrCreateFolder(rootFolder, finalRdc);
         processInvoiceRouter(file, finalRdc, carrierName, haulierData[finalRdc] || {}, masterData, additionalCostsData, tmstData, discrepancyData, glConfig, targetFolder, rdcAliases, emailTemplate, headerAliases, haulierUpdates);
         
         // Move the ORIGINAL FILE (the shortcut we own), not the Target File (which we don't own).
         try { moveFileToFolder(originalFile, targetFolder); } catch(e) { Logger.log("Failed to move file."); }
      }
    }
  }
  return count;
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function moveFileToFolder(file, newParent) {
  try {
    file.moveTo(newParent);
  } catch (e) {
    Logger.log(`[WARNING] Native moveTo failed: ${e.message}. Attempting Advanced Drive API fallback...`);
    
    try {
      const fileId = file.getId();
      let previousParents = [];
      const parents = file.getParents();
      while (parents.hasNext()) {
        previousParents.push(parents.next().getId());
      }
      
      Drive.Files.update({}, fileId, null, {
        addParents: newParent.getId(),
        removeParents: previousParents.join(','),
        supportsAllDrives: true
      });
    } catch (err) {
      Logger.log(`[ERROR] Advanced Drive API Move failed: ${err.message}`);
    }
  }
}

// --- INVOICE PROCESSING ROUTER ---

function processInvoiceRouter(file, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, glConfig, targetFolder, rdcAliases, emailTemplate, headerAliases, haulierUpdates) {
  let sheetId = file.getId();
  let isTemp = false;
  let fileCostSummary = []; 
  const initialDiscrepancyCount = discrepancyData.length; 
  
  // Extract clean invoice name to act as the Invoice Number for writing back
  const fallbackInvoiceName = file.getName().replace(/\.xlsx$/i, '');
  
  if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    try {
      sheetId = convertExcelToGoogleSheet(file.getId(), targetFolder.getId());
      isTemp = true;
    } catch(e) { return; }
  }
  
  try {
    const invoiceSS = SpreadsheetApp.openById(sheetId);
    const invoiceNumber = extractInvoiceNumber(invoiceSS, fallbackInvoiceName);
    Logger.log(`      -> Extracted Invoice Number: '${invoiceNumber}'`);
    
    let isSchByStructure = false;
    if (!carrierName.toUpperCase().includes('SCH')) {
       for(let s=0; s<invoiceSS.getSheets().length; s++) {
           let sheetNameLower = invoiceSS.getSheets()[s].getName().toLowerCase();
           if(sheetNameLower.includes('order detail') || sheetNameLower.includes('customer detail') || sheetNameLower.includes('schneider')) {
               isSchByStructure = true; carrierName = 'SCH'; break;
           }
       }
    }

    if (carrierName.toUpperCase().includes('CRE')) {
       processCreInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates);
    } else if (carrierName.toUpperCase().includes('HB')) {
       processHbInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates);
    } else if (carrierName.toUpperCase().includes('SCH') || isSchByStructure) {
       processSchInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates);
    } else {
       // Automatic Fallback logic for dynamic carriers (Werner, etc)
       processHbInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates);
    }

    const fileDiscrepancies = discrepancyData.slice(initialDiscrepancyCount);

    Logger.log(`    -> Discrepancy Check for ${file.getName()}: Found ${fileDiscrepancies.length} issues.`);

    if (fileDiscrepancies.length > 0) {
      Logger.log(`    -> DRAFTING DISCREPANCY EMAIL (Skipping GL PDF) for ${file.getName()}`);
      fileDiscrepancies.forEach((d, i) => {
         Logger.log(`       [Error ${i+1}] Date: ${d[1]}, TU: '${d[2]}', Issue: ${d[5]}`);
      });
      createDiscrepancyEmailDraft(carrierName, file.getName(), fileDiscrepancies, emailTemplate);
    } 
    else if (rdcName !== 'UNKNOWN' && fileCostSummary.length > 0) {
      Logger.log(`    -> Applying GL Codes & generating PDF for ${file.getName()}. Total cost items found: ${fileCostSummary.length}`);
      applyGLCodesAndExportPDF(invoiceSS, file.getName(), rdcName, fileCostSummary, glConfig, targetFolder);
    } 
    else {
      Logger.log(`    -> SKIPPING GL PDF for ${file.getName()}: RDC is '${rdcName}' (Must not be UNKNOWN) AND Cost Summary length is ${fileCostSummary.length} (Must be > 0).`);
    }

  } catch (e) { Logger.log(`Error processing sheet: ${e.message}`); }
  
  if (isTemp) {
    try { DriveApp.getFileById(sheetId).setTrashed(true); } catch(e) { }
  }
}

// --- EMAIL DRAFTING LOGIC ---

function createDiscrepancyEmailDraft(carrierName, fileName, discrepancies, emailTemplate) {
  const subjTemplate = (emailTemplate && emailTemplate['Subject']) ? emailTemplate['Subject'] : `Discrepancy Notice: Invoice {FileName}`;
  const greetTemplate = (emailTemplate && emailTemplate['Greeting']) ? emailTemplate['Greeting'] : `Hello {CarrierName} Team,\n\nThe following items are showing as discrepancies in our system. Please advise:`;
  const outroTemplate = (emailTemplate && emailTemplate['Outro']) ? emailTemplate['Outro'] : `Thank you.`;

  const subject = subjTemplate.replace(/{FileName}/g, fileName).replace(/{CarrierName}/g, carrierName);
  const greetingText = greetTemplate.replace(/{CarrierName}/g, carrierName).replace(/\n/g, '<br>');
  const outroText = outroTemplate.replace(/\n/g, '<br>');
  
  let plainBody = `${greetingText.replace(/<br>/g, '\n')}\n\n`;
  
  let htmlBody = `<div style="font-family: Arial, sans-serif; color: #333;">`;
  htmlBody += `<p>${greetingText}</p>`;
  htmlBody += `<table style="border-collapse: collapse; width: 100%; max-width: 800px; margin-top: 15px;">`;
  htmlBody += `<thead>`;
  htmlBody += `<tr style="background-color: #f2f2f2;">`;
  htmlBody += `<th style="border: 1px solid #000; padding: 10px; text-align: left;">Date</th>`;
  htmlBody += `<th style="border: 1px solid #000; padding: 10px; text-align: left;">TU</th>`;
  htmlBody += `<th style="border: 1px solid #000; padding: 10px; text-align: left;">Issue</th>`;
  htmlBody += `<th style="border: 1px solid #000; padding: 10px; text-align: left;">Billed Store/Tour</th>`;
  htmlBody += `</tr>`;
  htmlBody += `</thead>`;
  htmlBody += `<tbody>`;

  discrepancies.forEach(d => {
    let dateStr = d[1];
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    }
    
    plainBody += `• Date: ${dateStr} | TU: ${d[2]} | Issue: ${d[5]} (Billed: ${d[4]})\n`;
    
    htmlBody += `<tr>`;
    htmlBody += `<td style="border: 1px solid #000; padding: 8px;">${dateStr}</td>`;
    htmlBody += `<td style="border: 1px solid #000; padding: 8px;">${d[2]}</td>`;
    htmlBody += `<td style="border: 1px solid #000; padding: 8px;">${d[5]}</td>`;
    htmlBody += `<td style="border: 1px solid #000; padding: 8px;">${d[4]}</td>`;
    htmlBody += `</tr>`;
  });

  plainBody += `\n${outroText.replace(/<br>/g, '\n')}`;
  
  htmlBody += `</tbody>`;
  htmlBody += `</table>`;
  htmlBody += `<p style="margin-top: 20px;">${outroText}</p>`;
  htmlBody += `</div>`;

  GmailApp.createDraft("", subject, plainBody, { htmlBody: htmlBody });
}

// --- GL CODING & PDF EXPORT LOGIC ---

function applyGLCodesAndExportPDF(invoiceSS, originalFileName, rdcName, fileCostSummary, glConfig, targetFolder) {
  if (!glConfig || fileCostSummary.length === 0) return;
  
  const costCenter = glConfig.costCenters[rdcName] || 'XXXXX';
  const glTotals = {};
  let grandTotal = 0;
  
  fileCostSummary.forEach(item => {
    const descLower = item.desc.toLowerCase();
    let glPrefix = glConfig.defaultCategory || 'XXXXXX, XXXXXXXX';
    
    for (let cat of glConfig.categories) {
      if (descLower.includes(cat.keyword)) {
        glPrefix = cat.glPrefix; break;
      }
    }
    
    const fullGlString = `${glPrefix}, ${costCenter}`;
    if (!glTotals[fullGlString]) glTotals[fullGlString] = 0;
    glTotals[fullGlString] += item.amount;
    grandTotal += item.amount;
  });
  
  let stampRows = [];
  stampRows.push(['GL CODING SUMMARY', 'GL ACCOUNT / COST CENTER']);
  for (let gl in glTotals) {
     if (glTotals[gl] > 0) {
        stampRows.push([`$${glTotals[gl].toFixed(2)}`, gl]);
     }
  }
  stampRows.push([`$${grandTotal.toFixed(2)}`, `TOTAL INVOICE AMOUNT`]);
  
  let summarySheet = null;
  for (let sheet of invoiceSS.getSheets()) {
    const name = sheet.getName().toLowerCase();
    if (name.includes('summary') || name.includes('invoice')) { summarySheet = sheet; break; }
  }
  if (!summarySheet) summarySheet = invoiceSS.getSheets()[0];
  
  summarySheet.insertRowsBefore(1, stampRows.length + 1);
  const range = summarySheet.getRange(1, 1, stampRows.length, 2);
  range.setValues(stampRows);
  
  // Styling the stamp
  range.setFontWeight('bold').setFontColor('#cc0000').setWrap(false);
  range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  summarySheet.getRange(1, 1, 1, 2).setBackground('#ffff00').setFontColor('#000000');
  
  // Auto-resize and enforce minimum widths so SCH/narrow tabs don't squish the PDF
  summarySheet.autoResizeColumn(1);
  summarySheet.autoResizeColumn(2);
  if (summarySheet.getColumnWidth(1) < 160) summarySheet.setColumnWidth(1, 160);
  if (summarySheet.getColumnWidth(2) < 280) summarySheet.setColumnWidth(2, 280);
  
  const lastCol = summarySheet.getLastColumn();
  const maxCols = summarySheet.getMaxColumns();
  if (maxCols > lastCol) {
    summarySheet.hideColumns(lastCol + 1, maxCols - lastCol);
  }
  
  SpreadsheetApp.flush(); 
  
  const pdfName = originalFileName.replace(/\.xlsx$/i, '') + ' - CODED.pdf';
  exportSheetToPDF(invoiceSS.getId(), summarySheet.getSheetId(), pdfName, targetFolder);
}

function exportSheetToPDF(spreadsheetId, sheetId, pdfName, folder) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?exportFormat=pdf&format=pdf&size=letter&portrait=true&scale=2&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false&gid=${sheetId}`;
  const token = ScriptApp.getOAuthToken();
  const options = { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true };
  
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    const blob = response.getBlob().setName(pdfName);
    folder.createFile(blob);
  } else {
    Logger.log(`[ERROR] PDF Export failed: ${response.getContentText()}`);
  }
}

// --- UNIVERSAL TMST & DISCREPANCY LOGIC ---

function extractInvoiceNumber(invoiceSS, fallbackName) {
  try {
    const sheets = invoiceSS.getSheets();
    
    for (let s = 0; s < sheets.length; s++) {
      const sheetName = sheets[s].getName().toLowerCase();
      if (!sheetName.includes('summary') && !sheetName.includes('invoice')) continue;
      
      const data = sheets[s].getDataRange().getValues();
      for (let r = 0; r < Math.min(data.length, 40); r++) {
        for (let c = 0; c < Math.min(data[r].length, 10); c++) {
          let cellVal = String(data[r][c]).trim();
          if (!cellVal) continue;
          
          let match = cellVal.match(/Invoice\s*(?:#|ID|Number)\s*:?\s*([A-Za-z0-9_-]+)/i);
          if (match && match[1]) return match[1].trim();
          
          if (cellVal.match(/^Invoice\s*(?:#|ID|Number)\s*:?$/i)) {
            for (let nextC = c + 1; nextC < Math.min(data[r].length, c + 4); nextC++) {
              let nextVal = String(data[r][nextC]).trim();
              if (nextVal) return nextVal;
            }
          }
        }
      }
    }
    
    const data = sheets[0].getDataRange().getValues();
    for (let r = 0; r < Math.min(data.length, 30); r++) {
      for (let c = 0; c < Math.min(data[r].length, 10); c++) {
          let cellVal = String(data[r][c]).trim();
          if (!cellVal) continue;
          
          let match = cellVal.match(/Invoice\s*(?:#|ID|Number)\s*:?\s*([A-Za-z0-9_-]+)/i);
          if (match && match[1]) return match[1].trim();
          
          if (cellVal.match(/^Invoice\s*(?:#|ID|Number)\s*:?$/i)) {
            for (let nextC = c + 1; nextC < Math.min(data[r].length, c + 4); nextC++) {
              let nextVal = String(data[r][nextC]).trim();
              if (nextVal) return nextVal;
            }
          }
      }
    }
  } catch (e) {
    Logger.log(`[WARNING] Failed to extract invoice number from sheet: ${e.message}`);
  }
  return fallbackName;
}

function cleanTuNumber(tu) {
  if (!tu) return "";
  let cleanTu = String(tu).trim().toUpperCase();
  // Strip out "US" followed by exactly 4 digits (e.g., US0005) which represents the warehouse code
  cleanTu = cleanTu.replace(/^US\s*\d{4}/i, '');
  return cleanTu;
}

function evaluateTmstRow(rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, carrierName, haulierInfo, discrepancyData, invoiceNumber, haulierUpdates) {
  let type = 'INVOICE'; let match = 'NO'; let shiftMatch = 'UNMATCHED'; let storeMatch = 'UNMATCHED'; let tourMatch = 'UNMATCHED';
  const safeTuNumber = cleanTuNumber(tuNumber);
  
  // Extract intelligent Tour number based on TU Suffix length
  let invoiceTour = '1';
  let baseTu = safeTuNumber;
  if (safeTuNumber.length >= 8) {
     const lastChar = safeTuNumber.slice(-1);
     if (!isNaN(lastChar) && safeTuNumber.length === 9) { // Strict logic: base 8-digit TU + 1-digit tour identifier
       if (lastChar === '0') { invoiceTour = '1'; baseTu = safeTuNumber.slice(0, -1); }
       else if (lastChar === '1') { invoiceTour = '2'; baseTu = safeTuNumber.slice(0, -1); }
       else if (lastChar === '2') { invoiceTour = '3'; baseTu = safeTuNumber.slice(0, -1); }
       else if (lastChar === '3') { invoiceTour = '4'; baseTu = safeTuNumber.slice(0, -1); }
     }
  }

  let haulierRecord = null;
  const haulierMap = haulierInfo.records || {};
  
  Logger.log(`      -> Evaluating TU: '${tuNumber}' (Cleaned: '${safeTuNumber}', Base: '${baseTu}', Tour: ${invoiceTour})`);
  
  if (haulierMap[safeTuNumber]) {
    haulierRecord = haulierMap[safeTuNumber];
  } else if (haulierMap[baseTu]) {
    haulierRecord = haulierMap[baseTu];
  } else {
    for (let haulierKey in haulierMap) {
      const safeKey = String(haulierKey).trim().toUpperCase();
      if (safeKey.length > 4 && (safeTuNumber.includes(safeKey) || safeKey.includes(safeTuNumber) || baseTu.includes(safeKey) || safeKey.includes(baseTu))) {
        haulierRecord = haulierMap[haulierKey]; break;
      }
    }
  }
  
  if (haulierRecord) {
    match = 'YES'; type = haulierRecord.deliveryType || 'MATCHED'; 
    
    // Strict Store matching (TEMPORARILY DISABLED PER USER REQUEST)
    const invoiceStoreStr = String(store).trim();
    const storeRegexMatch = invoiceStoreStr.match(/(?:US)?(\d{4})/i);
    const invoiceStoreNum = storeRegexMatch ? storeRegexMatch[1] : invoiceStoreStr;
    
    // TEMPORARILY DISABLED: Force to MATCHED so it doesn't block write-backs or trigger discrepancies
    storeMatch = 'MATCHED'; 
    /*
    if (String(haulierRecord.store).includes(invoiceStoreNum) || invoiceStoreNum === "") {
       storeMatch = 'MATCHED';
    } else {
       storeMatch = 'UNMATCHED';
    }
    */
    
    // Strict Tour matching 
    let haulierTourStr = String(haulierRecord.tour).toUpperCase().trim();
    if (haulierTourStr === invoiceTour || haulierTourStr.includes(invoiceTour) || haulierTourStr.includes('0' + invoiceTour)) {
        tourMatch = 'MATCHED';
    } else if (!haulierTourStr) {
        tourMatch = 'MATCHED'; // Default to matched if Haulier report has no Tour column configured
    } else {
        tourMatch = 'UNMATCHED';
    }
    
    // Strict Shift matching
    let invoiceShiftStr = String(shift).toUpperCase();
    let haulierShiftStr = String(haulierRecord.shift).toUpperCase();
    
    let invIsShift1 = invoiceShiftStr.includes('1') || invoiceShiftStr.includes('FIRST');
    let invIsShift2 = invoiceShiftStr.includes('2') || invoiceShiftStr.includes('SECOND');
    let hauIsShift1 = haulierShiftStr.includes('1') || haulierShiftStr.includes('FIRST');
    let hauIsShift2 = haulierShiftStr.includes('2') || haulierShiftStr.includes('SECOND');
    
    if ((invIsShift1 && hauIsShift1) || (invIsShift2 && hauIsShift2) || (invoiceShiftStr === haulierShiftStr && invoiceShiftStr !== "")) {
        shiftMatch = 'MATCHED';
    } else if (!haulierShiftStr) {
        shiftMatch = 'MATCHED'; // Default if haulier report has no Shift column
    } else {
        shiftMatch = 'UNMATCHED';
    }
    
    // Discrepancy Logging for valid TU but mismatched data
    if (storeMatch === 'UNMATCHED') discrepancyData.push([rdcName, date, tuNumber, haulierRecord.store, store, 'Store Mismatch']);
    if (shiftMatch === 'UNMATCHED') discrepancyData.push([rdcName, date, tuNumber, haulierRecord.shift, shift, 'Shift Mismatch']);
    if (tourMatch === 'UNMATCHED') discrepancyData.push([rdcName, date, tuNumber, haulierRecord.tour, invoiceTour, 'Tour Mismatch']);
    
    // Push updates to be safely written to the Haulier Report later
    if (haulierUpdates && haulierUpdates[rdcName]) {
      let writeAmount = totalCost;
      
      // Core rule: Block the amount write-back if NOT a perfect match
      if (storeMatch === 'UNMATCHED' || tourMatch === 'UNMATCHED' || shiftMatch === 'UNMATCHED') {
         writeAmount = ""; 
      }
      
      haulierUpdates[rdcName].push({
        rowNumber: haulierRecord.rowNumber,
        invoiceNumber: invoiceNumber,
        amount: writeAmount
      });
    }

  } else {
    Logger.log(`      -> [DISCREPANCY ALERT] TU: '${safeTuNumber}' not found in Haulier report. Date: ${date}, Store: ${store}`);
    discrepancyData.push([rdcName, date, tuNumber, 'N/A', store, 'TU Not Found in Haulier Report']);
  }
  
  return [rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, type, match, shiftMatch, storeMatch, tourMatch, carrierName];
}

// --- CARRIER SPECIFIC MODULES ---

function processHbInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates) {
  invoiceSS.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return;
    
    let isShiftTab = false; let headerRowIndex = -1; let headers = [];
    for (let r = 0; r < Math.min(data.length, 10); r++) {
      if (rowHasHeader(data[r], 'TU', headerAliases) && (rowHasHeader(data[r], 'Date', headerAliases) || rowHasHeader(data[r], 'Shift', headerAliases) || rowHasHeader(data[r], 'Store', headerAliases))) {
        isShiftTab = true; headerRowIndex = r; headers = data[r]; break;
      }
    }
    if (!isShiftTab && sheetName.toLowerCase().includes('shift')) { isShiftTab = true; headerRowIndex = 0; headers = data[0]; }

    if (isShiftTab && headers.length > 0) {
      let dateIdx = -1, shiftIdx = -1, tuIdx = -1, storeIdx = -1, milesIdx = -1, nyPayIdx = -1, tollsIdx = -1, totalCostIdx = -1;
      
      // Pass 1: Exact matches first
      for (let j = 0; j < headers.length; j++) {
        const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
        if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, true)) totalCostIdx = j;
        else if (shiftIdx === -1 && matchHeader(h, 'Shift', headerAliases, true) && !h.includes('total')) shiftIdx = j;
        else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, true)) dateIdx = j;
        else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, true)) tuIdx = j;
        else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, true)) storeIdx = j;
        else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, true)) milesIdx = j;
        else if (nyPayIdx === -1 && matchHeader(h, 'NY Pay', headerAliases, true)) nyPayIdx = j;
        else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, true)) tollsIdx = j;
      }

      // Pass 2: Partial Matches
      for (let j = 0; j < headers.length; j++) {
        const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
        if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, false)) totalCostIdx = j;
        else if (shiftIdx === -1 && matchHeader(h, 'Shift', headerAliases, false) && !h.includes('total')) shiftIdx = j;
        else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, false)) dateIdx = j;
        else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, false)) tuIdx = j;
        else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, false)) storeIdx = j;
        else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, false)) milesIdx = j;
        else if (nyPayIdx === -1 && matchHeader(h, 'NY Pay', headerAliases, false)) nyPayIdx = j;
        else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, false)) tollsIdx = j;
      }

      for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        const date = dateIdx !== -1 ? row[dateIdx] : "";
        const tuNumber = tuIdx !== -1 ? cleanTuNumber(row[tuIdx]) : "";
        
        // Skip junk data, empty rows, or generic subtotal labels
        if (!tuNumber || String(tuNumber).includes('TOTAL') || String(tuNumber).length < 4) continue;
        
        const shift = shiftIdx !== -1 && row[shiftIdx] !== undefined ? row[shiftIdx] : "";
        const store = storeIdx !== -1 && row[storeIdx] !== undefined ? row[storeIdx] : "";
        const miles = milesIdx !== -1 && row[milesIdx] !== undefined ? row[milesIdx] : "";
        const nyPay = nyPayIdx !== -1 && row[nyPayIdx] !== undefined ? row[nyPayIdx] : "";
        const tolls = tollsIdx !== -1 && row[tollsIdx] !== undefined ? row[tollsIdx] : "";
        const totalCost = totalCostIdx !== -1 && row[totalCostIdx] !== undefined ? row[totalCostIdx] : "";
        
        masterData.push([rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost]);
        tmstData.push(evaluateTmstRow(rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, carrierName, haulierInfo, discrepancyData, invoiceNumber, haulierUpdates));
      }
    }
    
    if (sheetName.toLowerCase().includes('invoice')) {
      Logger.log(`        -> Processing HB Invoice tab for additional costs...`);
      let costsFound = 0;
      let inChargesSection = false;
      
      for (let i = 0; i < data.length; i++) {
        let rowStr = data[i].join(" ").toLowerCase();
        
        // Auto-detect start of charges to avoid reading invoice # or dates as costs
        if (!inChargesSection) {
           if (rowStr.includes('dedicated charges') || rowStr.includes('qty') || rowStr.includes('rate')) {
              inChargesSection = true;
           }
        }
        
        let desc = "";
        // Safely scan the first 3 columns for the description to handle empty column formatting issues
        for (let col = 0; col < Math.min(data[i].length, 3); col++) {
          if (String(data[i][col]).trim() !== "") {
            desc = String(data[i][col]).trim();
            break;
          }
        }
        
        let descLower = desc.toLowerCase();
        
        // Highly strict metadata skipping
        if (!desc || descLower.includes('total') || descLower.includes('invoice') || descLower.includes('week ending') || descLower.includes('diesel') || descLower.includes('truck count') || descLower.includes('@lidl') || descLower.includes('lidl us') || descLower.includes('arlington') || descLower.includes('http') || descLower === 'qty' || descLower.includes('dedicated charges')) continue;
        
        // Failsafe: if we haven't hit the charges section yet, only accept obvious charge lines
        if (!inChargesSection && !descLower.includes('charge') && !descLower.includes('shift') && !descLower.includes('mileage') && !descLower.includes('toll')) continue;
        
        // Apply dynamic Ignore Rules
        let shouldIgnore = false;
        if (glConfig && glConfig.ignoreKeywords) {
           for (let kw of glConfig.ignoreKeywords) {
              if (kw && descLower.includes(kw)) { shouldIgnore = true; break; }
           }
        }
        if (shouldIgnore) continue;
        
        let amount = 0;
        // Parse from right to left to grab the final numerical value
        for (let col = data[i].length - 1; col >= 0; col--) {
          const rawVal = String(data[i][col]).replace(/[^0-9.-]+/g, "");
          if (rawVal !== "" && rawVal !== "-") {
            const val = parseFloat(rawVal);
            if (!isNaN(val)) { amount = val; break; }
          }
        }
        
        // Apply dynamic Tolerance logic
        const tolerance = (glConfig && glConfig.tolerance) ? glConfig.tolerance : 0.01;
        if (amount >= tolerance) {
          fileCostSummary.push({ desc: desc, amount: amount });
          costsFound++;
          const isBaseCharge = /mileage|fuel|tolls|shift/i.test(desc);
          if (!isBaseCharge) additionalCostsData.push([rdcName, carrierName, desc, amount, invoiceSS.getName()]);
        }
      }
      Logger.log(`        -> Found ${costsFound} valid cost line items on HB Invoice tab.`);
    }
  });
}

function processCreInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates) {
  invoiceSS.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return;
    
    if (sheetName.toLowerCase().includes('detail')) {
      let headerRowIndex = -1; let headers = [];
      for (let r = 0; r < Math.min(data.length, 10); r++) {
        if (rowHasHeader(data[r], 'TU', headerAliases) && rowHasHeader(data[r], 'Date', headerAliases)) { headerRowIndex = r; headers = data[r]; break; }
      }
      
      if (headerRowIndex !== -1) {
        let dateIdx = -1, tuIdx = -1, storeIdx = -1, milesIdx = -1, nyPayIdx = -1, tollsIdx = -1, totalCostIdx = -1, firstShiftIdx = -1, secondShiftIdx = -1;
        
        // Pass 1: Exact matches
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
          if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, true)) totalCostIdx = j;
          else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, true)) dateIdx = j;
          else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, true)) tuIdx = j;
          else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, true)) storeIdx = j;
          else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, true)) milesIdx = j;
          else if (nyPayIdx === -1 && matchHeader(h, 'NY Pay', headerAliases, true)) nyPayIdx = j;
          else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, true)) tollsIdx = j;
          else if (firstShiftIdx === -1 && h === 'first shift') firstShiftIdx = j;
          else if (secondShiftIdx === -1 && h === 'second shift') secondShiftIdx = j;
        }

        // Pass 2: Partial matches
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
          if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, false)) totalCostIdx = j;
          else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, false)) dateIdx = j;
          else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, false)) tuIdx = j;
          else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, false)) storeIdx = j;
          else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, false)) milesIdx = j;
          else if (nyPayIdx === -1 && matchHeader(h, 'NY Pay', headerAliases, false)) nyPayIdx = j;
          else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, false)) tollsIdx = j;
          else if (firstShiftIdx === -1 && h.includes('first shift')) firstShiftIdx = j;
          else if (secondShiftIdx === -1 && h.includes('second shift')) secondShiftIdx = j;
        }

        for (let i = headerRowIndex + 1; i < data.length; i++) {
          const row = data[i];
          const date = dateIdx !== -1 ? row[dateIdx] : "";
          const tuNumber = tuIdx !== -1 ? cleanTuNumber(row[tuIdx]) : "";
          
          if (!tuNumber || String(tuNumber).includes('TOTAL') || String(tuNumber).length < 4) continue; 
          
          let shift = "Shift 1";
          if (firstShiftIdx !== -1 && parseFloat(row[firstShiftIdx]) > 0) shift = "Shift 1";
          else if (secondShiftIdx !== -1 && parseFloat(row[secondShiftIdx]) > 0) shift = "Shift 2";
          
          const store = storeIdx !== -1 && row[storeIdx] !== undefined ? row[storeIdx] : "";
          const miles = milesIdx !== -1 && row[milesIdx] !== undefined ? row[milesIdx] : "";
          const nyPay = nyPayIdx !== -1 && row[nyPayIdx] !== undefined ? row[nyPayIdx] : "";
          const tolls = tollsIdx !== -1 && row[tollsIdx] !== undefined ? row[tollsIdx] : "";
          const totalCost = totalCostIdx !== -1 && row[totalCostIdx] !== undefined ? row[totalCostIdx] : "";
          
          masterData.push([rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost]);
          tmstData.push(evaluateTmstRow(rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, carrierName, haulierInfo, discrepancyData, invoiceNumber, haulierUpdates));
        }
      }
    }
    
    if (sheetName.toLowerCase().includes('summary')) {
      Logger.log(`        -> Processing CRE Summary tab for additional costs...`);
      let headerRowIndex = -1; let headers = [];
      for (let r = 0; r < Math.min(data.length, 25); r++) {
        const rowStr = data[r].join(" ").toLowerCase();
        if (rowStr.includes('description') && rowStr.includes('total')) { headerRowIndex = r; headers = data[r]; break; }
      }
      
      if (headerRowIndex !== -1) {
        let descIdx = -1, totalIdx = -1;
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim();
          if (descIdx === -1 && h.includes('description')) descIdx = j;
          else if (totalIdx === -1 && h === 'total') totalIdx = j;
        }
        
        let costsFound = 0;
        for (let i = headerRowIndex + 1; i < data.length; i++) {
          let desc = "";
          for (let col = 0; col < Math.min(data[i].length, 3); col++) {
            if (String(data[i][col]).trim() !== "") { desc = String(data[i][col]).trim(); break; }
          }
          let descLower = desc.toLowerCase();
          
          if (!desc || descLower.includes('fixed charges') || descLower.includes('variable charges') || descLower === 'total') continue;
          
          // Apply dynamic Ignore Rules
          let shouldIgnore = false;
          if (glConfig && glConfig.ignoreKeywords) {
             for (let kw of glConfig.ignoreKeywords) {
                if (kw && descLower.includes(kw)) { shouldIgnore = true; break; }
             }
          }
          if (shouldIgnore) continue;

          let amount = 0;
          for (let col = data[i].length - 1; col >= 0; col--) {
            const rawVal = String(data[i][col]).replace(/[^0-9.-]+/g, "");
            if (rawVal !== "" && rawVal !== "-") {
              const val = parseFloat(rawVal);
              if (!isNaN(val)) { amount = val; break; }
            }
          }
          
          const tolerance = (glConfig && glConfig.tolerance) ? glConfig.tolerance : 0.01;
          if (!isNaN(amount) && amount >= tolerance) {
            fileCostSummary.push({ desc: desc, amount: amount });
            costsFound++;
            const isBaseCharge = /first shift|second shift|linehaul|tolls/i.test(desc);
            if (!isBaseCharge) additionalCostsData.push([rdcName, carrierName, desc, amount, invoiceSS.getName()]);
          }
        }
        Logger.log(`        -> Found ${costsFound} valid cost line items on CRE Summary tab.`);
      }
    }
  });
}

function processSchInvoice(invoiceSS, rdcName, carrierName, haulierInfo, masterData, additionalCostsData, tmstData, discrepancyData, fileCostSummary, headerAliases, glConfig, invoiceNumber, haulierUpdates) {
  invoiceSS.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return;
    
    if (sheetName.toLowerCase().includes('order detail') || sheetName.toLowerCase().includes('customer detail')) {
      let headerRowIndex = -1; let headers = [];
      for (let r = 0; r < Math.min(data.length, 10); r++) {
        if (rowHasHeader(data[r], 'TU', headerAliases) && rowHasHeader(data[r], 'Miles', headerAliases)) { headerRowIndex = r; headers = data[r]; break; }
      }
      
      if (headerRowIndex !== -1) {
        let dateIdx = -1, tuIdx = -1, storeIdx = -1, milesIdx = -1, nyPayIdx1 = -1, nyPayIdx2 = -1, tollsIdx = -1, totalCostIdx = -1;
        
        // Pass 1: Exact matches
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
          if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, true)) totalCostIdx = j;
          else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, true)) dateIdx = j;
          else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, true)) tuIdx = j;
          else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, true)) storeIdx = j;
          else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, true)) milesIdx = j;
          else if (nyPayIdx1 === -1 && h === 'dhu $ amt') nyPayIdx1 = j;
          else if (nyPayIdx2 === -1 && h === 'stp $ amt') nyPayIdx2 = j;
          else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, true)) tollsIdx = j;
        }

        // Pass 2: Partial matches
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim().replace(/\s+/g, ' ');
          if (totalCostIdx === -1 && matchHeader(h, 'Total Cost', headerAliases, false)) totalCostIdx = j;
          else if (dateIdx === -1 && matchHeader(h, 'Date', headerAliases, false)) dateIdx = j;
          else if (tuIdx === -1 && matchHeader(h, 'TU', headerAliases, false)) tuIdx = j;
          else if (storeIdx === -1 && matchHeader(h, 'Store', headerAliases, false)) storeIdx = j;
          else if (milesIdx === -1 && matchHeader(h, 'Miles', headerAliases, false)) milesIdx = j;
          else if (nyPayIdx1 === -1 && h.includes('dhu $ amt')) nyPayIdx1 = j;
          else if (nyPayIdx2 === -1 && h.includes('stp $ amt')) nyPayIdx2 = j;
          else if (tollsIdx === -1 && matchHeader(h, 'Tolls', headerAliases, false)) tollsIdx = j;
        }

        for (let i = headerRowIndex + 1; i < data.length; i++) {
          const row = data[i];
          const tuNumber = tuIdx !== -1 ? cleanTuNumber(row[tuIdx]) : "";
          
          if (!tuNumber || String(tuNumber).includes('TOTAL') || String(tuNumber).length < 4) continue; 
          
          let date = dateIdx !== -1 ? row[dateIdx] : "";
          if (date && typeof date === 'string') date = date.split(' ')[0]; 
          
          let shift = "Shift 1"; 
          const store = storeIdx !== -1 && row[storeIdx] !== undefined ? row[storeIdx] : "";
          const miles = milesIdx !== -1 && row[milesIdx] !== undefined ? row[milesIdx] : "";
          const tolls = tollsIdx !== -1 && row[tollsIdx] !== undefined ? row[tollsIdx] : "";
          const totalCost = totalCostIdx !== -1 && row[totalCostIdx] !== undefined ? row[totalCostIdx] : "";
          
          let nyPay = 0;
          if (nyPayIdx1 !== -1 && row[nyPayIdx1]) nyPay += parseFloat(row[nyPayIdx1]) || 0;
          if (nyPayIdx2 !== -1 && row[nyPayIdx2]) nyPay += parseFloat(row[nyPayIdx2]) || 0;
          nyPay = nyPay > 0 ? nyPay : "";

          masterData.push([rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost]);
          tmstData.push(evaluateTmstRow(rdcName, date, shift, tuNumber, store, miles, nyPay, tolls, totalCost, carrierName, haulierInfo, discrepancyData, invoiceNumber, haulierUpdates));
        }
      }
    }
    
    if (sheetName.toLowerCase().includes('summary')) {
      Logger.log(`        -> Processing SCH Summary tab for additional costs...`);
      let headerRowIndex = -1; let headers = [];
      for (let r = 0; r < Math.min(data.length, 35); r++) {
        const rowStr = data[r].join(" ").toLowerCase();
        if (rowStr.includes('description') && rowStr.includes('total')) { headerRowIndex = r; headers = data[r]; break; }
      }
      
      if (headerRowIndex !== -1) {
        let descIdx = -1, totalIdx = -1;
        for (let j = 0; j < headers.length; j++) {
          const h = String(headers[j]).toLowerCase().trim();
          if (descIdx === -1 && h.includes('description')) descIdx = j;
          else if (totalIdx === -1 && h === 'total') totalIdx = j;
        }
        
        let costsFound = 0;
        for (let i = headerRowIndex + 1; i < data.length; i++) {
          let desc = "";
          for (let col = 0; col < Math.min(data[i].length, 3); col++) {
            if (String(data[i][col]).trim() !== "") { desc = String(data[i][col]).trim(); break; }
          }
          let descLower = desc.toLowerCase();
          
          if (!desc || descLower === 'total') continue;
          
          // Apply dynamic Ignore Rules
          let shouldIgnore = false;
          if (glConfig && glConfig.ignoreKeywords) {
             for (let kw of glConfig.ignoreKeywords) {
                if (kw && descLower.includes(kw)) { shouldIgnore = true; break; }
             }
          }
          if (shouldIgnore) continue;
          
          let amount = 0;
          for (let col = data[i].length - 1; col >= 0; col--) {
            const rawVal = String(data[i][col]).replace(/[^0-9.-]+/g, "");
            if (rawVal !== "" && rawVal !== "-") {
              const val = parseFloat(rawVal);
              if (!isNaN(val)) { amount = val; break; }
            }
          }
          
          const tolerance = (glConfig && glConfig.tolerance) ? glConfig.tolerance : 0.01;
          if (!isNaN(amount) && amount >= tolerance) {
            fileCostSummary.push({ desc: desc, amount: amount });
            costsFound++;
            const isBaseCharge = /linehaul|driver charge|slip seat|fuel surcharge|tolls/i.test(desc);
            if (!isBaseCharge) additionalCostsData.push([rdcName, carrierName, desc, amount, invoiceSS.getName()]);
          }
        }
        Logger.log(`        -> Found ${costsFound} valid cost line items on SCH Summary tab.`);
      }
    }
  });
}

// --- HELPER FUNCTIONS ---

function fetchHaulierData(spreadsheetId, headerAliases, rdcLogName = "Unknown RDC") {
  const info = { spreadsheetId: spreadsheetId, sheetName: null, records: {}, invoiceColIdx: -1, amountColIdx: -1, maxCol: 0 };
  if (!spreadsheetId) return info;
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = null;
    for (let s = 0; s < ss.getSheets().length; s++) {
      const name = ss.getSheets()[s].getName().toLowerCase();
      if (name.includes('weekly') && (name.includes('haulier') || name.includes('hauler'))) { sheet = ss.getSheets()[s]; break; }
    }
    if (!sheet) {
      Logger.log(`[WARNING] Haulier sheet for ${rdcLogName} is missing 'Weekly Haulier' tab.`);
      return info;
    }
    
    info.sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    let tuIndex = -1, storeIndex = -1, typeIndex = -1, shiftIndex = -1, tourIndex = -1;
    let invCol = -1, amtCol = -1;
    const headers = data[0];
    
    // Pass 1: Exact matches
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i]).toLowerCase().trim().replace(/\s+/g, ' ');
      if (typeIndex === -1 && matchHeader(h, 'Type', headerAliases, true)) typeIndex = i;
      else if (storeIndex === -1 && matchHeader(h, 'Store', headerAliases, true)) storeIndex = i;
      else if (tourIndex === -1 && matchHeader(h, 'Tour', headerAliases, true)) tourIndex = i;
      else if (shiftIndex === -1 && matchHeader(h, 'Shift', headerAliases, true)) shiftIndex = i;
      else if (tuIndex === -1 && matchHeader(h, 'TU', headerAliases, true) && !h.includes('status') && !h.includes('return')) tuIndex = i;
      else if (invCol === -1 && matchHeader(h, 'Haulier Invoice', headerAliases, true)) invCol = i;
      else if (amtCol === -1 && matchHeader(h, 'Haulier Amount', headerAliases, true)) amtCol = i;
    }
    
    // Pass 2: Partial matches
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i]).toLowerCase().trim().replace(/\s+/g, ' ');
      if (typeIndex === -1 && matchHeader(h, 'Type', headerAliases, false)) typeIndex = i;
      else if (storeIndex === -1 && matchHeader(h, 'Store', headerAliases, false)) storeIndex = i;
      else if (tourIndex === -1 && matchHeader(h, 'Tour', headerAliases, false)) tourIndex = i;
      else if (shiftIndex === -1 && matchHeader(h, 'Shift', headerAliases, false)) shiftIndex = i;
      else if (tuIndex === -1 && matchHeader(h, 'TU', headerAliases, false) && !h.includes('status') && !h.includes('return')) tuIndex = i;
      else if (invCol === -1 && matchHeader(h, 'Haulier Invoice', headerAliases, false)) invCol = i;
      else if (amtCol === -1 && matchHeader(h, 'Haulier Amount', headerAliases, false)) amtCol = i;
    }
    
    if (tuIndex === -1) tuIndex = 3; // Enforced fallback to Column D (Index 3)
    if (storeIndex === -1) storeIndex = 4; 
    
    // Apply RDC specific columns for Invoice write-backs
    if (invCol === -1) {
      invCol = (rdcLogName === 'GRM') ? 20 : 19; // U for GRM, T for FRG/PYE
    }
    if (amtCol === -1) {
      amtCol = (rdcLogName === 'GRM') ? 21 : 20; // V for GRM, U for FRG/PYE
    }
    
    info.invoiceColIdx = invCol;
    info.amountColIdx = amtCol;
    info.maxCol = headers.length;
    
    let loadedCount = 0;
    const map = {};
    for (let i = 1; i < data.length; i++) {
      if (!data[i][tuIndex]) continue;
      const tu = cleanTuNumber(data[i][tuIndex]);
      map[tu] = {
        store: data[i][storeIndex],
        tour: tourIndex !== -1 ? String(data[i][tourIndex]).trim() : '',
        shift: shiftIndex !== -1 ? String(data[i][shiftIndex]).trim() : '',
        deliveryType: typeIndex !== -1 ? String(data[i][typeIndex]).trim() : 'MATCHED',
        rawRow: data[i],
        rowNumber: i + 1 
      };
      loadedCount++;
    }
    info.records = map;
    Logger.log(`      -> Loaded ${loadedCount} valid TU records from Haulier Report for ${rdcLogName}.`);
  } catch(e) {
    Logger.log(`[ERROR] Could not load haulier sheet ID ${spreadsheetId} for ${rdcLogName}: ${e.message}`);
  }
  return info;
}

function parseRDCName(text, rdcAliases) {
  const upper = text.toUpperCase();
  
  if (rdcAliases) {
     for (let rdc in rdcAliases) {
        for (let alias of rdcAliases[rdc]) {
           if (alias && upper.includes(alias)) return rdc;
        }
     }
  }
  
  // Fallback if aliases are missing or nothing matched
  if (upper.includes('FREDERICKSBURG') || upper.includes('FRG') || upper.includes(' VA') || upper.includes('LIDL VA') || upper.includes('_VA')) return 'FRG';
  if (upper.includes('GRAHAM') || upper.includes('GRM') || upper.includes(' NC') || upper.includes('LIDL NC') || upper.includes('_NC') || upper.includes('MEBANE')) return 'GRM';
  if (upper.includes('PERRYVILLE') || upper.includes('PYE') || upper.includes(' MD') || upper.includes('LIDL MD') || upper.includes('_MD') || upper.includes('PER')) return 'PYE';
  
  return 'UNKNOWN';
}

function convertExcelToGoogleSheet(excelFileId, parentFolderId) {
  const file = DriveApp.getFileById(excelFileId);
  const blob = file.getBlob();
  try {
    const resource = { name: "[TEMP] " + file.getName().replace(/\.xlsx$/i, ""), parents: [parentFolderId], mimeType: MimeType.GOOGLE_SHEETS };
    return Drive.Files.create(resource, blob, { supportsAllDrives: true }).id;
  } catch (e) {
    if (typeof Drive.Files.insert === "function") {
      const config = { title: "[TEMP] " + file.getName().replace(/\.xlsx$/i, ""), parents: [{id: parentFolderId}], mimeType: MimeType.GOOGLE_SHEETS };
      return Drive.Files.insert(config, blob, { supportsAllDrives: true }).id;
    } else { throw e; }
  }
}

function writeDataToSheet(ss, sheetName, dataArray, targetColumnCount) {
  if (dataArray.length === 0) return;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = [];
    for(let i=1; i<=targetColumnCount; i++) headers.push("Column " + i);
    sheet.appendRow(headers);
  }
  
  const cleanData = dataArray.map(row => {
    const newRow = [...row];
    while (newRow.length < targetColumnCount) newRow.push("");
    return newRow.slice(0, targetColumnCount);
  });
  
  sheet.getRange(sheet.getLastRow() + 1, 1, cleanData.length, targetColumnCount).setValues(cleanData);
}

function archiveAndClearSheets(ss) {
  const sheetsToProcess = ['Master Input', 'TMST', 'Additonal Costs', 'Discrepancy Tracker'];
  sheetsToProcess.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow > 1) {
        const dataToArchive = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
        let archiveSheet = ss.getSheetByName('Archive - ' + sheetName);
        if (!archiveSheet) {
          archiveSheet = ss.insertSheet('Archive - ' + sheetName);
          const header = sheet.getRange(1, 1, 1, lastCol).getValues();
          archiveSheet.appendRow(header[0]);
          archiveSheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#efefef');
        }
        archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, dataToArchive.length, lastCol).setValues(dataToArchive);
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    }
  });
}

function clearTrackerData(isWebApp = false) {
  if (!isWebApp) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Warning', 'Are you sure you want to clear and archive all current data in Master Input, TMST, Additional Costs, and Discrepancy Tracker?', ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
      archiveAndClearSheets(SpreadsheetApp.getActiveSpreadsheet());
      ui.alert('Data archived and cleared successfully.');
    }
  } else {
    archiveAndClearSheets(SpreadsheetApp.getActiveSpreadsheet());
    return 'Data archived and cleared successfully.';
  }
}

// --- WEB APP LOGIC ---

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('IMT Control Tower')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Helper: build a sheet summary object with embed/open URLs
function getSheetSummary_(ss, sheetName, label, description, tone) {
  var ssId  = ss.getId();
  var ssUrl = ss.getUrl();
  var sheet = ss.getSheetByName(sheetName);
  var exists = !!sheet;
  var gid = exists ? sheet.getSheetId() : 0;
  return {
    name:        sheetName,
    label:       label       || sheetName,
    description: description || '',
    tone:        tone        || 'neutral',
    exists:      exists,
    rows:        exists ? Math.max(0, sheet.getLastRow() - 1) : 0,
    openUrl:     exists ? ssUrl + '#gid=' + gid : '',
    embedUrl:    exists ? 'https://docs.google.com/spreadsheets/d/' + ssId + '/htmlview?gid=' + gid : ''
  };
}

function getDashboardData() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var ssId   = ss.getId();
  var ssUrl  = ss.getUrl();
  var ssName = ss.getName();

  // Config (safe: returns {} if System Config tab is missing)
  var config = {};
  try { config = getConfig(); } catch(e) { config = {}; }

  // Build carrier objects from all ROOT_FOLDER and HAULIER_ID config keys
  var allKeys  = Object.keys(config);
  var carriers = [];

  allKeys.filter(function(k) { return k.endsWith('_ROOT_FOLDER'); }).forEach(function(k) {
    var name     = k.replace('_ROOT_FOLDER', '');
    var folderId = String(config[k] || '').trim();
    carriers.push({
      name:                 name,
      classification:       'Invoice Carrier',
      primaryLane:          folderId !== '',
      rootFolderConfigured: folderId !== '',
      folderId:             folderId,
      driveUrl:             folderId ? 'https://drive.google.com/drive/folders/' + folderId : '',
      haulierLinked:        false,
      haulierId:            '',
      haulierUrl:           ''
    });
  });

  allKeys.filter(function(k) { return k.endsWith('_HAULIER_ID'); }).forEach(function(k) {
    var name    = k.replace('_HAULIER_ID', '');
    var haulId  = String(config[k] || '').trim();
    carriers.push({
      name:                 name,
      classification:       'Haulier RDC',
      primaryLane:          haulId !== '',
      rootFolderConfigured: false,
      folderId:             '',
      driveUrl:             '',
      haulierLinked:        haulId !== '',
      haulierId:            haulId,
      haulierUrl:           haulId ? 'https://docs.google.com/spreadsheets/d/' + haulId : ''
    });
  });

  // Sheet groups
  var operationsSheets = [
    getSheetSummary_(ss, 'Master Input',        'Master Input',        'Consolidated extracted invoice lines from all carriers', 'primary'),
    getSheetSummary_(ss, 'TMST',                'TMST',                'Haulier reconciliation and store/tour match output',     'primary'),
    getSheetSummary_(ss, 'Additonal Costs',     'Additional Costs',    'Non-base line items and surcharges',                     'secondary'),
    getSheetSummary_(ss, 'Discrepancy Tracker', 'Discrepancy Tracker', 'Invoice lines with unresolved cost or data mismatches',  'warning')
  ];

  var configSheets = [
    getSheetSummary_(ss, 'System Config',  'System Config',  'Carrier root folders, haulier IDs and processing settings', 'config'),
    getSheetSummary_(ss, 'GL Config',      'GL Config',      'GL account mapping rules, cost centre codes and tolerances', 'config'),
    getSheetSummary_(ss, 'RDC Aliases',    'RDC Aliases',    'RDC code alias mappings used to normalise sheet headers',   'config'),
    getSheetSummary_(ss, 'Email Template', 'Email Template', 'Discrepancy notification email subject and body templates', 'config'),
    getSheetSummary_(ss, 'Header Config',  'Header Config',  'Column header alias mappings for dynamic field detection',  'config')
  ];

  var viewerSheets = operationsSheets.concat(configSheets).filter(function(s) { return s.exists; });

  // Readiness scoring
  var rootCarriers           = carriers.filter(function(c) { return c.classification === 'Invoice Carrier'; });
  var haulierCarriers        = carriers.filter(function(c) { return c.classification === 'Haulier RDC'; });
  var configuredCarrierCount = rootCarriers.filter(function(c) { return c.rootFolderConfigured; }).length;
  var primaryCarrierCount    = rootCarriers.length;
  var linkedHaulierCount     = haulierCarriers.filter(function(c) { return c.haulierLinked; }).length;
  var configReadyCount       = configSheets.filter(function(s) { return s.exists; }).length;
  var missingConfigCount     = configSheets.filter(function(s) { return !s.exists; }).length;

  var masterRows      = operationsSheets[0].rows;
  var tmstRows        = operationsSheets[1].rows;
  var discrepancies   = operationsSheets[3].rows;
  var additionalCosts = operationsSheets[2].rows;
  var discrepancyRate = masterRows > 0 ? Math.round((discrepancies / masterRows) * 100) : 0;

  // Weighted readiness score: config 40%, carriers 30%, hauliers 30%
  var cfgScore    = configReadyCount / 5;
  var carrScore   = primaryCarrierCount  > 0 ? configuredCarrierCount / primaryCarrierCount  : 0;
  var haulScore   = haulierCarriers.length > 0 ? linkedHaulierCount   / haulierCarriers.length : 1;
  var readinessScore = Math.round(cfgScore * 40 + carrScore * 30 + haulScore * 30);

  // Alerts
  var alerts = [];
  if (missingConfigCount > 0) {
    alerts.push({ severity: 'critical', title: missingConfigCount + ' configuration tab(s) missing', detail: 'Run Initialize / Repair Config to create them.' });
  }
  if (configuredCarrierCount === 0 && primaryCarrierCount > 0) {
    alerts.push({ severity: 'warning', title: 'No carrier root folders configured', detail: 'Add Drive folder IDs to System Config.' });
  }
  if (discrepancyRate > 20) {
    alerts.push({ severity: 'warning', title: 'High discrepancy rate: ' + discrepancyRate + '%', detail: 'Review Discrepancy Tracker for open items.' });
  }
  if (configReadyCount === 5 && configuredCarrierCount > 0 && linkedHaulierCount > 0) {
    alerts.push({ severity: 'success', title: 'System fully operational', detail: 'All config tabs present and at least one carrier and haulier configured.' });
  }

  var consoleSeed = [
    { type: 'info',    message: 'Workbook: ' + ssName },
    { type: 'info',    message: 'Config sheets present: ' + configReadyCount + '/5' },
    { type: configuredCarrierCount > 0 ? 'success' : 'warning', message: 'Carrier lanes configured: ' + configuredCarrierCount + '/' + primaryCarrierCount },
    { type: linkedHaulierCount     > 0 ? 'success' : 'warning', message: 'Haulier RDCs linked: ' + linkedHaulierCount + '/' + haulierCarriers.length },
    { type: discrepancyRate        > 20 ? 'warning' : 'info',   message: 'Discrepancy rate: ' + discrepancyRate + '% (' + discrepancies + '/' + masterRows + ' rows)' }
  ];

  return {
    generatedAt:     new Date().toISOString(),
    spreadsheetName: ssName,
    sheetUrl:        ssUrl,
    workbook:        { name: ssName, openUrl: ssUrl, embedUrl: 'https://docs.google.com/spreadsheets/d/' + ssId + '/htmlview' },
    masterRows:      masterRows,
    tmstRows:        tmstRows,
    discrepancies:   discrepancies,
    additionalCosts: additionalCosts,
    carriers:        carriers,
    operationsSheets: operationsSheets,
    configSheets:    configSheets,
    viewerSheets:    viewerSheets,
    alerts:          alerts,
    summary: {
      readinessScore:        readinessScore,
      configuredCarrierCount: configuredCarrierCount,
      primaryCarrierCount:   primaryCarrierCount,
      linkedHaulierCount:    linkedHaulierCount,
      configReadyCount:      configReadyCount,
      missingConfigCount:    missingConfigCount,
      discrepancyRate:       discrepancyRate
    },
    consoleSeed: consoleSeed
  };
}

// Called by the UI "Initialize / Repair Config" button — identical sheet creation
// logic to setupConfigTab() but returns a structured result instead of a UI alert.
function setupConfigWeb() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var created = [];

  // 1. System Config
  if (!ss.getSheetByName('System Config')) {
    var sc = ss.insertSheet('System Config');
    sc.getRange(1, 1, 1, 3).setValues([['Setting Name', 'Value / ID', 'Description']]).setFontWeight('bold').setBackground('#d9ead3');
    sc.getRange(2, 1, 8, 3).setValues([
      ['FRG_HAULIER_ID', '1VyvVZ3BB2-tArldhuNKKBzfjU7FATVMJbtzhZP3P8pA', 'FRG Weekly Haulier Sheet ID'],
      ['GRM_HAULIER_ID', '13lvErboGKp-cd06XV1PvhB7zSrCEXXVm0POZlWPBZV4', 'GRM Weekly Haulier Sheet ID'],
      ['PYE_HAULIER_ID', '1AP_qyToLfhyuGihvfV3sUbLpj4WYx-wgAUVjtluHCoY', 'PYE Weekly Haulier Sheet ID'],
      ['CRE_ROOT_FOLDER', '1BCbMyWUu0npRGIQvN1lRtrL5oBh1PdAi', 'Root Folder ID for CRE Invoices'],
      ['HB_ROOT_FOLDER', '1awkIvApl6iyExTnrNTmD1Pt1oM1d-2GO', 'Root Folder ID for HB Invoices'],
      ['SCH_ROOT_FOLDER', '', 'Root Folder ID for SCH Invoices (Optional)'],
      ['WERNER_ROOT_FOLDER', '', 'Root Folder ID for Werner Invoices (Optional)'],
      ['TEMP_PROCESSING_FOLDER', '', 'Optional: Folder ID for temp converted files']
    ]);
    sc.setColumnWidth(1, 250); sc.setColumnWidth(2, 350); sc.setColumnWidth(3, 400);
    created.push('System Config');
  }

  // 2. GL Config
  if (!ss.getSheetByName('GL Config')) {
    var gc = ss.insertSheet('GL Config');
    gc.getRange(1, 1, 1, 4).setValues([['Rule Type', 'Keyword / RDC', 'GL Account / Center Code / Value', 'Description']]).setFontWeight('bold').setBackground('#fff2cc');
    gc.getRange(2, 1, 9, 4).setValues([
      ['Cost Center', 'FRG', '50001', 'Fredericksburg Cost Center Code'],
      ['Cost Center', 'GRM', '60001', 'Graham Cost Center Code'],
      ['Cost Center', 'PYE', '70001', 'Perryville Cost Center Code'],
      ['Category Mapping', 'toll', '471000, 47100002', 'Catches any line item containing "toll"'],
      ['Category Mapping', 'fuel', '471000, 47100099', 'Catches any line item containing "fuel"'],
      ['Default Category', 'BASE', '471000, 47100001', 'Base freight applied to all other costs'],
      ['Ignore Rule', 'discount', '', 'Ignores line items containing this word'],
      ['Ignore Rule', 'rebate', '', 'Ignores line items containing this word'],
      ['Tolerance', 'MIN_COST', '0.01', 'Ignores line items below this value']
    ]);
    gc.setColumnWidth(1, 150); gc.setColumnWidth(2, 150); gc.setColumnWidth(3, 220); gc.setColumnWidth(4, 350);
    created.push('GL Config');
  }

  // 3. RDC Aliases
  if (!ss.getSheetByName('RDC Aliases')) {
    var ra = ss.insertSheet('RDC Aliases');
    ra.getRange(1, 1, 1, 2).setValues([['RDC Code', 'Aliases (Comma Separated)']]).setFontWeight('bold').setBackground('#cfe2f3');
    ra.getRange(2, 1, 3, 2).setValues([
      ['FRG', 'FREDERICKSBURG, FRG, VA, LIDL VA, _VA'],
      ['GRM', 'GRAHAM, GRM, NC, LIDL NC, _NC, MEBANE'],
      ['PYE', 'PERRYVILLE, PYE, MD, LIDL MD, _MD, PER']
    ]);
    ra.setColumnWidth(1, 150); ra.setColumnWidth(2, 500);
    created.push('RDC Aliases');
  }

  // 4. Email Template
  if (!ss.getSheetByName('Email Template')) {
    var et = ss.insertSheet('Email Template');
    et.getRange(1, 1, 1, 3).setValues([['Setting', 'Template Text', 'Available Variables']]).setFontWeight('bold').setBackground('#fce5cd');
    et.getRange(2, 1, 3, 3).setValues([
      ['Subject', 'Discrepancy Notice: Invoice {FileName}', '{FileName}, {CarrierName}'],
      ['Greeting', 'Hello {CarrierName} Team,\n\nThe following items are showing as discrepancies. Please advise:', '{CarrierName}'],
      ['Outro', 'Thank you.', '']
    ]);
    et.setColumnWidth(1, 150); et.setColumnWidth(2, 500); et.setColumnWidth(3, 200);
    created.push('Email Template');
  }

  // 5. Header Config
  if (!ss.getSheetByName('Header Config')) {
    var hc = ss.insertSheet('Header Config');
    hc.getRange(1, 1, 1, 2).setValues([['Target Field', 'Column Name Aliases (Comma Separated)']]).setFontWeight('bold').setBackground('#e6b8af');
    hc.getRange(2, 1, 12, 2).setValues([
      ['Date', 'date, pickup dt, delivery date'],
      ['TU', 'tu, mb number'],
      ['Store', 'store, dest location, dest city, destination name'],
      ['Tour', 'tour, route'],
      ['Miles', 'miles, total miles'],
      ['NY Pay', 'ny pay, ny, new york, borough fee, dhu $ amt, stp $ amt'],
      ['Tolls', 'toll, tol $ amt'],
      ['Total Cost', 'total $ amt, total cost, cost, total'],
      ['Shift', 'shift, shifts'],
      ['Type', 'delivery type, tour type, shift type'],
      ['Haulier Invoice', 'invoice, invoice #, invoice number'],
      ['Haulier Amount', 'amount, actual amount, invoice amount, total cost']
    ]);
    hc.setColumnWidth(1, 150); hc.setColumnWidth(2, 500);
    created.push('Header Config');
  }

  var msg = created.length > 0
    ? 'Created ' + created.length + ' tab(s): ' + created.join(', ') + '.'
    : 'All configuration tabs already present — no changes made.';
  return { success: true, message: msg, created: created };
}

function uploadInvoiceWeb(base64Data, filename, mimeType, carrierKey) {
  try {
    const config = getConfig();
    const folderId = config[carrierKey + '_ROOT_FOLDER'];
    if (!folderId) throw new Error("Carrier root folder not configured.");
    
    const folder = DriveApp.getFolderById(folderId);
    
    // Decode base64 
    const splitBase = base64Data.split(',');
    const data = splitBase.length > 1 ? splitBase[1] : splitBase[0];
    
    const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, filename);
    folder.createFile(blob);
    return { success: true, message: filename + " uploaded successfully to " + carrierKey + " root folder!" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function runProcessorWeb() {
  return runMainProcess(true);
}

function clearTrackerWeb() {
  return clearTrackerData(true);
}