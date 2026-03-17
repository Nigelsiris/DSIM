// Serve the User Management page
function getUserManagementPage(token) {
  const session = getSession(token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  const template = HtmlService.createTemplateFromFile('UserManagement');
  template.token = token;
  return template.evaluate().getContent();
}

// Return all users for the user management page
function getAllUsers(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  let users = [];
  if (userSheet) {
    const lastRow = userSheet.getLastRow();
    if (lastRow > 1) {
      const usersData = userSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      users = usersData.map(row => ({ username: row[0], role: row[2], carrier: row[3] || '' })).filter(u => u.username);
    }
  }
  return users;
}
// --- CONFIGURATION ---
const LOG_SHEET_NAME = 'Form Responses 1';
const STATUS_SHEET_NAME = 'EPJ_Status';
const USER_SHEET_NAME = 'Users';
const MAINT_LOG_SHEET_NAME = 'Maintenance_Log';
const ZONES_SHEET_NAME = 'Zones';
const LOGIN_LOG_SHEET_NAME = 'Login_Log';
const ANNOUNCEMENTS_SHEET_NAME = 'Announcements';
const SETTINGS_SHEET_NAME = 'Site_Settings';

// --- PASTE YOUR WAREHOUSE COORDINATES HERE ---
const WAREHOUSE_LAT = 39.58390517747175;
const WAREHOUSE_LON = -76.02613486224995;
const GEOFENCE_RADIUS_METERS = 500;
// --- END OF CONFIGURATION ---

function doGet(e) {
  // This serves the main entry point, likely the login page.
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Warehouse Sign-In System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- UTILITIES ---
// Targeted cache invalidation functions for better performance
function clearStateCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['activeTrips', 'epjInfoMap', 'activeDrivers', 'zoneOptions', 'epjStatuses', 'activeDriverNames', 'admin_users', 'admin_maintlog']);
}

// Clear only trip-related caches (used after check-in/check-out)
function clearTripCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['activeTrips', 'activeDrivers', 'activeDriverNames']);
}

// Clear only EPJ-related caches (used after EPJ status changes)
function clearEpjCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['epjStatuses', 'epjInfoMap']);
}

// Numeric-aware sort comparator for EPJ strings (e.g. "12" before "100")
function epjSortCompare(a, b) {
  const numA = parseFloat(a);
  const numB = parseFloat(b);
  if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
  if (!isNaN(numA)) return -1;
  if (!isNaN(numB)) return 1;
  return String(a).localeCompare(String(b));
}

// Clear only user-related caches (used after user management changes)
function clearUserCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['admin_users']);
}

// Clear only maintenance-related caches
function clearMaintenanceCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['admin_maintlog', 'epjStatuses']);
}

function updateAllEpjStatuses() {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
  if (!logSheet || !statusSheet) return;

  const logData = logSheet.getDataRange().getValues();
  const lastRow = statusSheet.getLastRow();
  if (lastRow < 2) return;
  
  const epjList = statusSheet.getRange("A2:A" + lastRow).getValues().flat().filter(v => v).map(v => String(v).trim());
  const statusMap = {};
  epjList.forEach(epj => { if (epj) statusMap[epj] = 'Available'; });

  const processedEpjs = new Set();
  for (let i = logData.length - 1; i >= 1; i--) {
    if (processedEpjs.size === epjList.length) break;
    const row = logData[i];
    const epj = String(row[6] || '').trim();
    const status = String(row[9] || '').trim();
    if (epj && statusMap.hasOwnProperty(epj) && !processedEpjs.has(epj)) {
      switch (status) {
        case 'Check-Out': statusMap[epj] = 'Checked Out'; processedEpjs.add(epj); break;
        case 'Check-In': case 'Maintenance End': statusMap[epj] = 'Available'; processedEpjs.add(epj); break;
        case 'Maintenance Start': statusMap[epj] = 'Maintenance'; processedEpjs.add(epj); break;
      }
    }
  }
  
  const newStatuses = epjList.map(epj => [statusMap[epj] || 'Available']);
  if (newStatuses.length > 0) {
    statusSheet.getRange(2, 2, newStatuses.length, 1).setValues(newStatuses);
  }
  clearEpjCache(); // Only need to invalidate EPJ-related caches
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371e3;
  const φ1 = lat1 * Math.PI/180;
  const φ2 = lat2 * Math.PI/180;
  const Δφ = (lat2-lat1) * Math.PI/180;
  const Δλ = (lon2-lon1) * Math.PI/180;
  const a = Math.sin(Δφ/2) * Math.sin(Δφ/2) + Math.cos(φ1) * Math.cos(φ2) * Math.sin(Δλ/2) * Math.sin(Δλ/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}

// --- PAGE GETTERS ---
function getLoginPageHtml() {
  // Don't load activeDrivers here - let the client fetch them asynchronously
  const template = HtmlService.createTemplateFromFile('LoginPage');
  return template.evaluate().getContent();
}

function getUserView(token) {
  const session = getSession(token);
  if (!session) return null;

  if (session.role === 'Admin') {
    return getAdminPageHtml(token);
  } else if (session.role === 'Load Support') {
    return getLoadSupportPageHtml(token);
  } else { // Role is 'Driver'
    const activeTrip = findActiveTrip(session.username);
    if (activeTrip) {
      // Ensure tripInfo has valid epj property
      if (!activeTrip.epj) {
        Logger.log('Warning: activeTrip missing epj property: ' + JSON.stringify(activeTrip));
        activeTrip.epj = 'Unknown';
      }
      
      // Check if this is an Overspill or Pre-Load trip (no real EPJ)
      const isOverspillTrip = activeTrip.epj === 'N/A - Overspill' || 
                              activeTrip.epj === 'N/A - Pre-Load' || 
                              activeTrip.epj === 'N/A - None Required' ||
                              (activeTrip.epj && activeTrip.epj.startsWith('N/A'));
      
      try {
        const template = HtmlService.createTemplateFromFile('CheckInForm');
        template.username = session.username;
        // Pass individual values instead of object for better template compatibility
        template.currentEpj = activeTrip.epj || 'Unknown';
        template.tripId = activeTrip.tripId || '';
        template.token = token;
        // Preload zone options here to avoid async issues
        template.zoneOptions = getZoneOptions();
        template.isOverspill = isOverspillTrip;
        template.checkoutZone = activeTrip.zone || '';
        template.currentTruck = activeTrip.truck || '';
        template.currentTrailer = activeTrip.trailer || '';
        template.currentRoute = activeTrip.route || '';
        return template.evaluate().getContent();
      } catch (e) {
        Logger.log('CheckInForm template error: ' + e.message + ' | activeTrip: ' + JSON.stringify(activeTrip));
        throw e;
      }
    } else {
      const template = HtmlService.createTemplateFromFile('CheckOutForm');
      template.username = session.username;
      template.token = token;
      template.zoneOptions = ''; // Provide empty default for backward compatibility
      template.availableEpjs = '[]'; // Provide empty array for backward compatibility
      template.epjInfoMap = {}; // Provide empty object for backward compatibility
      // Don't load heavy data here - let client fetch it asynchronously
      return template.evaluate().getContent();
    }
  }
}

/**
 * --- FIX ---
 * This function now passes the session token to the HTML template.
 * This allows the client-side script to use the token for subsequent API calls.
 */
function getAdminPageHtml(token) {
    const session = getSession(token);
    if (!session) return null;
    const template = HtmlService.createTemplateFromFile('AdminDashboard');
    template.username = session.username;
    template.token = token; // Add this line to pass the token to the template
    return template.evaluate().getContent();
}

function getLoadSupportPageHtml(token) {
  const session = getSession(token);
  if (!session) return null;
  const template = HtmlService.createTemplateFromFile('LoadSupportDashboard');
  template.username = session.username;
  template.token = token;
  // Don't load heavy data here - let client fetch it asynchronously
  return template.evaluate().getContent();
}

function getProfilePageHtml(token) {
  const session = getSession(token);
  if (!session) return null;
  const template = HtmlService.createTemplateFromFile('ProfilePage');
  template.username = session.username;
  template.token = token; // Recommended: Do the same for other user roles
  return template.evaluate().getContent();
}

function getEquipmentStatusPageHtml(token) {
  const session = getSession(token);
  if (!session) return null;
  const template = HtmlService.createTemplateFromFile('EquipmentStatusPage');
  template.equipmentData = JSON.stringify(getEquipmentStatusViewData());
  template.token = token; // Recommended: Do the same for other user roles
  return template.evaluate().getContent();
}

// --- AUTHENTICATION & SESSIONS ---
function loginAndGetUserView(loginData) {
  const { username, password, latitude, longitude } = loginData;
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  
  // Optimize: Only read the columns we need (username, password hash, role)
  const lastRow = userSheet.getLastRow();
  if (lastRow < 2) return null; // No users
  
  const data = userSheet.getRange(2, 1, lastRow - 1, 3).getValues(); // Only columns A-C
  const passwordHash = sha256(password);
  let user = null;

  for (let i = 0; i < data.length; i++) {
    const cellUsername = String(data[i][0] || '').trim();
    if (cellUsername && cellUsername.toLowerCase() === username.toLowerCase() && data[i][1] === passwordHash) {
      user = { username: data[i][0], role: data[i][2] };
      break;
    }
  }
  
  if (!user) return null;
  
  // Log login asynchronously AFTER authentication succeeds (don't block response)
  // Create trigger to log after response is sent
  try {
    const loginLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOGIN_LOG_SHEET_NAME);
    if (loginLogSheet) {
      let isAtWarehouse = false;
      if (latitude && longitude) {
        const distance = calculateDistance(latitude, longitude, WAREHOUSE_LAT, WAREHOUSE_LON);
        isAtWarehouse = (distance <= GEOFENCE_RADIUS_METERS);
        loginLogSheet.appendRow([new Date(), username, latitude, longitude, isAtWarehouse]);
      } else {
        loginLogSheet.appendRow([new Date(), username, "Not Provided", "Not Provided", false]);
      }
      SpreadsheetApp.flush(); // Ensure write completes
    }
  } catch(e) {
    // Don't let logging errors block login
    Logger.log('Login logging error: ' + e);
  }
  
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(token, JSON.stringify(user), 86400);
  // Don't generate HTML here - let client fetch it separately for faster login
  return { token: token, role: user.role };
}

function getSession(token) {
  if (!token) return null;
  const sessionData = CacheService.getScriptCache().get(token);
  return sessionData ? JSON.parse(sessionData) : null;
}

function logoutUser(data) {
    if (data && data.token) { CacheService.getScriptCache().remove(data.token); }
    // Only clear trip cache on logout - user cache and EPJ info don't change
    clearTripCache();
    return true;
}

// --- DRIVER WORKFLOWS ---
function processCheckOut(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    // Invalidate EPJ status cache before checking availability
    CacheService.getScriptCache().remove('epjStatuses');
    const session = getSession(data.token);
    if (!session) throw new Error("Invalid session.");
    
    const isOverspill = data.isOverspill === true;
    const isPreload = data.isPreload === true;
    const isNaEpj = data.epjNumber === 'N/A - None Required';
    let startingZone = 'Overspill';
    
    // Only validate EPJ if not overspill and not N/A
    if (!isOverspill && !isNaEpj && data.epjNumber) {
      let availableEpjs = getEpjsByStatus('Available');
      availableEpjs = availableEpjs.map(String);
      const requestedEpj = String(data.epjNumber);
      if (!availableEpjs.includes(requestedEpj)) {
        return `Error: EPJ ${data.epjNumber} is no longer available. It may have just been checked out.`;
      }
      const epjInfoMap = getEpjInfoMap();
      startingZone = epjInfoMap[data.epjNumber] ? epjInfoMap[data.epjNumber].location : 'Unknown';
    }
    
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    
    // Determine trip ID prefix and EPJ value
    let tripIdPrefix = "TRIP-";
    let epjValue = data.epjNumber;
    
    if (isOverspill) {
      tripIdPrefix = "OS-";
      epjValue = 'N/A - Overspill';
      startingZone = 'Overspill';
    } else if (isPreload) {
      tripIdPrefix = "PRE-";
      if (isNaEpj || !data.epjNumber) {
        epjValue = 'N/A - Pre-Load';
        startingZone = 'Pre-Load';
      }
    }
    
    const tripId = tripIdPrefix + Utilities.getUuid().substring(0, 8).toUpperCase();
    
    // Log the checkout - Pre-Load trips stay active until admin checks them in from dashboard
    logSheet.appendRow([tripId, new Date(), session.username, data.driverName, data.truckNumber, data.trailerNumber, epjValue, data.route, startingZone, "Check-Out", data.faultReport, "", ""]);
    SpreadsheetApp.flush();
    
    // Fast-path: directly mark the single EPJ as 'Checked Out' instead of
    // recalculating every EPJ status from the full log.  This turns a
    // multi-second scan into a single cell write.
    if (!isOverspill && !isNaEpj && data.epjNumber) {
      const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
      if (statusSheet) {
        const epjCol = statusSheet.getRange('A2:A' + statusSheet.getLastRow()).getValues().flat();
        for (let i = 0; i < epjCol.length; i++) {
          if (String(epjCol[i]).trim() === String(data.epjNumber).trim()) {
            statusSheet.getRange(i + 2, 2).setValue('Checked Out');
            break;
          }
        }
      }
      clearEpjCache();
      clearTripCache();
      return `Successfully checked out EPJ ${data.epjNumber}.`;
    }

    clearTripCache();
    if (isOverspill) {
      return `Successfully checked out for Overspill (no EPJ assigned).`;
    } else if (isPreload && (isNaEpj || !data.epjNumber)) {
      return `Successfully checked out for Pre-Load (no EPJ assigned).`;
    } else {
      return `Successfully checked out${data.epjNumber ? ' EPJ ' + data.epjNumber : ''}.`;
    }
  } finally {
    lock.releaseLock();
  }
}

function processCheckIn(data) {
    const session = getSession(data.token);
    if (!session) throw new Error("Invalid session.");
    const activeTrip = findActiveTrip(session.username); 
    if (!activeTrip) return "Error: No active trip found.";
    // Sanitise check-in zone – never store the old placeholder value
    let checkInZone = String(data.checkInZone || '').trim();
    if (!checkInZone || checkInZone === 'Overspill - No EPJ') {
      checkInZone = activeTrip.zone || 'Unknown';
    }
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    logSheet.appendRow([
        activeTrip.tripId, new Date(), session.username, activeTrip.driver, "", "",
        activeTrip.epj, "", activeTrip.zone, "Check-In", "",
        checkInZone, data.faultReport, data.pluggedIn
    ]);

  // Fast-path: directly set the single EPJ back to 'Available'
  const epjStr = String(activeTrip.epj || '').trim();
  const isRealEpj = epjStr && !epjStr.startsWith('N/A');
  if (isRealEpj) {
    const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
    if (statusSheet) {
      const epjCol = statusSheet.getRange('A2:A' + statusSheet.getLastRow()).getValues().flat();
      for (let i = 0; i < epjCol.length; i++) {
        if (String(epjCol[i]).trim() === epjStr) {
          statusSheet.getRange(i + 2, 2).setValue('Available');
          break;
        }
      }
    }
  }
  clearEpjCache();
  clearTripCache();
  return `Successfully checked in EPJ ${activeTrip.epj}.`;
}

/**
 * Driver EPJ swap - allows a driver to swap their own EPJ mid-trip
 */
function driverSwapEpj(data) {
  const session = getSession(data.token);
  if (!session) throw new Error("Invalid session.");
  
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    // Find driver's active trip
    const activeTrip = findActiveTrip(session.username);
    if (!activeTrip) {
      return "Error: No active trip found.";
    }
    
    // Verify the new EPJ is available
    CacheService.getScriptCache().remove('epjStatuses');
    let availableEpjs = getEpjsByStatus('Available').map(String);
    if (!availableEpjs.includes(String(data.newEpj))) {
      return `Error: EPJ ${data.newEpj} is no longer available.`;
    }
    
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const now = new Date();
    
    // Step 1: Check in the old EPJ
    logSheet.appendRow([
      activeTrip.tripId, now, session.username, activeTrip.driver, "", "",
      activeTrip.epj, "", activeTrip.zone, "Check-In", "", 
      data.currentLocation || activeTrip.zone,
      `Driver swapped EPJ - returning ${activeTrip.epj}`
    ]);
    
    // Step 2: If there's an issue with the old EPJ, log it as a fault report only
    // Note: Only admins can put EPJs into maintenance mode
    if (data.issueReport && data.issueReport.trim()) {
      // Log the issue as a fault report in the main log for admin visibility
      logSheet.appendRow([
        '', now, session.username, activeTrip.driver, '', '',
        activeTrip.epj, '', '', 'Fault Report', data.issueReport, '', ''
      ]);
    }
    
    // Step 3: Check out the new EPJ with a new trip ID
    const newTripId = "SWAP-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    const epjInfoMap = getEpjInfoMap();
    const newEpjZone = epjInfoMap[data.newEpj] ? epjInfoMap[data.newEpj].location : 'Unknown';
    
    logSheet.appendRow([
      newTripId, now, session.username, activeTrip.driver, 
      activeTrip.truck, activeTrip.trailer,
      data.newEpj, activeTrip.route, newEpjZone, "Check-Out", 
      `Driver swap - Replaced ${activeTrip.epj}`, '', ''
    ]);
    
    SpreadsheetApp.flush();

    // Fast-path: directly update both EPJ statuses instead of full recalc
    const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
    if (statusSheet) {
      const epjCol = statusSheet.getRange('A2:A' + statusSheet.getLastRow()).getValues().flat();
      for (let i = 0; i < epjCol.length; i++) {
        const cell = String(epjCol[i]).trim();
        if (cell === String(activeTrip.epj).trim()) {
          statusSheet.getRange(i + 2, 2).setValue('Available');
        } else if (cell === String(data.newEpj).trim()) {
          statusSheet.getRange(i + 2, 2).setValue('Checked Out');
        }
      }
    }
    clearTripCache();
    clearEpjCache();
    
    return `Success! Swapped to EPJ ${data.newEpj}. Your old EPJ ${activeTrip.epj} has been checked in.`;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Allow an Overspill driver to get an EPJ and optionally update their trip info
 */
function overspillGetEpj(data) {
  const session = getSession(data.token);
  if (!session) throw new Error("Invalid session.");
  
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    // Find the driver's active overspill trip
    const activeTrip = findActiveTrip(session.username);
    if (!activeTrip) {
      return "Error: No active trip found for your account.";
    }
    
    // Verify this is an overspill/pre-load trip
    const isOverspillTrip = activeTrip.epj === 'N/A - Overspill' || 
                            activeTrip.epj === 'N/A - Pre-Load' || 
                            activeTrip.epj === 'N/A - None Required' ||
                            (activeTrip.epj && activeTrip.epj.startsWith('N/A'));
    
    if (!isOverspillTrip) {
      return "Error: You already have an EPJ checked out. Use the swap function instead.";
    }
    
    if (!data.epjNumber) {
      return "Error: Please select an EPJ.";
    }
    
    // Verify the EPJ is available
    CacheService.getScriptCache().remove('epjStatuses');
    let availableEpjs = getEpjsByStatus('Available').map(String);
    if (!availableEpjs.includes(String(data.epjNumber))) {
      return `Error: EPJ ${data.epjNumber} is no longer available.`;
    }
    
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const now = new Date();
    
    // Step 1: Check in the overspill trip
    logSheet.appendRow([
      activeTrip.tripId, now, session.username, activeTrip.driver, "", "",
      activeTrip.epj, "", "Overspill", "Check-In", "", 
      "Overspill - Getting EPJ",
      `Overspill trip ended - driver getting EPJ ${data.epjNumber}`
    ]);
    
    // Step 2: Create new checkout with the EPJ
    const newTripId = "TRIP-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    const epjInfoMap = getEpjInfoMap();
    const epjZone = epjInfoMap[data.epjNumber] ? epjInfoMap[data.epjNumber].location : 'Unknown';
    
    // Use updated trip info if provided, otherwise keep original
    const truckNumber = data.truckNumber || activeTrip.truck || '';
    const trailerNumber = data.trailerNumber || activeTrip.trailer || '';
    const route = data.route || activeTrip.route || '';
    
    logSheet.appendRow([
      newTripId, now, session.username, activeTrip.driver, 
      truckNumber, trailerNumber,
      data.epjNumber, route, epjZone, "Check-Out", 
      `Overspill driver picked up EPJ`, '', ''
    ]);
    
    SpreadsheetApp.flush();

    // Fast-path: directly mark the new EPJ as 'Checked Out'
    const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
    if (statusSheet) {
      const epjCol = statusSheet.getRange('A2:A' + statusSheet.getLastRow()).getValues().flat();
      for (let i = 0; i < epjCol.length; i++) {
        if (String(epjCol[i]).trim() === String(data.epjNumber).trim()) {
          statusSheet.getRange(i + 2, 2).setValue('Checked Out');
          break;
        }
      }
    }
    clearTripCache();
    clearEpjCache();
    
    return `Success! You now have EPJ ${data.epjNumber}. Drive safe!`;
  } finally {
    lock.releaseLock();
  }
}

function driverChangePassword(data) {
  const session = getSession(data.token);
  if (!session) throw new Error("Invalid session.");
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  const users = userSheet.getDataRange().getValues();
  const currentPasswordHash = sha256(data.currentPassword);
  for (let i = 1; i < users.length; i++) {
    if (users[i][0] === session.username) {
      if (users[i][1] !== currentPasswordHash) {
        return "Error: Incorrect current password.";
      }
      userSheet.getRange(i + 1, 2).setValue(sha256(data.newPassword));
      return "Password updated successfully!";
    }
  }
  return "Error: Could not find user profile.";
}


// --- ADMIN & LOAD SUPPORT WORKFLOWS ---

/**
 * Get recent driver logins for admin notifications
 * Returns logins from the last N minutes
 */
function getRecentDriverLogins(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  
  const minutesBack = data.minutesBack || 2; // Default to last 2 minutes
  const cutoffTime = new Date(Date.now() - (minutesBack * 60 * 1000));
  
  // Get user info to filter drivers only
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  const userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 3).getValues();
  const userRoles = {};
  userData.forEach(row => { 
    const uname = String(row[0] || '').trim();
    if (uname) userRoles[uname.toLowerCase()] = row[2]; 
  });
  
  const recentLogins = [];
  
  // Check login log for fresh logins
  const loginLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOGIN_LOG_SHEET_NAME);
  if (loginLogSheet && loginLogSheet.getLastRow() >= 2) {
    const lastRow = loginLogSheet.getLastRow();
    const startRow = Math.max(2, lastRow - 49);
    const numRows = lastRow - startRow + 1;
    const logData = loginLogSheet.getRange(startRow, 1, numRows, 2).getValues(); // Timestamp, Username
    
    for (let i = logData.length - 1; i >= 0; i--) {
      const timestamp = new Date(logData[i][0]);
      const username = String(logData[i][1] || '').trim();
      
      if (timestamp >= cutoffTime && username) {
        const role = userRoles[username.toLowerCase()];
        if (role === 'Driver') {
          recentLogins.push({
            username: username,
            timestamp: timestamp.toLocaleTimeString(),
            timestampMs: timestamp.getTime(),
            eventType: 'login',
            isSwap: false
          });
        }
      } else if (timestamp < cutoffTime) {
        break; // Older entries, stop checking
      }
    }
  }
  
  // Check main log for recent EPJ swaps
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (logSheet && logSheet.getLastRow() >= 2) {
    const lastRow = logSheet.getLastRow();
    const startRow = Math.max(2, lastRow - 49);
    const numRows = lastRow - startRow + 1;
    const swapData = logSheet.getRange(startRow, 1, numRows, 11).getValues(); // TripID, Timestamp, Username, Driver, ... EPJ, ... Status, Notes
    
    for (let i = swapData.length - 1; i >= 0; i--) {
      const tripId = String(swapData[i][0] || '');
      const timestamp = new Date(swapData[i][1]);
      const username = String(swapData[i][2] || '').trim();
      const driverName = String(swapData[i][3] || '').trim();
      const epj = String(swapData[i][6] || '').trim();
      const status = String(swapData[i][9] || '').trim();
      const notes = String(swapData[i][10] || '');
      
      if (timestamp >= cutoffTime && tripId.startsWith('SWAP-') && status === 'Check-Out') {
        recentLogins.push({
          username: driverName || username,
          timestamp: timestamp.toLocaleTimeString(),
          timestampMs: timestamp.getTime(),
          eventType: 'swap',
          isSwap: true,
          newEpj: epj,
          notes: notes
        });
      } else if (timestamp < cutoffTime) {
        break;
      }
    }
  }
  
  // Sort by timestamp descending and return
  return recentLogins.sort((a, b) => b.timestampMs - a.timestampMs);
}

function getDashboardData(data) {
  Logger.log('getDashboardData called with: ' + JSON.stringify(data));
  const session = getSession(data.token);
  Logger.log('Session: ' + JSON.stringify(session));
  if (!session || session.role !== 'Admin') {
    Logger.log('Permission denied or session missing.');
    throw new Error("Permission denied.");
  }
  // Use cache for users and maintenance log
  const cache = CacheService.getScriptCache();
  let users = [];
  let userMap = {};
  let usersCached = cache.get('admin_users');
  if (usersCached) {
    users = JSON.parse(usersCached);
    users.forEach(u => { userMap[u.username] = u; });
  } else {
    const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
    if (userSheet) {
      const lastRow = userSheet.getLastRow();
      if (lastRow > 1) {
        const usersData = userSheet.getRange(2, 1, lastRow - 1, 4).getValues();
        users = usersData.map(row => {
          const userObject = { username: row[0], role: row[2], carrier: row[3] || '' };
          userMap[row[0]] = userObject;
          return userObject;
        }).filter(u => u.username);
        cache.put('admin_users', JSON.stringify(users), 600); // cache for 10 min
      }
    }
  }
  let maintenanceLog = [];
  let maintCached = cache.get('admin_maintlog');
  if (maintCached) {
    maintenanceLog = JSON.parse(maintCached);
  } else {
    const maintSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAINT_LOG_SHEET_NAME);
    if (maintSheet && maintSheet.getLastRow() > 1) {
      const lastMaintRow = maintSheet.getLastRow();
      const startRow = Math.max(2, lastMaintRow - 19);
      const numRows = lastMaintRow - startRow + 1;
      maintenanceLog = maintSheet.getRange(startRow, 1, numRows, 5).getValues().reverse();
      cache.put('admin_maintlog', JSON.stringify(maintenanceLog), 300); // cache for 5 min
    }
  }
  const activeCheckouts = getActiveCheckouts().map(checkout => {
    const driverInfo = userMap[checkout.driverUsername];
    checkout.carrier = driverInfo ? driverInfo.carrier : 'N/A';
    return checkout;
  });
  const result = {
    epjStatuses: getEpjsByStatus(null, true),
    users: users,
    maintenanceLog: maintenanceLog,
    activeCheckouts: activeCheckouts
  };
  Logger.log('Returning dashboard data: ' + JSON.stringify(result));
  return result;
}

function adminForceCheckIn(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  const logData = logSheet.getDataRange().getValues();
  let originalTrip = null;
  for(let i = logData.length - 1; i >= 1; i--) {
    if (logData[i][0] === data.tripId) {
      originalTrip = logData[i];
      break;
    }
  }
  if (originalTrip) {
    const originalDriver = originalTrip[3];
    const originalZone = originalTrip[8];
    const checkInLocation = data.location || originalZone || 'Unknown';
    logSheet.appendRow([
        data.tripId, new Date(), session.username, originalDriver, "", "",
        data.epj, "", originalZone, "Check-In", "", checkInLocation,
        `Forced check-in by admin ${session.username}`
    ]);
  updateAllEpjStatuses();
  CacheService.getScriptCache().remove('epjStatuses');
  clearTripCache();
  return `Successfully checked in EPJ ${data.epj} at ${checkInLocation}.`;
  }
  return `Error: Could not find original trip ID ${data.tripId}.`;
}

/**
 * Admin checkout - allows admin to check out an EPJ on behalf of a driver
 */
function adminCheckoutForDriver(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    CacheService.getScriptCache().remove('epjStatuses');
    
    const isOverspill = data.isOverspill === true;
    let startingZone = 'Overspill';
    
    // Only validate EPJ if not overspill
    if (!isOverspill && data.epjNumber) {
      let availableEpjs = getEpjsByStatus('Available').map(String);
      if (!availableEpjs.includes(String(data.epjNumber))) {
        return `Error: EPJ ${data.epjNumber} is no longer available.`;
      }
      const epjInfoMap = getEpjInfoMap();
      startingZone = epjInfoMap[data.epjNumber] ? epjInfoMap[data.epjNumber].location : 'Unknown';
    }
    
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const tripId = "ADMIN-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    const epjValue = isOverspill ? 'N/A - Overspill' : data.epjNumber;
    
    // Log with driverUsername for tracking, but use data.driverName for display
    logSheet.appendRow([
      tripId, 
      new Date(), 
      data.driverUsername || data.driverName, // Username for system tracking
      data.driverName,  // Display name
      data.truckNumber || '', 
      data.trailerNumber || '', 
      epjValue, 
      data.route || 'Admin Checkout', 
      startingZone, 
      "Check-Out", 
      `Admin checkout by ${session.username}`, 
      "", 
      ""
    ]);
    
    SpreadsheetApp.flush();
    
    if (!isOverspill && data.epjNumber) {
      updateAllEpjStatuses();
    }
    // Clear trip and EPJ caches since both changed
    clearTripCache();
    if (data.epjNumber) clearEpjCache();
    
    if (isOverspill) {
      return `Successfully checked out ${data.driverName} for Overspill (no EPJ).`;
    }
    return `Successfully checked out EPJ ${data.epjNumber} for ${data.driverName}.`;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Swap a driver's EPJ mid-trip
 * - Checks in the old EPJ (optionally putting it in maintenance)
 * - Checks out a new EPJ for the same trip
 */
function adminSwapEpj(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const logData = logSheet.getDataRange().getValues();
    
    // Find the original trip to get details
    let originalTrip = null;
    for (let i = logData.length - 1; i >= 1; i--) {
      if (logData[i][0] === data.tripId) {
        originalTrip = logData[i];
        break;
      }
    }
    
    if (!originalTrip) {
      return `Error: Could not find trip ${data.tripId}.`;
    }
    
    // Verify the new EPJ is available
    CacheService.getScriptCache().remove('epjStatuses');
    let availableEpjs = getEpjsByStatus('Available').map(String);
    if (!availableEpjs.includes(String(data.newEpj))) {
      return `Error: EPJ ${data.newEpj} is no longer available.`;
    }
    
    const originalDriver = originalTrip[3];
    const originalZone = originalTrip[8];
    const truck = originalTrip[4];
    const trailer = originalTrip[5];
    const route = originalTrip[7];
    const now = new Date();
    
    // Step 1: Check in the old EPJ
    logSheet.appendRow([
      data.tripId, now, session.username, originalDriver, "", "",
      data.oldEpj, "", originalZone, "Check-In", "", "EPJ Swap",
      `EPJ swapped by admin ${session.username} - Old EPJ returned`
    ]);
    
    // Step 2: If requested, put old EPJ in maintenance
    if (data.putInMaintenance) {
      const maintSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAINT_LOG_SHEET_NAME);
      const reason = data.maintenanceReason || 'Issue reported during swap';
      maintSheet.appendRow([now, data.oldEpj, 'Maintenance Start', reason, '']);
      logSheet.appendRow([
        '', now, session.username, 'ADMIN', '', '',
        data.oldEpj, '', '', 'Maintenance Start', reason, '', ''
      ]);
    }
    
    // Step 3: Check out the new EPJ with a new trip ID
    const newTripId = "SWAP-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    const epjInfoMap = getEpjInfoMap();
    const newEpjZone = epjInfoMap[data.newEpj] ? epjInfoMap[data.newEpj].location : 'Unknown';
    
    logSheet.appendRow([
      newTripId, now, data.driverUsername, data.driverName, truck, trailer,
      data.newEpj, route, newEpjZone, "Check-Out", 
      `EPJ swap - Replaced ${data.oldEpj}`, '', ''
    ]);
    
    SpreadsheetApp.flush();
    updateAllEpjStatuses();
    // Clear trip and EPJ caches since both changed
    clearTripCache();
    clearEpjCache();
    
    return `Success! Swapped ${data.oldEpj} → ${data.newEpj} for ${data.driverName}`;
  } finally {
    lock.releaseLock();
  }
}

function updateEpjLocation(data) {
    const session = getSession(data.token);
    if (!session || (session.role !== 'Admin' && session.role !== 'Load Support')) { throw new Error("Permission denied."); }
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    logSheet.appendRow([ '', new Date(), session.username, 'LOAD SUPPORT', '', '', data.epj, '', data.newLocation, 'Location Update', `Updated by ${session.role}`, data.newLocation, '' ]);
  updateAllEpjStatuses();
  CacheService.getScriptCache().remove('epjStatuses');
  return `Location for EPJ ${data.epj} updated to ${data.newLocation}.`;
}

/**
 * Single EPJ maintenance status update (kept for backward compatibility)
 */
function adminSetMaintenanceStatus(data) {
    const session = getSession(data.token);
    if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const maintSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAINT_LOG_SHEET_NAME);
    
    if (data.status === 'Maintenance Start') {
      const reason = data.reason || 'Reason pending';
      maintSheet.appendRow([new Date(), data.epj, 'Maintenance Start', reason, '']);
      logSheet.appendRow(['', new Date(), session.username, 'ADMIN', '', '', data.epj, '', '', 'Maintenance Start', reason, '', '']);
    } else {
      const resolution = data.resolution || 'Returned to service';
      maintSheet.appendRow([new Date(), data.epj, 'Maintenance End', '', resolution]);
      logSheet.appendRow(['', new Date(), session.username, 'ADMIN', '', '', data.epj, '', '', 'Maintenance End', resolution, '', '']);
    }
    
    updateAllEpjStatuses();
    CacheService.getScriptCache().remove('epjStatuses');
    return `EPJ ${data.epj} status updated.`;
}

/**
 * BATCH maintenance status update - much faster for multiple EPJs
 */
function adminBatchMaintenanceStatus(data) {
    const session = getSession(data.token);
    if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
    
    const epjs = data.epjs; // Array of EPJ names
    if (!epjs || epjs.length === 0) throw new Error("No EPJs provided.");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    const maintSheet = ss.getSheetByName(MAINT_LOG_SHEET_NAME);
    const now = new Date();
    
    // Prepare batch data
    const maintRows = [];
    const logRows = [];
    
    if (data.status === 'Maintenance Start') {
        const reason = data.reason || 'Reason pending';
        epjs.forEach(function(epj) {
            maintRows.push([now, epj, 'Maintenance Start', reason, '']);
            logRows.push(['', now, session.username, 'ADMIN', '', '', epj, '', '', 'Maintenance Start', reason, '', '']);
        });
    } else {
        const resolution = data.resolution || 'Returned to service';
        epjs.forEach(function(epj) {
            maintRows.push([now, epj, 'Maintenance End', '', resolution]);
            logRows.push(['', now, session.username, 'ADMIN', '', '', epj, '', '', 'Maintenance End', resolution, '', '']);
        });
    }
    
    // Batch write to maintenance log
    if (maintRows.length > 0) {
        const maintStartRow = maintSheet.getLastRow() + 1;
        maintSheet.getRange(maintStartRow, 1, maintRows.length, 5).setValues(maintRows);
    }
    
    // Batch write to main log
    if (logRows.length > 0) {
        const logStartRow = logSheet.getLastRow() + 1;
        logSheet.getRange(logStartRow, 1, logRows.length, 13).setValues(logRows);
    }
    
    updateAllEpjStatuses();
    CacheService.getScriptCache().remove('epjStatuses');
    
    return { success: true, count: epjs.length, message: `${epjs.length} EPJ(s) updated successfully.` };
}

/**
 * Get maintenance history for the admin dashboard
 */
function getMaintenanceHistory(data) {
    const session = getSession(data.token);
    if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
    
    const maintSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAINT_LOG_SHEET_NAME);
    if (!maintSheet || maintSheet.getLastRow() < 2) {
        return [];
    }
    
    const rows = maintSheet.getRange(2, 1, maintSheet.getLastRow() - 1, 5).getValues();
    
    // Return most recent first
    return rows.map(function(row) {
        return {
            timestamp: row[0],
            epj: row[1],
            event: row[2],
            reason: row[3] || '',
            resolution: row[4] || ''
        };
    }).reverse();
}

function adminMassCreateUsers(data) {
    const session = getSession(data.token);
    if (!session || session.role !== 'Admin') { throw new Error("Permission denied."); }
    const csvData = data.csvData;
    const lines = csvData.split('\n');
    const newUsers = [];
    const errors = [];
    const validRoles = ['Admin', 'Driver', 'Load Support'];
    lines.forEach((line, index) => {
        const trimmedLine = line.trim();
        if (trimmedLine === '') return;
        const fields = trimmedLine.split(',').map(field => field.trim());
        if (fields.length !== 4) {
            errors.push(`Line ${index + 1}: Incorrect number of fields. Expected 4, got ${fields.length}.`);
            return;
        }
        const [username, password, role, carrier] = fields;
        if (!username || !password || !role) {
            errors.push(`Line ${index + 1}: Username, password, and role are required.`);
            return;
        }
        if (validRoles.indexOf(role) === -1) {
            errors.push(`Line ${index + 1}: Invalid role "${role}". Must be Admin, Driver, or Load Support.`);
            return;
        }
        newUsers.push([username, sha256(password), role, carrier]);
    });
    if (newUsers.length > 0) {
        const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
        const startRow = userSheet.getLastRow() + 1;
        userSheet.getRange(startRow, 1, newUsers.length, 4).setValues(newUsers);
    }
    let message = `Batch process complete. Successfully created ${newUsers.length} users.`;
    if (errors.length > 0) {
        message += ` Skipped ${errors.length} rows due to errors. First error: ${errors[0]}`;
    }
    return message;
}

function adminAddUser(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  userSheet.appendRow([data.username, sha256(data.password), data.role, data.carrier]);
  CacheService.getScriptCache().remove('admin_users'); // Invalidate user cache
  return `User "${data.username}" created successfully.`;
}

function adminEditUser(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  const usernames = userSheet.getRange("A:A").getValues();
  
  // If changing username, check if new username already exists
  if (data.newUsername && data.newUsername.toLowerCase() !== data.username.toLowerCase()) {
    for (let i = 1; i < usernames.length; i++) {
      const cellUsername = String(usernames[i][0] || '').trim();
      if (cellUsername && cellUsername.toLowerCase() === data.newUsername.toLowerCase()) {
        return `Error: Username "${data.newUsername}" already exists.`;
      }
    }
  }
  
  for (let i = 1; i < usernames.length; i++) {
    const cellUsername = String(usernames[i][0] || '').trim();
    if (cellUsername && cellUsername.toLowerCase() === data.username.toLowerCase()) {
      // Update username if provided and different
      if (data.newUsername && data.newUsername !== data.username) {
        userSheet.getRange(i + 1, 1).setValue(data.newUsername);
      }
      userSheet.getRange(i + 1, 3).setValue(data.role);
      userSheet.getRange(i + 1, 4).setValue(data.carrier);
      CacheService.getScriptCache().remove('admin_users');
      
      const displayName = data.newUsername || data.username;
      return `User "${displayName}" updated successfully.`;
    }
  }
  return `Error: User "${data.username}" not found.`;
}

// --- DATA FETCH FUNCTIONS FOR ASYNC LOADING ---
function getCheckoutFormData(data) {
  const session = getSession(data.token);
  if (!session) throw new Error("Invalid session.");
  
  // Get all EPJ statuses with store-only flags
  const allStatuses = getEpjsByStatus(null, true);
  const availableEpjs = allStatuses
    .filter(item => item.status === 'Available')
    .map(item => ({
      epj: item.epj,
      storeOnly: item.storeOnly || false
    }))
    .sort((a, b) => epjSortCompare(a.epj, b.epj));
  
  // Include site settings for default overspill/preload toggles
  const siteSettings = getSiteSettings();
  
  return {
    availableEpjs: availableEpjs,
    zoneOptions: getZoneOptions(),
    epjInfoMap: getEpjInfoMap(),
    siteSettings: {
      defaultOverspill: siteSettings.defaultOverspill,
      defaultPreload: siteSettings.defaultPreload,
      overspillStartHour: siteSettings.overspillStartHour,
      overspillEndHour: siteSettings.overspillEndHour,
      preloadStartHour: siteSettings.preloadStartHour,
      preloadEndHour: siteSettings.preloadEndHour,
      morningModeStartHour: siteSettings.morningModeStartHour,
      morningModeEndHour: siteSettings.morningModeEndHour,
      driverCanReportFaults: siteSettings.driverCanReportFaults,
      siteMessage: siteSettings.siteMessage
    }
  };
}

function getCheckinFormData(data) {
  try {
    const session = getSession(data.token);
    if (!session) throw new Error("Invalid session.");
    
    // Get zone options first (this is fast due to caching)
    const zoneOptions = getZoneOptions();
    
    // Return early with just zones if no EPJ to check
    if (!data.epj || data.epj.startsWith('N/A')) {
      return {
        zoneOptions: zoneOptions,
        currentEpjIsStoreOnly: false
      };
    }
    
    // Check if the current EPJ is store-only - use cached data only to avoid slowness
    let currentEpjIsStoreOnly = false;
    try {
      const cache = CacheService.getScriptCache();
      const cachedStatuses = cache.get('epjStatuses');
      
      if (cachedStatuses) {
        const allStatuses = JSON.parse(cachedStatuses);
        const epjData = allStatuses.find(item => item.epj === data.epj);
        if (epjData) {
          currentEpjIsStoreOnly = epjData.storeOnly || false;
        }
      }
      // Skip direct sheet lookup - it's too slow and store-only is not critical
    } catch (e) {
      Logger.log('Error checking EPJ store-only status: ' + e.message);
    }
    
    return {
      zoneOptions: zoneOptions,
      currentEpjIsStoreOnly: currentEpjIsStoreOnly
    };
  } catch (e) {
    Logger.log('Error in getCheckinFormData: ' + e.message);
    throw e;
  }
}

function getLoadSupportData(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Load Support') throw new Error("Permission denied.");
  return {
    epjStatuses: getEpjsByStatus(null, true),
    epjInfoMap: getEpjInfoMap(),
    zoneOptions: getZoneOptions()
  };
}

function adminResetPassword(data) {
    const session = getSession(data.token);
    if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
    const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
    const usernames = userSheet.getRange("A:A").getValues();
    for (let i = 1; i < usernames.length; i++) {
        const cellUsername = String(usernames[i][0] || '').trim();
        if (cellUsername && cellUsername.toLowerCase() === data.username.toLowerCase()) {
            userSheet.getRange(i + 1, 2).setValue(sha256(data.newPassword));
            return `Password reset for user "${data.username}".`;
        }
    }
    return `Error: User "${data.username}" not found.`;
}

function adminDeleteUser(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  const usernames = userSheet.getRange("A:A").getValues();
  for (let i = usernames.length - 1; i > 0; i--) { 
    const cellUsername = String(usernames[i][0] || '').trim();
    if (cellUsername && cellUsername.toLowerCase() === data.username.toLowerCase()) {
      userSheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove('admin_users');
      return `User "${data.username}" has been deleted.`;
    }
  }
  return `Error: User "${data.username}" not found.`;
}

// Bulk edit users - change role for multiple users at once
function adminBulkEditUsers(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  
  const usernames = data.usernames || [];
  const newRole = data.role;
  
  if (!usernames.length) return 'No users selected';
  if (!newRole) return 'No role specified';
  
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  const allUsernames = userSheet.getRange("A:A").getValues();
  let updatedCount = 0;
  
  for (let i = 1; i < allUsernames.length; i++) {
    if (allUsernames[i][0] && usernames.includes(allUsernames[i][0])) {
      userSheet.getRange(i + 1, 3).setValue(newRole);
      updatedCount++;
    }
  }
  
  CacheService.getScriptCache().remove('admin_users');
  return `Successfully updated ${updatedCount} user(s) to role "${newRole}".`;
}

// Bulk delete users
function adminBulkDeleteUsers(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  
  const usernames = data.usernames || [];
  if (!usernames.length) return 'No users selected';
  
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  const allUsernames = userSheet.getRange("A:A").getValues();
  let deletedCount = 0;
  
  // Delete from bottom to top to avoid row index issues
  for (let i = allUsernames.length - 1; i > 0; i--) {
    if (allUsernames[i][0] && usernames.includes(allUsernames[i][0])) {
      userSheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  
  CacheService.getScriptCache().remove('admin_users');
  return `Successfully deleted ${deletedCount} user(s).`;
}

// --- HELPERS & DATA GETTERS ---
function getEpjInfoMap() {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('epjInfoMap');
    if (cached != null) { return JSON.parse(cached); }
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    const logData = logSheet.getDataRange().getValues();
    const epjStatusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
    const epjs = epjStatusSheet.getRange("A2:A").getValues().flat().filter(v => v).map(v => String(v).trim());
    const infoMap = {};
    const foundLocations = new Set();
    const foundFaults = new Set();
    for (const epj of epjs) { infoMap[epj] = { location: "N/A", fault: "No issues reported" }; }
    for (let i = logData.length - 1; i >= 1; i--) {
        if (foundLocations.size === epjs.length && foundFaults.size === epjs.length) { break; }
        const row = logData[i];
        const epj = String(row[6] || '').trim();
        if (!epj || !infoMap[epj] || (foundLocations.has(epj) && foundFaults.has(epj))) { continue; }
        if (!foundLocations.has(epj)) {
            const eventType = row[9];
            const checkInLocation = row[11];
            const checkOutZone = row[8];
            let currentLocation = null;
            if (checkInLocation) { currentLocation = checkInLocation; }
            else if (eventType === 'Location Update' && checkOutZone) { currentLocation = checkOutZone; }
            else if (checkOutZone) { currentLocation = checkOutZone; }
            if (currentLocation) {
                infoMap[epj].location = currentLocation;
                foundLocations.add(epj);
            }
        }
        if (!foundFaults.has(epj)) {
            let currentFault = row[12] || row[10];
            if (currentFault) {
                infoMap[epj].fault = currentFault;
                foundFaults.add(epj);
            }
        }
    }
    cache.put('epjInfoMap', JSON.stringify(infoMap), 21600);
    return infoMap;
}

function getActiveCheckouts() {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('activeTrips');
    let result;
    
    if (cached != null) {
        result = JSON.parse(cached);
    } else {
        const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
        if (!logSheet) { return []; }
        const lastRow = logSheet.getLastRow();
        if (lastRow < 2) return [];
        const data = logSheet.getRange(2, 1, lastRow - 1, 15).getValues(); // Extended to column O (15)
        const activeTrips = {};
        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          const tripId = row[0];
          if (!tripId) continue;
          const status = String(row[9] || '').trim();
          if (status === 'Check-Out') {
            const checkoutTime = new Date(row[1]);
            const notes = String(row[10] || ''); // Notes field contains swap info
            
            // Detect if this is a swap checkout - check tripId prefix OR notes content
            const tripIdStr = String(tripId);
            const isSwap = tripIdStr.startsWith('SWAP-') || notes.toLowerCase().includes('epj swap') || notes.toLowerCase().includes('replaced');
            let swappedFrom = null;
            if (isSwap && notes) {
              // Extract old EPJ from notes like "EPJ swap - Replaced EPJ-001" or "Replaced EPJ001"
              const swapMatch = notes.match(/Replaced\s+([^\s,]+)/i);
              if (swapMatch) swappedFrom = swapMatch[1];
            }
            
            // Column O (index 14) contains the worked status
            const workedStatus = String(row[14] || '').toUpperCase();
            const isWorked = workedStatus === 'YES' || workedStatus === 'TRUE';
            
            activeTrips[tripId] = {
              tripId: tripId, timestamp: checkoutTime.toLocaleString(), timestampMs: checkoutTime.getTime(),
              driverUsername: String(row[2] || ''),
              driver: String(row[3] || ''), truck: String(row[4] || ''), trailer: String(row[5] || ''), 
              epj: String(row[6] || ''), route: String(row[7] || ''), zone: String(row[8] || ''),
              isSwap: isSwap,
              swappedFrom: swappedFrom,
              worked: isWorked
            };
          } else if (status === 'Check-In') {
            if(activeTrips[tripId]) { delete activeTrips[tripId]; }
          }
        }
        result = Object.values(activeTrips);
        cache.put('activeTrips', JSON.stringify(result), 30);
    }
    
    // Always sort by timestamp descending (newest first) - ensures consistency
    return result.sort((a, b) => (b.timestampMs || 0) - (a.timestampMs || 0));
}

/**
 * Get active checkouts for admin dashboard - lightweight call for frequent polling
 */
function getActiveCheckoutsForAdmin(data) {
    const session = getSession(data.token);
    if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
    
    // Get user map for carrier info
    const cache = CacheService.getScriptCache();
    let userMap = {};
    let usersCached = cache.get('admin_users');
    if (usersCached) {
        JSON.parse(usersCached).forEach(u => { userMap[u.username] = u; });
    } else {
        const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
        if (userSheet) {
            const lastRow = userSheet.getLastRow();
            if (lastRow > 1) {
                const usersData = userSheet.getRange(2, 1, lastRow - 1, 4).getValues();
                usersData.forEach(row => {
                    if (row[0]) userMap[row[0]] = { username: row[0], role: row[2], carrier: row[3] || '' };
                });
            }
        }
    }
    
    // Force fresh data by clearing active trips cache
    if (data.forceRefresh) {
        cache.remove('activeTrips');
    }
    
    // Get EPJ store-only status map
    const epjStatuses = getEpjsByStatus(null, true);
    const storeOnlyMap = {};
    epjStatuses.forEach(item => { storeOnlyMap[item.epj] = item.storeOnly || false; });
    
    // Auto-expire Overspill and Pre-Load trips older than 1 hour
    autoExpireNonEpjTrips();
    
    const activeCheckouts = getActiveCheckouts().map(checkout => {
        const driverInfo = userMap[checkout.driverUsername];
        checkout.carrier = driverInfo ? driverInfo.carrier : 'N/A';
        checkout.storeOnly = storeOnlyMap[checkout.epj] || false;
        return checkout;
    });
    
    return activeCheckouts;
}

/**
 * Auto-expire Overspill and Pre-Load trips after 1 hour
 * These are morning drivers who don't take an EPJ
 */
function autoExpireNonEpjTrips() {
    const settings = getSiteSettings();
    const expireMinutes = settings.autoExpireMinutes || 60;
    if (expireMinutes <= 0) return; // 0 = disabled
    const expireMs = expireMinutes * 60 * 1000;
    const now = Date.now();
    
    const activeCheckouts = getActiveCheckouts();
    const tripsToExpire = activeCheckouts.filter(trip => {
        // Only auto-expire Overspill and Pre-Load trips
        const isOverspill = trip.epj === 'N/A - Overspill';
        const isPreload = trip.epj === 'N/A - Pre-Load';
        if (!isOverspill && !isPreload) return false;
        
        // Check if older than configured expire time
        const tripAge = now - (trip.timestampMs || 0);
        return tripAge > expireMs;
    });
    
    if (tripsToExpire.length === 0) return;
    
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) return;
    
    tripsToExpire.forEach(trip => {
        const tripType = trip.epj === 'N/A - Overspill' ? 'Overspill' : 'Pre-Load';
        logSheet.appendRow([
            trip.tripId,
            new Date(),
            'SYSTEM',
            trip.driver,
            '',
            '',
            trip.epj,
            '',
            trip.zone || '',
            'Check-In',
            '',
            'Auto-Expired',
            `${tripType} trip auto-expired after 1 hour`
        ]);
    });
    
    // Clear cache to reflect changes
    clearTripCache();
}

function findActiveTrip(username) {
  const allActiveTrips = getActiveCheckouts();
  return allActiveTrips.find(trip => trip.driverUsername === username) || null;
}

function getActiveDriverNames() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('activeDriverNames');
  if (cached != null) { return JSON.parse(cached); }
  
  const checkouts = getActiveCheckouts();
  const driverMap = new Map();
  checkouts.forEach(trip => {
    driverMap.set(trip.driverUsername, trip.driver);
  });
  const result = Array.from(driverMap, ([username, displayName]) => ({ username, displayName }));
  cache.put('activeDriverNames', JSON.stringify(result), 300); // cache for 5 min
  return result;
}

function getZoneOptions() {
    try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get('zoneOptions');
        if (cached != null && cached !== '') { return cached; }
        
        const zoneSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ZONES_SHEET_NAME);
        if (!zoneSheet) {
            Logger.log('Zones sheet not found: ' + ZONES_SHEET_NAME);
            return '<option value="Unknown">Unknown Location</option>';
        }
        
        const lastRow = zoneSheet.getLastRow();
        if (lastRow < 1) {
            Logger.log('Zones sheet is empty');
            return '<option value="Unknown">Unknown Location</option>';
        }
        
        const zones = zoneSheet.getRange(1, 1, lastRow, 1).getValues().flat().filter(String);
        if (zones.length === 0) {
            Logger.log('No zones found in sheet');
            return '<option value="Unknown">Unknown Location</option>';
        }
        
        let options = '';
        zones.forEach(zone => { options += `<option value="${zone}">${zone}</option>`; });
        cache.put('zoneOptions', options, 21600);
        return options;
    } catch (e) {
        Logger.log('Error in getZoneOptions: ' + e.message);
        return '<option value="Unknown">Unknown Location</option>';
    }
}

/**
 * Get list of zones as an array (for dropdowns)
 */
function getZonesList(data) {
    const session = getSession(data.token);
    if (!session) throw new Error('Invalid session.');
    
    const zoneSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ZONES_SHEET_NAME);
    if (!zoneSheet) return [];
    
    const zones = zoneSheet.getRange("A1:A").getValues().flat().filter(String);
    return zones;
}

function getEpjsByStatus(status, all = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'epjStatuses';
  const cached = cache.get(cacheKey);
  let allStatuses;
  if (cached != null) {
    allStatuses = JSON.parse(cached);
  } else {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];
    // Updated to include column C for store-only flag
    const values = sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
    allStatuses = values.map(row => ({
      epj: String(row[0] || '').trim(), 
      status: String(row[1] || '').trim(),
      storeOnly: row[2] === true || row[2] === 'TRUE' || row[2] === 'Yes'
    })).filter(item => item.epj);
    cache.put(cacheKey, JSON.stringify(allStatuses), 21600);
  }
  if (all) {
    return allStatuses.sort((a, b) => epjSortCompare(a.epj, b.epj));
  }
  return allStatuses.filter(item => item.status === status).map(item => item.epj).sort(epjSortCompare);
}

// Get all EPJs with full details including store-only flag
function getAllEpjsWithDetails(data) {
  const session = getSession(data.token);
  if (!session || (session.role !== 'Admin' && session.role !== 'Load Support')) {
    throw new Error("Permission denied.");
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const values = sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
  const epjInfoMap = getEpjInfoMap();
  
  return values.map(row => {
    const epjStr = String(row[0] || '').trim();
    return {
      epj: epjStr,
      status: String(row[1] || '').trim(),
      storeOnly: row[2] === true || row[2] === 'TRUE' || row[2] === 'Yes',
      location: epjInfoMap[epjStr] ? epjInfoMap[epjStr].location : 'N/A',
      fault: epjInfoMap[epjStr] ? epjInfoMap[epjStr].fault : 'No issues reported'
    };
  }).filter(item => item.epj).sort((a, b) => epjSortCompare(a.epj, b.epj));
}

// Quick update EPJ status (for right-click menu)
function adminQuickUpdateEpjStatus(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  
  const epj = data.epj;
  const newStatus = data.status;
  
  const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
  const epjList = statusSheet.getRange("A2:A" + statusSheet.getLastRow()).getValues().flat();
  
  for (let i = 0; i < epjList.length; i++) {
    if (epjList[i] === epj) {
      statusSheet.getRange(i + 2, 2).setValue(newStatus);
      
      // Log the status change with proper event type that updateAllEpjStatuses() recognizes
      const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
      
      // Map status to event type
      let eventType = 'Admin Status Override';
      if (newStatus === 'Checked Out') {
        eventType = 'Check-Out';
      } else if (newStatus === 'Available') {
        eventType = 'Check-In';
      } else if (newStatus === 'Maintenance') {
        eventType = 'Maintenance Start';
      }
      
      // Generate a trip ID for Check-Out events so they can be force checked in later
      const tripId = (eventType === 'Check-Out') ? "ADMIN-" + Utilities.getUuid().substring(0, 8).toUpperCase() : '';
      
      logSheet.appendRow([
        tripId, new Date(), session.username, 'ADMIN OVERRIDE', '', '', 
        epj, '', '', eventType, 
        `Admin set status to ${newStatus}`, '', ''
      ]);
      
      // Clear EPJ and trip caches since status and possibly trip changed
      clearEpjCache();
      if (eventType === 'Check-Out' || eventType === 'Check-In') clearTripCache();
      return `EPJ ${epj} status updated to ${newStatus}`;
    }
  }
  
  return `Error: EPJ ${epj} not found.`;
}

// Toggle store-only flag for an EPJ
function adminToggleStoreOnly(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  
  const epj = data.epj;
  const storeOnly = data.storeOnly;
  
  const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
  const epjList = statusSheet.getRange("A2:A" + statusSheet.getLastRow()).getValues().flat();
  
  for (let i = 0; i < epjList.length; i++) {
    if (epjList[i] === epj) {
      statusSheet.getRange(i + 2, 3).setValue(storeOnly);
      CacheService.getScriptCache().remove('epjStatuses');
      return `EPJ ${epj} ${storeOnly ? 'marked for' : 'removed from'} store delivery only`;
    }
  }
  
  return `Error: EPJ ${epj} not found.`;
}

function getEquipmentStatusViewData() {
  const statuses = getEpjsByStatus(null, true);
  const infoMap = getEpjInfoMap();
  return statuses.map(item => {
    const info = infoMap[item.epj] || { location: 'N/A' };
    return { epj: item.epj, status: item.status, location: info.location };
  });
}

// Add a new EPJ to the system
function adminAddEpj(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const epjNumber = (data.epjNumber || '').toString().trim().toUpperCase();
  if (!epjNumber) throw new Error('EPJ number is required.');
  
  const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
  if (!statusSheet) throw new Error('EPJ Status sheet not found.');
  
  // Check if EPJ already exists
  const lastRow = statusSheet.getLastRow();
  if (lastRow >= 2) {
    const existingEpjs = statusSheet.getRange('A2:A' + lastRow).getValues().flat();
    if (existingEpjs.some(e => e.toString().toUpperCase() === epjNumber)) {
      throw new Error('EPJ ' + epjNumber + ' already exists.');
    }
  }
  
  // Add new EPJ with default status 'Available'
  const storeOnly = data.storeOnly === true || data.storeOnly === 'true';
  statusSheet.appendRow([epjNumber, 'Available', storeOnly]);
  
  // Clear caches
  CacheService.getScriptCache().remove('epjStatuses');
  CacheService.getScriptCache().remove('epjInfoMap');
  
  return 'EPJ ' + epjNumber + ' added successfully.';
}

// Remove an EPJ from the system
function adminRemoveEpj(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const epjNumber = (data.epjNumber || '').toString().trim().toUpperCase();
  if (!epjNumber) throw new Error('EPJ number is required.');
  
  const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
  if (!statusSheet) throw new Error('EPJ Status sheet not found.');
  
  const lastRow = statusSheet.getLastRow();
  if (lastRow < 2) throw new Error('No EPJs found in the system.');
  
  const epjData = statusSheet.getRange('A2:C' + lastRow).getValues();
  let rowToDelete = -1;
  
  for (let i = 0; i < epjData.length; i++) {
    if (epjData[i][0].toString().toUpperCase() === epjNumber) {
      rowToDelete = i + 2; // +2 because of header row and 0-index
      break;
    }
  }
  
  if (rowToDelete === -1) {
    throw new Error('EPJ ' + epjNumber + ' not found.');
  }
  
  // Check if EPJ is currently checked out
  if (epjData[rowToDelete - 2][1] === 'Checked Out') {
    throw new Error('Cannot remove EPJ ' + epjNumber + ' - it is currently checked out.');
  }
  
  statusSheet.deleteRow(rowToDelete);
  
  // Clear caches
  CacheService.getScriptCache().remove('epjStatuses');
  CacheService.getScriptCache().remove('epjInfoMap');
  
  return 'EPJ ' + epjNumber + ' removed successfully.';
}

// Get list of all EPJs for management
/**
 * Update the worked status for a checkout (column O)
 */
function adminUpdateWorkedStatus(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const tripId = data.tripId;
  const worked = data.worked === true;
  
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) throw new Error('Log sheet not found.');
  
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) throw new Error('No data in log sheet.');
  
  // Find the row with this tripId and Check-Out status
  const tripIds = logSheet.getRange('A2:A' + lastRow).getValues().flat();
  const statuses = logSheet.getRange('J2:J' + lastRow).getValues().flat();
  
  for (let i = 0; i < tripIds.length; i++) {
    if (tripIds[i] === tripId && statuses[i] === 'Check-Out') {
      // Update column O (column 15)
      logSheet.getRange(i + 2, 15).setValue(worked ? 'YES' : 'NO');
      
      // Clear active trips cache so the change shows up immediately
      CacheService.getScriptCache().remove('activeTrips');
      
      return { success: true, tripId: tripId, worked: worked };
    }
  }
  
  throw new Error('Trip ' + tripId + ' not found.');
}

function adminGetAllEpjs(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
  if (!statusSheet || statusSheet.getLastRow() < 2) return [];
  
  const values = statusSheet.getRange('A2:C' + statusSheet.getLastRow()).getValues();
  return values.filter(row => row[0]).map(row => ({
    epj: String(row[0] || '').trim(),
    status: String(row[1] || '').trim(),
    storeOnly: row[2] === true || row[2] === 'TRUE' || row[2] === 'Yes'
  }));
}

function sha256(input) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
  let hash = '';
   for (let i = 0; i < raw.length; i++) {
    let hex = (raw[i] < 0 ? raw[i] + 256 : raw[i]).toString(16);
    hash += (hex.length == 1 ? '0' : '') + hex;
  }
  return hash;
}

// ==================== ANNOUNCEMENTS ====================

/**
 * Get or create the Announcements sheet
 */
function getAnnouncementsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ANNOUNCEMENTS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ANNOUNCEMENTS_SHEET_NAME);
    sheet.appendRow(['ID', 'Created', 'CreatedBy', 'Message', 'Priority', 'ExpiresAt', 'Active']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Create a new announcement (Admin only)
 */
function adminCreateAnnouncement(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const message = (data.message || '').trim();
  if (!message) throw new Error('Message is required.');
  
  const sheet = getAnnouncementsSheet();
  const id = 'ANN-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const priority = data.priority || 'normal'; // normal, important, urgent
  
  // Calculate expiration (default 24 hours, or custom)
  let expiresAt = null;
  if (data.expiresInHours && data.expiresInHours > 0) {
    expiresAt = new Date(Date.now() + (data.expiresInHours * 60 * 60 * 1000));
  } else {
    expiresAt = new Date(Date.now() + (24 * 60 * 60 * 1000)); // Default 24 hours
  }
  
  sheet.appendRow([id, new Date(), session.username, message, priority, expiresAt, true]);
  
  return { success: true, id: id, message: 'Announcement created successfully.' };
}

/**
 * Get active announcements (for drivers and admins)
 */
function getActiveAnnouncements(data) {
  const session = getSession(data.token);
  if (!session) throw new Error('Invalid session.');
  
  const sheet = getAnnouncementsSheet();
  if (sheet.getLastRow() < 2) return [];
  
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const now = new Date();
  const announcements = [];
  
  rows.forEach(row => {
    const isActive = row[6] === true || row[6] === 'TRUE';
    const expiresAt = row[5] ? new Date(row[5]) : null;
    const isExpired = expiresAt && expiresAt < now;
    
    if (isActive && !isExpired) {
      announcements.push({
        id: row[0],
        created: new Date(row[1]).toLocaleString(),
        createdBy: row[2],
        message: row[3],
        priority: row[4] || 'normal',
        expiresAt: expiresAt ? expiresAt.toLocaleString() : null
      });
    }
  });
  
  // Sort by priority (urgent first) then by date (newest first)
  const priorityOrder = { urgent: 0, important: 1, normal: 2 };
  return announcements.sort((a, b) => {
    if (priorityOrder[a.priority] !== priorityOrder[b.priority]) {
      return priorityOrder[a.priority] - priorityOrder[b.priority];
    }
    return new Date(b.created) - new Date(a.created);
  });
}

/**
 * Delete/deactivate an announcement (Admin only)
 */
function adminDeleteAnnouncement(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const sheet = getAnnouncementsSheet();
  if (sheet.getLastRow() < 2) return 'Announcement not found.';
  
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  
  for (let i = 0; i < ids.length; i++) {
    if (ids[i] === data.announcementId) {
      // Set Active to false instead of deleting
      sheet.getRange(i + 2, 7).setValue(false);
      return 'Announcement deleted successfully.';
    }
  }
  
  return 'Announcement not found.';
}

/**
 * Get all announcements for admin management
 */
function adminGetAnnouncements(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const sheet = getAnnouncementsSheet();
  if (sheet.getLastRow() < 2) return [];
  
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const now = new Date();
  
  return rows.map(row => {
    const expiresAt = row[5] ? new Date(row[5]) : null;
    return {
      id: row[0],
      created: new Date(row[1]).toLocaleString(),
      createdBy: row[2],
      message: row[3],
      priority: row[4] || 'normal',
      expiresAt: expiresAt ? expiresAt.toLocaleString() : null,
      active: row[6] === true || row[6] === 'TRUE',
      expired: expiresAt && expiresAt < now
    };
  }).reverse(); // Newest first
}

// ==================== REPORTING ====================

/**
 * Get checkout history with optional date filtering
 */
function getCheckoutHistory(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet || logSheet.getLastRow() < 2) return { checkouts: [], checkins: [] };
  
  const startDate = data.startDate ? new Date(data.startDate) : new Date(Date.now() - 7 * 24 * 60 * 60 * 1000); // Default: last 7 days
  const endDate = data.endDate ? new Date(data.endDate) : new Date();
  endDate.setHours(23, 59, 59, 999); // Include full end day
  
  const rows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 13).getValues();
  const checkouts = [];
  const checkins = [];
  
  rows.forEach(row => {
    const timestamp = new Date(row[1]);
    if (timestamp < startDate || timestamp > endDate) return;
    
    const status = String(row[9] || '').trim();
    const entry = {
      tripId: row[0],
      timestamp: timestamp.toLocaleString(),
      timestampMs: timestamp.getTime(),
      username: row[2],
      driver: row[3],
      truck: row[4],
      trailer: row[5],
      epj: row[6],
      route: row[7],
      zone: row[8],
      status: status,
      notes: row[10],
      location: row[11]
    };
    
    if (status === 'Check-Out') {
      checkouts.push(entry);
    } else if (status === 'Check-In') {
      checkins.push(entry);
    }
  });
  
  return {
    checkouts: checkouts.sort((a, b) => b.timestampMs - a.timestampMs),
    checkins: checkins.sort((a, b) => b.timestampMs - a.timestampMs),
    dateRange: { start: startDate.toLocaleDateString(), end: endDate.toLocaleDateString() }
  };
}

/**
 * Get peak hours analysis (checkouts by hour of day)
 */
function getPeakHoursAnalysis(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet || logSheet.getLastRow() < 2) return { hourly: [], daily: [] };
  
  const daysBack = data.daysBack || 30;
  const cutoff = new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000);
  
  const rows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 10).getValues();
  
  // Initialize counters
  const hourlyCheckouts = Array(24).fill(0);
  const dailyCheckouts = { Sun: 0, Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0 };
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  let totalCheckouts = 0;
  
  rows.forEach(row => {
    const timestamp = new Date(row[1]);
    const status = String(row[9] || '').trim();
    
    if (status === 'Check-Out' && timestamp >= cutoff) {
      hourlyCheckouts[timestamp.getHours()]++;
      dailyCheckouts[dayNames[timestamp.getDay()]]++;
      totalCheckouts++;
    }
  });
  
  // Find peak hour
  let peakHour = 0;
  let peakCount = 0;
  hourlyCheckouts.forEach((count, hour) => {
    if (count > peakCount) {
      peakCount = count;
      peakHour = hour;
    }
  });
  
  return {
    hourly: hourlyCheckouts.map((count, hour) => ({
      hour: hour,
      label: `${hour.toString().padStart(2, '0')}:00`,
      count: count
    })),
    daily: Object.entries(dailyCheckouts).map(([day, count]) => ({ day, count })),
    peakHour: `${peakHour.toString().padStart(2, '0')}:00`,
    peakHourCount: peakCount,
    totalCheckouts: totalCheckouts,
    daysAnalyzed: daysBack
  };
}

/**
 * Get EPJ downtime/maintenance report
 */
function getEpjDowntimeReport(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const maintSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAINT_LOG_SHEET_NAME);
  if (!maintSheet || maintSheet.getLastRow() < 2) return { epjStats: [], totalDowntime: 0 };
  
  const daysBack = data.daysBack || 30;
  const cutoff = new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000);
  
  const rows = maintSheet.getRange(2, 1, maintSheet.getLastRow() - 1, 5).getValues();
  
  // Track maintenance periods per EPJ
  const epjMaintenance = {};
  const openMaintenance = {}; // Track ongoing maintenance
  
  rows.forEach(row => {
    const timestamp = new Date(row[0]);
    if (timestamp < cutoff) return;
    
    const epj = String(row[1] || '').trim();
    const eventType = String(row[2] || '').trim();
    
    if (!epj) return;
    if (!epjMaintenance[epj]) {
      epjMaintenance[epj] = { totalMinutes: 0, incidents: 0, reasons: [] };
    }
    
    if (eventType === 'Maintenance Start' || eventType === 'Issue Reported') {
      openMaintenance[epj] = timestamp;
      epjMaintenance[epj].incidents++;
      const reason = String(row[3] || 'Unknown').trim();
      if (reason && !epjMaintenance[epj].reasons.includes(reason)) {
        epjMaintenance[epj].reasons.push(reason);
      }
    } else if (eventType === 'Maintenance End' && openMaintenance[epj]) {
      const duration = (timestamp - openMaintenance[epj]) / (1000 * 60); // minutes
      epjMaintenance[epj].totalMinutes += duration;
      delete openMaintenance[epj];
    }
  });
  
  // Add ongoing maintenance time
  const now = new Date();
  Object.entries(openMaintenance).forEach(([epj, startTime]) => {
    const duration = (now - startTime) / (1000 * 60);
    epjMaintenance[epj].totalMinutes += duration;
  });
  
  // Convert to array and sort by downtime
  const epjStats = Object.entries(epjMaintenance).map(([epj, stats]) => ({
    epj: epj,
    totalMinutes: Math.round(stats.totalMinutes),
    totalHours: Math.round(stats.totalMinutes / 60 * 10) / 10,
    incidents: stats.incidents,
    reasons: stats.reasons.slice(0, 3), // Top 3 reasons
    isCurrentlyDown: !!openMaintenance[epj]
  })).sort((a, b) => b.totalMinutes - a.totalMinutes);
  
  const totalDowntime = epjStats.reduce((sum, s) => sum + s.totalMinutes, 0);
  
  return {
    epjStats: epjStats,
    totalDowntimeMinutes: totalDowntime,
    totalDowntimeHours: Math.round(totalDowntime / 60 * 10) / 10,
    daysAnalyzed: daysBack,
    epjsWithDowntime: epjStats.filter(s => s.totalMinutes > 0).length
  };
}

/**
 * Get daily summary statistics
 */
function getDailySummary(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet || logSheet.getLastRow() < 2) return {};
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  const rows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 10).getValues();
  
  let todayCheckouts = 0;
  let todayCheckins = 0;
  const driversToday = new Set();
  const epjsUsedToday = new Set();
  let totalTripMinutes = 0;
  let completedTrips = 0;
  
  // Track checkout times for calculating average trip duration
  const checkoutTimes = {};
  
  rows.forEach(row => {
    const timestamp = new Date(row[1]);
    const status = String(row[9] || '').trim();
    const tripId = row[0];
    const driver = row[3];
    const epj = row[6];
    
    if (timestamp >= today && timestamp < tomorrow) {
      if (status === 'Check-Out') {
        todayCheckouts++;
        if (driver) driversToday.add(driver);
        if (epj && !epj.startsWith('N/A')) epjsUsedToday.add(epj);
        checkoutTimes[tripId] = timestamp;
      } else if (status === 'Check-In') {
        todayCheckins++;
        if (checkoutTimes[tripId]) {
          const duration = (timestamp - checkoutTimes[tripId]) / (1000 * 60);
          totalTripMinutes += duration;
          completedTrips++;
        }
      }
    }
  });
  
  const avgTripMinutes = completedTrips > 0 ? Math.round(totalTripMinutes / completedTrips) : 0;
  
  return {
    date: today.toLocaleDateString(),
    checkouts: todayCheckouts,
    checkins: todayCheckins,
    activeTrips: todayCheckouts - todayCheckins,
    uniqueDrivers: driversToday.size,
    uniqueEpjs: epjsUsedToday.size,
    avgTripMinutes: avgTripMinutes,
    avgTripFormatted: avgTripMinutes > 60 
      ? `${Math.floor(avgTripMinutes / 60)}h ${avgTripMinutes % 60}m`
      : `${avgTripMinutes}m`
  };
}

/**
 * Get driver statistics for admin dashboard
 */
function getDriverStatistics(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const days = data.days || 30;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  cutoff.setHours(0, 0, 0, 0);
  
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet || logSheet.getLastRow() < 2) return { topDrivers: [], avgTripDuration: '0m', peakDay: 'N/A', totalTrips: 0 };
  
  const rows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 10).getValues();
  
  // Track driver checkouts
  const driverCheckouts = {};
  const driverTripMinutes = {};
  const checkoutTimes = {};
  const dayCheckouts = {};
  let totalTripMinutes = 0;
  let completedTrips = 0;
  
  rows.forEach(row => {
    const timestamp = new Date(row[1]);
    if (timestamp < cutoff) return;
    
    const tripId = row[0];
    const driver = String(row[3] || '').trim();
    const status = String(row[9] || '').trim();
    
    if (!driver) return;
    
    if (status === 'Check-Out') {
      // Count checkouts per driver
      driverCheckouts[driver] = (driverCheckouts[driver] || 0) + 1;
      checkoutTimes[tripId] = { timestamp, driver };
      
      // Track checkouts by day of week
      const dayName = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][timestamp.getDay()];
      dayCheckouts[dayName] = (dayCheckouts[dayName] || 0) + 1;
    } else if (status === 'Check-In' && checkoutTimes[tripId]) {
      // Calculate trip duration
      const duration = (timestamp - checkoutTimes[tripId].timestamp) / (1000 * 60);
      if (duration > 0 && duration < 1440) { // Ignore trips over 24 hours as likely errors
        totalTripMinutes += duration;
        completedTrips++;
        
        // Track per-driver durations
        const tripDriver = checkoutTimes[tripId].driver;
        if (!driverTripMinutes[tripDriver]) driverTripMinutes[tripDriver] = { total: 0, count: 0 };
        driverTripMinutes[tripDriver].total += duration;
        driverTripMinutes[tripDriver].count++;
      }
    }
  });
  
  // Sort drivers by checkout count and get top 10
  const topDrivers = Object.entries(driverCheckouts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([driver, checkouts]) => {
      const driverStats = driverTripMinutes[driver];
      let avgMinutes = 0;
      if (driverStats && driverStats.count > 0) {
        avgMinutes = Math.round(driverStats.total / driverStats.count);
      }
      return {
        driver,
        checkouts,
        avgTrip: avgMinutes > 60 
          ? `${Math.floor(avgMinutes / 60)}h ${avgMinutes % 60}m`
          : `${avgMinutes}m`
      };
    });
  
  // Calculate overall average trip duration
  const avgMinutes = completedTrips > 0 ? Math.round(totalTripMinutes / completedTrips) : 0;
  const avgTripFormatted = avgMinutes > 60 
    ? `${Math.floor(avgMinutes / 60)}h ${avgMinutes % 60}m`
    : `${avgMinutes}m`;
  
  // Find peak day
  let peakDay = 'N/A';
  let maxDayCheckouts = 0;
  Object.entries(dayCheckouts).forEach(([day, count]) => {
    if (count > maxDayCheckouts) {
      maxDayCheckouts = count;
      peakDay = day;
    }
  });
  
  // Total trips in period
  const totalTrips = Object.values(driverCheckouts).reduce((sum, c) => sum + c, 0);
  
  return {
    topDrivers,
    avgTripDuration: avgTripFormatted,
    peakDay,
    peakDayCount: maxDayCheckouts,
    totalTrips,
    completedTrips,
    periodDays: days
  };
}

/**
 * Export checkout data as CSV (returns CSV string)
 */
function exportCheckoutDataCsv(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const history = getCheckoutHistory(data);
  
  // Combine and sort all entries
  const allEntries = [...history.checkouts, ...history.checkins]
    .sort((a, b) => a.timestampMs - b.timestampMs);
  
  // Build CSV
  const headers = ['Timestamp', 'Trip ID', 'Driver', 'Username', 'EPJ', 'Truck', 'Trailer', 'Route', 'Zone', 'Status', 'Location', 'Notes'];
  const rows = allEntries.map(e => [
    e.timestamp,
    e.tripId,
    e.driver,
    e.username,
    e.epj,
    e.truck,
    e.trailer,
    e.route,
    e.zone,
    e.status,
    e.location || '',
    (e.notes || '').replace(/,/g, ';').replace(/\n/g, ' ')
  ]);
  
  const csv = [headers.join(','), ...rows.map(r => r.map(c => `"${c}"`).join(','))].join('\n');
  
  return csv;
}

// ==================== SITE SETTINGS ====================

/**
 * Default site-wide settings (admin-configurable)
 */
const DEFAULT_SITE_SETTINGS = {
  defaultOverspill: false,           // Default the checkout form to overspill mode
  defaultPreload: false,             // Default the checkout form to pre-load mode
  overspillStartHour: -1,            // Overspill schedule start hour (24h, -1 = disabled)
  overspillEndHour: -1,              // Overspill schedule end hour (24h, -1 = disabled)
  preloadStartHour: -1,              // Pre-load schedule start hour (24h, -1 = disabled)
  preloadEndHour: -1,                // Pre-load schedule end hour (24h, -1 = disabled)
  requireFaultReport: false,         // Require fault report on check-in
  autoExpireMinutes: 60,             // Auto-expire overspill trips after N minutes
  morningModeStartHour: 6,           // Morning mode start hour (24h)
  morningModeEndHour: 12,            // Morning mode end hour (24h)
  driverCanReportFaults: true,       // Allow drivers to report faults (but not set maintenance)
  maxCheckoutHours: 24,              // Alert if checkout exceeds N hours
  showCarrierColumn: true,           // Show carrier column in live checkouts
  siteMessage: ''                    // Global message shown on all pages
};

/**
 * Get or create the Site_Settings sheet
 */
function getSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.appendRow(['Setting', 'Value']);
    sheet.setFrozenRows(1);
    // Write defaults
    Object.entries(DEFAULT_SITE_SETTINGS).forEach(([key, value]) => {
      sheet.appendRow([key, String(value)]);
    });
  }
  return sheet;
}

/**
 * Get all site settings (cached for 5 minutes)
 */
function getSiteSettings(data) {
  if (data && data.token) {
    const session = getSession(data.token);
    if (!session) throw new Error('Invalid session.');
  }
  
  const cache = CacheService.getScriptCache();
  const cached = cache.get('site_settings');
  if (cached) {
    return JSON.parse(cached);
  }
  
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  const settings = { ...DEFAULT_SITE_SETTINGS };
  
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    rows.forEach(row => {
      const key = String(row[0]).trim();
      const value = String(row[1]).trim();
      if (key && DEFAULT_SITE_SETTINGS.hasOwnProperty(key)) {
        // Parse value back to correct type
        if (typeof DEFAULT_SITE_SETTINGS[key] === 'boolean') {
          settings[key] = value === 'true';
        } else if (typeof DEFAULT_SITE_SETTINGS[key] === 'number') {
          settings[key] = Number(value) || DEFAULT_SITE_SETTINGS[key];
        } else {
          settings[key] = value;
        }
      }
    });
  }
  
  cache.put('site_settings', JSON.stringify(settings), 300); // 5 min cache
  return settings;
}

/**
 * Update a site setting (Admin only)
 */
function adminUpdateSiteSetting(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const key = data.key;
  const value = data.value;
  
  if (!DEFAULT_SITE_SETTINGS.hasOwnProperty(key)) {
    throw new Error('Unknown setting: ' + key);
  }
  
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  let found = false;
  
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i]).trim() === key) {
        sheet.getRange(i + 2, 2).setValue(String(value));
        found = true;
        break;
      }
    }
  }
  
  if (!found) {
    sheet.appendRow([key, String(value)]);
  }
  
  // Clear cached settings
  CacheService.getScriptCache().remove('site_settings');
  
  return { success: true, key: key, value: value };
}

/**
 * Update multiple site settings at once (Admin only)
 */
function adminUpdateSiteSettings(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error('Permission denied.');
  
  const settings = data.settings;
  if (!settings || typeof settings !== 'object') throw new Error('Invalid settings data.');
  
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  const existingKeys = {};
  
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    keys.forEach((k, i) => { existingKeys[String(k).trim()] = i + 2; });
  }
  
  let updated = 0;
  Object.entries(settings).forEach(([key, value]) => {
    if (!DEFAULT_SITE_SETTINGS.hasOwnProperty(key)) return;
    
    if (existingKeys[key]) {
      sheet.getRange(existingKeys[key], 2).setValue(String(value));
    } else {
      sheet.appendRow([key, String(value)]);
    }
    updated++;
  });
  
  CacheService.getScriptCache().remove('site_settings');
  
  return { success: true, updated: updated, message: updated + ' setting(s) saved.' };
}

