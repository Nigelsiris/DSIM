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
  const directory = getUserDirectory();
  return directory.list
    .map(record => ({ username: record.username, role: record.role, carrier: record.carrier }))
    .filter(user => user.username);
}
// --- CONFIGURATION ---
const LOG_SHEET_NAME = 'Form Responses 1';
const STATUS_SHEET_NAME = 'EPJ_Status';
const USER_SHEET_NAME = 'Users';
const MAINT_LOG_SHEET_NAME = 'Maintenance_Log';
const ZONES_SHEET_NAME = 'Zones';
const LOGIN_LOG_SHEET_NAME = 'Login_Log';

const CACHE_KEYS = Object.freeze({
  ACTIVE_TRIPS: 'activeTrips',
  EPJ_INFO_MAP: 'epjInfoMap',
  ACTIVE_DRIVERS: 'activeDrivers',
  ZONE_OPTIONS: 'zoneOptions',
  EPJ_STATUSES: 'epjStatuses',
  USER_DIRECTORY: 'user_directory_v1'
});

let activeSpreadsheet = null;

function getSpreadsheet() {
  if (!activeSpreadsheet) {
    activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  return activeSpreadsheet;
}

function getSheet(sheetName) {
  const spreadsheet = getSpreadsheet();
  return spreadsheet ? spreadsheet.getSheetByName(sheetName) : null;
}

function getUserDirectory(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const cached = cache.get(CACHE_KEYS.USER_DIRECTORY);
    if (cached) return JSON.parse(cached);
  }

  const directory = { list: [], byUsername: {} };
  const userSheet = getSheet(USER_SHEET_NAME);
  if (userSheet) {
    const lastRow = userSheet.getLastRow();
    if (lastRow > 1) {
      const values = userSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      values.forEach((row, index) => {
        const username = row[0];
        if (!username) return;
        const record = {
          username: username,
          passwordHash: row[1],
          role: row[2],
          carrier: row[3] || '',
          rowIndex: index + 2
        };
        directory.list.push(record);
        directory.byUsername[username.toLowerCase()] = record;
      });
    }
  }

  cache.put(CACHE_KEYS.USER_DIRECTORY, JSON.stringify(directory), 600);
  return directory;
}

function invalidateUserCaches() {
  CacheService.getScriptCache().removeAll([CACHE_KEYS.USER_DIRECTORY, 'admin_users']);
}

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
function clearStateCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll([
    CACHE_KEYS.ACTIVE_TRIPS,
    CACHE_KEYS.EPJ_INFO_MAP,
    CACHE_KEYS.ACTIVE_DRIVERS,
    CACHE_KEYS.ZONE_OPTIONS,
    CACHE_KEYS.EPJ_STATUSES
  ]);
}

function updateAllEpjStatuses() {
  const logSheet = getSheet(LOG_SHEET_NAME);
  const statusSheet = getSheet(STATUS_SHEET_NAME);
  if (!logSheet || !statusSheet) return;

  const logData = logSheet.getDataRange().getValues();
  const lastRow = statusSheet.getLastRow();
  if (lastRow < 2) return;
  
  const epjList = statusSheet.getRange("A2:A" + lastRow).getValues().flat().filter(String);
  const statusMap = {};
  epjList.forEach(epj => { if (epj) statusMap[epj] = 'Available'; });

  const processedEpjs = new Set();
  for (let i = logData.length - 1; i >= 1; i--) {
    if (processedEpjs.size === epjList.length) break;
    const row = logData[i];
    const epj = row[6];
    const status = row[9] ? row[9].trim() : "";
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
  clearStateCache();
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
  const template = HtmlService.createTemplateFromFile('LoginPage');
  template.activeDrivers = getActiveDriverNames();
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
      const template = HtmlService.createTemplateFromFile('CheckInForm');
      template.username = session.username;
      template.tripInfo = activeTrip;
      template.zoneOptions = getZoneOptions();
      return template.evaluate().getContent();
    } else {
      const template = HtmlService.createTemplateFromFile('CheckOutForm');
      template.username = session.username;
      template.availableEpjs = JSON.stringify(getEpjsByStatus('Available'));
      template.zoneOptions = getZoneOptions();
      template.epjInfoMap = getEpjInfoMap();
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
  template.token = token; // Recommended: Do the same for other user roles
  template.epjStatuses = JSON.stringify(getEpjsByStatus(null, true));
  template.epjInfoMap = JSON.stringify(getEpjInfoMap());
  template.zoneOptions = getZoneOptions();
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
  const directory = getUserDirectory();
  const passwordHash = sha256(password);
  let user = null;

  const userRecord = directory.byUsername[username.toLowerCase()];
  if (userRecord && userRecord.passwordHash === passwordHash) {
    user = { username: userRecord.username, role: userRecord.role };
  }
  const loginLogSheet = getSheet(LOGIN_LOG_SHEET_NAME);
  if (loginLogSheet) {
    let isAtWarehouse = false;
    if (latitude && longitude) {
      const distance = calculateDistance(latitude, longitude, WAREHOUSE_LAT, WAREHOUSE_LON);
      isAtWarehouse = (distance <= GEOFENCE_RADIUS_METERS);
      loginLogSheet.appendRow([new Date(), username, latitude, longitude, isAtWarehouse]);
    } else {
      loginLogSheet.appendRow([new Date(), username, "Not Provided", "Not Provided", false]);
    }
  }
  if (!user) return null;
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(token, JSON.stringify(user), 86400);
  let htmlContent = getUserView(token);
  return { token: token, role: user.role, html: htmlContent };
}

function getSession(token) {
  if (!token) return null;
  const sessionData = CacheService.getScriptCache().get(token);
  return sessionData ? JSON.parse(sessionData) : null;
}

function logoutUser(data) {
    if (data && data.token) { CacheService.getScriptCache().remove(data.token); }
    clearStateCache();
    return true;
}

// --- DRIVER WORKFLOWS ---
function processCheckOut(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    // Invalidate EPJ status cache before checking availability
    CacheService.getScriptCache().remove(CACHE_KEYS.EPJ_STATUSES);
    Logger.log('processCheckOut called with: ' + JSON.stringify(data));
    const session = getSession(data.token);
    Logger.log('Session: ' + JSON.stringify(session));
    if (!session) throw new Error("Invalid session.");
    let availableEpjs = getEpjsByStatus('Available');
    Logger.log('Available EPJs: ' + JSON.stringify(availableEpjs));
    Logger.log('Requested EPJ: ' + data.epjNumber);
    // Convert both availableEpjs and data.epjNumber to strings for comparison
    availableEpjs = availableEpjs.map(String);
    const requestedEpj = String(data.epjNumber);
    if (!availableEpjs.includes(requestedEpj)) {
      Logger.log('EPJ not available error triggered.');
      return `Error: EPJ ${data.epjNumber} is no longer available. It may have just been checked out.`;
    }
    const epjInfoMap = getEpjInfoMap();
    const startingZone = epjInfoMap[data.epjNumber] ? epjInfoMap[data.epjNumber].location : 'Unknown';
    const logSheet = getSheet(LOG_SHEET_NAME);
    const tripId = "TRIP-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    logSheet.appendRow([tripId, new Date(), session.username, data.driverName, data.truckNumber, data.trailerNumber, data.epjNumber, data.route, startingZone, "Check-Out", data.faultReport, "", ""]);
    SpreadsheetApp.flush();
    updateAllEpjStatuses();
    // Invalidate again after status update
    CacheService.getScriptCache().remove(CACHE_KEYS.EPJ_STATUSES);
    Logger.log('Check-Out successful for EPJ: ' + data.epjNumber);
    return `Successfully checked out EPJ ${data.epjNumber}.`;
  } finally {
    lock.releaseLock();
  }
}

function processCheckIn(data) {
    const session = getSession(data.token);
    if (!session) throw new Error("Invalid session.");
    const activeTrip = findActiveTrip(session.username); 
    if (!activeTrip) return "Error: No active trip found.";
    const logSheet = getSheet(LOG_SHEET_NAME);
    logSheet.appendRow([
        activeTrip.tripId, new Date(), session.username, activeTrip.driver, "", "",
        activeTrip.epj, "", activeTrip.zone, "Check-In", "",
        data.checkInZone, data.faultReport, data.pluggedIn
    ]);
  updateAllEpjStatuses();
  CacheService.getScriptCache().remove(CACHE_KEYS.EPJ_STATUSES);
  return `Successfully checked in EPJ ${activeTrip.epj}.`;
}

function driverChangePassword(data) {
  const session = getSession(data.token);
  if (!session) throw new Error("Invalid session.");
  const currentPasswordHash = sha256(data.currentPassword);
  const directory = getUserDirectory();
  const record = directory.byUsername[session.username.toLowerCase()];
  if (!record) {
    return "Error: Could not find user profile.";
  }
  if (record.passwordHash !== currentPasswordHash) {
    return "Error: Incorrect current password.";
  }
  const userSheet = getSheet(USER_SHEET_NAME);
  if (!userSheet) {
    return "Error: Could not access user directory.";
  }
  userSheet.getRange(record.rowIndex, 2).setValue(sha256(data.newPassword));
  invalidateUserCaches();
  return "Password updated successfully!";
}


// --- ADMIN & LOAD SUPPORT WORKFLOWS ---
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
  const userDirectory = getUserDirectory();
  const users = userDirectory.list
    .map(record => ({ username: record.username, role: record.role, carrier: record.carrier }))
    .filter(user => user.username);
  const userMap = {};
  userDirectory.list.forEach(record => {
    userMap[record.username.toLowerCase()] = record;
  });
  let maintenanceLog = [];
  let maintCached = cache.get('admin_maintlog');
  if (maintCached) {
    maintenanceLog = JSON.parse(maintCached);
  } else {
    const maintSheet = getSheet(MAINT_LOG_SHEET_NAME);
    if (maintSheet && maintSheet.getLastRow() > 1) {
      const lastMaintRow = maintSheet.getLastRow();
      const startRow = Math.max(2, lastMaintRow - 19);
      const numRows = lastMaintRow - startRow + 1;
      maintenanceLog = maintSheet.getRange(startRow, 1, numRows, 5).getValues().reverse();
      cache.put('admin_maintlog', JSON.stringify(maintenanceLog), 300); // cache for 5 min
    }
  }
  const activeCheckouts = getActiveCheckouts().map(checkout => {
    const driverInfo = userMap[(checkout.driverUsername || '').toLowerCase()];
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
  const logSheet = getSheet(LOG_SHEET_NAME);
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
    logSheet.appendRow([
        data.tripId, new Date(), session.username, originalDriver, "", "",
        data.epj, "", originalZone, "Check-In", "", "Admin Override",
        `Forced check-in by admin ${session.username}`
    ]);
  updateAllEpjStatuses();
  CacheService.getScriptCache().remove(CACHE_KEYS.EPJ_STATUSES);
  return `Successfully checked in EPJ ${data.epj}.`;
  }
  return `Error: Could not find original trip ID ${data.tripId}.`;
}

function updateEpjLocation(data) {
    const session = getSession(data.token);
    if (!session || (session.role !== 'Admin' && session.role !== 'Load Support')) { throw new Error("Permission denied."); }
    const logSheet = getSheet(LOG_SHEET_NAME);
    logSheet.appendRow([ '', new Date(), session.username, 'LOAD SUPPORT', '', '', data.epj, '', data.newLocation, 'Location Update', `Updated by ${session.role}`, data.newLocation, '' ]);
  updateAllEpjStatuses();
  CacheService.getScriptCache().remove(CACHE_KEYS.EPJ_STATUSES);
  return `Location for EPJ ${data.epj} updated to ${data.newLocation}.`;
}

function adminSetMaintenanceStatus(data) {
    const session = getSession(data.token);
    if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
    const logSheet = getSheet(LOG_SHEET_NAME);
    const maintSheet = getSheet(MAINT_LOG_SHEET_NAME);
    if (data.status === 'Maintenance Start') {
      maintSheet.appendRow([new Date(), data.epj, 'Maintenance Start', data.reason || 'Reason pending', '']);
      logSheet.appendRow(['', new Date(), session.username, 'ADMIN', '', '', data.epj, '', '', 'Maintenance Start', data.reason || 'Reason pending', '', '']);
    } else {
      maintSheet.appendRow([new Date(), data.epj, 'Maintenance End', '', 'Returned to service']);
      logSheet.appendRow(['', new Date(), session.username, 'ADMIN', '', '', data.epj, '', '', 'Maintenance End', 'Returned to service', '', '']);
    }
  updateAllEpjStatuses();
  CacheService.getScriptCache().remove(CACHE_KEYS.EPJ_STATUSES);
  return `EPJ ${data.epj} status updated.`;
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
        const userSheet = getSheet(USER_SHEET_NAME);
        if (!userSheet) {
            errors.push('Unable to locate user sheet.');
        } else {
            const startRow = userSheet.getLastRow() + 1;
            userSheet.getRange(startRow, 1, newUsers.length, 4).setValues(newUsers);
            invalidateUserCaches();
        }
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
  const userSheet = getSheet(USER_SHEET_NAME);
  if (!userSheet) {
    return `Error: Unable to locate user sheet.`;
  }
  userSheet.appendRow([data.username, sha256(data.password), data.role, data.carrier]);
  invalidateUserCaches();
  return `User "${data.username}" created successfully.`;
}

function adminEditUser(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  const userSheet = getSheet(USER_SHEET_NAME);
  if (!userSheet) {
    return `Error: Unable to locate user sheet.`;
  }
  const directory = getUserDirectory();
  const record = directory.byUsername[data.username.toLowerCase()];
  if (!record) {
    return `Error: User "${data.username}" not found.`;
  }
  userSheet.getRange(record.rowIndex, 3).setValue(data.role);
  userSheet.getRange(record.rowIndex, 4).setValue(data.carrier);
  invalidateUserCaches();
  return `User "${data.username}" updated successfully.`;
}

function adminResetPassword(data) {
    const session = getSession(data.token);
    if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
        return `Error: Unable to locate user sheet.`;
    }
    const directory = getUserDirectory();
    const record = directory.byUsername[data.username.toLowerCase()];
    if (!record) {
        return `Error: User "${data.username}" not found.`;
    }
    userSheet.getRange(record.rowIndex, 2).setValue(sha256(data.newPassword));
    invalidateUserCaches();
    return `Password reset for user "${data.username}".`;
}

function adminDeleteUser(data) {
  const session = getSession(data.token);
  if (!session || session.role !== 'Admin') throw new Error("Permission denied.");
  const userSheet = getSheet(USER_SHEET_NAME);
  if (!userSheet) {
    return `Error: Unable to locate user sheet.`;
  }
  const directory = getUserDirectory();
  const record = directory.byUsername[data.username.toLowerCase()];
  if (!record) {
    return `Error: User "${data.username}" not found.`;
  }
  userSheet.deleteRow(record.rowIndex);
  invalidateUserCaches();
  return `User "${data.username}" has been deleted.`;
}

// --- HELPERS & DATA GETTERS ---
function getEpjInfoMap() {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('epjInfoMap');
    if (cached != null) { return JSON.parse(cached); }
    const logSheet = getSheet(LOG_SHEET_NAME);
    const epjStatusSheet = getSheet(STATUS_SHEET_NAME);
    if (!logSheet || !epjStatusSheet) { return {}; }
    const logData = logSheet.getDataRange().getValues();
    const epjs = epjStatusSheet.getRange("A2:A").getValues().flat().filter(String);
    const infoMap = {};
    const foundLocations = new Set();
    const foundFaults = new Set();
    for (const epj of epjs) { infoMap[epj] = { location: "N/A", fault: "No issues reported" }; }
    for (let i = logData.length - 1; i >= 1; i--) {
        if (foundLocations.size === epjs.length && foundFaults.size === epjs.length) { break; }
        const row = logData[i];
        const epj = row[6];
        if (!infoMap[epj] || (foundLocations.has(epj) && foundFaults.has(epj))) { continue; }
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
    if (cached != null) { return JSON.parse(cached); }
    const logSheet = getSheet(LOG_SHEET_NAME);
    if (!logSheet) { return []; }
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return [];
    const data = logSheet.getRange(2, 1, lastRow - 1, 13).getValues();
    const activeTrips = {};
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const tripId = row[0];
      if (!tripId) continue;
      const status = row[9];
      if (status === 'Check-Out') {
        activeTrips[tripId] = {
          tripId: tripId, timestamp: new Date(row[1]).toLocaleString(), driverUsername: row[2],
          driver: row[3], truck: row[4], trailer: row[5], epj: row[6], route: row[7], zone: row[8]
        };
      } else if (status === 'Check-In') {
        if(activeTrips[tripId]) { delete activeTrips[tripId]; }
      }
    }
    const result = Object.values(activeTrips);
    cache.put('activeTrips', JSON.stringify(result), 21600);
    return result;
}

function findActiveTrip(username) {
  const allActiveTrips = getActiveCheckouts();
  return allActiveTrips.find(trip => trip.driverUsername === username) || null;
}

function getActiveDriverNames() {
  const checkouts = getActiveCheckouts();
  const driverMap = new Map();
  checkouts.forEach(trip => {
    driverMap.set(trip.driverUsername, trip.driver);
  });
  return Array.from(driverMap, ([username, displayName]) => ({ username, displayName }));
}

function getZoneOptions() {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('zoneOptions');
    if (cached != null) { return cached; }
    const zoneSheet = getSheet(ZONES_SHEET_NAME);
    if (!zoneSheet) return "";
    const zones = zoneSheet.getRange("A2:A").getValues().flat().filter(String);
    let options = '';
    zones.forEach(zone => { options += `<option value="${zone}">${zone}</option>`; });
    cache.put('zoneOptions', options, 21600);
    return options;
}

function getEpjsByStatus(status, all = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = CACHE_KEYS.EPJ_STATUSES;
  const cached = cache.get(cacheKey);
  let allStatuses;
  if (cached != null) {
    allStatuses = JSON.parse(cached);
  } else {
    const sheet = getSheet(STATUS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const values = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
    allStatuses = values.map(row => ({epj: row[0], status: row[1]})).filter(item => item.epj);
    cache.put(cacheKey, JSON.stringify(allStatuses), 21600);
  }
  if (all) { return allStatuses; }
  return allStatuses.filter(item => item.status === status).map(item => item.epj);
}

function getEquipmentStatusViewData() {
  const statuses = getEpjsByStatus(null, true);
  const infoMap = getEpjInfoMap();
  return statuses.map(item => {
    const info = infoMap[item.epj] || { location: 'N/A' };
    return { epj: item.epj, status: item.status, location: info.location };
  });
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

