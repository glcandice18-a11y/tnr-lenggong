// Monstera TNR Tracker — Google Apps Script Backend
//
// DEPLOY INSTRUCTIONS:
// 1. Create a new Google Sheet at https://sheets.new
// 2. Go to Extensions → Apps Script
// 3. Delete ALL code in Code.gs and paste this entire file
// 4. From the dropdown, select setupSheets and click Run
// 5. Grant permissions when prompted
// 6. Deploy → New deployment → Web app → Execute as: Me, Access: Anyone
// 7. Copy the Web app URL and paste into index.html settings

var SHEET_NAME = 'Colonies';
var LOG_SHEET_NAME = 'Activity Log';

// ============================================================
// SETUP — Run setupSheets() once from the Apps Script editor
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var coloniesSheet = ss.getSheetByName(SHEET_NAME);
  if (!coloniesSheet) {
    coloniesSheet = ss.insertSheet(SHEET_NAME);
  }
  coloniesSheet.clear();
  
  var colonyHeaders = [
    'ID', 'Type', 'Name', 'Location', 'CatCount', 'ColonyType',
    'FeedingTimes', 'Status', 'DateSeen', 'Notes', 'Reporter',
    'NeuteredCats', 'NeuterDate', 'CreatedAt', 'UpdatedAt'
  ];
  coloniesSheet.appendRow(colonyHeaders);
  coloniesSheet.getRange(1, 1, 1, colonyHeaders.length).setFontWeight('bold');
  coloniesSheet.getRange(1, 1, 1, colonyHeaders.length).setBackground('#1a3a2a');
  coloniesSheet.getRange(1, 1, 1, colonyHeaders.length).setFontColor('#ffffff');
  
  var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
  }
  logSheet.clear();
  
  var logHeaders = ['Timestamp', 'Action', 'Message', 'By'];
  logSheet.appendRow(logHeaders);
  logSheet.getRange(1, 1, 1, logHeaders.length).setFontWeight('bold');
  logSheet.getRange(1, 1, 1, logHeaders.length).setBackground('#1a3a2a');
  logSheet.getRange(1, 1, 1, logHeaders.length).setFontColor('#ffffff');
  
  Logger.log('Setup complete! Sheets created: ' + SHEET_NAME + ', ' + LOG_SHEET_NAME);
}

// ============================================================
// GET ALL COLONIES
// ============================================================
function getAllColonies() {
  var sheet = getColoniesSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r[0] === '') continue;
    result.push({
      id: r[0],
      type: r[1],
      name: r[2],
      location: r[3],
      catCount: parseInt(r[4]) || 1,
      colonyType: r[5],
      feedingTimes: parseCSV(r[6]),
      status: r[7],
      dateSeen: r[8],
      notes: r[9],
      reporter: r[10],
      neuteredCats: parseCSV(r[11]),
      neuterDate: r[12],
      createdAt: r[13],
      updatedAt: r[14]
    });
  }
  return result;
}

// ============================================================
// CREATE NEW COLONY
// ============================================================
function createColony(entry) {
  var sheet = getColoniesSheet();
  var now = new Date();
  var id = 'c' + Date.now();
  
  var row = [
    id,
    entry.type || 'colony',
    entry.name,
    entry.location,
    entry.catCount || 1,
    entry.colonyType || 'unknown',
    (entry.feedingTimes || []).join(','),
    entry.status || 'unknown',
    entry.dateSeen || formatDate(now),
    entry.notes || '',
    entry.reporter || 'Anonymous',
    (entry.status === 'neutered' && entry.name) ? entry.name : '',
    entry.status === 'neutered' ? entry.dateSeen || formatDate(now) : '',
    formatDate(now),
    formatDate(now)
  ];
  
  sheet.appendRow(row);
  addLog('report', 'New report: "' + entry.name + '" at ' + entry.location, entry.reporter || 'App');
  
  return { success: true, id: id };
}

// ============================================================
// UPDATE COLONY STATUS
// ============================================================
function updateColonyStatus(id, updates) {
  var sheet = getColoniesSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      var row = i + 1;
      
      if (updates.status) {
        setCell(sheet, row, getColumnIndex(headers, 'Status'), updates.status);
      }
      if (updates.neuterDate) {
        setCell(sheet, row, getColumnIndex(headers, 'NeuterDate'), updates.neuterDate);
      }
      if (updates.neuteredCats) {
        setCell(sheet, row, getColumnIndex(headers, 'NeuteredCats'), updates.neuteredCats.join(','));
      }
      if (updates.notes) {
        setCell(sheet, row, getColumnIndex(headers, 'Notes'), updates.notes);
      }
      if (updates.colonyType) {
        setCell(sheet, row, getColumnIndex(headers, 'ColonyType'), updates.colonyType);
      }
      if (updates.feedingTimes) {
        setCell(sheet, row, getColumnIndex(headers, 'FeedingTimes'), updates.feedingTimes.join(','));
      }
      if (updates.location) {
        setCell(sheet, row, getColumnIndex(headers, 'Location'), updates.location);
      }
      if (updates.name) {
        setCell(sheet, row, getColumnIndex(headers, 'Name'), updates.name);
      }
      
      setCell(sheet, row, getColumnIndex(headers, 'UpdatedAt'), formatDate(new Date()));
      
      addLog('status_update', 'Updated colony #' + id, 'App');
      return { success: true };
    }
  }
  
  return { success: false, error: 'Colony not found' };
}

// ============================================================
// DELETE COLONY
// ============================================================
function deleteColony(id) {
  var sheet = getColoniesSheet();
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      addLog('delete', 'Deleted colony #' + id, 'App');
      return { success: true };
    }
  }
  
  return { success: false, error: 'Colony not found' };
}

// ============================================================
// SEARCH COLONIES
// ============================================================
function searchColonies(query) {
  var colonies = getAllColonies();
  if (!query) return colonies;
  
  var q = query.toLowerCase();
  var result = [];
  for (var i = 0; i < colonies.length; i++) {
    var c = colonies[i];
    if ((c.name || '').toLowerCase().includes(q) ||
        (c.location || '').toLowerCase().includes(q) ||
        (c.notes || '').toLowerCase().includes(q) ||
        (c.colonyType || '').toLowerCase().includes(q) ||
        (c.status || '').toLowerCase().includes(q) ||
        (c.reporter || '').toLowerCase().includes(q)) {
      result.push(c);
    }
  }
  return result;
}

// ============================================================
// GET ACTIVITY LOG
// ============================================================
function getLog() {
  var sheet = getLogSheet();
  var data = sheet.getDataRange().getValues();
  var rows = data.slice(1);
  
  var result = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    if (r[0] === '') continue;
    result.push({
      timestamp: r[0],
      action: r[1],
      message: r[2],
      by: r[3]
    });
  }
  return result;
}

// ============================================================
// STATS
// ============================================================
function getStats() {
  var colonies = getAllColonies();
  var total = 0, neutered = 0, intact = 0, pending = 0, unknown = 0;
  
  for (var i = 0; i < colonies.length; i++) {
    var c = colonies[i];
    var count = c.catCount || 1;
    total += count;
    if (c.status === 'neutered') neutered += count;
    else if (c.status === 'intact') intact += count;
    else if (c.status === 'pending') pending += count;
    else unknown += count;
  }
  
  return { total: total, neutered: neutered, intact: intact, pending: pending, unknown: unknown, colonyCount: colonies.length };
}

// ============================================================
// GOOGLE SCRIPTS WEB APP HANDLERS
// These handle HTTP requests from the frontend app
// ============================================================
function doGet(e) {
  var params = e.parameters;
  var cmd = params['_cmd'];
  var result;
  
  if (cmd === 'getAllColonies') {
    result = getAllColonies();
  } else if (cmd === 'getStats') {
    result = getStats();
  } else if (cmd === 'search') {
    result = searchColonies(params['_q']);
  } else if (cmd === 'getLog') {
    result = getLog();
  } else {
    result = { success: false, error: 'Unknown command' };
  }
  
  return ContentService.createTextOutput(JSON.stringify({ success: true, result: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var params = e.parameter;
  var cmd = params['_cmd'];
  var result;
  
  if (cmd === 'create') {
    var payload = JSON.parse(params['_data']);
    result = createColony(payload);
  } else if (cmd === 'update') {
    var payload = JSON.parse(params['_data']);
    result = updateColonyStatus(payload.id, payload.updates);
  } else if (cmd === 'delete') {
    result = deleteColony(params['_id']);
  } else {
    result = { success: false, error: 'Unknown command' };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================
function getColoniesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID', 'Type', 'Name', 'Location', 'CatCount', 'ColonyType', 'FeedingTimes', 'Status', 'DateSeen', 'Notes', 'Reporter', 'NeuteredCats', 'NeuterDate', 'CreatedAt', 'UpdatedAt']);
    sheet.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  }
  return sheet;
}

function getLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Action', 'Message', 'By']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1a3a2a').setFontColor('#ffffff');
  }
  return sheet;
}

function addLog(action, message, by) {
  var sheet = getLogSheet();
  sheet.appendRow([new Date(), action, message, by]);
}

function formatDate(date) {
  if (!date) return '';
  var d = new Date(date);
  return d.toISOString().split('T')[0];
}

function parseCSV(str) {
  if (!str) return [];
  if (typeof str === 'string') {
    return str.split(',').filter(function(s) { return s.trim(); });
  }
  return [];
}

function setCell(sheet, row, col, value) {
  sheet.getRange(row, col).setValue(value);
}

function getColumnIndex(headers, columnName) {
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === columnName) return i + 1;
  }
  return 1;
}
