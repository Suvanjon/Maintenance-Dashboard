// ============================================================================
// GOOGLE APPS SCRIPT — Preventive Maintenance data relay
// Deploy as a Web App (Execute as: Me, Who has access: Anyone)
// Then paste the deployed URL into pm_data_entry.html (WEB_APP_URL)
// ============================================================================

var SHEET_ID = '1W76G3n90DZh5pqcowYbCqByezaFSIiFokO7Ic2Uvmzs';
var SHEET_NAME = 'Sheet1';

function doGet(e) {
  return ContentService.createTextOutput('Preventive Maintenance Relay OK')
      .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    var p = e.parameter || {};

    // Build row: Timestamp | Section | Date | Machine | checks...
    var row = [];
    row.push(new Date());                                    // Timestamp
    row.push(p.section || 'All Departments');                // Section
    row.push(p.date || '');                                  // Date
    row.push(p.machine || '');                               // Machine Name

    // All other fields are checklist answers (entry.* keys)
    Object.keys(p).forEach(function (k) {
      if (k === 'section' || k === 'date' || k === 'machine') return;
      row.push(p[k]);
    });

    // Ensure header row exists
    var lastRow = sheet.getLastRow();
    if (lastRow === 0 || sheet.getRange(1, 1).getValue() === '') {
      var headers = ['Timestamp', 'Section', 'Date', 'Machine Name'];
      Object.keys(p).forEach(function (k) {
        if (k === 'section' || k === 'date' || k === 'machine') return;
        headers.push(k);
      });
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#005151')
          .setFontColor('#fff');
    }

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', rows: sheet.getLastRow() }))
        .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
  }
}
