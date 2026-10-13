/**
 * Distributions for OOD Time.
 *
 * Paste this whole file into the Apps Script project as a new file
 * (Extensions > Apps Script > + > Script, name it "distributions"),
 * then add two lines to the existing doGet(e) dispatch, next to the
 * 'addExpense' and 'getExpenses' cases:
 *
 *   if (action === 'addDistribution')  return addDistribution_(e);
 *   if (action === 'getDistributions') return getDistributions_();
 *
 * Then Deploy > Manage deployments > edit > Version: New version > Deploy.
 *
 * Rows go to the Time Log with Type "Distribution":
 *   A Date | B Who | C Hours (blank) | D Project (blank) | E Notes |
 *   F Rate (blank) | G Amount | H Submitted At | I Type | J Phase (blank) | K Account (blank)
 * Project stays blank so no project, invoice or hours total ever picks it up.
 */

var DIST_SHEET_NAME = 'Time Log';

function distJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function addDistribution_(e) {
  try {
    var p = e.parameter;
    var amount = parseFloat(p.amount);
    if (!p.date || !p.person || !(amount > 0)) {
      return distJson_({ status: 'error', message: 'Date, person and a positive amount are required.' });
    }
    var date = new Date(p.date);
    if (isNaN(date.getTime())) {
      return distJson_({ status: 'error', message: 'Could not read the date.' });
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DIST_SHEET_NAME);
    var submitted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy, h:mm:ss a');
    sheet.appendRow([date, p.person, '', '', p.notes || 'Draw', '', amount, submitted, 'Distribution', '', '']);
    return distJson_({ status: 'ok' });
  } catch (err) {
    return distJson_({ status: 'error', message: String(err) });
  }
}

function getDistributions_() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DIST_SHEET_NAME);
    var rows = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (String(r[8]).toLowerCase() !== 'distribution') continue;
      var d = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'MMM d, yyyy') : String(r[0]);
      out.push({ date: d, person: r[1], amount: r[6], notes: r[4] });
    }
    return distJson_({ status: 'ok', distributions: out });
  } catch (err) {
    return distJson_({ status: 'error', message: String(err) });
  }
}
