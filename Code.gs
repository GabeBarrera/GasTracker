/**
 * Gas Tracker - Apps Script backend
 * -----------------------------------------------------------------------
 * Deploy this bound to a Google Sheet (Extensions > Apps Script), together
 * with the "index" HTML file, as a Web App.
 *
 * AUTHENTICATION:
 * Authentication is handled entirely by the deployment settings, not by
 * any code in this file. When you deploy (Deploy > New deployment):
 *   - Execute as:  Me
 *   - Who has access:  Only myself   (or "Only specific people" and list
 *     just your own Google account)
 * Google will require sign-in with YOUR account before this page loads
 * for anyone, including you on a different device. There is no separate
 * password to manage or leak.
 * -----------------------------------------------------------------------
 */

var SHEET_NAME = 'GasTracker';
var HEADERS = ['Date', 'Day', 'Cost', 'Gallons', 'Miles', 'TotalCost', 'TotalGallons', 'TotalMiles'];

/**
 * Serves the HTML page. Requires a file named "index.html" in the same
 * Apps Script project.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Gas Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Gets (or creates) the sheet, ensuring the header row exists.
 */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Formats a Date object (or passthrough string) as yyyy-MM-dd.
 */
function formatDate_(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    var tz = Session.getScriptTimeZone() || 'America/Los_Angeles';
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  return d;
}

/**
 * Returns the weekday name for a yyyy-MM-dd string, computed locally
 * (not via Date.parse) to avoid timezone off-by-one errors.
 */
function dayOfWeek_(dateStr) {
  var parts = String(dateStr).split('-');
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return days[d.getDay()];
}

/**
 * Validates an incoming entry payload from the client.
 */
function validateEntry_(entry) {
  if (!entry || !entry.date || !/^\d{4}-\d{2}-\d{2}$/.test(entry.date)) {
    throw new Error('A valid date is required.');
  }
  var cost = Number(entry.cost);
  var gallons = Number(entry.gallons);
  var miles = Number(entry.miles);
  if (isNaN(cost) || cost < 0) throw new Error('Cost must be a non-negative number.');
  if (isNaN(gallons) || gallons < 0) throw new Error('Gallons must be a non-negative number.');
  if (isNaN(miles) || miles < 0) throw new Error('Miles must be a non-negative number.');
}

/**
 * Returns every entry as an array of plain objects, in sheet row order.
 */
function getAllEntries() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    out.push({
      row: i + 2,
      date: formatDate_(r[0]),
      day: r[1],
      cost: Number(r[2]) || 0,
      gallons: Number(r[3]) || 0,
      miles: Number(r[4]) || 0,
      totalCost: Number(r[5]) || 0,
      totalGallons: Number(r[6]) || 0,
      totalMiles: Number(r[7]) || 0
    });
  }
  return out;
}

/**
 * Appends a new entry, re-sorts the whole sheet by Date so a backfilled
 * entry lands in its correct chronological spot, then recomputes totals.
 */
function addEntry(entry) {
  validateEntry_(entry);
  var sheet = getSheet_();
  var day = dayOfWeek_(entry.date);
  sheet.appendRow([entry.date, day, Number(entry.cost), Number(entry.gallons), Number(entry.miles), 0, 0, 0]);
  sortByDate_();
  return getAllEntries();
}

/**
 * Updates an existing entry's Date/Cost/Gallons/Miles (Day is recomputed),
 * then re-sorts by Date (in case the edited date moved it) and recomputes
 * totals for the whole sheet, since any change can shift every running
 * total below the affected row.
 */
function updateEntry(rowNumber, entry) {
  validateEntry_(entry);
  var sheet = getSheet_();
  rowNumber = Number(rowNumber);
  var lastRow = sheet.getLastRow();
  if (!rowNumber || rowNumber < 2 || rowNumber > lastRow) {
    throw new Error('That entry no longer exists. Please refresh.');
  }
  var day = dayOfWeek_(entry.date);
  sheet.getRange(rowNumber, 1, 1, 5).setValues([[entry.date, day, Number(entry.cost), Number(entry.gallons), Number(entry.miles)]]);
  sortByDate_();
  return getAllEntries();
}

/**
 * Sorts all data rows chronologically by Date (ascending). Entries sharing
 * the same date keep their existing relative order (stable sort), so
 * same-day fill-ups aren't shuffled. Always finishes by recomputing totals,
 * since a sort changes which row is "above" which.
 */
function sortByDate_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    // 0 or 1 data rows: nothing to reorder.
    recomputeTotals_();
    return;
  }
  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 5).getValues(); // Date, Day, Cost, Gallons, Miles
  var indexed = data.map(function (row, i) {
    return { row: row, idx: i, key: formatDate_(row[0]) };
  });
  indexed.sort(function (a, b) {
    if (a.key < b.key) return -1;
    if (a.key > b.key) return 1;
    return a.idx - b.idx; // stable tie-break: preserve original order for same-day entries
  });
  var sortedData = indexed.map(function (x) { return x.row; });
  sheet.getRange(2, 1, numRows, 5).setValues(sortedData);
  recomputeTotals_();
}

/**
 * Deletes an entry and recomputes totals for the rows that shift up.
 * No re-sort needed: removing a row can't put the remaining rows out
 * of date order.
 */
function deleteEntry(rowNumber) {
  var sheet = getSheet_();
  rowNumber = Number(rowNumber);
  var lastRow = sheet.getLastRow();
  if (!rowNumber || rowNumber < 2 || rowNumber > lastRow) {
    throw new Error('That entry no longer exists. Please refresh.');
  }
  sheet.deleteRow(rowNumber);
  recomputeTotals_();
  return getAllEntries();
}

/**
 * Recomputes TotalCost / TotalGallons / TotalMiles for every row as a
 * running sum from the top of the data down. Run after any add / edit /
 * delete so the totals always reflect current row order.
 */
function recomputeTotals_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 3, numRows, 3).getValues(); // Cost, Gallons, Miles
  var totals = [];
  var runningCost = 0, runningGallons = 0, runningMiles = 0;
  for (var i = 0; i < data.length; i++) {
    runningCost += Number(data[i][0]) || 0;
    runningGallons += Number(data[i][1]) || 0;
    runningMiles += Number(data[i][2]) || 0;
    totals.push([runningCost, runningGallons, runningMiles]);
  }
  sheet.getRange(2, 6, numRows, 3).setValues(totals);
}
