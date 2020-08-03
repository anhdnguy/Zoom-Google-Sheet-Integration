// Using the Start Time and End Time of the meeting,
// this script will calculate the Duration every time users
// input data.
function onEdit(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  var numRows = sheet.getLastRow() - 1;
  var numColumns = sheet.getLastColumn();
  var duration_range = sheet.getRange(2, 4, numRows, 2).getValues();
  var topic_range = sheet.getRange(2, 1, numRows, 12).getValues();
  var duration_update = [];
  var topic_update = [];
  var texttosplit;
  
  for (var i = 0; i < numRows; i++) {
    if (duration_range[i][1] != 0 && duration_range[i][0] != 0) {
      duration_update[i] = [(duration_range[i][1] - duration_range[i][0]) / 60000];
    }
  }
  if (duration_update.length != 0) {
    sheet.getRange(2, 6, numRows, 1).setValues(duration_update);
  }
 }
}
