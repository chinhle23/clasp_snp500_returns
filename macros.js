/** @OnlyCurrentDoc */

function addRows() {
  var spreadsheet = SpreadsheetApp.getActive();
  var numberOfRows = spreadsheet.getRange('A2').getValue();
  spreadsheet.getRange('7:8').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), numberOfRows);
  spreadsheet.getActiveRange().offset(0, 0, numberOfRows, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A2').activate();
};