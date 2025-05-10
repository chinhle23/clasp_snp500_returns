/** @OnlyCurrentDoc */

function addRows() {
  var spreadsheet = SpreadsheetApp.getActive();
  var numberOfRows = spreadsheet.getRange('A2').getValue();
  spreadsheet.getRange('7:8').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), numberOfRows);
  spreadsheet.getActiveRange().offset(0, 0, numberOfRows, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A2').activate();
};

function CopyAndPasteFormulas() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H29').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('H7:H29').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('H29'));
  spreadsheet.getRange('H29:Z29').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};