function updateSnpData() {
  const snpReturnsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('S&P500_RETURNS');
  const holidaysSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Holidays');
  const numOfRows = holidaysSheet.getRange(16, 11, 1, 1).getValue();
  const newSnpData = holidaysSheet.getRange(18, 7, numOfRows, 5).getValues();

  snpReturnsSheet.insertRowsBefore(7, numOfRows);
  snpReturnsSheet.getRange(7, 1, numOfRows, 5).setValues(newSnpData);

  const copyToRange = snpReturnsSheet.getRange(7, 8, numOfRows, 1);

  snpReturnsSheet.getRange(7 + numOfRows, 8, 1, 19).copyTo(copyToRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}
