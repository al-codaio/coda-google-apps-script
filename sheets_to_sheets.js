var sourceSpreadsheetID = "TO UPDATE";
var sourceWorksheetName = "TO UPDATE";
var targetSpreadsheetID = "TO UPDATE";
var targetWorksheetName = "TO UPDATE";

function importData() {
  var thisSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetID);
  var thisWorksheet = thisSpreadsheet.getSheetByName(sourceWorksheetName);
  var thisData = thisWorksheet.getDataRange();
  //Uncomment line 11 below and comment out line 9 if you want to sync a named range. Replace "teamBugs" with your named range.
  //var thisData = thisSpreadsheet.getRangeByName("teamBugs");

  var toSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetID);
  var toWorksheet = toSpreadsheet.getSheetByName(targetWorksheetName);
  var toRange = toWorksheet.getRange(1, 1, thisData.getNumRows(), thisData.getNumColumns())
  toRange.setValues(thisData.getValues()); 
}