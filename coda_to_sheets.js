// One-way data sync from Coda to Google Sheets using Google Apps Script
// Author: Al Chen (al@coda.io)
// Last Updated: July 16, 2020
// Notes: Assumes you are using the V8 runtime (https://developers.google.com/apps-script/guides/v8-runtime)
// Coda's library for Google Apps Script: 15IQuWOk8MqT50FDWomh57UqWGH23gjsWVWYFms3ton6L-UHmefYHS9Vl

//////////////// Setup and global variables ////////////////////////////////

CodaAPI.authenticate('YOUR_API_KEY')
SOURCE_DOC_ID = 'YOUR_SOURCE_DOC_ID'
SOURCE_TABLE_ID = 'YOUR_SOURCE_TABLE_ID'
TARGET_SHEET_ID = 'YOUR_GOOGLE_SHEETS_ID' 
TARGET_WORKSHEET_NAME = 'YOUR_GOOGLE_SHEETS_WORKSHEET_NAME'
TARGET_SHEET_SOURCE_ROW_COLUMN = 'YOUR_SOURCE_ROW_URL_COLUMN_NAME'

////////////////////////////////////////////////////////////////////////////

toSpreadsheet = SpreadsheetApp.openById(TARGET_SHEET_ID);
toWorksheet = toSpreadsheet.getSheetByName(TARGET_WORKSHEET_NAME);
headerRow = toWorksheet.getDataRange().offset(0, 0, 1).getValues()[0];
rowURLIndex = headerRow.indexOf(TARGET_SHEET_SOURCE_ROW_COLUMN);

// Run main sync functions
function runSync() {
  addDeleteToSheet();
  updateSheet();
}

// Updates existing rows in Sheet if any changes in Coda table
function updateSheet() {
  var matchingRows = [];
  var diffRowURLs = [];
  var diffRows = [];
  var allRows = prepRows();
  var sourceRows = allRows['sourceRows'];
  var targetRows = allRows['targetRows'];
  var sortedTargetRows = targetRows.sort(sortArray);
  var targetRowURLs = toWorksheet.getRange(2, rowURLIndex + 1, targetRows.length).getValues().flat();  
  var editCount = 0;

  // Find rows in Coda table only if it exists in Sheets
  sourceRows.map(function(row) {
    var rowURL = row['cells'].slice(-1)[0]['value'];
    if (targetRowURLs.indexOf(rowURL) != -1) {
      matchingRows.push(row)
    }        
  })
  var sortedMatchingRows = convertValues(sortCodaTableCols(matchingRows)).sort(sortArray)
  var numCols = sortedMatchingRows[0].length;
  
  // Create array of rows that need to be updated in Sheets
  for (var i = 0; i < sortedMatchingRows.length; i++) {
    for (var j = 0; j < numCols - 1; j++) {
      if (sortedMatchingRows[i][j] == null) {
        continue;
      }
      else if (sortedMatchingRows[i][j].length != sortedTargetRows[i][j].length) {      
        if (diffRowURLs.indexOf(sortedMatchingRows[i][rowURLIndex]) == -1) { diffRowURLs.push(sortedMatchingRows[i][rowURLIndex]); }
      } 
      else if (sortedMatchingRows[i][j] != sortedTargetRows[i][j]) {
        if (diffRowURLs.indexOf(sortedMatchingRows[i][rowURLIndex]) == -1) { diffRowURLs.push(sortedMatchingRows[i][rowURLIndex]); }
      }
    }
    
    // Get the full row from source Coda table if one of the row URLs needs updating in the Sheets file
    diffRowURLs.map(function(row) {
      if (sortedMatchingRows[i][rowURLIndex] == row) {
        diffRows.push(sortedMatchingRows[i]);
      }
    })
  } 
  
  // Update row in Sheets
  diffRows.map(function(row) {   
    diffRowIndex = targetRowURLs.indexOf(row[rowURLIndex]);
    var sheetRow = toWorksheet.getRange(diffRowIndex + 2, 1, 1, numCols).getValues().flat();
    for (var i = 0; i < numCols - 1; i++) {
      if (row[i] == null) {
        continue;
      }
      else if (row[i] != sheetRow[i]) {
        editCount++;
        toWorksheet.getRange(diffRowIndex + 2, i + 1, 1).setValue(row[i]);
      }      
    }
  })
  Logger.log('::::: %s values changed in Coda => Updating "%s" in Google Sheets...', editCount, TARGET_WORKSHEET_NAME);
  editCount = 0;
}

// Append new data from Coda table to Sheets and delete any rows from Sheets if in Coda table
function addDeleteToSheet() {
  var allRows = prepRows();
  if (allRows['targetRows'].length > 0) {
    var targetRowURLs = toWorksheet.getRange(2, rowURLIndex + 1, allRows['targetRows'].length).getValues().flat();
    var deletedRows = findDeletedRows(allRows['sourceRows'], targetRowURLs);    
  }
  else {
    targetRowURLs = [];
    deletedRows = [];
  }
  var sourceRows = findNewRows(allRows['sourceRows'], targetRowURLs);  
  
  // Add rows to Sheets only if new rows exist
  if (sourceRows.length != 0) {
    Logger.log('::::: Adding %s new rows from Coda => "%s" in Google Sheets...', sourceRows.length, TARGET_WORKSHEET_NAME);
    var sortedSourceRows = sortCodaTableCols(sourceRows)
    var convertedSourceRows = convertValues(sortedSourceRows);
    toWorksheet.getRange(toWorksheet.getLastRow() + 1, 1, convertedSourceRows.length, convertedSourceRows[0].length).setValues(convertedSourceRows)
  }
  
  // Remove deleted rows
  if (deletedRows.length != 0) {
    Logger.log('::::: %s deleted rows in Coda => Deleting these row in "%s" in Google Sheets...', deletedRows.length, TARGET_WORKSHEET_NAME);
    deletedRows.map(function(row) {
      toWorksheet.deleteRow(targetRowURLs.indexOf(row) + 2);
    })
  }
}

// Pre-processing step for retrieving/cleaning rows from source and target
function prepRows() {
  var sourceRows = retrieveRows();
  var targetRows = getSheetValues();
  targetRows.shift(); // Remove header row from Sheets range
  return {sourceRows: sourceRows, targetRows: targetRows}
}

function findDeletedRows(sourceRows, targetRowURLs) {
  var deletedRows = [];
  var sourceRowURLs = sourceRows.map(function(row) {
    return row['cells'].slice(-1)[0]['value'];
  })
  targetRowURLs.map(function(row) {
    if (sourceRowURLs.indexOf(row) == -1) {
      deletedRows.push(row)
    }
  })  
  return deletedRows;                   
}

// Finds new rows in Coda table to sync
function findNewRows(sourceRows, targetRowURLs) {
  var newRows = [];
  sourceRows.map(function(row) {
    var rowURL = row['cells'].slice(-1)[0]['value'];
    if (targetRowURLs.indexOf(rowURL) == -1) {
      newRows.push(row)
    }
  })
  return newRows;  
}

// Converts Coda table rows to 2D array of values for Sheets
function convertValues(rows) {
  var values = rows.map(function(row) {
    var rowValues = []
    row['cells'].map(function(rowValue) {
      rowValues.push(rowValue['value'])  
    })      
    return rowValues;
  })
  return values;
}

// Sort's Coda's table columns by column order in Sheet
function sortCodaTableCols(sourceRows) {
  var headerCodaTable = sourceRows[0]['cells'].map(function(row) { return row['column'] });
  var sheetsColOrder = [];
  
  headerRow.map(function(col) {
    sheetsColOrder.push(headerCodaTable.indexOf(col))
  })
 
  var sortedSourceRows = sourceRows.map(function(row) {
    var cells = sheetsColOrder.map(function(col) {
      if (col == -1) {
        return {        
          column: null,
          value: null,
        }
      } 
      else {
        return {
          column: headerCodaTable[col], 
          value: row['cells'][col]['value'],
        }       
      }
    });
    return {cells: cells}
  })
  return sortedSourceRows;
}

// Get values from Sheets
function getSheetValues() {
  var values = toWorksheet.getDataRange().getValues();
  return values;
}

// Get all Coda table rows
function retrieveRows() {
  var sourceTable = CodaAPI.getTable(SOURCE_DOC_ID, SOURCE_TABLE_ID);
  var sourceRows = [];
  var pageToken;
  var sourceColumns = CodaAPI.listColumns(SOURCE_DOC_ID, SOURCE_TABLE_ID).items.map(function(item) { return item.name; });
  
  do {
    var response = CodaAPI.listRows(SOURCE_DOC_ID, SOURCE_TABLE_ID, {limit: 500, pageToken: pageToken, useColumnNames: true, sortBy: 'natural'});
    var sourceRows = sourceRows.concat(response.items);
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  var upsertBodyRows = sourceRows.map(function(row) {
    var cells = sourceColumns.map(function(colName) {
      return {
        column: colName,
        value: row.values[colName],
      };
    });
    cells.push({column: TARGET_SHEET_SOURCE_ROW_COLUMN, value: row.browserLink});
    return {cells: cells};
  });  
  return upsertBodyRows;
}

////// Helper functions //////

function prettyPrint(value) {
  return JSON.stringify(value, null, 2);
}

function printDocTables() {
  var tables = CodaAPI.listTables(SOURCE_DOC_ID).items;
  Logger.log('Tables are: ' + prettyPrint(tables));
}

// Sort function for array of arrays which sorts Google Sheets rows by SOURCE_SHEET_SOURCE_ROW_COLUMN in alphabetical order
function sortArray(a, b) {
  var x = a[rowURLIndex];
  var y = b[rowURLIndex];
  if (x === y) {
    return 0;
  }
  else {
    return (x < y) ? -1 : 1;
  }
}