// One-way data sync from Coda to Google Sheets using Google Apps Script
// Author: Al Chen (al@coda.io)
// Last Updated: March 6th, 2024
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
  var allRows = prepRows();
  addDeleteToSheet(allRows);
  updateSheet(allRows);
}

// Updates existing rows in Sheet if any changes in Coda table
function updateSheet(allRows) {
  let sourceRows = convertValues(sortCodaTableCols(allRows['sourceRows']));
  let targetRows = allRows['targetRows'];
  let numCols = targetRows[0].length;
  let editCount = 0;
  
  for (let i = 0; i < targetRows.length; i++) {
    let targetRow = targetRows[i];
    let rowUrl = targetRow[rowURLIndex];
    let sourceRow = sourceRows.find(row => row[rowURLIndex] == rowUrl);
    if (!sourceRow) {
      continue;
    }
    for (let j = 0; j < numCols - 1; j++) {
      let sourceValue = sourceRow[j];
      let targetValue = targetRow[j];
      if (sourceValue != null && sourceValue != targetValue) {
        editCount++;
        toWorksheet.getRange(i + 2, j + 1, 1).setValue(sourceValue);
      }
    }
  }
  Logger.log('::::: %s values changed in Coda => Updating "%s" in Google Sheets...', editCount, TARGET_WORKSHEET_NAME);
}

// Append new data from Coda table to Sheets and delete any rows from Sheets if in Coda table
function addDeleteToSheet(allRows) {
  if (allRows['targetRows'].length > 0) {
    var targetRowURLs = allRows['targetRows'].map(row => row[rowURLIndex]);
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
  var sourceColumns = getAllColumns().map(function(column) { return column.name; });

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

function getAllColumns() {
  let result = [];
  let pageToken;
  do {
    let page = CodaAPI.listColumns(SOURCE_DOC_ID, SOURCE_TABLE_ID, {
      pageToken: pageToken,
      limit: 100,
    });
    result = result.concat(page.items);
    pageToken = page.nextPageToken;
  } while (pageToken);
  return result;
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
