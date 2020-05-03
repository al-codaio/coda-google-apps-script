// One-way data sync from Google Sheets to Coda using Google Apps Script
// Author: Al Chen (al@coda.io)
// Last Updated: April 30, 2020
// Notes: Assumes you are using the V8 runtime (https://developers.google.com/apps-script/guides/v8-runtime)
// Coda's library for Google Apps Script: 15IQuWOk8MqT50FDWomh57UqWGH23gjsWVWYFms3ton6L-UHmefYHS9Vl

//////////////// Setup and global variables ////////////////////////////////

CodaAPI.authenticate('YOUR_API_KEY')

// Coda settings
TARGET_DOC_ID = 'YOUR_TARGET_DOC_ID'
TARGET_TABLE_ID = 'YOUR_TARGET_TABLE_ID'
TARGET_ROW_ID_COLUMN = 'Coda Row ID' // You cannot have a column in your Coda table with this name

// Sheets Settings
SOURCE_SHEET_ID = 'YOUR_SOURCE_GOOGLE_SHEETS_ID'
SOURCE_WORKSHEET_NAME = 'YOUR_SOURCE_GOOGLE_SHEETS_WORKSHEET_NAME'
SOURCE_SHEET_SOURCE_ROW_COLUMN = 'YOUR_SOURCE_ROW_URL_COLUMN_NAME'  // Column name in Google Sheets to store source row URLs from Coda

// Optional Settings
DO_NOT_DELETE_COLUMN = 'Do not delete' // If you want to add rows directly to Coda table without rows getting deleted on sync. Must be a checkbox column in your Coda table and cannot exist in Google Sheets.
REWRITE_CODA_TABLE = false // Set as true if you want the sync to completely overwrite the Coda table each sync
 
////////////////////////////////////////////////////////////////////////////

fromSpreadsheet = SpreadsheetApp.openById(SOURCE_SHEET_ID);
fromWorksheet = fromSpreadsheet.getSheetByName(SOURCE_WORKSHEET_NAME);
sheetsHeaderRow = fromWorksheet.getDataRange().offset(0, 0, 1).getValues()[0];
codaHeaderRow = CodaAPI.listColumns(TARGET_DOC_ID, TARGET_TABLE_ID).items;
codaTableName = CodaAPI.getTable(TARGET_DOC_ID, TARGET_TABLE_ID).name;
rowURLIndex = sheetsHeaderRow.indexOf(SOURCE_SHEET_SOURCE_ROW_COLUMN);

// Run main sync functions
function runSync() {
  sheetsPermissions();
  addDeleteToCoda();
  if (!REWRITE_CODA_TABLE) {
    updateCoda();
  }
}

// Determine if you have edit or view access to to the Google Sheet
function sheetsPermissions() {
  try {
    fromSpreadsheet.addEditor(Session.getActiveUser());
  } 
  catch (e) {
    REWRITE_CODA_TABLE = true; // If no access automatically rewrite Coda tables each sync
  }
}

// Updates existing rows in Coda if any changes in Sheets
function updateCoda() {
  var matchingRows = [];
  var diffRowURLs = [];
  var diffRows = [];
  var cleanedSourceRows = [];
  var allRows = prepRows();
  var sortedSourceRows = allRows['sourceRows'].sort(sortArray);
  var targetRows = allRows['targetRows'];

  // Remove rows with empty source URLs in sortedSourceRows
  for (var i = 0; i < sortedSourceRows.length; i++) {
    if (sortedSourceRows[i][rowURLIndex].length != 0) {
      cleanedSourceRows.push(sortedSourceRows[i]);
    }
  }
  
  // Get relevant target rows with columns that match Sheets columns
  var cleanedTargetRows = targetRows.map(function(row) {
    var cells = row['cells'].map(function(cell) {
      if (sheetsHeaderRow.indexOf(cell['column']) != -1) {
        return {
          column: cell['column'],
          value:  cell['value'],
        }  
      }    
    })
    return { //filter out undefined cells
      cells: cells.filter(function(x) {
        return x !== undefined;    
      })
    }
  })
    
  // Find rows in Sheets only if it exists in Coda
  cleanedSourceRows.map(function(row) {
    var rowURL = row[rowURLIndex];
    cleanedTargetRows.map(function(targetRow) {
      if (targetRow['cells'].slice(-1)[0]['value'] == rowURL) {
        matchingRows.push(targetRow)
      }            
    })
  })
  
  var convertedMatchingRows = convertValuesForSheets(matchingRows);  
  
  // Create array of rows that need to be updated in Coda
  for (var i = 0; i < cleanedSourceRows.length; i++) {
    sourceRowURL = cleanedSourceRows[i][rowURLIndex]; 
    for (var j = 0; j < cleanedSourceRows[0].length; j++) {
      if (cleanedSourceRows[i][j].length != convertedMatchingRows[i][j].length) { 
        if(diffRowURLs.indexOf(sourceRowURL) == -1) { diffRowURLs.push(sourceRowURL); } 
      }
      else if (cleanedSourceRows[i][j] != convertedMatchingRows[i][j]) { 
        if(diffRowURLs.indexOf(sourceRowURL) == -1) { diffRowURLs.push(sourceRowURL); }
      }    
    }    
  }
   
  // Get full rows from source that have changes
  cleanedSourceRows.map(function(row) {
    if (diffRowURLs.indexOf(row[rowURLIndex]) != -1) {
      diffRows.push(row);
    }
  }) 
  
  // Get row IDs from target and splice into diffRows
  targetRows.map(function(targetRow) {
    diffRows.map(function(diffRow) {
      if (diffRow[rowURLIndex].indexOf(targetRow['cells'].slice(-1)[0]['value']) != -1) {
        diffRow.splice(-1, 0, targetRow['cells'].slice(-2)[0]['value'])
      }
    })
  })
  
  // Convert diffRows into format for Coda
  var updateTargetRows = sortSheetsTableCols(diffRows, true);
  
  // Update row in Coda
  updateTargetRows.map(function(row) {
    var body = {
      row: { cells: row['cells'] }
    }
    CodaAPI.updateRow(TARGET_DOC_ID, TARGET_TABLE_ID, row['rowID'][0], body);
  })
  Logger.log('::::: %s values changed in Gogole Sheets => Updating "%s" in Coda...', updateTargetRows.length, codaTableName);
}

// Append new data from Sheets to a Coda table and delete any rows from the Coda table if in Google Sheets
function addDeleteToCoda() {
  var allRows = prepRows();
  
  if (REWRITE_CODA_TABLE) {
    deleteAllTargetRows(allRows['targetRows']);
    var sourceRows = allRows['sourceRows'];
  }
  else {
    if (allRows['targetRows'].length > 0) {
      var targetRowURLs = getTargetRowValues(allRows['targetRows'], -1);
      var deletedRows = findDeletedRows(allRows['sourceRows'], allRows['targetRows']);    
    }
    else {
      targetRowURLs = [];
      deletedRows = [];
    }
    var sourceRows = findNewRows(allRows['sourceRows'], targetRowURLs);
  }
  
  // Add rows to Coda only if new rows exist
  if (sourceRows.length != 0) {
    Logger.log('::::: Adding %s new rows from Google Sheets => "%s" in Coda...', sourceRows.length, codaTableName);
    var sortedSourceRows = sortSheetsTableCols(sourceRows)
    var timer = 0;
    CodaAPI.upsertRows(TARGET_DOC_ID, TARGET_TABLE_ID, {rows: sortedSourceRows});
    
    if (!REWRITE_CODA_TABLE) {
      // Following three commands to see if new rows have propagated to Coda table and therefore new row URLs exist (will be repeated in time delay loop).
      var currentCodaRows = retrieveRows();
      var currentTargetRowURLs = getTargetRowValues(currentCodaRows, -1);
      var newSourceRows = findNewRows(sourceRows, currentTargetRowURLs);
    
      // Time delay to wait for Coda API to propagate rows into Coda. Times out after 30 seconds.
      while(currentCodaRows.length <= allRows['targetRows'].length) {
        timer += 2;
        if (timer == 60) { break; }
        Utilities.sleep(2000);
        currentCodaRows = retrieveRows(); 
        currentTargetRowURLs = getTargetRowValues(currentCodaRows, -1);
        newSourceRows = findNewRows(sourceRows, currentTargetRowURLs);      
      }
       
      //Write new Source Row URLs to Google Sheets
      var rowURLs = [];
      currentCodaRows.slice(0, sourceRows.length).map(function(row) {  // Only look at the number of new rows in the Sheet
        rowURLs.push(row['cells'].slice(-1)[0]['value']);
      })                                                                  
      fromWorksheet.getRange(allRows['sourceRows'].length - sourceRows.length + 2, rowURLIndex + 1, rowURLs.length, 1).setValues(convertValues(rowURLs));  
    }
  }
  
  if (!REWRITE_CODA_TABLE) {
    // Remove deleted rows
    if (deletedRows.length != 0) {
      Logger.log('::::: %s deleted rows in Sheets => Deleting these row in "%s" in Coda...', deletedRows.length, codaTableName);
      var body = {
        'rowIds': deletedRows,
      };
      CodaAPI.deleteRows(TARGET_DOC_ID, TARGET_TABLE_ID, body);
    }
  }
}

// Delete all rows in target
function deleteAllTargetRows(targetRows) {
  var rowIDs = [];
  targetRows.map(function(row){
    rowIDs.push(row['cells'].slice(-2)[0]['value']);    
  })  
  var body = { 'rowIds': rowIDs, };
  CodaAPI.deleteRows(TARGET_DOC_ID, TARGET_TABLE_ID, body);
}

// Pre-processing step for retrieving/cleaning rows from source and target
function prepRows() {      
  var targetRows = retrieveRows();
  var sourceRows = getSheetValues();
  sourceRows.shift(); // Remove header row from Sheets range
  return {targetRows: targetRows, sourceRows: sourceRows}
}

function findDeletedRows(sourceRows, targetRows) {
  var deletedRowIDs = [];
  var sourceRowURLs = [];
  var deleteColNum;
  sourceRows.map(function(row) { 
    sourceRowURLs.push(row[rowURLIndex]); 
  });
  for (var i = 0; i < codaHeaderRow.length; i++) {
    if (codaHeaderRow[i]['name'] == DO_NOT_DELETE_COLUMN) {
      deleteColNum = i;
    }
  }
  targetRows.map(function(row) {
    if (deleteColNum == null) {
      if (sourceRowURLs.indexOf(row['cells'].slice(-1)[0]['value']) == -1) {
        deletedRowIDs.push(row['cells'].slice(-2)[0]['value'])
      }
    } 
    else {
      if (sourceRowURLs.indexOf(row['cells'].slice(-1)[0]['value']) == -1 && row['cells'][deleteColNum]['value'] != true) {
        deletedRowIDs.push(row['cells'].slice(-2)[0]['value'])
      }      
    }
  })  
  return deletedRowIDs;                   
}

// Finds new rows in Sheets table to sync
function findNewRows(sourceRows, targetRowURLs) {
  var newRows = [];
  sourceRows.map(function(row) {
    var rowURL = row[rowURLIndex];
    if (targetRowURLs.indexOf(rowURL) == -1) {
      newRows.push(row)
    }
  })
  return newRows;  
}

// Finds a value in target Coda table based on number of columns from the end of the table (last two columns are TARGET_ROW_ID_COLUMN, SOURCE_SHEET_SOURCE_ROW_COLUMN).
function getTargetRowValues(targetRows, indexFromEnd) {
  var newRowURLs = [];
  targetRows.map(function(row) {
    newRowURLs.push(row['cells'].slice(indexFromEnd)[0]['value']);
  })
  return newRowURLs;
}

// Converts array of values to 2D array of values for Sheets
function convertValues(values) {
  var x = [];
  for (var i = 0; i < values.length; i++) {
    x[i] = [values[i]];
  }  
  return x;
}

// Converts Coda table rows to 2D array of values for Sheets in same column order as Sheets
function convertValuesForSheets(rows) {  
  // Get order of columns from Sheets
  var sheetsColOrder = [];
  rows[0]['cells'].map(function(cell) {
    sheetsColOrder.push(sheetsHeaderRow.indexOf(cell['column']));    
  }) 
  
  var values = rows.map(function(row) {
    var rowValues = [];
    for (var i = 0; i < sheetsColOrder.length; i++) {
      rowValues.push(row['cells'][sheetsColOrder[i]]['value']);
    }
    return rowValues;
  })
  return values;
}

// Sort's Sheets' table columns by column order in Coda
function sortSheetsTableCols(sourceRows, opt_rowIDs) {
  var cleanedCodaColOrder = cleanCodaColumns();
  var sortedSourceRows = sourceRows.map(function(row) {
    var cells = cleanedCodaColOrder.map(function(col) {
      return {
        column: col['colId'],
        value:  row[col['colNum']],
      }     
    });
    return {
      cells: cells,
      rowID: row.slice(-2, -1),
    }    
  })
  return sortedSourceRows;
}

// Remove any columns to sync that exist in Coda table but not in Sheets 
function cleanCodaColumns() {  
  var cleanedCodaColOrder = [];
  var codaColOrder = codaHeaderRow.map(function(col) {
    return {
      colNum: sheetsHeaderRow.indexOf(col['name']),
      colId:  col['id'],
    }
  })
  
  for (var i = 0; i < codaColOrder.length; i++) {
    if (codaColOrder[i]['colNum'] != -1) {
      cleanedCodaColOrder.push(codaColOrder[i]);
    }
  }
  return cleanedCodaColOrder
}

// Get values from Sheets
function getSheetValues() {
  var values = fromWorksheet.getDataRange().getValues();
  return values;
}

// Get all Coda table rows
function retrieveRows() {
  var targetTable = CodaAPI.getTable(TARGET_DOC_ID, TARGET_TABLE_ID);
  var targetRows = [];
  var pageToken;
  var targetColumns = CodaAPI.listColumns(TARGET_DOC_ID, TARGET_TABLE_ID).items.map(function(item) { return item.name; });
  
  do {
    var response = CodaAPI.listRows(TARGET_DOC_ID, TARGET_TABLE_ID, {limit: 500, pageToken: pageToken, useColumnNames: true});
    var targetRows = targetRows.concat(response.items);
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  var upsertBodyRows = targetRows.map(function(row) {
    var cells = targetColumns.map(function(colName) {
      return {
        column: colName,
        value: row.values[colName],
      };
    });
    cells.push({column: TARGET_ROW_ID_COLUMN, value: row.id});
    cells.push({column: SOURCE_SHEET_SOURCE_ROW_COLUMN, value: row.browserLink});
    return {cells: cells};
  });  
  return upsertBodyRows;
}

////// Helper functions //////

function prettyPrint(value) {
  return JSON.stringify(value, null, 2);
}

function printDocTables() {
  var tables = CodaAPI.listTables(TARGET_DOC_ID).items;
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