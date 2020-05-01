// Setup and global variables
CodaAPI.authenticate('TO UPDATE');
SOURCE_DOC_ID = 'TO UPDATE'
TARGET_DOC_ID = 'TO UPDATE' 
TARGET_TABLE_SOURCE_ROW_COLUMN = 'Source Row URL';

/////////////////////// Define tables to sync between source and target docs below //////////////////////////////////

var TABLES = [  
  
  //1st table to sync
  [
    {
      doc: SOURCE_DOC_ID,
      table: 'TO UPDATE', //1st table from source doc
    },
    {
      doc: TARGET_DOC_ID,
      table: 'TO UPDATE', //1st table from target doc
    }
  ],
  
  //2nd table to sync
  [
    {
      doc: SOURCE_DOC_ID,
      table: 'TO UPDATE', //2nd table from source doc
    },
    {
      doc: TARGET_DOC_ID,
      table: 'TO UPDATE', //2nd table from target doc
    }
  ]

];

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function oneWaySync() {
  for each (var table in TABLES) {      //If you are using the V8 runtime, delete the word "each" in this line
    syncSpecificTable(table[0], table[1]);
  }
}

function syncSpecificTable(source, target) {  
  var sourceTable = CodaAPI.getTable(source.doc, source.table);
  var targetTable = CodaAPI.getTable(target.doc, target.table);
  Logger.log('::::: Syncing "%s" => "%s"...', sourceTable.name, targetTable.name);

  var sourceColumns = CodaAPI.listColumns(source.doc, source.table).items.map(function(item) { return item.name; });
  var targetColumns = CodaAPI.listColumns(target.doc, target.table).items.map(function(item) { return item.name; });
  var commonColumns = intersection(sourceColumns, targetColumns);
  Logger.log('Syncing columns: %s', commonColumns.join(', '));
  
  var sourceRows = CodaAPI.listRows(source.doc, source.table, {limit: 500, useColumnNames: true}).items;
  Logger.log('Source table has %s rows', sourceRows.length);

  var sourceSourceRowURLs = [];
  var upsertBodyRows = sourceRows.map(function(row) {
    var cells = commonColumns.map(function(colName) {
      return {
        column: colName,
        value: row.values[colName],
      };
    });

    cells.push({column: TARGET_TABLE_SOURCE_ROW_COLUMN, value: row.browserLink});
    sourceSourceRowURLs.push(row.browserLink);
    return {cells: cells};
  });
  
  var targetRows = CodaAPI.listRows(target.doc, target.table, {limit: 500, useColumnNames: true}).items;
  Logger.log('Target table has %s rows', targetRows.length);
  
  targetRows.map(function(row) {
    if (sourceSourceRowURLs.indexOf(row.values[TARGET_TABLE_SOURCE_ROW_COLUMN]) == -1) {
      CodaAPI.deleteRow(TARGET_DOC_ID, target.table, row['id']);
    }
  });     
  
  CodaAPI.upsertRows(target.doc, target.table, {rows: upsertBodyRows, keyColumns: [TARGET_TABLE_SOURCE_ROW_COLUMN]});
  Logger.log('Updated %s!', targetTable.name);
}

//Helper functions
function intersection(a, b) {
  var result = [];
  for each (var x in a) {      //If you are using the V8 runtime, delete the word "each" in this line
    if (b.indexOf(x) !== -1) {
      result.push(x);
    }
  }
  return result;
}

function prettyPrint(value) {
  return JSON.stringify(value, null, 2);
}

function printDocTables() {
  var tables = CodaAPI.listTables(TARGET_DOC_ID).items;
  Logger.log('Tables are: ' + prettyPrint(tables));
}