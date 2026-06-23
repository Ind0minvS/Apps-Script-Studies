/**
 * Web app that reads an external spreadsheet, lists its sheets as left-side
 * tabs, and shows each sheet's non-empty data fields when a tab is clicked.
 */

// The spreadsheet to read from (from the URL the user provided).
var SPREADSHEET_ID = '1Y0ZH_q_lvVS5AyvZL7KVBPEmFS7tR3IU0Ezw1lVo7bo';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Visualizador de Planilha')
    // REQUIRED so Google Sites can embed this in an iframe:
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Pull a sub-file's content into the template: <?!= include('Styles') ?> */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Return the list of sheet (tab) names in the spreadsheet.
 * @return {string[]}
 */
function getSheetNames() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheets().map(function (sh) { return sh.getName(); });
}

/**
 * Return the data of a single sheet, stripped of fully-empty rows/columns.
 * @param {string} name  the sheet/tab name
 * @return {{header: string[], rows: string[][]}}
 */
function getSheetData(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error('Aba não encontrada: ' + name);
  }

  // One bulk read of the used range (fast, quota-friendly).
  var values = sh.getDataRange().getDisplayValues();
  if (!values.length) {
    return { header: [], rows: [] };
  }

  // Find which columns have at least one non-empty cell.
  var width = values[0].length;
  var keepCol = [];
  for (var c = 0; c < width; c++) {
    var hasData = false;
    for (var r = 0; r < values.length; r++) {
      if (String(values[r][c]).trim() !== '') { hasData = true; break; }
    }
    keepCol.push(hasData);
  }

  // Keep only rows that have at least one non-empty cell (within kept columns).
  function filterCols(row) {
    return row.filter(function (_, c) { return keepCol[c]; });
  }

  var nonEmptyRows = values.filter(function (row) {
    return row.some(function (cell, c) {
      return keepCol[c] && String(cell).trim() !== '';
    });
  });

  if (!nonEmptyRows.length) {
    return { header: [], rows: [] };
  }

  var header = filterCols(nonEmptyRows[0]);
  var rows = nonEmptyRows.slice(1).map(filterCols);

  return { header: header, rows: rows };
}
/**
 * Web app that reads an external spreadsheet, lists its sheets as left-side
 * tabs, and shows each sheet's non-empty data fields when a tab is clicked.
 */

// The spreadsheet to read from (from the URL the user provided).
var SPREADSHEET_ID = '1Y0ZH_q_lvVS5AyvZL7KVBPEmFS7tR3IU0Ezw1lVo7bo';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate() 
    .setTitle('Visualizador de Planilha')
    // REQUIRED so Google Sites can embed this in an iframe:
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Pull a sub-file's content into the template: <?!= include('Styles') ?> */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Return the list of sheet (tab) names in the spreadsheet.
 * @return {string[]}
 */
function getSheetNames() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheets().map(function (sh) { return sh.getName(); });
}

/**
 * Return the data of a single sheet plus its merged-cell layout, so the client
 * can reproduce merges (merged cells come back with the value only in the
 * top-left cell and empties in the rest).
 * @param {string} name  the sheet/tab name
 * @return {{rows: string[][], merges: Object[], numCols: number}}
 */
function getSheetData(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error('Aba não encontrada: ' + name);
  }

  // One bulk read of the used range (fast, quota-friendly).
  var range = sh.getDataRange();
  var values = range.getDisplayValues();
  if (!values.length || !values[0].length) {
    return { rows: [], merges: [], numCols: 0 };
  }

  var startRow = range.getRow();     // 1-based top of the data range
  var startCol = range.getColumn();  // 1-based left of the data range
  var numCols = values[0].length;

  // Merged ranges within the data range, normalized to 0-based offsets so the
  // client can place colspan/rowspan on the right anchor cell.
  var merges = range.getMergedRanges().map(function (m) {
    return {
      row: m.getRow() - startRow,
      col: m.getColumn() - startCol,
      numRows: m.getNumRows(),
      numCols: m.getNumColumns()
    };
  });

  return { rows: values, merges: merges, numCols: numCols };
}