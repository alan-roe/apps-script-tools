import Properties from './properties.js';

/**
 * Sets a value in a specific column for a row identified by ID
 * @param {string} sheetName - Name of the sheet
 * @param {string} id - ID to find the row
 * @param {number} col - Column number to set value
 * @param {string} value - Value to set
 * @throws {Error} If row with ID not found
 */
export function setRowValue(sheetName, id, col, value) {
  let sheet = getSheet(sheetName);
  let row = sheet.createTextFinder(id).findNext();
  if (!row) throw new Error(`Row ${id} not found in sheet ${sheetName}`);

  sheet.getRange(row.getRow(), col, 1, 1).setValue(value);
}

/**
 * Gets a row by ID
 * @param {string} sheetName - Name of the sheet
 * @param {string} id - ID to find the row
 * @returns {any[]} - The row containing strings or Date objects
 */
export function getRowByID(sheetName, id) {
  const sheet = getSheet(sheetName);
  const data = sheet
    .getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  return data.find(row => row[0] === id);
}

let _spreadSheet = null;

/**
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} - The spreadsheet
 */
function getSpreadSheet() {
  if (!_spreadSheet) {
    _spreadSheet = SpreadsheetApp.openById(Properties.get('spreadSheetId'));

    if (!_spreadSheet) {
      throw new Error('Failed to open spreadsheet');
    }
  }

  return _spreadSheet;
}

function getSheet(name) {
  let sheet = getSpreadSheet().getSheetByName(name);

  if (!sheet) {
    sheet = getSpreadSheet().insertSheet(name);
  }

  return sheet;
}

/**
 * @param {string} sheetName - Name of the sheet
 * @param {string} columnName - Name of the column
 * @param {string} id - ID to find the row
 * @returns {any} - The cell value
 */
function getCell(sheetName, columnName, id) {
  const sheet = getSheet(sheetName);
  let schema = loadSchema(sheetName);
  const row = sheet.createTextFinder(id).findNext();
  if (!row) throw new Error(`Row ${id} not found in sheet ${sheetName}`);
  Logger.log(
    `Getting cell ${columnName} in sheet ${sheetName} with id ${id}\n${JSON.stringify(schema)}`
  );
  return sheet.getRange(row.getRow(), schema[columnName] + 1).getValue(); // +1 because schema is 0-indexed
}

let _schemas = {};

/**
 * Loads the schema for specified sheet from the spreadsheet
 * @param {string} sheetName - Name of the sheet
 * @returns {Object} - The schema
 */
export function loadSchema(sheetName) {
  if (_schemas[sheetName]) {
    return _schemas[sheetName];
  }

  const sheet = getSheet(sheetName);
  const schema = {};
  sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .forEach((column, index) => {
      schema[column] = index;
    });
  if (Object.keys(schema).length == 0) {
    throw new Error(`No columns found in sheet ${sheetName}`);
  }
  Logger.log(`Loaded schema for sheet ${sheetName}: ${JSON.stringify(schema)}`);
  _schemas[sheetName] = schema;
  return schema;
}
