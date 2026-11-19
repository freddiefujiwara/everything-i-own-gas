// Shared helpers with no single home: which spreadsheet to open,
// sheet access, script properties, and sheet-value conversion.

const SPREADSHEET_ID = '1yxnHhfAT2cEF50hvajHoLMsW_sf7JlxnjFKeDFLNAkc';

// Environment separation (staging): when the SPREADSHEET_ID script
// property is set, it overrides the built-in production spreadsheet ID
// above. When the property is unset, behavior is identical to
// production.
const spreadsheetId_ = () =>
  PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') ||
  SPREADSHEET_ID;

// Simple fetch JSON helper
const fetchJson = url => JSON.parse(UrlFetchApp.fetch(url).getContentText());

const getSheet = name => SpreadsheetApp.openById(spreadsheetId_()).getSheetByName(name);

const scriptProperty_ = name =>
  PropertiesService.getScriptProperties().getProperty(name);

// Helper to convert 2D array from a sheet to an array of objects
const sheetValuesToObjects = values => {
  const headers = values.shift();
  return values.map(row =>
    headers.reduce((obj, header, index) => {
      obj[header] = row[index];
      return obj;
    }, {})
  );
};

export { fetchJson, getSheet, scriptProperty_, sheetValuesToObjects, spreadsheetId_ };
