var CONFIG = {
  sourceSpreadsheetId: '1Pol0prbO-4MZjfITGqlhOJZdypnA2rA5cMfyrCbyutg',
  sourceSheetName: 'Validated FD Tickets',
  destSpreadsheetId: '1NHrBbGEPLTOvrFywTMdNKDdSoQxU1IGrFPSZZgtZ6Xo',
  destSheetName: '[nod_ph3]_fuel_delivery_ticket (48)',
  outputColumnCount: 41,
  sourceColumnCount: 33,
  headerRow: 1,
  dataStartRow: 2
};

function backupOnOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DATA SYNC AUTOMATION')
    .addItem('Sync Data', 'syncValidatedTickets')
    .addToUi();
}

function backupOnInstall(e) {
  backupOnOpen(e);
}

function syncValidatedTickets() {
  try {
    var config = getConfig_();
    var sourceSpreadsheet = SpreadsheetApp.openById(config.sourceSpreadsheetId);
    var destSpreadsheet = SpreadsheetApp.openById(config.destSpreadsheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName(config.sourceSheetName);
    var destSheet = destSpreadsheet.getSheetByName(config.destSheetName);
  } catch(e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error accessing spreadsheets: ' + e.message, 'ERROR', 10);
    return {error: e.message};
  }

  if (!sourceSheet) {
    throw new Error('Source sheet not found: ' + config.sourceSheetName);
  }
  if (!destSheet) {
    throw new Error('Destination sheet not found: ' + config.destSheetName);
  }

  var sourceRange = sourceSheet.getDataRange();
  var sourceValues = sourceRange.getValues();
  var sourceDisplayValues = sourceRange.getDisplayValues();
  var sourceLastRow = sourceSheet.getLastRow();

  var existingRowsByGpo = buildExistingDestinationMap_(destSheet);
  var outputRows = [];
  var skippedRows = 0;

  for (var rowIndex = config.dataStartRow - 1; rowIndex < sourceValues.length; rowIndex++) {
    var sourceRow = sourceValues[rowIndex];
    var sourceDisplayRow = sourceDisplayValues[rowIndex];
    var gpoNumber = textFromDisplay_(sourceDisplayRow[2]);

    if (!gpoNumber) {
      skippedRows++;
      continue;
    }

    var existingRow = existingRowsByGpo[gpoNumber] || null;
    outputRows.push(buildDestinationRow_(sourceRow, sourceDisplayRow, existingRow));
  }

  clearDestinationData_(destSheet);

  if (outputRows.length > 0) {
    writeRowsInChunks_(destSheet, outputRows);
    applyDestinationFormatting_(destSheet, outputRows.length);
  }

  var summary = 'Sync complete: ' + outputRows.length + ' rows written, ' + skippedRows + ' rows skipped.';
  SpreadsheetApp.getActiveSpreadsheet().toast(summary, 'DATA SYNC AUTOMATION', 5);
  return {
    written: outputRows.length,
    skipped: skippedRows,
    sourceRows: Math.max(0, sourceLastRow - 1)
  };
}

function getConfig_() {
  return CONFIG;
}

function buildExistingDestinationMap_(destSheet) {
  var lastRow = destSheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }

  var range = destSheet.getRange(2, 1, lastRow - 1, CONFIG.outputColumnCount);
  var values = range.getValues();
  var map = {};

  for (var i = 0; i < values.length; i++) {
    var row = padRow_(values[i], CONFIG.outputColumnCount);
    var gpo = textFromValue_(row[8]);
    if (gpo && !map[gpo]) {
      map[gpo] = row;
    }
  }

  return map;
}

function buildDestinationRow_(sourceRow, sourceDisplayRow, existingRow) {
  var row = padRow_(existingRow ? existingRow.slice() : [], CONFIG.outputColumnCount);

  setText_(row, 0, textFromDisplay_(sourceDisplayRow[0]));
  setText_(row, 1, firstNonEmpty_(textFromDisplay_(sourceDisplayRow[12]), textFromDisplay_(sourceDisplayRow[30])));
  setText_(row, 4, textFromDisplay_(sourceDisplayRow[9]));
  setText_(row, 5, textFromDisplay_(sourceDisplayRow[24]));
  setText_(row, 6, textFromDisplay_(sourceDisplayRow[19]));
  setDate_(row, 7, parseDateFromValue_(sourceRow[3], sourceDisplayRow[3]));
  setText_(row, 8, textFromDisplay_(sourceDisplayRow[2]));
  setText_(row, 9, textFromDisplay_(sourceDisplayRow[8]));
  setText_(row, 10, textFromDisplay_(sourceDisplayRow[1]));
  setText_(row, 11, textFromDisplay_(sourceDisplayRow[4]));
  setText_(row, 12, textFromDisplay_(sourceDisplayRow[17]));
  setText_(row, 13, textFromDisplay_(sourceDisplayRow[6]));
  setText_(row, 14, textFromDisplay_(sourceDisplayRow[5]));
  setNumber_(row, 15, toNumber_(sourceRow[14], sourceDisplayRow[14]));
  setNumber_(row, 16, toNumber_(sourceRow[18], sourceDisplayRow[18]));
  setText_(row, 17, textFromDisplay_(sourceDisplayRow[22]));
  setText_(row, 20, firstNonEmpty_(textFromDisplay_(sourceDisplayRow[26]), textFromDisplay_(sourceDisplayRow[31])));
  setText_(row, 21, textFromDisplay_(sourceDisplayRow[13]));
  setText_(row, 22, textFromDisplay_(sourceDisplayRow[20]));
  setText_(row, 23, firstNonEmpty_(textFromDisplay_(sourceDisplayRow[32]), textFromDisplay_(sourceDisplayRow[13])));
  setText_(row, 24, textFromDisplay_(sourceDisplayRow[21]));
  setNumber_(row, 25, toNumber_(sourceRow[7], sourceDisplayRow[7]));
  setNumber_(row, 26, toNumber_(sourceRow[25], sourceDisplayRow[25]));
  setText_(row, 28, textFromDisplay_(sourceDisplayRow[11]));
  setText_(row, 30, textFromDisplay_(sourceDisplayRow[15]));
  setText_(row, 31, textFromDisplay_(sourceDisplayRow[16]));
  setText_(row, 33, textFromDisplay_(sourceDisplayRow[28]));
  setText_(row, 34, textFromDisplay_(sourceDisplayRow[29]));

  return row;
}

function clearDestinationData_(destSheet) {
  var lastRow = destSheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  destSheet.getRange(2, 1, lastRow - 1, CONFIG.outputColumnCount).clearContent();
}

function writeRowsInChunks_(sheet, rows) {
  var chunkSize = 500;
  var startRow = 2;

  for (var i = 0; i < rows.length; i += chunkSize) {
    var chunk = rows.slice(i, i + chunkSize);
    sheet.getRange(startRow + i, 1, chunk.length, CONFIG.outputColumnCount).setValues(chunk);
  }
}

function applyDestinationFormatting_(sheet, rowCount) {
  if (rowCount < 1) {
    return;
  }

  sheet.getRange(2, 8, rowCount, 1).setNumberFormat('MM/dd/yyyy h:mm');
  sheet.getRange(2, 21, rowCount, 1).setNumberFormat('MM/dd/yyyy h:mm');
  sheet.getRange(2, 16, rowCount, 1).setNumberFormat('0.###');
  sheet.getRange(2, 17, rowCount, 1).setNumberFormat('0.###');
  sheet.getRange(2, 26, rowCount, 1).setNumberFormat('0.###');
  sheet.getRange(2, 27, rowCount, 1).setNumberFormat('0.###');
  sheet.getRange(2, 28, rowCount, 1).setNumberFormat('0.###');
}

function setText_(row, index, value) {
  if (value !== '' && value != null) {
    row[index] = String(value);
  } else {
    row[index] = '';
  }
}

function setNumber_(row, index, value) {
  if (value === '' || value == null || isNaN(value)) {
    row[index] = '';
  } else {
    row[index] = Number(value);
  }
}

function setDate_(row, index, value) {
  row[index] = value instanceof Date && !isNaN(value.getTime()) ? value : '';
}

function padRow_(row, length) {
  var output = row ? row.slice(0) : [];
  while (output.length < length) {
    output.push('');
  }
  return output;
}

function firstNonEmpty_() {
  for (var i = 0; i < arguments.length; i++) {
    var value = arguments[i];
    if (value !== '' && value != null) {
      return value;
    }
  }
  return '';
}

function textFromValue_(value) {
  if (value === '' || value == null) {
    return '';
  }
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'MM/dd/yyyy H:mm');
  }
  return String(value).trim();
}

function textFromDisplay_(value) {
  if (value === '' || value == null) {
    return '';
  }
  return String(value).trim();
}

function toNumber_(value, displayValue) {
  if (value === '' || value == null) {
    if (displayValue === '' || displayValue == null) {
      return '';
    }
    var parsedFromDisplay = Number(String(displayValue).replace(/,/g, ''));
    return isNaN(parsedFromDisplay) ? '' : parsedFromDisplay;
  }

  if (typeof value === 'number') {
    return value;
  }

  if (value instanceof Date) {
    return '';
  }

  var parsed = Number(String(value).replace(/,/g, ''));
  return isNaN(parsed) ? '' : parsed;
}

function parseDateFromValue_(value, displayValue) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  var parsed = parseDateString_(displayValue);
  if (parsed) {
    return parsed;
  }

  if (value && typeof value === 'string') {
    return parseDateString_(value);
  }

  return '';
}

function parseDateString_(text) {
  if (!text) {
    return null;
  }

  var value = String(text).trim();
  if (!value) {
    return null;
  }

  var match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!match) {
    return null;
  }

  var month = Number(match[1]) - 1;
  var day = Number(match[2]);
  var year = Number(match[3]);
  var hour = match[4] ? Number(match[4]) : 0;
  var minute = match[5] ? Number(match[5]) : 0;
  var second = match[6] ? Number(match[6]) : 0;

  var date = new Date(year, month, day, hour, minute, second);
  return isNaN(date.getTime()) ? null : date;
}
