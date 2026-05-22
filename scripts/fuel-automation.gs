// Combined Script: DATA SYNC AUTOMATION + FUEL DASHBOARD LOOKUP AUTOMATION

var CONFIG = {
  sourceSpreadsheetId: "1Pol0prbO-4MZjfITGqlhOJZdypnA2rA5cMfyrCbyutg",
  sourceSheetName: "Validated FD Tickets",
  destSpreadsheetId: "1NHrBbGEPLTOvrFywTMdNKDdSoQxU1IGrFPSZZgtZ6Xo",
  destSheetName: "[nod_ph3]_fuel_delivery_ticket (48)",
  outputColumnCount: 41,
  sourceColumnCount: 33,
  headerRow: 1,
  dataStartRow: 2,
};

const PROTECTED_MAX_COL = 32; // A:AF

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  buildDataSyncMenu_(ui);
  if (typeof buildFuelMenu_ === "function") {
    try {
      buildFuelMenu_(ui);
    } catch (err) {
      Logger.log("buildFuelMenu_ failed: " + (err && err.message));
    }
  }
}

function onInstall(e) {
  onOpen(e);
}

function buildDataSyncMenu_(ui) {
  ui.createMenu("DATA SYNC AUTOMATION")
    .addItem("Sync Data", "syncValidatedTickets")
    .addItem("Show Instructions", "showDataSyncInstructions")
    .addToUi();
}

function buildFuelMenu_(ui) {
  ui.createMenu("FUEL DASHBOARD LOOKUP AUTOMATION")
    .addItem("Start Process", "startProcess")
    .addItem("Show Instructions", "showInstructions")
    .addToUi();
}

function showDataSyncInstructions() {
  var ui = SpreadsheetApp.getUi();
  var message = [
    "Please follow these steps before starting the data sync:",
    "",
    "1. Confirm that this spreadsheet is connected to the intended destination sheet.",
    "2. Verify that the source sheet 'Validated FD Tickets' is up to date.",
    "3. Review the mapped output columns before syncing, because the process overwrites rows starting from row 2.",
    "4. Make sure any manual edits you want to keep are backed up before running Sync Data.",
    "5. Run Sync Data only after the source records are finalized.",
  ].join("\n");

  ui.alert("Data Sync Instructions", message, ui.ButtonSet.OK);
}

/* ---------------- FUEL DASHBOARD LOOKUP AUTOMATION ---------------- */
function showInstructions() {
  var ui = SpreadsheetApp.getUi();
  var message = [
    "Please follow these steps before starting the process:",
    "",
    "1. Ensure that AMCEO list worksheet exists with columns: PLA ID, TERRITORY, SITE CLASS, & TOWERCO TAGGING",
    "2. Ensure that Responsible worksheet exists with columns: STATUS, RESPONSIBLE",
    "3. Ensure that main worksheet exists with columns: Region, PLA ID, Area, etc.",
    "4. Ensure that the main worksheet is updated with the latest values",
    "5. Ensure that you are currently viewing the main worksheet before starting the lookup process",
  ].join("\n");

  ui.alert("Instructions", message, ui.ButtonSet.OK);
}

function startProcess() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var mainSheet = ss.getActiveSheet();

  if (!isValidMainSheet(mainSheet)) {
    ui.alert("Unable to process. Invalid main worksheet.");
    return;
  }

  var amceoSheet = ss.getSheetByName("AMCEO LIST");
  if (!amceoSheet || !isValidAmceoSheet(amceoSheet)) {
    ui.alert("Unable to process. Invalid AMCEO list.");
    return;
  }

  var responsibleSheet = ss.getSheetByName("Responsible");
  if (!responsibleSheet || !isValidResponsibleSheet(responsibleSheet)) {
    ui.alert("Unable to process. Invalid Responsible worksheet.");
    return;
  }

  var lastRow = mainSheet.getLastRow();
  var lastCol = Math.max(mainSheet.getLastColumn(), PROTECTED_MAX_COL);
  var headerRow = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = buildHeaderMap(headerRow);

  var destCol = headerMap["DESTINATION PLA ID"];
  var statusCol = headerMap["STATUS"];

  var outputHeaders = [
    "Territory Tagging",
    "Ticket Status",
    "TAGGING",
    "PROJECT TAGGING",
    "OWNER",
    "RESPONSIBLE",
    "SITE CLASS",
  ];

  var outputColumns = ensureOutputColumns(
    mainSheet,
    headerRow,
    outputHeaders,
    PROTECTED_MAX_COL,
  );

  if (lastRow < 2) {
    return;
  }

  var amceoData = amceoSheet.getDataRange().getValues();
  var amceoHeaderMap = buildHeaderMap(amceoData[0] || []);
  var amceoIndexes = {
    plaId: amceoHeaderMap["PLA ID"],
    territory: amceoHeaderMap["TERRITORY"],
    siteClass: amceoHeaderMap["SITE CLASS"],
    tagging: amceoHeaderMap["TOWERCO TAGGING"],
  };

  var amceoLookup = {};
  for (var i = 1; i < amceoData.length; i += 1) {
    var row = amceoData[i];
    var key = normalizePlaIdKey(row[amceoIndexes.plaId]);
    if (!key) continue;
    var candidate = {
      territory: sanitizeOutput(row[amceoIndexes.territory]),
      siteClass: sanitizeOutput(row[amceoIndexes.siteClass]),
      tagging: sanitizeOutput(row[amceoIndexes.tagging]),
    };
    if (!amceoLookup[key]) {
      amceoLookup[key] = candidate;
    } else {
      amceoLookup[key] = mergeLookupRecord(amceoLookup[key], candidate);
    }
  }

  var responsibleData = responsibleSheet.getDataRange().getValues();
  var responsibleHeaderMap = buildHeaderMap(responsibleData[0] || []);
  var responsibleIndexes = {
    status: responsibleHeaderMap["STATUS"],
    responsible: responsibleHeaderMap["RESPONSIBLE"],
  };

  var responsibleLookup = {};
  for (var j = 1; j < responsibleData.length; j += 1) {
    var rrow = responsibleData[j];
    var k = normalizeKey(rrow[responsibleIndexes.status]);
    if (!k) continue;
    if (!responsibleLookup[k])
      responsibleLookup[k] = sanitizeOutput(
        rrow[responsibleIndexes.responsible],
      );
  }

  var rowCount = lastRow - 1;
  if (rowCount < 1) return;

  var mainData = mainSheet.getRange(2, 1, rowCount, lastCol).getValues();

  var outputValues = {};
  for (var h = 0; h < outputHeaders.length; h++)
    outputValues[outputHeaders[h]] = new Array(rowCount).fill("");

  var lookupIssues = [];
  var responsibleIssues = [];

  for (var r = 0; r < rowCount; r += 1) {
    var crow = mainData[r];
    var destinationKey =
      typeof destCol === "number" ? normalizePlaIdKey(crow[destCol]) : "";
    var statusValue = typeof statusCol === "number" ? crow[statusCol] : "";
    var statusKey = normalizeKey(statusValue);

    var territoryTagging = "";
    var tagging = "";
    var siteClass = "";

    if (destinationKey && amceoLookup[destinationKey]) {
      var lookup = amceoLookup[destinationKey];
      territoryTagging = lookup.territory || "";
      tagging = lookup.tagging || "";
      siteClass = lookup.siteClass || "";
    }
    if (destinationKey && !amceoLookup[destinationKey]) {
      if (lookupIssues.length < 200)
        lookupIssues.push([r + 2, crow[destCol], destinationKey, statusValue]);
    }

    var ticketStatus = computeTicketStatus(statusValue);
    var projectTagging = computeProjectTagging(tagging);
    var owner = computeOwner(projectTagging);
    var responsible =
      statusKey && responsibleLookup[statusKey]
        ? responsibleLookup[statusKey]
        : "";
    if (
      statusKey &&
      shouldExpectResponsible(statusKey) &&
      !responsibleLookup[statusKey]
    ) {
      if (responsibleIssues.length < 200)
        responsibleIssues.push([r + 2, statusValue, statusKey]);
    }

    outputValues["Territory Tagging"][r] = territoryTagging;
    outputValues["Ticket Status"][r] = ticketStatus;
    outputValues["TAGGING"][r] = tagging;
    outputValues["PROJECT TAGGING"][r] = projectTagging;
    outputValues["OWNER"][r] = owner;
    outputValues["RESPONSIBLE"][r] = responsible;
    outputValues["SITE CLASS"][r] = siteClass;
  }

  var BATCH_SIZE = 1000;
  outputHeaders.forEach(function (header) {
    var column = outputColumns[header];
    if (!column) return;
    try {
      for (var start = 0; start < rowCount; start += BATCH_SIZE) {
        var end = Math.min(start + BATCH_SIZE, rowCount);
        var slice = outputValues[header].slice(start, end);
        var values = slice.map(function (v) {
          return [sanitizeOutput(v)];
        });
        if (values.length === 0) continue;
        mainSheet
          .getRange(2 + start, column, values.length, 1)
          .setValues(values)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
      }
    } catch (e) {
      Logger.log(
        "Error writing header %s at column %s: %s",
        header,
        column,
        e && e.message,
      );
    }
  });

  if (lookupIssues.length > 0 || responsibleIssues.length > 0) {
    var reportName = "LOOKUP_ISSUES";
    var reportSheet = ss.getSheetByName(reportName);
    if (!reportSheet) reportSheet = ss.insertSheet(reportName);
    else reportSheet.clear();

    var writeRows = [];
    if (lookupIssues.length > 0) {
      writeRows.push([
        "--- Destination PLA ID mismatches (row, original, normalized, status) ---",
      ]);
      writeRows = writeRows.concat(lookupIssues);
    }
    if (responsibleIssues.length > 0) {
      writeRows.push([""]);
      writeRows.push([
        "--- STATUS mismatches where RESPONSIBLE is expected (row, original, normalized) ---",
      ]);
      writeRows = writeRows.concat(responsibleIssues);
    }

    var maxCols =
      writeRows.reduce(function (m, r) {
        return Math.max(m, (r && r.length) || 0);
      }, 0) || 1;
    var padded = writeRows.map(function (r) {
      var row = Array.isArray(r) ? r.slice() : [r];
      while (row.length < maxCols) row.push("");
      return row.map(function (c) {
        return c === undefined ? "" : c;
      });
    });
    reportSheet.getRange(1, 1, padded.length, maxCols).setValues(padded);
  }
}

/* ---------------- DATA SYNC AUTOMATION ---------------- */
function syncValidatedTickets() {
  try {
    var config = getConfig_();
    var sourceSpreadsheet = SpreadsheetApp.openById(config.sourceSpreadsheetId);
    var destSpreadsheet = SpreadsheetApp.openById(config.destSpreadsheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName(config.sourceSheetName);
    var destSheet = destSpreadsheet.getSheetByName(config.destSheetName);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Error accessing spreadsheets: " + e.message,
      "ERROR",
      10,
    );
    return { error: e.message };
  }

  if (!sourceSheet)
    throw new Error("Source sheet not found: " + config.sourceSheetName);
  if (!destSheet)
    throw new Error("Destination sheet not found: " + config.destSheetName);

  var sourceRange = sourceSheet.getDataRange();
  var sourceValues = sourceRange.getValues();
  var sourceDisplayValues = sourceRange.getDisplayValues();
  var sourceLastRow = sourceSheet.getLastRow();

  var existingRowsByGpo = buildExistingDestinationMap_(destSheet);
  var outputRows = [];
  var skippedRows = 0;

  for (
    var rowIndex = config.dataStartRow - 1;
    rowIndex < sourceValues.length;
    rowIndex++
  ) {
    var sourceRow = sourceValues[rowIndex];
    var sourceDisplayRow = sourceDisplayValues[rowIndex];
    var gpoNumber = textFromDisplay_(sourceDisplayRow[2]);
    if (!gpoNumber) {
      skippedRows++;
      continue;
    }
    var existingRow = existingRowsByGpo[gpoNumber] || null;
    outputRows.push(
      buildDestinationRow_(sourceRow, sourceDisplayRow, existingRow),
    );
  }

  clearDestinationData_(destSheet);
  if (outputRows.length > 0) {
    writeRowsInChunks_(destSheet, outputRows);
    applyDestinationFormatting_(destSheet, outputRows.length);
  }

  var summary =
    "Sync complete: " +
    outputRows.length +
    " rows written, " +
    skippedRows +
    " rows skipped.";
  SpreadsheetApp.getActiveSpreadsheet().toast(
    summary,
    "DATA SYNC",
    5,
  );
  return {
    written: outputRows.length,
    skipped: skippedRows,
    sourceRows: Math.max(0, sourceLastRow - 1),
  };
}

function getConfig_() {
  return CONFIG;
}

function buildExistingDestinationMap_(destSheet) {
  var lastRow = destSheet.getLastRow();
  if (lastRow < 2) return {};
  var range = destSheet.getRange(2, 1, lastRow - 1, CONFIG.outputColumnCount);
  var values = range.getValues();
  var map = {};
  for (var i = 0; i < values.length; i++) {
    var row = padRow_(values[i], CONFIG.outputColumnCount);
    var gpo = textFromValue_(row[8]);
    if (gpo && !map[gpo]) map[gpo] = row;
  }
  return map;
}

function buildDestinationRow_(sourceRow, sourceDisplayRow, existingRow) {
  var row = padRow_(
    existingRow ? existingRow.slice() : [],
    CONFIG.outputColumnCount,
  );
  setText_(row, 0, textFromDisplay_(sourceDisplayRow[0]));
  setText_(
    row,
    1,
    firstNonEmpty_(
      textFromDisplay_(sourceDisplayRow[12]),
      textFromDisplay_(sourceDisplayRow[30]),
    ),
  );
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
  setText_(
    row,
    20,
    firstNonEmpty_(
      textFromDisplay_(sourceDisplayRow[26]),
      textFromDisplay_(sourceDisplayRow[31]),
    ),
  );
  setText_(row, 21, textFromDisplay_(sourceDisplayRow[13]));
  setText_(row, 22, textFromDisplay_(sourceDisplayRow[20]));
  setText_(
    row,
    23,
    firstNonEmpty_(
      textFromDisplay_(sourceDisplayRow[32]),
      textFromDisplay_(sourceDisplayRow[13]),
    ),
  );
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
  if (lastRow < 2) return;
  destSheet
    .getRange(2, 1, lastRow - 1, CONFIG.outputColumnCount)
    .clearContent();
}

function writeRowsInChunks_(sheet, rows) {
  var chunkSize = 500;
  var startRow = 2;
  for (var i = 0; i < rows.length; i += chunkSize) {
    var chunk = rows.slice(i, i + chunkSize);
    sheet
      .getRange(startRow + i, 1, chunk.length, CONFIG.outputColumnCount)
      .setValues(chunk);
  }
}

function applyDestinationFormatting_(sheet, rowCount) {
  if (rowCount < 1) return;
  sheet.getRange(2, 8, rowCount, 1).setNumberFormat("MM/dd/yyyy h:mm");
  sheet.getRange(2, 21, rowCount, 1).setNumberFormat("MM/dd/yyyy h:mm");
  sheet.getRange(2, 16, rowCount, 1).setNumberFormat("0.###");
  sheet.getRange(2, 17, rowCount, 1).setNumberFormat("0.###");
  sheet.getRange(2, 26, rowCount, 1).setNumberFormat("0.###");
  sheet.getRange(2, 27, rowCount, 1).setNumberFormat("0.###");
  sheet.getRange(2, 28, rowCount, 1).setNumberFormat("0.###");
}

function setText_(row, index, value) {
  row[index] = value !== "" && value != null ? String(value) : "";
}
function setNumber_(row, index, value) {
  row[index] =
    value === "" || value == null || isNaN(value) ? "" : Number(value);
}
function setDate_(row, index, value) {
  row[index] = value instanceof Date && !isNaN(value.getTime()) ? value : "";
}

function padRow_(row, length) {
  var output = row ? row.slice(0) : [];
  while (output.length < length) output.push("");
  return output;
}
function firstNonEmpty_() {
  for (var i = 0; i < arguments.length; i++) {
    var v = arguments[i];
    if (v !== "" && v != null) return v;
  }
  return "";
}

function textFromValue_(value) {
  if (value === "" || value == null) return "";
  if (value instanceof Date)
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "MM/dd/yyyy H:mm",
    );
  return String(value).trim();
}

function textFromDisplay_(value) {
  return value === "" || value == null ? "" : String(value).trim();
}

function toNumber_(value, displayValue) {
  if (value === "" || value == null) {
    if (displayValue === "" || displayValue == null) return "";
    var parsedFromDisplay = Number(String(displayValue).replace(/,/g, ""));
    return isNaN(parsedFromDisplay) ? "" : parsedFromDisplay;
  }
  if (typeof value === "number") return value;
  if (value instanceof Date) return "";
  var parsed = Number(String(value).replace(/,/g, ""));
  return isNaN(parsed) ? "" : parsed;
}

function parseDateFromValue_(value, displayValue) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var parsed = parseDateString_(displayValue);
  if (parsed) return parsed;
  if (value && typeof value === "string") return parseDateString_(value);
  return "";
}

function parseDateString_(text) {
  if (!text) return null;
  var value = String(text).trim();
  if (!value) return null;
  var match = value.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/,
  );
  if (!match) return null;
  var month = Number(match[1]) - 1;
  var day = Number(match[2]);
  var year = Number(match[3]);
  var hour = match[4] ? Number(match[4]) : 0;
  var minute = match[5] ? Number(match[5]) : 0;
  var second = match[6] ? Number(match[6]) : 0;
  var date = new Date(year, month, day, hour, minute, second);
  return isNaN(date.getTime()) ? null : date;
}

/* ---------------- Shared helpers used by both scripts ---------------- */
function ensureOutputColumns(sheet, headerRow, outputHeaders, protectedMaxCol) {
  var headerMap = buildHeaderMap(headerRow);
  var lastCol = headerRow.length;
  var outputColumns = {};
  var headersToWrite = [];
  outputHeaders.forEach(function (header) {
    var normalized = normalizeHeader(header);
    var existingIndex = headerMap[normalized];
    if (typeof existingIndex === "number" && existingIndex >= protectedMaxCol) {
      var column = existingIndex + 1;
      outputColumns[header] = column;
      headersToWrite.push({ column: column, value: header });
      return;
    }
    lastCol += 1;
    outputColumns[header] = lastCol;
    headersToWrite.push({ column: lastCol, value: header });
  });
  var maxColumns = sheet.getMaxColumns();
  if (lastCol > maxColumns)
    sheet.insertColumnsAfter(maxColumns, lastCol - maxColumns);
  headersToWrite.forEach(function (item) {
    var cell = sheet.getRange(1, item.column);
    cell.setValue(item.value);
    cell.setBackground("#000000");
    cell.setFontColor("#ffffff");
  });
  return outputColumns;
}

function buildHeaderMap(headerRow) {
  var map = {};
  headerRow.forEach(function (value, index) {
    var normalized = normalizeHeader(value);
    if (normalized && map[normalized] === undefined) map[normalized] = index;
  });
  return map;
}

function normalizeHeader(value) {
  if (value === null || value === undefined) return "";
  var text = String(value).trim();
  if (!text) return "";
  return text.toUpperCase();
}

function normalizeKey(value) {
  if (isBlankish(value)) return "";
  var cleaned = String(value)
    .replace(/\u200B|\uFEFF|\u00A0/g, "")
    .trim();
  return cleaned.toUpperCase();
}

function normalizePlaIdKey(value) {
  var normalized = normalizeKey(value);
  if (!normalized) return "";
  var alnum = normalized.replace(/[^A-Z0-9]/g, "");
  if (!alnum) return "";
  return alnum.replace(/\d+/g, function (digits) {
    var parsed = parseInt(digits, 10);
    return isNaN(parsed) ? digits : String(parsed);
  });
}

function shouldExpectResponsible(statusKey) {
  if (!statusKey) return false;
  return (
    statusKey !== "CLOSED" &&
    statusKey !== "REJECTED" &&
    statusKey !== "CANCELLED" &&
    statusKey !== "DECLINED" &&
    statusKey !== "NO STATUS"
  );
}

function isBlankish(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === "string") {
    var trimmed = value.trim();
    if (!trimmed) return true;
    var upper = trimmed.toUpperCase();
    if (upper === "N/A" || upper === "#N/A" || upper === "NA") return true;
    if (upper.indexOf("#") === 0) return true;
    return false;
  }
  if (typeof value === "number" && isNaN(value)) return true;
  return false;
}

function sanitizeOutput(value) {
  if (isBlankish(value)) return "";
  if (typeof value === "string") return value.trim();
  return value;
}

function choosePreferredValue(currentValue, incomingValue) {
  var current = sanitizeOutput(currentValue);
  var incoming = sanitizeOutput(incomingValue);
  return !current && incoming ? incoming : current;
}

function mergeLookupRecord(existing, incoming) {
  return {
    territory: choosePreferredValue(
      existing && existing.territory,
      incoming && incoming.territory,
    ),
    siteClass: choosePreferredValue(
      existing && existing.siteClass,
      incoming && incoming.siteClass,
    ),
    tagging: choosePreferredValue(
      existing && existing.tagging,
      incoming && incoming.tagging,
    ),
  };
}

function computeTicketStatus(statusValue) {
  if (isBlankish(statusValue)) return "";
  var normalized = normalizeKey(statusValue);
  return normalized === "CLOSED" ||
    normalized === "REJECTED" ||
    normalized === "CANCELLED" ||
    normalized === "DECLINED"
    ? "CLOSED"
    : "OPEN";
}

function computeProjectTagging(taggingValue) {
  if (isBlankish(taggingValue)) return "";
  var normalized = normalizeKey(taggingValue);
  if (normalized === "BAU") return "Globe Sites";
  if (normalized === "BUILD GLOBE-ACQUIRED SITES")
    return "BUILD GLOBE-ACQUIRED SITES";
  if (normalized === "COLOCATION") return "COLOCATION";
  if (normalized === "SLB") return "SLB";
  if (normalized === "BUILD TO SUIT") return "BTS";
  return "";
}

function computeOwner(projectTaggingValue) {
  if (isBlankish(projectTaggingValue)) return "";
  var normalized = normalizeKey(projectTaggingValue);
  if (
    normalized === "GLOBE SITES" ||
    normalized === "BUILD GLOBE-ACQUIRED SITES"
  )
    return "GLOBE";
  if (
    normalized === "BTS" ||
    normalized === "COLOCATION" ||
    normalized === "SLB"
  )
    return "TOWERCO";
  return "";
}

function isValidMainSheet(sheet) {
  var headerValues = sheet.getRange(1, 1, 1, 3).getValues()[0];
  var region = String(headerValues[0] || "").trim();
  var plaId = String(headerValues[1] || "").trim();
  var area = String(headerValues[2] || "").trim();
  return region === "Region" && plaId === "PLA ID" && area === "Area";
}

function isValidAmceoSheet(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return false;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = buildHeaderMap(headers);
  var required = ["PLA ID", "TERRITORY", "SITE CLASS", "TOWERCO TAGGING"];
  return required.every(function (h) {
    return headerMap[normalizeHeader(h)] !== undefined;
  });
}

function isValidResponsibleSheet(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return false;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = buildHeaderMap(headers);
  var required = ["STATUS", "RESPONSIBLE"];
  return required.every(function (h) {
    return headerMap[normalizeHeader(h)] !== undefined;
  });
}

function testOpenSourceNoUi() {
  var id = "1Pol0prbO-4MZjfITGqlhOJZdypnA2rA5cMfyrCbyutg";
  try {
    var ss = SpreadsheetApp.openById(id);
    Logger.log("Opened: " + ss.getName());
    return ss.getName();
  } catch (e) {
    Logger.log("openById error: " + e.toString());
    return "Error: " + e.toString();
  }
}
