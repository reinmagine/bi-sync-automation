const PROTECTED_MAX_COL = 32; // A:AF

function backupOnOpen() {
  SpreadsheetApp.getUi()
    .createMenu("FUEL DASHBOARD LOOKUP AUTOMATION")
    .addItem("Start Process", "startProcess")
    .addItem("Show Instructions", "showInstructions")
    .addToUi();
}

function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  const message = [
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
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const mainSheet = ss.getActiveSheet();

  if (!isValidMainSheet(mainSheet)) {
    ui.alert("Unable to process. Invalid main worksheet.");
    return;
  }

  const amceoSheet = ss.getSheetByName("AMCEO LIST");
  if (!amceoSheet || !isValidAmceoSheet(amceoSheet)) {
    ui.alert("Unable to process. Invalid AMCEO list.");
    return;
  }

  const responsibleSheet = ss.getSheetByName("Responsible");
  if (!responsibleSheet || !isValidResponsibleSheet(responsibleSheet)) {
    ui.alert("Unable to process. Invalid Responsible worksheet.");
    return;
  }

  const lastRow = mainSheet.getLastRow();
  const lastCol = Math.max(mainSheet.getLastColumn(), PROTECTED_MAX_COL);
  const headerRow = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = buildHeaderMap(headerRow);

  const destCol = headerMap["DESTINATION PLA ID"];
  const statusCol = headerMap["STATUS"];

  const outputHeaders = [
    "Territory Tagging",
    "Ticket Status",
    "TAGGING",
    "PROJECT TAGGING",
    "OWNER",
    "RESPONSIBLE",
    "SITE CLASS",
  ];

  const outputColumns = ensureOutputColumns(
    mainSheet,
    headerRow,
    outputHeaders,
    PROTECTED_MAX_COL,
  );

  if (lastRow < 2) {
    return;
  }

  const amceoData = amceoSheet.getDataRange().getValues();
  const amceoHeaderMap = buildHeaderMap(amceoData[0] || []);
  const amceoIndexes = {
    plaId: amceoHeaderMap["PLA ID"],
    territory: amceoHeaderMap["TERRITORY"],
    siteClass: amceoHeaderMap["SITE CLASS"],
    tagging: amceoHeaderMap["TOWERCO TAGGING"],
  };

  const amceoLookup = {};
  for (let i = 1; i < amceoData.length; i += 1) {
    const row = amceoData[i];
    const key = normalizePlaIdKey(row[amceoIndexes.plaId]);
    if (!key) {
      continue;
    }
    const candidate = {
      territory: sanitizeOutput(row[amceoIndexes.territory]),
      siteClass: sanitizeOutput(row[amceoIndexes.siteClass]),
      tagging: sanitizeOutput(row[amceoIndexes.tagging]),
    };
    if (!amceoLookup[key]) {
      amceoLookup[key] = candidate;
      continue;
    }
    amceoLookup[key] = mergeLookupRecord(amceoLookup[key], candidate);
  }

  const responsibleData = responsibleSheet.getDataRange().getValues();
  const responsibleHeaderMap = buildHeaderMap(responsibleData[0] || []);
  const responsibleIndexes = {
    status: responsibleHeaderMap["STATUS"],
    responsible: responsibleHeaderMap["RESPONSIBLE"],
  };

  const responsibleLookup = {};
  for (let i = 1; i < responsibleData.length; i += 1) {
    const row = responsibleData[i];
    const key = normalizeKey(row[responsibleIndexes.status]);
    if (!key) {
      continue;
    }
    if (!responsibleLookup[key]) {
      responsibleLookup[key] = sanitizeOutput(
        row[responsibleIndexes.responsible],
      );
    }
  }

  const rowCount = lastRow - 1;
  if (rowCount < 1) {
    return;
  }

  const mainData = mainSheet.getRange(2, 1, rowCount, lastCol).getValues();

  // Prepare per-header arrays and fill them in memory
  const outputValues = {};
  outputHeaders.forEach((header) => {
    outputValues[header] = new Array(rowCount).fill("");
  });

  const lookupIssues = [];
  const responsibleIssues = [];

  for (let r = 0; r < rowCount; r += 1) {
    const row = mainData[r];
    const destinationKey =
      typeof destCol === "number" ? normalizePlaIdKey(row[destCol]) : "";
    const statusValue = typeof statusCol === "number" ? row[statusCol] : "";
    const statusKey = normalizeKey(statusValue);

    let territoryTagging = "";
    let tagging = "";
    let siteClass = "";

    if (destinationKey && amceoLookup[destinationKey]) {
      const lookup = amceoLookup[destinationKey];
      territoryTagging = lookup.territory || "";
      tagging = lookup.tagging || "";
      siteClass = lookup.siteClass || "";
    }
    // record any destination keys that looked valid but had no lookup hit
    if (destinationKey && !amceoLookup[destinationKey]) {
      if (lookupIssues.length < 200) {
        lookupIssues.push([r + 2, row[destCol], destinationKey, statusValue]);
      }
    }

    const ticketStatus = computeTicketStatus(statusValue);
    const projectTagging = computeProjectTagging(tagging);
    const owner = computeOwner(projectTagging);
    const responsible =
      statusKey && responsibleLookup[statusKey]
        ? responsibleLookup[statusKey]
        : "";
    if (
      statusKey &&
      shouldExpectResponsible(statusKey) &&
      !responsibleLookup[statusKey]
    ) {
      if (responsibleIssues.length < 200) {
        responsibleIssues.push([r + 2, statusValue, statusKey]);
      }
    }

    outputValues["Territory Tagging"][r] = territoryTagging;
    outputValues["Ticket Status"][r] = ticketStatus;
    outputValues["TAGGING"][r] = tagging;
    outputValues["PROJECT TAGGING"][r] = projectTagging;
    outputValues["OWNER"][r] = owner;
    outputValues["RESPONSIBLE"][r] = responsible;
    outputValues["SITE CLASS"][r] = siteClass;
  }

  // Write values in batches to avoid timeouts and large single writes.
  const BATCH_SIZE = 1000;
  outputHeaders.forEach((header) => {
    const column = outputColumns[header];
    if (!column) {
      return;
    }
    try {
      for (let start = 0; start < rowCount; start += BATCH_SIZE) {
        const end = Math.min(start + BATCH_SIZE, rowCount);
        const slice = outputValues[header].slice(start, end);
        const values = slice.map((v) => [sanitizeOutput(v)]);
        if (values.length === 0) continue;
        mainSheet
          .getRange(2 + start, column, values.length, 1)
          .setValues(values)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
      }
    } catch (e) {
      // If a batch fails, log to the execution log for diagnosis and continue
      Logger.log(
        "Error writing header %s at column %s: %s",
        header,
        column,
        e && e.message,
      );
    }
  });

  // If there were lookup misses, write a short report to a sheet for diagnosis
  if (lookupIssues.length > 0 || responsibleIssues.length > 0) {
    const reportName = "LOOKUP_ISSUES";
    let reportSheet = ss.getSheetByName(reportName);
    if (!reportSheet) {
      reportSheet = ss.insertSheet(reportName);
    } else {
      reportSheet.clear();
    }

    let writeRows = [];
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

    // Ensure all rows have the same number of columns by padding shorter rows
    const maxCols =
      writeRows.reduce((m, r) => Math.max(m, r.length || 0), 0) || 1;
    const padded = writeRows.map((r) => {
      const row = Array.isArray(r) ? r.slice() : [r];
      while (row.length < maxCols) row.push("");
      return row.map((c) => (c === undefined ? "" : c));
    });
    reportSheet.getRange(1, 1, padded.length, maxCols).setValues(padded);
  }
}

function isValidMainSheet(sheet) {
  const headerValues = sheet.getRange(1, 1, 1, 3).getValues()[0];
  const region = String(headerValues[0] || "").trim();
  const plaId = String(headerValues[1] || "").trim();
  const area = String(headerValues[2] || "").trim();
  return region === "Region" && plaId === "PLA ID" && area === "Area";
}

function isValidAmceoSheet(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    return false;
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = buildHeaderMap(headers);
  const required = ["PLA ID", "TERRITORY", "SITE CLASS", "TOWERCO TAGGING"];
  return required.every(
    (header) => headerMap[normalizeHeader(header)] !== undefined,
  );
}

function isValidResponsibleSheet(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    return false;
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = buildHeaderMap(headers);
  const required = ["STATUS", "RESPONSIBLE"];
  return required.every(
    (header) => headerMap[normalizeHeader(header)] !== undefined,
  );
}

function ensureOutputColumns(sheet, headerRow, outputHeaders, protectedMaxCol) {
  const headerMap = buildHeaderMap(headerRow);
  let lastCol = headerRow.length;
  const outputColumns = {};
  const headersToWrite = [];

  outputHeaders.forEach((header) => {
    const normalized = normalizeHeader(header);
    const existingIndex = headerMap[normalized];

    if (typeof existingIndex === "number" && existingIndex >= protectedMaxCol) {
      const column = existingIndex + 1;
      outputColumns[header] = column;
      headersToWrite.push({ column, value: header });
      return;
    }

    lastCol += 1;
    outputColumns[header] = lastCol;
    headersToWrite.push({ column: lastCol, value: header });
  });

  const maxColumns = sheet.getMaxColumns();
  if (lastCol > maxColumns) {
    sheet.insertColumnsAfter(maxColumns, lastCol - maxColumns);
  }

  headersToWrite.forEach((item) => {
    const cell = sheet.getRange(1, item.column);
    cell.setValue(item.value);
    cell.setBackground("#000000"); // black fill
    cell.setFontColor("#ffffff"); // optional: make text readable
  });

  return outputColumns;
}

function buildHeaderMap(headerRow) {
  const map = {};
  headerRow.forEach((value, index) => {
    const normalized = normalizeHeader(value);
    if (normalized && map[normalized] === undefined) {
      map[normalized] = index;
    }
  });
  return map;
}

function normalizeHeader(value) {
  if (value === null || value === undefined) {
    return "";
  }
  const text = String(value).trim();
  if (!text) {
    return "";
  }
  return text.toUpperCase();
}

function normalizeKey(value) {
  if (isBlankish(value)) {
    return "";
  }
  // Remove common invisible / non-breaking characters then normalize
  const cleaned = String(value)
    .replace(/\u200B|\uFEFF|\u00A0/g, "") // zero-width, bom, nbsp
    .trim();
  return cleaned.toUpperCase();
}

function normalizePlaIdKey(value) {
  const normalized = normalizeKey(value);
  if (!normalized) {
    return "";
  }
  // Compare PLA IDs using only letters and numbers (remove punctuation/symbols/spaces)
  const alnum = normalized.replace(/[^A-Z0-9]/g, "");
  if (!alnum) {
    return "";
  }

  // Normalize numeric segments so SAN02 == SAN2, 0007 == 7
  return alnum.replace(/\d+/g, (digits) => {
    const parsed = parseInt(digits, 10);
    return isNaN(parsed) ? digits : String(parsed);
  });
}

function shouldExpectResponsible(statusKey) {
  if (!statusKey) {
    return false;
  }
  // These statuses are expected to keep RESPONSIBLE blank.
  return (
    statusKey !== "CLOSED" &&
    statusKey !== "REJECTED" &&
    statusKey !== "CANCELLED" &&
    statusKey !== "NO STATUS"
  );
}

function isBlankish(value) {
  if (value === null || value === undefined) {
    return true;
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return true;
    }
    const upper = trimmed.toUpperCase();
    if (upper === "N/A" || upper === "#N/A" || upper === "NA") {
      return true;
    }
    if (upper.startsWith("#")) {
      return true;
    }
    return false;
  }
  if (typeof value === "number" && isNaN(value)) {
    return true;
  }
  return false;
}

function sanitizeOutput(value) {
  if (isBlankish(value)) {
    return "";
  }
  if (typeof value === "string") {
    return value.trim();
  }
  return value;
}

function choosePreferredValue(currentValue, incomingValue) {
  const current = sanitizeOutput(currentValue);
  const incoming = sanitizeOutput(incomingValue);
  if (!current && incoming) {
    return incoming;
  }
  return current;
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
  if (isBlankish(statusValue)) {
    return "";
  }
  const normalized = normalizeKey(statusValue);
  return normalized === "CLOSED" ||
    normalized === "REJECTED" ||
    normalized === "CANCELLED"
    ? "CLOSED"
    : "OPEN";
}

function computeProjectTagging(taggingValue) {
  if (isBlankish(taggingValue)) {
    return "";
  }
  const normalized = normalizeKey(taggingValue);
  if (normalized === "BAU") {
    return "Globe Sites";
  }
  if (normalized === "BUILD GLOBE-ACQUIRED SITES") {
    return "BUILD GLOBE-ACQUIRED SITES";
  }
  if (normalized === "COLOCATION") {
    return "COLOCATION";
  }
  if (normalized === "SLB") {
    return "SLB";
  }
  if (normalized === "BUILD TO SUIT") {
    return "BTS";
  }
  return "";
}

function computeOwner(projectTaggingValue) {
  if (isBlankish(projectTaggingValue)) {
    return "";
  }
  const normalized = normalizeKey(projectTaggingValue);
  if (
    normalized === "GLOBE SITES" ||
    normalized === "BUILD GLOBE-ACQUIRED SITES"
  ) {
    return "GLOBE";
  }
  if (
    normalized === "BTS" ||
    normalized === "COLOCATION" ||
    normalized === "SLB"
  ) {
    return "TOWERCO";
  }
  return "";
}
