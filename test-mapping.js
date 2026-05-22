const assert = require('assert');

function padRow(row, length) {
  const output = row ? row.slice(0) : [];
  while (output.length < length) output.push('');
  return output;
}

function firstNonEmpty(...values) {
  for (const value of values) {
    if (value !== '' && value != null) return value;
  }
  return '';
}

function textFromDisplay(value) {
  if (value === '' || value == null) return '';
  return String(value).trim();
}

function toNumber(value, displayValue) {
  if (value === '' || value == null) {
    if (displayValue === '' || displayValue == null) return '';
    const parsedFromDisplay = Number(String(displayValue).replace(/,/g, ''));
    return Number.isNaN(parsedFromDisplay) ? '' : parsedFromDisplay;
  }
  if (typeof value === 'number') return value;
  const parsed = Number(String(value).replace(/,/g, ''));
  return Number.isNaN(parsed) ? '' : parsed;
}

function parseDateString(text) {
  if (!text) return null;
  const value = String(text).trim();
  const match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!match) return null;
  const month = Number(match[1]) - 1;
  const day = Number(match[2]);
  const year = Number(match[3]);
  const hour = match[4] ? Number(match[4]) : 0;
  const minute = match[5] ? Number(match[5]) : 0;
  const second = match[6] ? Number(match[6]) : 0;
  const date = new Date(year, month, day, hour, minute, second);
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeKey(value) {
  if (value === null || value === undefined) return '';
  const cleaned = String(value).replace(/\u200B|\uFEFF|\u00A0/g, '').trim();
  return cleaned.toUpperCase();
}

function computeTicketStatus(statusValue) {
  if (statusValue === '' || statusValue == null) return '';
  const normalized = normalizeKey(statusValue);
  return normalized === 'CLOSED' ||
    normalized === 'REJECTED' ||
    normalized === 'CANCELLED' ||
    normalized === 'DECLINED'
    ? 'CLOSED'
    : 'OPEN';
}

function buildDestinationRow(sourceRow, sourceDisplayRow, existingRow) {
  const row = padRow(existingRow ? existingRow.slice() : [], 41);

  row[0] = textFromDisplay(sourceDisplayRow[0]);
  row[1] = firstNonEmpty(textFromDisplay(sourceDisplayRow[12]), textFromDisplay(sourceDisplayRow[30]));
  row[4] = textFromDisplay(sourceDisplayRow[9]);
  row[5] = textFromDisplay(sourceDisplayRow[24]);
  row[6] = textFromDisplay(sourceDisplayRow[19]);
  row[7] = parseDateString(sourceDisplayRow[3]);
  row[8] = textFromDisplay(sourceDisplayRow[2]);
  row[9] = textFromDisplay(sourceDisplayRow[8]);
  row[10] = textFromDisplay(sourceDisplayRow[1]);
  row[11] = textFromDisplay(sourceDisplayRow[4]);
  row[12] = textFromDisplay(sourceDisplayRow[17]);
  row[13] = textFromDisplay(sourceDisplayRow[6]);
  row[14] = textFromDisplay(sourceDisplayRow[5]);
  row[15] = toNumber(sourceRow[14], sourceDisplayRow[14]);
  row[16] = toNumber(sourceRow[18], sourceDisplayRow[18]);
  row[17] = textFromDisplay(sourceDisplayRow[22]);
  row[20] = firstNonEmpty(textFromDisplay(sourceDisplayRow[26]), textFromDisplay(sourceDisplayRow[31]));
  row[21] = textFromDisplay(sourceDisplayRow[13]);
  row[22] = textFromDisplay(sourceDisplayRow[20]);
  row[23] = firstNonEmpty(textFromDisplay(sourceDisplayRow[32]), textFromDisplay(sourceDisplayRow[13]));
  row[24] = textFromDisplay(sourceDisplayRow[21]);
  row[25] = toNumber(sourceRow[7], sourceDisplayRow[7]);
  row[26] = toNumber(sourceRow[25], sourceDisplayRow[25]);
  row[28] = textFromDisplay(sourceDisplayRow[11]);
  row[30] = textFromDisplay(sourceDisplayRow[15]);
  row[31] = textFromDisplay(sourceDisplayRow[16]);
  row[33] = textFromDisplay(sourceDisplayRow[28]);
  row[34] = textFromDisplay(sourceDisplayRow[29]);

  return row;
}

const sourceValues = [
  ['NCR', 'RAENELLE ANDREA ESQUILON', 'GPO-00065177', '07/01/2025 9:47', 'EMMANUEL IVAN MAMASIG', 'CLOSED', 'MICHAEL GAPASIN', 25, 231826, '', 'DRP Gasoline', 'Manual Purchase', 'NCR6516', '', 50, 'No', 'GOODEARTHPLZMLANCRID', 'RITCHARD BARTOLOME', 50, 'NCR6516', '', '', '5784', '', '738766030276552003', 0, '07/01/2025 11:58', 'Jul 2025', 'CLOSED', 'BAU', 'NCR6516', '07/01/2025 11:58', '']
];

const sourceDisplayValues = [
  ['NCR', 'RAENELLE ANDREA ESQUILON', 'GPO-00065177', '07/01/2025 9:47', 'EMMANUEL IVAN MAMASIG', 'CLOSED', 'MICHAEL GAPASIN', '25', '231826', '', 'DRP Gasoline', 'Manual Purchase', 'NCR6516', '', '50', 'No', 'GOODEARTHPLZMLANCRID', 'RITCHARD BARTOLOME', '50', 'NCR6516', '', '', '5784', '', '738766030276552003', '0', '07/01/2025 11:58', 'Jul 2025', 'CLOSED', 'BAU', 'NCR6516', '07/01/2025 11:58', '']
];

const existingRow = padRow([], 41);
existingRow[2] = 'GMA Area 4';
existingRow[3] = 'preserve fleet card id';
existingRow[18] = 'preserve manual receipt';
existingRow[19] = 'preserve transaction type';
existingRow[27] = 'preserve last fuel delivery hours';
existingRow[32] = 'preserve territory tagging';
existingRow[35] = 'preserve site owner';
existingRow[36] = 'preserve project tagging';
existingRow[37] = 'preserve owner';
existingRow[38] = 'preserve responsible';
existingRow[39] = 'preserve site class';
existingRow[40] = 'preserve severity';

const output = buildDestinationRow(sourceValues[0], sourceDisplayValues[0], existingRow);

assert.strictEqual(output[0], 'NCR');
assert.strictEqual(output[1], 'NCR6516');
assert.strictEqual(output[2], 'GMA Area 4');
assert.strictEqual(output[3], 'preserve fleet card id');
assert.strictEqual(output[8], 'GPO-00065177');
assert.strictEqual(output[24], '');
assert.strictEqual(output[25], 25);
assert.strictEqual(output[26], 0);
assert.strictEqual(output[27], 'preserve last fuel delivery hours');
assert.strictEqual(output[28], 'Manual Purchase');
assert.strictEqual(output[30], 'No');
assert.strictEqual(output[32], 'preserve territory tagging');
assert.strictEqual(output[33], 'CLOSED');
assert.strictEqual(output[34], 'BAU');

assert.strictEqual(computeTicketStatus('DECLINED'), 'CLOSED');
assert.strictEqual(computeTicketStatus('dispatched'), 'OPEN');

console.log('Mapping test passed.');
