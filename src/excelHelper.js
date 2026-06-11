// src/excelHelper.js
import ExcelJS from 'exceljs';
import fs      from 'fs';

// ── Column layout (now 11 cols: added Email2, Phone2) ─────────────────────────
export const HEADERS = [
  'Name', 'Chapter', 'Company', 'City', 'Industry and Classification',
  'Email', 'Email2', 'PhoneNo', 'Phone2', 'Website', 'Status'
];
// Col indices (1-based)
export const COL = {
  Name: 1, Chapter: 2, Company: 3, City: 4, Industry: 5,
  Email: 6, Email2: 7, Phone: 8, Phone2: 9, Website: 10, Status: 11,
};

// ── Sanitize Excel sheet name ─────────────────────────────────────────────────
export function sanitizeSheetName(name) {
  return name.replace(/[*?:\\/\[\]]/g, '-').substring(0, 31).trim();
}

// ── Read all rows from a specific sheet (row 1 = headers) ────────────────────
export async function readAllRows(filePath, sheetName = null) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = sheetName
    ? (wb.getWorksheet(sanitizeSheetName(sheetName)) || wb.worksheets[0])
    : wb.worksheets[0];
  if (!ws) return [];

  const headers = [];
  ws.getRow(1).eachCell({ includeEmpty: true }, (cell, colNum) => {
    headers[colNum] = String(cell.value ?? '').trim();
  });

  const rows = [];
  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (rowNum === 1) return;
    const obj = {};
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      obj[headers[colNum] || `col${colNum}`] = cell.value ?? '';
    });
    obj.__rowNum = rowNum;
    rows.push(obj);
  });
  return rows;
}

// ── Read Category_Country.xlsx ────────────────────────────────────────────────
export async function readCategoryCountry(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];

  const headers = [];
  ws.getRow(1).eachCell({ includeEmpty: true }, (cell, colNum) => {
    headers[colNum] = String(cell.value ?? '').trim();
  });

  const results = [];
  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (rowNum === 1) return;
    const country    = String(row.getCell(1).value ?? '').trim();
    const categories = [];
    for (let col = 2; col <= headers.length; col++) {
      const val = String(row.getCell(col).value ?? '').trim();
      if (val) categories.push(val);
    }
    if (country) results.push({ country, categories });
  });
  return results;
}

// ── Create multi-sheet workbook ───────────────────────────────────────────────
export async function createMultiSheetWorkbook(filePath, categories) {
  const wb = new ExcelJS.Workbook();

  for (const cat of categories) {
    const ws = wb.addWorksheet(sanitizeSheetName(cat));
    const headerRow = ws.addRow(HEADERS);
    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
    ws.columns = [
      { width: 25 }, // Name
      { width: 20 }, // Chapter
      { width: 25 }, // Company
      { width: 15 }, // City
      { width: 35 }, // Industry
      { width: 30 }, // Email
      { width: 30 }, // Email2
      { width: 18 }, // PhoneNo
      { width: 18 }, // Phone2
      { width: 30 }, // Website
      { width: 10 }, // Status
    ];
  }

  await wb.xlsx.writeFile(filePath);
  console.log(`[Excel] Created ${categories.length}-sheet workbook: ${filePath}`);
}

// ── Append member rows — SKIP duplicates by Name ──────────────────────────────
export async function appendRowsToSheet(filePath, sheetName, members) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  const safeName = sanitizeSheetName(sheetName);
  let ws = wb.getWorksheet(safeName);
  if (!ws) {
    console.warn(`[Excel] Sheet not found: "${safeName}" — creating it`);
    ws = wb.addWorksheet(safeName);
    ws.addRow(HEADERS);
  }

  // Build set of existing names for duplicate check
  const existingNames = new Set();
  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (rowNum === 1) return;
    const name = String(row.getCell(COL.Name).value ?? '').trim().toLowerCase();
    if (name) existingNames.add(name);
  });

  let added = 0, dupes = 0;
  for (const m of members) {
    const nameKey = (m.Name || '').trim().toLowerCase();
    if (!nameKey) continue;

    // SKIP if already exists in this sheet
    if (existingNames.has(nameKey)) {
      console.log(`  [Dupe] Skipping existing: "${m.Name}"`);
      dupes++;
      continue;
    }

    existingNames.add(nameKey);
    ws.addRow([
      m.Name || '', m.Chapter || '', m.Company || '',
      m.City || '', m.IndustryClassification || '',
      '', '', '', '', '', '',   // Email, Email2, Phone, Phone2, Website, Status
    ]);
    added++;
  }

  await wb.xlsx.writeFile(filePath);
  console.log(`[Excel] Sheet "${safeName}": ${added} added, ${dupes} duplicates skipped`);
}

// ── Write validated contact data to correct sheet row ────────────────────────
export async function writeCellInSheet(filePath, sheetName, memberName, contacts) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  const safeName = sanitizeSheetName(sheetName);
  const ws = wb.getWorksheet(safeName);
  if (!ws) { console.warn(`[Excel] Sheet not found: ${safeName}`); return; }

  // Find member row
  let targetRow = -1;
  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (targetRow !== -1) return;
    if (String(row.getCell(COL.Name).value ?? '').trim().toLowerCase()
        === memberName.trim().toLowerCase()) {
      targetRow = rowNum;
    }
  });

  if (targetRow === -1) {
    console.warn(`[Excel] Row not found for "${memberName}" in "${safeName}"`);
    return;
  }

  // Check if already DONE — skip re-writing
  const currentStatus = String(ws.getCell(targetRow, COL.Status).value ?? '').trim().toUpperCase();
  if (currentStatus === 'DONE') {
    console.log(`  [Skip] Already DONE in Excel: "${memberName}"`);
    return;
  }

  // Write validated contacts
  ws.getCell(targetRow, COL.Email).value   = contacts.email   || 'Not Found';
  ws.getCell(targetRow, COL.Email2).value  = contacts.email2  || '';
  ws.getCell(targetRow, COL.Phone).value   = contacts.phone   || 'Not Found';
  ws.getCell(targetRow, COL.Phone2).value  = contacts.phone2  || '';
  ws.getCell(targetRow, COL.Website).value = contacts.website || 'Not Found';
  ws.getCell(targetRow, COL.Status).value  = 'DONE';

  await wb.xlsx.writeFile(filePath);
  console.log(`[Excel] Row ${targetRow} → DONE in "${safeName}"`);
}

// ── Check if member row is already DONE ──────────────────────────────────────
export async function isRowDoneInSheet(filePath, sheetName, memberName) {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet(sanitizeSheetName(sheetName));
    if (!ws) return false;

    let done = false;
    ws.eachRow({ includeEmpty: false }, (row) => {
      if (done) return;
      if (String(row.getCell(COL.Name).value ?? '').trim().toLowerCase()
          === memberName.trim().toLowerCase()) {
        done = String(row.getCell(COL.Status).value ?? '').trim().toUpperCase() === 'DONE';
      }
    });
    return done;
  } catch {
    return false;
  }
}

// ── Legacy helpers (kept for compatibility) ───────────────────────────────────
export async function writeSheet(filePath, headers, dataRows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.addRow(headers);
  for (const row of dataRows) ws.addRow(row);
  await wb.xlsx.writeFile(filePath);
}

export async function writeCell(filePath, rowIndex, colLetter, value) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  ws.getCell(rowIndex, colLetterToNumber(colLetter)).value = value;
  await wb.xlsx.writeFile(filePath);
}

export async function findRowByValue(filePath, searchValue) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  let found = -1;
  ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (found !== -1) return;
    row.eachCell({ includeEmpty: false }, cell => {
      if (String(cell.value).trim() === String(searchValue).trim()) found = rowNum;
    });
  });
  return found;
}

function colLetterToNumber(letter) {
  let n = 0;
  for (const ch of letter.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}