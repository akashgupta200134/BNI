// // src/excelHelper.js
// import ExcelJS from 'exceljs';
// import fs      from 'fs';

// // ── Column layout (now 11 cols: added Email2, Phone2) ─────────────────────────
// export const HEADERS = [
//   'Name', 'Chapter', 'Company', 'City', 'Industry and Classification',
//   'Email', 'Email2', 'PhoneNo', 'Phone2', 'Website', 'Status'
// ];
// // Col indices (1-based)
// export const COL = {
//   Name: 1, Chapter: 2, Company: 3, City: 4, Industry: 5,
//   Email: 6, Email2: 7, Phone: 8, Phone2: 9, Website: 10, Status: 11,
// };

// // ── Sheet name registry — maps full category name → safe Excel sheet name ─────
// // Key: filePath  Value: Map<fullCategoryName, safeSheetName>
// const _sheetNameRegistry = new Map();

// export function initSheetRegistry(filePath) {
//   _sheetNameRegistry.set(filePath, new Map());
// }

// // Register ALL categories up-front before any sheet is created.
// // Guarantees every full name maps to a unique ≤31-char sheet name.
// export function registerCategories(filePath, categories) {
//   if (!_sheetNameRegistry.has(filePath)) initSheetRegistry(filePath);
//   const registry  = _sheetNameRegistry.get(filePath);
//   const usedNames = new Set();

//   for (const cat of categories) {
//     if (registry.has(cat)) continue;

//     // Strip illegal chars, collapse whitespace
//     let safe = cat
//       .replace(/[*?:\\/\[\]]/g, ' ')
//       .replace(/\s+/g, ' ')
//       .trim()
//       .substring(0, 31)
//       .trim();

//     // Deduplicate: if name collides, shorten base + append counter
//     if (usedNames.has(safe.toLowerCase())) {
//       let counter   = 2;
//       const base    = safe.substring(0, 28).trim();
//       let candidate = `${base} ${counter}`;
//       while (usedNames.has(candidate.toLowerCase())) {
//         counter++;
//         candidate = `${base} ${counter}`;
//       }
//       safe = candidate;
//     }

//     usedNames.add(safe.toLowerCase());
//     registry.set(cat, safe);
//   }
// }

// // Resolve full category name → safe sheet name
// // Always pass filePath so the registry is used; fallback only for legacy callers
// export function sanitizeSheetName(catName, filePath = null) {
//   if (filePath) {
//     const reg = _sheetNameRegistry.get(filePath);
//     if (reg && reg.has(catName)) return reg.get(catName);
//   }
//   // Legacy fallback (no registry available)
//   return catName
//     .replace(/[*?:\\/\[\]]/g, ' ')
//     .replace(/\s+/g, ' ')
//     .trim()
//     .substring(0, 31)
//     .trim();
// }

// // ── Read all rows from a specific sheet (row 1 = headers) ────────────────────
// export async function readAllRows(filePath, sheetName = null) {
//   const wb = new ExcelJS.Workbook();
//   await wb.xlsx.readFile(filePath);
//   const ws = sheetName
//     ? (wb.getWorksheet(sanitizeSheetName(sheetName, filePath)) || wb.worksheets[0])
//     : wb.worksheets[0];
//   if (!ws) return [];

//   const headers = [];
//   ws.getRow(1).eachCell({ includeEmpty: true }, (cell, colNum) => {
//     headers[colNum] = String(cell.value ?? '').trim();
//   });

//   const rows = [];
//   ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
//     if (rowNum === 1) return;
//     const obj = {};
//     row.eachCell({ includeEmpty: true }, (cell, colNum) => {
//       obj[headers[colNum] || `col${colNum}`] = cell.value ?? '';
//     });
//     obj.__rowNum = rowNum;
//     rows.push(obj);
//   });
//   return rows;
// }

// // ── Read Category_Country.xlsx ────────────────────────────────────────────────
// export async function readCategoryCountry(filePath) {
//   const wb = new ExcelJS.Workbook();
//   await wb.xlsx.readFile(filePath);
//   const ws = wb.worksheets[0];

//   const headers = [];
//   ws.getRow(1).eachCell({ includeEmpty: true }, (cell, colNum) => {
//     headers[colNum] = String(cell.value ?? '').trim();
//   });

//   const results = [];
//   ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
//     if (rowNum === 1) return;
//     const country    = String(row.getCell(1).value ?? '').trim();
//     const categories = [];
//     for (let col = 2; col <= headers.length; col++) {
//       const val = String(row.getCell(col).value ?? '').trim();
//       if (val) categories.push(val);
//     }
//     if (country) results.push({ country, categories });
//   });
//   return results;
// }

// // ── Create multi-sheet workbook ───────────────────────────────────────────────
// export async function createMultiSheetWorkbook(filePath, categories) {
//   // Register ALL categories before creating any sheet so names are stable
//   registerCategories(filePath, categories);

//   const wb = new ExcelJS.Workbook();

//   for (const cat of categories) {
//     const ws = wb.addWorksheet(sanitizeSheetName(cat, filePath));
//     const headerRow = ws.addRow(HEADERS);
//     headerRow.font = { bold: true };
//     headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
//     ws.columns = [
//       { width: 25 }, // Name
//       { width: 20 }, // Chapter
//       { width: 25 }, // Company
//       { width: 15 }, // City
//       { width: 35 }, // Industry
//       { width: 30 }, // Email
//       { width: 30 }, // Email2
//       { width: 18 }, // PhoneNo
//       { width: 18 }, // Phone2
//       { width: 30 }, // Website
//       { width: 10 }, // Status
//     ];
//   }

//   await wb.xlsx.writeFile(filePath);
//   console.log(`[Excel] Created ${categories.length}-sheet workbook: ${filePath}`);
// }

// // ── Append member rows — SKIP duplicates by Name ──────────────────────────────
// export async function appendRowsToSheet(filePath, sheetName, members) {
//   const wb = new ExcelJS.Workbook();
//   await wb.xlsx.readFile(filePath);

//   const safeName = sanitizeSheetName(sheetName, filePath);
//   let ws = wb.getWorksheet(safeName);
//   if (!ws) {
//     console.warn(`[Excel] Sheet not found: "${safeName}" — creating it`);
//     ws = wb.addWorksheet(safeName);
//     ws.addRow(HEADERS);
//   }

//   // Build set of existing names for duplicate check
//   const existingNames = new Set();
//   ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
//     if (rowNum === 1) return;
//     const name = String(row.getCell(COL.Name).value ?? '').trim().toLowerCase();
//     if (name) existingNames.add(name);
//   });

//   let added = 0, dupes = 0;
//   for (const m of members) {
//     const nameKey = (m.Name || '').trim().toLowerCase();
//     if (!nameKey) continue;

//     // SKIP if already exists in this sheet
//     if (existingNames.has(nameKey)) {
//       console.log(`  [Dupe] Skipping existing: "${m.Name}"`);
//       dupes++;
//       continue;
//     }

//     existingNames.add(nameKey);
//     ws.addRow([
//       m.Name || '', m.Chapter || '', m.Company || '',
//       m.City || '', m.IndustryClassification || '',
//       '', '', '', '', '', '',   // Email, Email2, Phone, Phone2, Website, Status
//     ]);
//     added++;
//   }

//   await wb.xlsx.writeFile(filePath);
//   console.log(`[Excel] Sheet "${safeName}": ${added} added, ${dupes} duplicates skipped`);
// }

// // ── Write validated contact data to correct sheet row ────────────────────────
// export async function writeCellInSheet(filePath, sheetName, memberName, contacts) {
//   const wb = new ExcelJS.Workbook();
//   await wb.xlsx.readFile(filePath);

//   const safeName = sanitizeSheetName(sheetName, filePath);
//   const ws = wb.getWorksheet(safeName);
//   if (!ws) { console.warn(`[Excel] Sheet not found: ${safeName}`); return; }

//   // Find member row
//   let targetRow = -1;
//   ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
//     if (targetRow !== -1) return;
//     if (String(row.getCell(COL.Name).value ?? '').trim().toLowerCase()
//         === memberName.trim().toLowerCase()) {
//       targetRow = rowNum;
//     }
//   });

//   if (targetRow === -1) {
//     console.warn(`[Excel] Row not found for "${memberName}" in "${safeName}"`);
//     return;
//   }

//   // Check if already DONE — skip re-writing
//   const currentStatus = String(ws.getCell(targetRow, COL.Status).value ?? '').trim().toUpperCase();
//   if (currentStatus === 'DONE') {
//     console.log(`  [Skip] Already DONE in Excel: "${memberName}"`);
//     return;
//   }

//   // null contacts = no data found — mark as NoData and leave contact cells blank
//   if (!contacts) {
//     ws.getCell(targetRow, COL.Status).value = 'NoData';
//     await wb.xlsx.writeFile(filePath);
//     console.log(`[Excel] Row ${targetRow} → NoData in "${safeName}"`);
//     return;
//   }

//   // Write validated contacts
//   ws.getCell(targetRow, COL.Email).value   = contacts.email   || 'Not Found';
//   ws.getCell(targetRow, COL.Email2).value  = contacts.email2  || '';
//   ws.getCell(targetRow, COL.Phone).value   = contacts.phone   || 'Not Found';
//   ws.getCell(targetRow, COL.Phone2).value  = contacts.phone2  || '';
//   ws.getCell(targetRow, COL.Website).value = contacts.website || 'Not Found';
//   ws.getCell(targetRow, COL.Status).value  = 'DONE';

//   await wb.xlsx.writeFile(filePath);
//   console.log(`[Excel] Row ${targetRow} → DONE in "${safeName}"`);
// }

// // ── Check if member row is already DONE ──────────────────────────────────────
// export async function isRowDoneInSheet(filePath, sheetName, memberName) {
//   try {
//     const wb = new ExcelJS.Workbook();
//     await wb.xlsx.readFile(filePath);
//     const ws = wb.getWorksheet(sanitizeSheetName(sheetName, filePath));
//     if (!ws) return false;

//     let done = false;
//     ws.eachRow({ includeEmpty: false }, (row) => {
//       if (done) return;
//       if (String(row.getCell(COL.Name).value ?? '').trim().toLowerCase()
//           === memberName.trim().toLowerCase()) {
//         done = String(row.getCell(COL.Status).value ?? '').trim().toUpperCase() === 'DONE';
//       }
//     });
//     return done;
//   } catch {
//     return false;
//   }
// }

// // ── Legacy helpers (kept for compatibility) ───────────────────────────────────
// export async function writeSheet(filePath, headers, dataRows) {
//   const wb = new ExcelJS.Workbook();
//   const ws = wb.addWorksheet('Sheet1');
//   ws.addRow(headers);
//   for (const row of dataRows) ws.addRow(row);
//   await wb.xlsx.writeFile(filePath);
// }

// export async function writeCell(filePath, rowIndex, colLetter, value) {
//   const wb = new ExcelJS.Workbook();
//   await wb.xlsx.readFile(filePath);
//   const ws = wb.worksheets[0];
//   ws.getCell(rowIndex, colLetterToNumber(colLetter)).value = value;
//   await wb.xlsx.writeFile(filePath);
// }

// export async function findRowByValue(filePath, searchValue) {
//   const wb = new ExcelJS.Workbook();
//   await wb.xlsx.readFile(filePath);
//   const ws = wb.worksheets[0];
//   let found = -1;
//   ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
//     if (found !== -1) return;
//     row.eachCell({ includeEmpty: false }, cell => {
//       if (String(cell.value).trim() === String(searchValue).trim()) found = rowNum;
//     });
//   });
//   return found;
// }

// function colLetterToNumber(letter) {
//   let n = 0;
//   for (const ch of letter.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
//   return n;
// }



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

// ── Sheet name registry — maps full category name → safe Excel sheet name ─────
// IMPORTANT: this is persisted INSIDE the workbook itself, on a hidden
// "_CategoryMap" sheet (col A = full category name, col B = actual sheet name).
// This fixes a real bug: an in-memory-only registry resets between process
// runs (e.g. resuming a bot.js run later, or the previous run being
// interrupted). If the registry was rebuilt fresh from a possibly-reordered
// or edited Category_Country.xlsx, it could compute a DIFFERENT truncated
// sheet name than what's actually saved in the .xlsx — causing
// wb.getWorksheet(name) to return null and silently skip writing data.
// Storing the mapping in the file itself means lookups always match reality.
const MAP_SHEET_NAME   = '_CategoryMap';
const _memoryRegistry   = new Map(); // filePath -> Map<fullName, safeName>  (fast path cache)

function sanitizeBase(cat) {
  return cat
    .replace(/[*?:\\/\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 31)
    .trim();
}

export function initSheetRegistry(filePath) {
  _memoryRegistry.set(filePath, new Map());
}

// Load the persisted mapping from the workbook's hidden _CategoryMap sheet,
// if the file already exists on disk. Populates the in-memory cache.
async function loadPersistedRegistry(filePath) {
  if (!fs.existsSync(filePath)) return new Map();
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const mapSheet = wb.getWorksheet(MAP_SHEET_NAME);
    const map = new Map();
    if (mapSheet) {
      mapSheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
        if (rowNum === 1) return; // header
        const full = String(row.getCell(1).value ?? '').trim();
        const safe = String(row.getCell(2).value ?? '').trim();
        if (full && safe) map.set(full, safe);
      });
    }
    return map;
  } catch (err) {
    console.warn(`[Excel] Could not load persisted category map: ${err.message}`);
    return new Map();
  }
}

// Register ALL categories up-front before any sheet is created.
// Guarantees every full name maps to a unique ≤31-char sheet name, and
// REUSES any mapping already persisted in the file on disk so re-running
// the bot (even after editing Category_Country.xlsx) never breaks lookups
// for categories that were already processed.
export async function registerCategories(filePath, categories) {
  // Start from whatever is already persisted in the actual file, if any
  const persisted = await loadPersistedRegistry(filePath);
  const registry   = persisted; // reuse the Map directly
  const usedNames  = new Set([...registry.values()].map(v => v.toLowerCase()));

  for (const cat of categories) {
    if (registry.has(cat)) continue; // already mapped — keep existing name stable

    let safe = sanitizeBase(cat);

    if (usedNames.has(safe.toLowerCase())) {
      let counter   = 2;
      const base    = safe.substring(0, 28).trim();
      let candidate = `${base} ${counter}`;
      while (usedNames.has(candidate.toLowerCase())) {
        counter++;
        candidate = `${base} ${counter}`;
      }
      safe = candidate;
    }

    usedNames.add(safe.toLowerCase());
    registry.set(cat, safe);
  }

  _memoryRegistry.set(filePath, registry);
  return registry;
}

// Persist the current in-memory registry into the workbook's hidden
// _CategoryMap sheet. Call this right after creating/writing the workbook.
function writeRegistryToWorkbook(wb, filePath) {
  const registry = _memoryRegistry.get(filePath);
  if (!registry || registry.size === 0) return;

  let mapSheet = wb.getWorksheet(MAP_SHEET_NAME);
  if (mapSheet) wb.removeWorksheet(mapSheet.id);
  mapSheet = wb.addWorksheet(MAP_SHEET_NAME);
  mapSheet.state = 'veryHidden'; // hide from normal Excel view
  mapSheet.addRow(['FullCategoryName', 'SheetName']);
  for (const [full, safe] of registry.entries()) {
    mapSheet.addRow([full, safe]);
  }
}

// Resolve full category name → safe sheet name.
// Synchronous fast path using the in-memory cache (populated by
// registerCategories). Always call registerCategories(filePath, categories)
// at the start of a run before relying on this.
export function sanitizeSheetName(catName, filePath = null) {
  if (filePath) {
    const reg = _memoryRegistry.get(filePath);
    if (reg && reg.has(catName)) return reg.get(catName);
  }
  // Legacy fallback (no registry available) — best-effort only
  return sanitizeBase(catName);
}

// ── Read all rows from a specific sheet (row 1 = headers) ────────────────────
export async function readAllRows(filePath, sheetName = null) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = sheetName
    ? (wb.getWorksheet(await resolveSheetName(filePath, sheetName)) || wb.worksheets[0])
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
  // Register ALL categories before creating any sheet so names are stable.
  // This also reuses any mapping already persisted on disk (for resumed runs).
  await registerCategories(filePath, categories);

  const wb = new ExcelJS.Workbook();

  for (const cat of categories) {
    const ws = wb.addWorksheet(sanitizeSheetName(cat, filePath));
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

  // Persist the category→sheet-name map INSIDE the workbook so future runs
  // (even from a fresh process / edited Category_Country.xlsx) always
  // resolve the same sheet names for categories already created here.
  writeRegistryToWorkbook(wb, filePath);

  await wb.xlsx.writeFile(filePath);
  console.log(`[Excel] Created ${categories.length}-sheet workbook: ${filePath}`);
}

// Self-healing resolver: if the in-memory cache doesn't have this category
// (e.g. registry wasn't populated in this process, or this is a resumed
// run that skipped createOutputFile), load the persisted map straight from
// the file on disk before falling back to fresh truncation. This guarantees
// we never silently miss a sheet that was actually created with a
// previously-registered (possibly different) truncated name.
async function resolveSheetName(filePath, catName) {
  let reg = _memoryRegistry.get(filePath);
  if (reg && reg.has(catName)) return reg.get(catName);

  // Cache miss — try loading from the actual file on disk
  const persisted = await loadPersistedRegistry(filePath);
  if (persisted.has(catName)) {
    _memoryRegistry.set(filePath, persisted); // warm the cache for next calls
    return persisted.get(catName);
  }

  // Truly unknown category — fall back to fresh truncation (best effort)
  return sanitizeBase(catName);
}

// ── Append member rows — SKIP duplicates by Name ──────────────────────────────
export async function appendRowsToSheet(filePath, sheetName, members) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  const safeName = await resolveSheetName(filePath, sheetName);
  let ws = wb.getWorksheet(safeName);
  if (!ws) {
    console.warn(`[Excel] Sheet not found: "${safeName}" — creating it`);
    ws = wb.addWorksheet(safeName);
    ws.addRow(HEADERS);

    // Make sure this new mapping is remembered for future calls/processes
    let reg = _memoryRegistry.get(filePath);
    if (!reg) { reg = new Map(); _memoryRegistry.set(filePath, reg); }
    reg.set(sheetName, safeName);
    writeRegistryToWorkbook(wb, filePath); // will be saved with this write below
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

  const safeName = await resolveSheetName(filePath, sheetName);
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

  // null contacts = no data found — mark as NoData and leave contact cells blank
  if (!contacts) {
    ws.getCell(targetRow, COL.Status).value = 'NoData';
    await wb.xlsx.writeFile(filePath);
    console.log(`[Excel] Row ${targetRow} → NoData in "${safeName}"`);
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
    const ws = wb.getWorksheet(await resolveSheetName(filePath, sheetName));
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