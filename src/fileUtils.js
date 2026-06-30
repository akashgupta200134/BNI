// // src/fileUtils.js
// import fs      from 'fs';
// import path    from 'path';
// import ExcelJS from 'exceljs';
// import cfg     from '../config.js';
// import { sanitizeSheetName, HEADERS, COL } from './excelHelper.js';

// // ─────────────────────────────────────────────────────────────────────────────
// //  VALIDATORS — used both during scraping AND during clean+archive
// // ─────────────────────────────────────────────────────────────────────────────

// // Valid Indian/international phone: 7-15 digits, not a date, not a fragment
// export function isValidPhone(raw) {
//   if (!raw || typeof raw !== 'string') return false;
//   const p = raw.trim();

//   // Reject date patterns: dd.mm.yyyy, dd-mm-yy, dd/mm/yyyy etc.
//   if (/^\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}\.?$/.test(p)) return false;

//   // Reject strings ending with a dot that are clearly reference numbers
//   if (/^\d+\.$/.test(p)) return false;

//   // Reject Facebook post IDs (15-16 digits starting with common prefixes)
//   const digits = p.replace(/[\s\-().+]/g, '');
//   if (digits.length > 15) return false;
//   if (digits.length < 7)  return false;
//   if (!/^\d+$/.test(digits)) return false;

//   // Reject PIN codes / short zip codes (exactly 6 digits with a dot or space)
//   if (/^\d{6}\.$/.test(p) || /^\d{3}\s\d{3}$/.test(p)) return false;

//   return true;
// }

// // Valid email: standard format, not a BNI internal address
// export function isValidEmail(raw) {
//   if (!raw || typeof raw !== 'string') return false;
//   const e = raw.trim().toLowerCase();
//   if (!/^[a-zA-Z0-9._+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(e)) return false;
//   if (e.includes('bniconnect') || e.includes('@bni.')) return false;
//   return true;
// }

// // Valid website: must parse as URL with a real TLD, not a BNI internal link
// export function isValidWebsite(raw) {
//   if (!raw || typeof raw !== 'string') return false;
//   const url = raw.trim().toLowerCase();
//   try {
//     const parsed   = new URL(url);
//     const hostname = parsed.hostname;
//     const parts    = hostname.split('.');
//     if (parts.length < 2) return false;
//     const tld = parts[parts.length - 1];
//     if (!/^[a-z]{2,6}$/.test(tld)) return false;
//     // Reject BNI domains
//     if (hostname.includes('bniconnectglobal') || hostname.includes('bni.com')
//         || hostname.includes('bni.in')) return false;
//     // Reject known incomplete/fake domains
//     const fakeDomains = ['comingsoon', 'coming soon', 'zendesk'];
//     if (fakeDomains.some(f => hostname.includes(f))) return false;
//     return true;
//   } catch {
//     return false;
//   }
// }

// // ─────────────────────────────────────────────────────────────────────────────
// //  parseContactDetails — now returns email, email2, phone, phone2, website
// // ─────────────────────────────────────────────────────────────────────────────
// export function parseContactDetails(text) {
//   // ── Email ─────────────────────────────────────────────────────────────────
//   const emailRx  = /[a-zA-Z0-9._+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
//   const allEmails = [...new Set(text.match(emailRx) || [])]
//     .filter(isValidEmail);

//   // ── Phone ─────────────────────────────────────────────────────────────────
//   // Match tel: links first (most reliable), then regex
//   const telLinks = [...text.matchAll(/Phone:\s*([^\n]+)/g)]
//     .map(m => m[1].trim())
//     .filter(isValidPhone);

//   const phoneRx   = /(?:\+91[\s-]?)?[6-9]\d{9}|\+?\d[\d\s\-()+.]{7,14}/g;
//   const allPhones = [
//     ...telLinks,
//     ...[...new Set(text.match(phoneRx) || [])]
//       .map(p => p.trim())
//       .filter(isValidPhone),
//   ];
//   // Deduplicate phones by their digit-only form
//   const seenDigits = new Set();
//   const uniquePhones = allPhones.filter(p => {
//     const d = p.replace(/\D/g, '');
//     if (seenDigits.has(d)) return false;
//     seenDigits.add(d);
//     return true;
//   });

//   // ── Website ───────────────────────────────────────────────────────────────
//   // Collect from Website: lines first, then regex
//   const siteLines = [...text.matchAll(/Website:\s*(https?:\/\/[^\n\s]+)/g)]
//     .map(m => m[1].trim())
//     .filter(isValidWebsite);

//   const urlRx       = /https?:\/\/(?:www\.)?[a-zA-Z0-9-]+\.[a-zA-Z]{2,}(?:\/[^\s<"')]*)?/g;
//   const allWebsites = [
//     ...siteLines,
//     ...[...new Set(text.match(urlRx) || [])].filter(isValidWebsite),
//   ];
//   const uniqueWebsites = [...new Set(allWebsites)];

//   return {
//     email:   allEmails[0]      || 'Not Found',
//     email2:  allEmails[1]      || '',           // second email → Email2 column
//     phone:   uniquePhones[0]   || 'Not Found',
//     phone2:  uniquePhones[1]   || '',           // second phone → Phone2 column
//     website: uniqueWebsites[0] || 'Not Found',
//   };
// }

// // ─────────────────────────────────────────────────────────────────────────────
// //  CLEAN THEN ARCHIVE
// //  Reads every .xlsx in PROFILE_DIR, cleans all sheets, moves to MOVED_DIR
// // ─────────────────────────────────────────────────────────────────────────────
// export async function cleanAndArchive() {
//   fs.mkdirSync(cfg.PROFILE_DIR, { recursive: true });
//   fs.mkdirSync(cfg.MOVED_DIR,   { recursive: true });

//   const files = fs.readdirSync(cfg.PROFILE_DIR).filter(f => f.endsWith('.xlsx'));
//   if (!files.length) { console.log('[Archive] Nothing to archive.'); return; }

//   const stamp = formatTimestamp(new Date());

//   for (const file of files) {
//     const srcPath = path.join(cfg.PROFILE_DIR, file);
//     console.log(`\n[Clean] Processing: ${file}`);

//     try {
//       const wb = new ExcelJS.Workbook();
//       await wb.xlsx.readFile(srcPath);

//       let totalFixed = 0;

//       // Clean every sheet (multi-sheet workbook)
//       for (const ws of wb.worksheets) {
//         // Find col indices from this sheet's header row
//         const headerRow = ws.getRow(1);
//         const colIdx    = {};
//         headerRow.eachCell((cell, colNum) => {
//           colIdx[String(cell.value ?? '').trim()] = colNum;
//         });

//         const emailCol   = colIdx['Email']   || COL.Email;
//         const email2Col  = colIdx['Email2']  || COL.Email2;
//         const phoneCol   = colIdx['PhoneNo'] || COL.Phone;
//         const phone2Col  = colIdx['Phone2']  || COL.Phone2;
//         const websiteCol = colIdx['Website'] || COL.Website;

//         let sheetFixed = 0;
//         ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
//           if (rowNum === 1) return;

//           const rawEmail   = String(row.getCell(emailCol).value   ?? '');
//           const rawEmail2  = String(row.getCell(email2Col).value  ?? '');
//           const rawPhone   = String(row.getCell(phoneCol).value   ?? '');
//           const rawPhone2  = String(row.getCell(phone2Col).value  ?? '');
//           const rawWebsite = String(row.getCell(websiteCol).value ?? '');

//           const cleanedEmail   = cleanEmailValue(rawEmail);
//           const cleanedEmail2  = cleanEmailValue(rawEmail2);
//           const cleanedPhone   = cleanPhoneValue(rawPhone);
//           const cleanedPhone2  = cleanPhoneValue(rawPhone2);
//           const cleanedWebsite = cleanWebsiteValue(rawWebsite);

//           if (cleanedEmail   !== rawEmail   ||
//               cleanedEmail2  !== rawEmail2  ||
//               cleanedPhone   !== rawPhone   ||
//               cleanedPhone2  !== rawPhone2  ||
//               cleanedWebsite !== rawWebsite) {
//             row.getCell(emailCol).value   = cleanedEmail;
//             row.getCell(email2Col).value  = cleanedEmail2;
//             row.getCell(phoneCol).value   = cleanedPhone;
//             row.getCell(phone2Col).value  = cleanedPhone2;
//             row.getCell(websiteCol).value = cleanedWebsite;
//             sheetFixed++;
//           }
//         });

//         if (sheetFixed > 0) {
//           console.log(`  [Clean] Sheet "${ws.name}": fixed ${sheetFixed} rows`);
//           totalFixed += sheetFixed;
//         }
//       }

//       console.log(`[Clean] Total fixed: ${totalFixed} rows`);
//       await wb.xlsx.writeFile(srcPath);

//     } catch (err) {
//       console.error(`[Clean] Error processing ${file}: ${err.message}`);
//       // Still move the file even if cleaning failed
//     }

//     const base     = path.basename(file, '.xlsx');
//     const destPath = path.join(cfg.MOVED_DIR, `${base}-${stamp}.xlsx`);
//     fs.renameSync(srcPath, destPath);
//     console.log(`[Archive] Moved → MOVED_FILES/${base}-${stamp}.xlsx`);
//   }
// }

// // ─────────────────────────────────────────────────────────────────────────────
// //  Cell-level cleaners (used by cleanAndArchive)
// // ─────────────────────────────────────────────────────────────────────────────
// function cleanEmailValue(raw) {
//   if (!raw || ['Not Found','ERROR','No URL','None',''].includes(raw.trim())) return raw;
//   const parts = raw.split(/[,;\n]+/).map(e => e.trim().toLowerCase()).filter(Boolean);
//   const valid  = parts.filter(isValidEmail);
//   return valid.length ? valid.join(', ') : 'Not Found';
// }

// function cleanPhoneValue(raw) {
//   if (!raw || ['Not Found','ERROR','No URL','None',''].includes(raw.trim())) return raw;

//   // Split on newlines — handles "9847500075\n\n9847500075"
//   const parts  = raw.split(/\n+/).map(p => p.trim()).filter(Boolean);
//   const unique = [...new Set(parts)];
//   const valid  = unique.filter(isValidPhone);

//   if (valid.length > 0) {
//     valid.sort((a, b) => b.replace(/\D/g, '').length - a.replace(/\D/g, '').length);
//     return valid[0];
//   }
//   // Try joining fragments: "1352\n\n 7346" → "13527346"
//   const joined = parts.join('').replace(/\s/g, '');
//   if (isValidPhone(joined)) return joined;

//   return 'Not Found';
// }

// function cleanWebsiteValue(raw) {
//   if (!raw || ['Not Found','ERROR','No URL','None',''].includes(raw.trim())) return raw;
//   const url = raw.trim().toLowerCase();
//   return isValidWebsite(url) ? url : 'Not Found';
// }

// // ─────────────────────────────────────────────────────────────────────────────
// //  Other utilities
// // ─────────────────────────────────────────────────────────────────────────────
// export function writeTextFile(filePath, content) {
//   fs.mkdirSync(path.dirname(filePath), { recursive: true });
//   fs.writeFileSync(filePath, content, 'utf8');
// }

// export function readTextFile(filePath) {
//   return fs.existsSync(filePath) ? fs.readFileSync(filePath, 'utf8') : '';
// }

// export function formatTimestamp(date) {
//   const p = n => String(n).padStart(2, '0');
//   return `${date.getFullYear()}${p(date.getMonth()+1)}${p(date.getDate())}_${p(date.getHours())}${p(date.getMinutes())}${p(date.getSeconds())}`;
// }

// export function ensureDir(dir) {
//   fs.mkdirSync(dir, { recursive: true });
// }




// src/fileUtils.js
import fs      from 'fs';
import path    from 'path';
import ExcelJS from 'exceljs';
import cfg     from '../config.js';
import { sanitizeSheetName, HEADERS, COL } from './excelHelper.js';

// ─────────────────────────────────────────────────────────────────────────────
//  VALIDATORS — used both during scraping AND during clean+archive
// ─────────────────────────────────────────────────────────────────────────────

// Valid Indian/international phone: 7-15 digits, not a date, not a fragment
export function isValidPhone(raw) {
  if (!raw || typeof raw !== 'string') return false;
  const p = raw.trim();

  // Reject date patterns: dd.mm.yyyy, dd-mm-yy, dd/mm/yyyy etc.
  if (/^\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}\.?$/.test(p)) return false;

  // Reject strings ending with a dot that are clearly reference numbers
  if (/^\d+\.$/.test(p)) return false;

  // Reject Facebook post IDs (15-16 digits starting with common prefixes)
  const digits = p.replace(/[\s\-().+]/g, '');
  if (digits.length > 15) return false;
  if (digits.length < 7)  return false;
  if (!/^\d+$/.test(digits)) return false;

  // Reject PIN codes / short zip codes (exactly 6 digits with a dot or space)
  if (/^\d{6}\.$/.test(p) || /^\d{3}\s\d{3}$/.test(p)) return false;

  return true;
}

// Valid email: standard format, not a BNI internal address
export function isValidEmail(raw) {
  if (!raw || typeof raw !== 'string') return false;
  const e = raw.trim().toLowerCase();
  if (!/^[a-zA-Z0-9._+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(e)) return false;
  if (e.includes('bniconnect') || e.includes('@bni.')) return false;
  return true;
}

// Valid website: must parse as URL with a real TLD, not a BNI internal link
export function isValidWebsite(raw) {
  if (!raw || typeof raw !== 'string') return false;
  const url = raw.trim().toLowerCase();
  try {
    const parsed   = new URL(url);
    const hostname = parsed.hostname;
    const parts    = hostname.split('.');
    if (parts.length < 2) return false;
    const tld = parts[parts.length - 1];
    if (!/^[a-z]{2,6}$/.test(tld)) return false;
    // Reject BNI domains
    if (hostname.includes('bniconnectglobal') || hostname.includes('bni.com')
        || hostname.includes('bni.in')) return false;
    // Reject known incomplete/fake domains
    const fakeDomains = ['comingsoon', 'coming soon', 'zendesk'];
    if (fakeDomains.some(f => hostname.includes(f))) return false;
    return true;
  } catch {
    return false;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  parseContactDetails — now returns email, email2, phone, phone2, website
// ─────────────────────────────────────────────────────────────────────────────
export function parseContactDetails(text) {
  // ── Email ─────────────────────────────────────────────────────────────────
  const emailRx  = /[a-zA-Z0-9._+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
  const allEmails = [...new Set(text.match(emailRx) || [])]
    .filter(isValidEmail);

  // ── Phone ─────────────────────────────────────────────────────────────────
  // Match tel: links first (most reliable), then regex
  const telLinks = [...text.matchAll(/Phone:\s*([^\n]+)/g)]
    .map(m => m[1].trim())
    .filter(isValidPhone);

  const phoneRx   = /(?:\+91[\s-]?)?[6-9]\d{9}|\+?\d[\d\s\-()+.]{7,14}/g;
  const allPhones = [
    ...telLinks,
    ...[...new Set(text.match(phoneRx) || [])]
      .map(p => p.trim())
      .filter(isValidPhone),
  ];
  // Deduplicate phones by their digit-only form
  const seenDigits = new Set();
  const uniquePhones = allPhones.filter(p => {
    const d = p.replace(/\D/g, '');
    if (seenDigits.has(d)) return false;
    seenDigits.add(d);
    return true;
  });

  // ── Website ───────────────────────────────────────────────────────────────
  // Collect from Website: lines first, then regex
  const siteLines = [...text.matchAll(/Website:\s*(https?:\/\/[^\n\s]+)/g)]
    .map(m => m[1].trim())
    .filter(isValidWebsite);

  const urlRx       = /https?:\/\/(?:www\.)?[a-zA-Z0-9-]+\.[a-zA-Z]{2,}(?:\/[^\s<"')]*)?/g;
  const allWebsites = [
    ...siteLines,
    ...[...new Set(text.match(urlRx) || [])].filter(isValidWebsite),
  ];
  const uniqueWebsites = [...new Set(allWebsites)];

  return {
    email:   allEmails[0]      || 'Not Found',
    email2:  allEmails[1]      || '',           // second email → Email2 column
    phone:   uniquePhones[0]   || 'Not Found',
    phone2:  uniquePhones[1]   || '',           // second phone → Phone2 column
    website: uniqueWebsites[0] || 'Not Found',
  };
}

// ─────────────────────────────────────────────────────────────────────────────
//  CLEAN THEN ARCHIVE
//  Reads every .xlsx in PROFILE_DIR, cleans all sheets, moves to MOVED_DIR
// ─────────────────────────────────────────────────────────────────────────────
export async function cleanAndArchive() {
  fs.mkdirSync(cfg.PROFILE_DIR, { recursive: true });
  fs.mkdirSync(cfg.MOVED_DIR,   { recursive: true });

  const files = fs.readdirSync(cfg.PROFILE_DIR).filter(f => f.endsWith('.xlsx'));
  if (!files.length) { console.log('[Archive] Nothing to archive.'); return; }

  const stamp = formatTimestamp(new Date());

  for (const file of files) {
    const srcPath = path.join(cfg.PROFILE_DIR, file);
    console.log(`\n[Clean] Processing: ${file}`);

    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(srcPath);

      let totalFixed = 0;

      // Clean every sheet (multi-sheet workbook)
      for (const ws of wb.worksheets) {
        if (ws.name === '_CategoryMap') continue; // skip internal metadata sheet
        // Find col indices from this sheet's header row
        const headerRow = ws.getRow(1);
        const colIdx    = {};
        headerRow.eachCell((cell, colNum) => {
          colIdx[String(cell.value ?? '').trim()] = colNum;
        });

        const emailCol   = colIdx['Email']   || COL.Email;
        const email2Col  = colIdx['Email2']  || COL.Email2;
        const phoneCol   = colIdx['PhoneNo'] || COL.Phone;
        const phone2Col  = colIdx['Phone2']  || COL.Phone2;
        const websiteCol = colIdx['Website'] || COL.Website;

        let sheetFixed = 0;
        ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
          if (rowNum === 1) return;

          const rawEmail   = String(row.getCell(emailCol).value   ?? '');
          const rawEmail2  = String(row.getCell(email2Col).value  ?? '');
          const rawPhone   = String(row.getCell(phoneCol).value   ?? '');
          const rawPhone2  = String(row.getCell(phone2Col).value  ?? '');
          const rawWebsite = String(row.getCell(websiteCol).value ?? '');

          const cleanedEmail   = cleanEmailValue(rawEmail);
          const cleanedEmail2  = cleanEmailValue(rawEmail2);
          const cleanedPhone   = cleanPhoneValue(rawPhone);
          const cleanedPhone2  = cleanPhoneValue(rawPhone2);
          const cleanedWebsite = cleanWebsiteValue(rawWebsite);

          if (cleanedEmail   !== rawEmail   ||
              cleanedEmail2  !== rawEmail2  ||
              cleanedPhone   !== rawPhone   ||
              cleanedPhone2  !== rawPhone2  ||
              cleanedWebsite !== rawWebsite) {
            row.getCell(emailCol).value   = cleanedEmail;
            row.getCell(email2Col).value  = cleanedEmail2;
            row.getCell(phoneCol).value   = cleanedPhone;
            row.getCell(phone2Col).value  = cleanedPhone2;
            row.getCell(websiteCol).value = cleanedWebsite;
            sheetFixed++;
          }
        });

        if (sheetFixed > 0) {
          console.log(`  [Clean] Sheet "${ws.name}": fixed ${sheetFixed} rows`);
          totalFixed += sheetFixed;
        }
      }

      console.log(`[Clean] Total fixed: ${totalFixed} rows`);
      await wb.xlsx.writeFile(srcPath);

    } catch (err) {
      console.error(`[Clean] Error processing ${file}: ${err.message}`);
      // Still move the file even if cleaning failed
    }

    const base     = path.basename(file, '.xlsx');
    const destPath = path.join(cfg.MOVED_DIR, `${base}-${stamp}.xlsx`);
    fs.renameSync(srcPath, destPath);
    console.log(`[Archive] Moved → MOVED_FILES/${base}-${stamp}.xlsx`);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  Cell-level cleaners (used by cleanAndArchive)
// ─────────────────────────────────────────────────────────────────────────────
function cleanEmailValue(raw) {
  if (!raw || ['Not Found','ERROR','No URL','None',''].includes(raw.trim())) return raw;
  const parts = raw.split(/[,;\n]+/).map(e => e.trim().toLowerCase()).filter(Boolean);
  const valid  = parts.filter(isValidEmail);
  return valid.length ? valid.join(', ') : 'Not Found';
}

function cleanPhoneValue(raw) {
  if (!raw || ['Not Found','ERROR','No URL','None',''].includes(raw.trim())) return raw;

  // Split on newlines — handles "9847500075\n\n9847500075"
  const parts  = raw.split(/\n+/).map(p => p.trim()).filter(Boolean);
  const unique = [...new Set(parts)];
  const valid  = unique.filter(isValidPhone);

  if (valid.length > 0) {
    valid.sort((a, b) => b.replace(/\D/g, '').length - a.replace(/\D/g, '').length);
    return valid[0];
  }
  // Try joining fragments: "1352\n\n 7346" → "13527346"
  const joined = parts.join('').replace(/\s/g, '');
  if (isValidPhone(joined)) return joined;

  return 'Not Found';
}

function cleanWebsiteValue(raw) {
  if (!raw || ['Not Found','ERROR','No URL','None',''].includes(raw.trim())) return raw;
  const url = raw.trim().toLowerCase();
  return isValidWebsite(url) ? url : 'Not Found';
}

// ─────────────────────────────────────────────────────────────────────────────
//  Other utilities
// ─────────────────────────────────────────────────────────────────────────────
export function writeTextFile(filePath, content) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content, 'utf8');
}

export function readTextFile(filePath) {
  return fs.existsSync(filePath) ? fs.readFileSync(filePath, 'utf8') : '';
}

export function formatTimestamp(date) {
  const p = n => String(n).padStart(2, '0');
  return `${date.getFullYear()}${p(date.getMonth()+1)}${p(date.getDate())}_${p(date.getHours())}${p(date.getMinutes())}${p(date.getSeconds())}`;
}

export function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}