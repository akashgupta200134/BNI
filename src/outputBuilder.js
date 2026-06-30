// // src/outputBuilder.js
// import path              from 'path';
// import cfg               from '../config.js';
// import * as excelHelper  from './excelHelper.js';
// import { registerCategories } from './excelHelper.js';
// import * as fileUtils    from './fileUtils.js';

// export function getOutputPath(countryName) {
//   fileUtils.ensureDir(cfg.PROFILE_DIR);
//   return path.join(cfg.PROFILE_DIR, `${countryName}.xlsx`);
// }

// export async function createOutputFile(countryName, categories) {
//   const filePath = getOutputPath(countryName);
//   // Register all category names FIRST so sheet name mapping is stable
//   // across the entire run — even if categories are added dynamically later
//   registerCategories(filePath, categories);
//   await excelHelper.createMultiSheetWorkbook(filePath, categories);
//   return filePath;
// }

// export async function appendMembersToSheet(countryName, categoryName, members) {
//   const filePath = getOutputPath(countryName);
//   console.log(`[Output] Writing ${members.length} members to sheet "${categoryName.substring(0, 31)}"`);
//   await excelHelper.appendRowsToSheet(filePath, categoryName, members);
// }

// export async function writeContactToRow(countryName, categoryName, memberName, contacts) {
//   const filePath = getOutputPath(countryName);

//   // Check which fields have real data
//   const hasEmail   = contacts.email   && contacts.email   !== 'Not Found';
//   const hasEmail2  = contacts.email2  && contacts.email2  !== '';
//   const hasPhone   = contacts.phone   && contacts.phone   !== 'Not Found';
//   const hasPhone2  = contacts.phone2  && contacts.phone2  !== '';
//   const hasWebsite = contacts.website && contacts.website !== 'Not Found';

//   // Nothing found at all — stamp NoData, leave contact cells blank
//   if (!hasEmail && !hasEmail2 && !hasPhone && !hasPhone2 && !hasWebsite) {
//     console.log(`  [NoData] No contact info for "${memberName}" — row left blank`);
//     await excelHelper.writeCellInSheet(filePath, categoryName, memberName, null);
//     return;
//   }

//   // Only pass real values — empty string for missing fields, no "Not Found" in cells
//   const cleaned = {
//     email:   hasEmail   ? contacts.email   : '',
//     email2:  hasEmail2  ? contacts.email2  : '',
//     phone:   hasPhone   ? contacts.phone   : '',
//     phone2:  hasPhone2  ? contacts.phone2  : '',
//     website: hasWebsite ? contacts.website : '',
//   };

//   await excelHelper.writeCellInSheet(filePath, categoryName, memberName, cleaned);
// }

// export async function isAlreadyDone(countryName, categoryName, memberName) {
//   const filePath = getOutputPath(countryName);
//   return excelHelper.isRowDoneInSheet(filePath, categoryName, memberName);
// }

// export async function readSheetRows(countryName, categoryName) {
//   const filePath = getOutputPath(countryName);
//   return excelHelper.readAllRows(filePath, categoryName);
// }





// src/outputBuilder.js
import path              from 'path';
import cfg               from '../config.js';
import * as excelHelper  from './excelHelper.js';
import * as fileUtils    from './fileUtils.js';

export function getOutputPath(countryName) {
  fileUtils.ensureDir(cfg.PROFILE_DIR);
  return path.join(cfg.PROFILE_DIR, `${countryName}.xlsx`);
}

export async function createOutputFile(countryName, categories) {
  const filePath = getOutputPath(countryName);
  // createMultiSheetWorkbook already calls registerCategories internally
  // (and reuses any mapping persisted in an existing file on disk).
  await excelHelper.createMultiSheetWorkbook(filePath, categories);
  return filePath;
}

export async function appendMembersToSheet(countryName, categoryName, members) {
  const filePath = getOutputPath(countryName);
  console.log(`[Output] Writing ${members.length} members to sheet "${categoryName.substring(0, 31)}"`);
  await excelHelper.appendRowsToSheet(filePath, categoryName, members);
}

export async function writeContactToRow(countryName, categoryName, memberName, contacts) {
  const filePath = getOutputPath(countryName);

  // Check which fields have real data
  const hasEmail   = contacts.email   && contacts.email   !== 'Not Found';
  const hasEmail2  = contacts.email2  && contacts.email2  !== '';
  const hasPhone   = contacts.phone   && contacts.phone   !== 'Not Found';
  const hasPhone2  = contacts.phone2  && contacts.phone2  !== '';
  const hasWebsite = contacts.website && contacts.website !== 'Not Found';

  // Nothing found at all — stamp NoData, leave contact cells blank
  if (!hasEmail && !hasEmail2 && !hasPhone && !hasPhone2 && !hasWebsite) {
    console.log(`  [NoData] No contact info for "${memberName}" — row left blank`);
    await excelHelper.writeCellInSheet(filePath, categoryName, memberName, null);
    return;
  }

  // Only pass real values — empty string for missing fields, no "Not Found" in cells
  const cleaned = {
    email:   hasEmail   ? contacts.email   : '',
    email2:  hasEmail2  ? contacts.email2  : '',
    phone:   hasPhone   ? contacts.phone   : '',
    phone2:  hasPhone2  ? contacts.phone2  : '',
    website: hasWebsite ? contacts.website : '',
  };

  await excelHelper.writeCellInSheet(filePath, categoryName, memberName, cleaned);
}

export async function isAlreadyDone(countryName, categoryName, memberName) {
  const filePath = getOutputPath(countryName);
  return excelHelper.isRowDoneInSheet(filePath, categoryName, memberName);
}

export async function readSheetRows(countryName, categoryName) {
  const filePath = getOutputPath(countryName);
  return excelHelper.readAllRows(filePath, categoryName);
}