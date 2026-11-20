// The admin API's item sheets: tidy-up flags and item CRUD on
// the hand-edited sheets (workman, variety, others), plus the
// read-only room listing whose flags live in the flags sheet.

import { getSheet, spreadsheetId_ } from './common.js';
import { clearItemsCache } from './items.js';

const ADMIN_SHEETS = ['workman', 'variety', 'others'];

// --- Tidy-up flags （整理対象） ---
// Entries the owner plans to sort out: possessions to dispose of,
// subscriptions to cancel. The item sheets and the subscriptions
// sheet carry the flag in their own extra last column (H and G).
// The room sheet is rewritten by the ROOM sync, so room flags live
// in a small flags sheet keyed by the item's URL and survive it.
const FLAG_VALUE = '整理対象';
const FLAG_SHEET_NAME = 'flags';
const FLAG_HEADERS = ['url'];

const getFlagSheet_ = () => {
  const ss = SpreadsheetApp.openById(spreadsheetId_());
  let sheet = ss.getSheetByName(FLAG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FLAG_SHEET_NAME);
    sheet.appendRow(FLAG_HEADERS);
  }
  return sheet;
};

const flaggedRoomUrls_ = () => {
  const values = getFlagSheet_().getDataRange().getValues();
  return new Set(
    values.slice(1).map(row => String(row[0] ?? '').trim()).filter(Boolean)
  );
};

const setRoomFlag_ = (url, flagged) => {
  const key = String(url ?? '').trim();
  if (!key) {
    throw new Error('url is required to flag a room item');
  }
  const sheet = getFlagSheet_();
  const values = sheet.getDataRange().getValues();
  const index = values.findIndex(
    (row, i) => i > 0 && String(row[0] ?? '').trim() === key
  );
  if (flagged && index < 0) {
    sheet.appendRow([key]);
  }
  if (!flagged && index >= 0) {
    sheet.deleteRow(index + 1);
  }
  return { url: key, flag: flagged ? FLAG_VALUE : '' };
};

// --- Item CRUD on the hand-edited sheets ---

// Only the hand-edited item sheets may be touched through the item
// actions; the room sheet is owned by crawlRoom(). The subscriptions
// sheet has its own shape and its own functions below.
const adminSheet_ = name => {
  if (!ADMIN_SHEETS.includes(name)) {
    throw new Error(`sheet is not editable: ${name}`);
  }
  return getSheet(name);
};

const readAdminItems_ = name => {
  const sheet = adminSheet_(name);
  const values = sheet.getDataRange().getValues();
  return values.slice(1).map((row, index) => ({
    row: index + 2,
    title: row[0] ?? '',
    image: row[1] ?? '',
    url: row[2] ?? '',
    price: row[3] ?? '',
    unit: row[4] ?? '',
    genre: row[5] ?? '',
    comment: row[6] ?? '',
    flag: row[7] ?? '',
  }));
};

// Room items are listed read-only: the sheet itself stays owned by
// crawlRoom(); only the tidy-up flag is writable, and it lives in
// the flags sheet (see above).
const readRoomItems_ = () => {
  const sheet = getSheet('room');
  const values = sheet.getDataRange().getValues();
  const flagged = flaggedRoomUrls_();
  return values.slice(1).map((row, index) => ({
    row: index + 2,
    title: row[0] ?? '',
    image: row[1] ?? '',
    url: row[2] ?? '',
    price: row[3] ?? '',
    unit: row[4] ?? '',
    genre: row[5] ?? '',
    comment: row[6] ?? '',
    flag: flagged.has(String(row[2] ?? '').trim()) ? FLAG_VALUE : '',
  }));
};

const setItemFlag_ = (name, row, flagged) => {
  const sheet = adminSheet_(name);
  const n = checkedRow_(sheet, row);
  const flag = flagged ? FLAG_VALUE : '';
  sheet.getRange(n, 8).setValue(flag);
  return { row: n, flag };
};

// Coerce one item payload into a sheet row in the fixed column order
// (title, image, url, price, unit, genre, comment). Price and unit are
// stored as numbers when they parse as one, blank otherwise.
const normalizeItemRow_ = item => {
  const src = item || {};
  // Text reads totally: a value whose String() throws (an object
  // with a non-callable toString, producible via the JSON body)
  // reads as blank instead of failing the whole mutation.
  const text = v => {
    if (v === null || v === undefined) return '';
    try {
      return String(v);
    } catch (e) {
      return '';
    }
  };
  // Price and unit are stored as numbers only when the payload
  // carries a number or a numeric string; a bare Number() would
  // store a boolean as 1 and an array as its single element.
  const number = v => {
    if (typeof v === 'number') return isFinite(v) ? v : '';
    if (typeof v === 'string' && v.trim() !== '') {
      const n = Number(v);
      return isFinite(n) ? n : '';
    }
    return '';
  };
  return [
    text(src.title),
    text(src.image),
    text(src.url),
    number(src.price),
    number(src.unit),
    text(src.genre),
    text(src.comment),
    text(src.flag),
  ];
};

const checkedRow_ = (sheet, row) => {
  const n = Number(row);
  if (!Number.isInteger(n) || n < 2 || n > sheet.getLastRow()) {
    throw new Error(`row out of range: ${row}`);
  }
  return n;
};

const createItem_ = (name, item) => {
  const sheet = adminSheet_(name);
  sheet.appendRow(normalizeItemRow_(item));
  clearItemsCache();
  return { row: sheet.getLastRow() };
};

const updateItem_ = (name, row, item) => {
  const sheet = adminSheet_(name);
  const n = checkedRow_(sheet, row);
  const cells = normalizeItemRow_(item);
  if (!item || item.flag === undefined) {
    // The flag is managed through set_flag; a plain edit keeps it.
    const current = sheet.getRange(n, 8).getValues()[0] || [];
    cells[7] = current[0] ?? '';
  }
  sheet.getRange(n, 1, 1, 8).setValues([cells]);
  clearItemsCache();
  return { row: n };
};

const deleteItem_ = (name, row) => {
  const sheet = adminSheet_(name);
  const n = checkedRow_(sheet, row);
  sheet.deleteRow(n);
  clearItemsCache();
  return { row: n };
};

// The distinct genre values across all item sheets (room included:
// the genres are shared by every item), in first-appearance order.
// The admin form offers them as a dropdown.
const listGenres_ = () => {
  const seen = [];
  ['room', 'workman', 'variety', 'others'].forEach(name => {
    const values = getSheet(name).getDataRange().getValues();
    values.slice(1).forEach(row => {
      const genre = row[5];
      if (genre && !seen.includes(genre)) {
        seen.push(genre);
      }
    });
  });
  return seen;
};

export { FLAG_VALUE, checkedRow_, createItem_, deleteItem_, flaggedRoomUrls_, listGenres_, normalizeItemRow_, readAdminItems_, readRoomItems_, setItemFlag_, setRoomFlag_, updateItem_ };
