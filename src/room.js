// Rakuten ROOM import: crawlRoom() rewrites the room sheet from the
// Rakuten ROOM API, falling back to the home data server's ROOM
// export when the direct fetch does not work, then drops the items
// cache.

import { fetchJson, getSheet, scriptProperty_ } from './common.js';
import { clearItemsCache } from './items.js';

const ROOM_API_BASE_URL = 'https://room.rakuten.co.jp/api';

// A fetch payload's data: a missing data property (an error
// payload like {"message": ...}) reads as an empty array, as it
// always has. A data property that is present but not an array
// means the source answered in a shape this import cannot use,
// which counts as that source failing — crawlRoom falls back, or
// keeps the previous sheet — instead of crashing deeper in the
// row building.
const dataArray_ = payload => {
  const data = payload && payload.data;
  if (data === null || data === undefined) {
    return [];
  }
  if (!Array.isArray(data)) {
    throw new Error('response data is not an array');
  }
  return data;
};

// Text as the sheet stores it: strings pass through and finite
// numbers stringify, while anything else (arrays, booleans, objects
// — including ones whose toString is not callable) reads as empty
// rather than landing in a cell, where setValues would reject it.
const roomText_ = value => {
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'number' && isFinite(value)) {
    return String(value);
  }
  return '';
};

const clearSheet = (sh = getSheet('room')) => {
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 7).clearContent();
  }
};

// Rakuten's image servers resize on the fly via an _ex=WxH query, so
// the sheet stores the resized URL: card thumbnails (a few hundred px
// wide) stop downloading the full-size product photo, and readers of
// the sheet — the everything-i-own front-end included — need no
// host-specific handling of their own. URLs on other hosts, and URLs
// that already carry a query string, are stored unchanged.
const rakutenThumbnailUrl = url => {
  if (
    typeof url !== 'string' ||
    !/^https:\/\/(?:tshop|shop)\.r10s\.jp\//.test(url) ||
    url.includes('?')
  ) {
    return url;
  }
  return `${url}?_ex=400x400`;
};

const createRow = (item, c, x) => {
  Logger.log(item.name);
  // The price keeps the sheet's convention (a number or a string,
  // blank when falsy); a price of any other type reads as blank.
  const price =
    typeof item.price === 'number' || typeof item.price === 'string'
      ? item.price || ''
      : '';
  return [
    roomText_(item.name),
    rakutenThumbnailUrl(roomText_(item.picture?.url)),
    roomText_(item.url),
    price,
    1,
    roomText_(c.name),
    roomText_(x.content),
  ];
};

// Turn a collections list into sheet rows. The two sources below
// differ only in fetchCollects, which returns one collection's
// collects array. Entries that are not objects at all (a null in a
// deformed payload) are skipped: one bad entry must not kill the
// whole crawl, and a non-object collect has no item worth a blank
// row in the sheet.
const buildRows = (collections, fetchCollects) =>
  collections
    .filter(c => c && typeof c === 'object')
    .flatMap(c => {
      Logger.log(c.name);
      return fetchCollects(c.id)
        .filter(x => x && typeof x === 'object')
        .map(x => createRow(x.item || {}, c, x));
    });

// Direct fetch from the Rakuten ROOM API. The user ID lives in the
// script property ROOM_USER_ID, not in this public source code. An
// unset property or an empty collections response counts as a
// failure, so crawlRoom falls back to the home data server.
const fetchRoomApiRows = () => {
  const userId = scriptProperty_('ROOM_USER_ID');
  if (!userId) {
    throw new Error('script property ROOM_USER_ID is not set');
  }
  const collections = dataArray_(
    fetchJson(`${ROOM_API_BASE_URL}/${userId}/collections`)
  );
  if (!collections.length) {
    throw new Error('ROOM API returned no collections');
  }
  return buildRows(collections, id =>
    dataArray_(fetchJson(`${ROOM_API_BASE_URL}/${id}/collects?limit=100`))
  );
};

// The home data server base URL lives in the script property
// DATA_BASE_URL, not in this public source code.
const dataBaseUrl_ = () => {
  const value = scriptProperty_('DATA_BASE_URL');
  if (!value) {
    throw new Error('script property DATA_BASE_URL is not set');
  }
  return value;
};

const fetchFallbackRows = () => {
  const base = dataBaseUrl_();
  const collections = dataArray_(fetchJson(`${base}/collections.json`));
  return buildRows(collections, id =>
    dataArray_(fetchJson(`${base}/collects_${id}.json`))
  );
};

// Fetch everything before touching the sheet: if every source fails,
// the previous data stays in place instead of leaving an emptied
// sheet behind (observed 2026-11-03, when one collects fetch
// returned a plain-text "OK" and the cleared sheet was never
// rewritten).
const fetchRows = () => {
  try {
    const rows = fetchRoomApiRows();
    Logger.log('source: room-api');
    return rows;
  } catch (e) {
    Logger.log(`ROOM API fetch failed (${e.message}); using the fallback`);
    const rows = fetchFallbackRows();
    Logger.log('source: fallback (a.ze.gs)');
    return rows;
  }
};

function crawlRoom() {
  const rows = fetchRows();

  Logger.log(`Fetched ${rows.length} items.`);
  clearSheet();
  if (rows.length) {
    getSheet('room').getRange(2, 1, rows.length, 7).setValues(rows);
  }

  // The room sheet was just rewritten, so drop the cached doGet response
  // instead of letting it age out via its TTL.
  clearItemsCache();
}

export { crawlRoom };
