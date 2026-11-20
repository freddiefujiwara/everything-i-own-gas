// Rakuten ROOM import: crawlRoom() rewrites the room sheet from the
// Rakuten ROOM API, falling back to the home data server's ROOM
// export when the direct fetch does not work, then drops the items
// cache.

import { fetchJson, getSheet, scriptProperty_ } from './common.js';
import { clearItemsCache } from './items.js';

const ROOM_API_BASE_URL = 'https://room.rakuten.co.jp/api';

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
  return [
    item.name || '',
    rakutenThumbnailUrl(item.picture?.url || ''),
    item.url || '',
    item.price || '',
    1,
    c.name || '',
    x.content || '',
  ];
};

// Turn a collections list into sheet rows. The two sources below
// differ only in fetchCollects, which returns one collection's
// collects array.
const buildRows = (collections, fetchCollects) =>
  collections.flatMap(c => {
    Logger.log(c.name);
    return fetchCollects(c.id).map(x => createRow(x.item || {}, c, x));
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
  const collections =
    fetchJson(`${ROOM_API_BASE_URL}/${userId}/collections`).data || [];
  if (!collections.length) {
    throw new Error('ROOM API returned no collections');
  }
  return buildRows(
    collections,
    id => fetchJson(`${ROOM_API_BASE_URL}/${id}/collects?limit=100`).data || []
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
  const collections = fetchJson(`${base}/collections.json`).data || [];
  return buildRows(
    collections,
    id => fetchJson(`${base}/collects_${id}.json`).data || []
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
