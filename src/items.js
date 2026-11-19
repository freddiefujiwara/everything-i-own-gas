// The public items API behind doGet: every item sheet read as
// JSON, served through a chunked CacheService cache.

import { splitTitle } from './article.js';
import { sheetValuesToObjects, spreadsheetId_ } from './common.js';

// A JSONP callback must be a plain JavaScript identifier path; anything
// else would be reflected verbatim into a JavaScript response (the same
// validation bucket-list-gas applies).
const isValidCallback_ = (callback) =>
  typeof callback === 'string' &&
  /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback);

// Helper to create the final output for doGet
const createJsonpOutput = (data, callback) => {
  const output = ContentService.createTextOutput();
  if (callback && isValidCallback_(callback)) {
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
    output.setContent(`${callback}&&${callback}(${JSON.stringify(data)});`);
  } else {
    output.setMimeType(ContentService.MimeType.JSON);
    output.setContent(JSON.stringify(data));
  }
  return output;
};

// Response cache for doGet, kept in the script cache so it is shared
// across all visitors and survives cold starts.
//
// The built response is about 101 KB of JSON, which exceeds the
// CacheService limit of 100 KB per value, so the JSON string is split
// into chunks stored under separate keys. A small manifest key records
// the chunk count and the total length, and a cached entry is only
// trusted when every chunk is present, the joined length matches, and
// the result parses as JSON; anything else falls back to the sheets.
//
// Freshness is TTL-based: a cache entry lives for
// ITEMS_CACHE_TTL_SECONDS, so sheet edits show up within that window
// at the latest. crawlRoom clears the cache immediately after it
// rewrites the room sheet. The TTL is the CacheService maximum (6h):
// a full sheet read takes seconds, and with this site's traffic a
// 10-minute TTL left most visits on the slow uncached path.
const ITEMS_CACHE_KEY_PREFIX = 'items-response:v2:';
const ITEMS_CACHE_META_KEY = `${ITEMS_CACHE_KEY_PREFIX}meta`;
const ITEMS_CACHE_TTL_SECONDS = 21600;
const ITEMS_CACHE_CHUNK_CHARS = 20000;

// Split a string into chunks of at most chunkSize characters, without
// splitting a surrogate pair across a chunk boundary.
const chunkString = (str, chunkSize) => {
  const chunks = [];
  let start = 0;
  while (start < str.length) {
    let end = Math.min(start + chunkSize, str.length);
    if (end < str.length) {
      const code = str.charCodeAt(end - 1);
      if (code >= 0xd800 && code <= 0xdbff) {
        end += 1;
      }
    }
    chunks.push(str.slice(start, end));
    start = end;
  }
  return chunks;
};

const readItemsCache = () => {
  try {
    const cache = CacheService.getScriptCache();
    const metaText = cache.get(ITEMS_CACHE_META_KEY);
    if (!metaText) {
      return null;
    }
    const meta = JSON.parse(metaText);
    const keys = [];
    for (let i = 0; i < meta.chunks; i++) {
      keys.push(`${ITEMS_CACHE_KEY_PREFIX}${i}`);
    }
    const values = cache.getAll(keys);
    const parts = keys.map(key => values[key]);
    if (parts.some(part => typeof part !== 'string')) {
      return null;
    }
    const json = parts.join('');
    if (json.length !== meta.length) {
      return null;
    }
    return JSON.parse(json);
  } catch (e) {
    // A broken cache must never break the response.
    return null;
  }
};

const writeItemsCache = data => {
  try {
    const cache = CacheService.getScriptCache();
    const json = JSON.stringify(data);
    const chunks = chunkString(json, ITEMS_CACHE_CHUNK_CHARS);
    const entries = {
      [ITEMS_CACHE_META_KEY]: JSON.stringify({
        chunks: chunks.length,
        length: json.length,
      }),
    };
    chunks.forEach((chunk, index) => {
      entries[`${ITEMS_CACHE_KEY_PREFIX}${index}`] = chunk;
    });
    cache.putAll(entries, ITEMS_CACHE_TTL_SECONDS);
  } catch (e) {
    // Caching is best-effort: the response is served either way.
  }
};

const clearItemsCache = () => {
  try {
    const cache = CacheService.getScriptCache();
    const metaText = cache.get(ITEMS_CACHE_META_KEY);
    const keys = [ITEMS_CACHE_META_KEY];
    if (metaText) {
      const meta = JSON.parse(metaText);
      for (let i = 0; i < meta.chunks; i++) {
        keys.push(`${ITEMS_CACHE_KEY_PREFIX}${i}`);
      }
    }
    cache.removeAll(keys);
  } catch (e) {
    // Best-effort: a leftover entry expires on its own via the TTL.
  }
};

// Read every item sheet and build the response data. Each item gains
// a shortTitle (just the product name) and props (the property-like
// remainder of the title), split by splitTitle, so card and article
// surfaces can show the short name and put the properties in the
// card description. The original title is left untouched.
const fetchAllItems = () => {
  const ss = SpreadsheetApp.openById(spreadsheetId_());
  const sheetNames = ["room", "workman", "variety", "others"];

  return sheetNames.flatMap(name => {
    const sheet = ss.getSheetByName(name);
    const values = sheet.getDataRange().getValues();
    return sheetValuesToObjects(values).map(item => {
      const { name: shortTitle, props } = splitTitle(item.title);
      return { ...item, shortTitle, props };
    });
  });
};

export { clearItemsCache, createJsonpOutput, fetchAllItems, readItemsCache, writeItemsCache };
