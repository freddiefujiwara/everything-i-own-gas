import { FLAG_VALUE, flaggedRoomUrls_ } from './admin.js';
import { sheetValuesToObjects, spreadsheetId_ } from './common.js';

// ---------------------------------------------------------------------------
// Blog article draft generation (prototype).
//
// generateListArticleDraft() turns the item sheets into a Google Doc, so
// blog-gas can render it as an article: items from every tab are merged
// and grouped under one heading per genre, and for each item the URL
// comes first as its own standalone paragraph (the shape blog-gas turns
// into a link card), followed by the item's comment as plain text below
// the card. The output is one fixed dummy document (a staging
// dummy article) whose ID is kept in the script property
// LIST_ARTICLE_DUMMY_DOC_ID; every run replaces its body. The generator
// is manual-run: setupListArticleTrigger() installs the every-6-hours
// trigger, and is itself run manually, only after the dummy output has
// been verified.
//
// Privacy rules for this public-facing draft: per-item prices are fine,
// but no grand total or sum is ever written, and the subscriptions
// section lists service names and official-site URLs only (no amounts,
// no school names).
// ---------------------------------------------------------------------------

const LIST_ARTICLE_DRAFT_FOLDER_ID = '1tTl57xXMEYIqKmo8L0EbtRj7N1VfrNAC';
const LIST_ARTICLE_DUMMY_DOC_PROPERTY = 'LIST_ARTICLE_DUMMY_DOC_ID';

const LIST_ARTICLE_SHEETS = ['room', 'workman', 'variety', 'others'];

// Total text coercion for values that come from outside this code
// (sheet cells, JSON payloads): null and absent read as empty, and a
// value whose String() throws — an object with a non-callable
// toString, which JSON.parse can produce — also reads as empty
// instead of killing the shaping it flows through.
const cellText_ = value => {
  if (value === null || value === undefined) {
    return '';
  }
  try {
    return String(value);
  } catch (e) {
    return '';
  }
};

// A number as a cell carries it: a finite number, or a string that
// parses as one. Everything else (booleans, arrays, Dates, objects)
// is not a number — a bare Number() would turn true into 1, [500]
// into 500 and a Date into its concatenated digits.
const cellNumber_ = value => {
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value === 'string' && value.trim() !== '') {
    return Number(value);
  }
  return NaN;
};

// Format a price cell ("2029", 2029, "¥4,380") for display; blank or
// non-numeric cells produce no price text. Only a number or a string
// cell can carry a price: any other type (a boolean, a Date, an
// object), and a string with no digits in it, produces no price text
// either — the digit-stripping below would otherwise turn them into
// a plausible-looking ¥0 or worse.
const formatItemPrice = value => {
  if (typeof value !== 'number' && typeof value !== 'string') {
    return '';
  }
  const cleaned = String(value).replace(/[^\d.-]/g, '');
  if (cleaned === '') {
    return '';
  }
  const n = Number(cleaned);
  if (!isFinite(n)) {
    return '';
  }
  return `¥${Math.round(n).toLocaleString('en-US')}`;
};

// One display line for an item: title, price, unit (only when > 1) and
// comment. Never any total.
const buildItemLine = item => {
  const title = cellText_(item.title).trim();
  const details = [];
  const price = formatItemPrice(item.price);
  if (price) {
    details.push(price);
  }
  const unit = cellNumber_(item.unit);
  if (isFinite(unit) && unit > 1) {
    details.push(`×${unit}`);
  }
  const head = details.length ? `${title}（${details.join(' ')}）` : title;
  const comment = cellText_(item.comment).trim();
  return comment ? `${head} ${comment}` : head;
};

// URLs in the sheet can contain raw Japanese text and spaces (Google
// search links). Written into the Doc as-is they are not recognised as
// standalone URL paragraphs, so they never become link cards. Encode
// exactly the characters that break URL recognition — spaces and
// non-ASCII — and leave everything else (including existing %XX
// sequences) untouched. The u flag matters: without it the character
// class matches one UTF-16 code unit at a time, and
// encodeURIComponent throws URIError on a lone surrogate half — one
// emoji in a sheet URL killed the whole regeneration.
const normalizeArticleUrl = url =>
  String(url).replace(/ |[^\x00-\x7F]/gu, ch =>
    ch === ' ' ? '%20' : encodeURIComponent(ch)
  );

// ---------------------------------------------------------------------------
// Title splitting: shop titles (Rakuten's especially) bury the product
// name under promotional brackets and property-like tails (model
// numbers, sizes, colours, compatibility lists). splitTitle separates
// the two so surfaces can show just the name and move the rest into
// the card description. The split is heuristic and deliberately
// conservative: a title it cannot split comes back whole as the name.
// ---------------------------------------------------------------------------

// Bracketed segments whose content matches this are pure promotion
// and are dropped; other bracketed segments (【新品】, 【Office 2019
// H&B】...) count as properties.
const TITLE_PROMO_PATTERN =
  /オフ|%|倍|クーポン|1位|送料無料|セール|ポイント|特典|獲得|応援|最大|割引|半額|入荷|SALE/i;

const TITLE_COLOR_TOKENS = new Set([
  'ホワイト', 'ブラック', 'ブルー', 'レッド', 'グリーン', 'ピンク',
  'グレー', 'シルバー', 'ゴールド', 'ネイビー', 'ベージュ', 'ブラウン',
  'イエロー', 'パープル', 'オレンジ', 'クリア', 'アイボリー', 'カーキ',
]);

const TITLE_PROPERTY_TOKENS = new Set([
  '特大', '大サイズ', '中サイズ', '小サイズ', '本体', '単品', '専用', '兼用',
]);

// A token is property-like when it carries a spec: a model number or
// a number with a unit (digits mixed with anything else), a colour, a
// size/state word, or a compatibility marker. A bare number stays in
// the name — it is usually part of a model name (Anker 543).
const isPropertyToken = token => {
  if (/\d/.test(token)) {
    if (/^\d+$/.test(token)) {
      return false;
    }
    // A number with Japanese counter words (3色ボールペン, 12個セット)
    // reads as part of the product name, not as a spec.
    if (
      !/[A-Za-z]/.test(token) &&
      /(色|個セット|本セット|枚セット|個入り|本入り|枚入り)/.test(token)
    ) {
      return false;
    }
    return true;
  }
  if (TITLE_COLOR_TOKENS.has(token) || TITLE_PROPERTY_TOKENS.has(token)) {
    return true;
  }
  return /対応|向け|専用/.test(token);
};

const TITLE_NAME_SOFT_LIMIT = 30;

const splitTitle = rawTitle => {
  const title = cellText_(rawTitle).trim();
  if (!title) {
    return { name: '', props: '' };
  }
  const bracketProps = [];
  const withoutBrackets = title.replace(
    /【([^】]*)】|\[([^\]]*)\]/g,
    (_, g1, g2) => {
      const content = String(g1 ?? g2 ?? '').trim();
      if (content && !TITLE_PROMO_PATTERN.test(content)) {
        bracketProps.push(content);
      }
      return ' ';
    }
  );
  const tokens = withoutBrackets
    .split(/[\s　]+/)
    // Standalone '&' separators and a bare 送料無料 token are noise,
    // not name or properties.
    .filter(token => token !== '' && token !== '&' && token !== '＆' && token !== '送料無料');

  const nameTokens = [];
  const propTokens = [];
  let inProps = false;
  tokens.forEach(token => {
    if (!inProps && isPropertyToken(token)) {
      inProps = true;
    }
    if (!inProps && nameTokens.join(' ').length >= TITLE_NAME_SOFT_LIMIT) {
      inProps = true;
    }
    (inProps ? propTokens : nameTokens).push(token);
  });
  // A title that starts with a property token keeps its first token
  // as the name rather than ending up nameless.
  if (nameTokens.length === 0 && propTokens.length > 0) {
    nameTokens.push(propTokens.shift());
  }
  return {
    name: nameTokens.join(' '),
    props: [...bracketProps, ...propTokens].join(' '),
  };
};

// Build the whole article as a flat list of blocks:
// { type: 'heading1' | 'heading2' | 'itemTitle' | 'text' | 'link' |
// 'toc', text }. An 'itemTitle' is written bold into the Doc, so the
// converted markdown holds it as a bold-only paragraph — the shape
// the blog-wide [[アンカーリンク]] marker binds to. Every item and
// subscription entry ends with that marker, so the front-ends give
// each entry its own "この項目へのリンク" fragment link.
// Items from all tabs are merged and grouped by genre; genres appear in
// first-seen order and items keep sheet order within each genre. An
// item with a URL renders as a link card, so only its comment follows
// as text; an item without a URL falls back to the full text line.
// The subscriptions section comes last: names and their purpose
// notes, plus a link card for services whose sheet row carries a
// URL. Amounts never appear.
const buildListArticleBlocks = (itemsByTab, subscriptions) => {
  const blocks = [];
  const push = (type, text) => blocks.push({ type, text });

  const genreOrder = [];
  const byGenre = new Map();
  LIST_ARTICLE_SHEETS.forEach(sheet => {
    (itemsByTab[sheet] || [])
      .filter(item => cellText_(item.title).trim() !== '')
      .forEach(item => {
        const genre = cellText_(item.genre).trim() || 'その他';
        if (!byGenre.has(genre)) {
          byGenre.set(genre, []);
          genreOrder.push(genre);
        }
        byGenre.get(genre).push(item);
      });
  });

  let itemNumber = 0;
  genreOrder.forEach(genre => {
    push('heading2', genre);
    push('toc');
    // Each item's name, numbered with one running sequence across
    // all genres, sits directly above its card as a bold paragraph —
    // deliberately NOT a heading, so the article's [[目次]] TOCs
    // (one per genre, from the 'toc' block above) list the genres
    // only, not every item. The comment follows below the card, and
    // the entry closes with the [[アンカーリンク]] marker bound to
    // the bold name.
    byGenre.get(genre).forEach(item => {
      itemNumber += 1;
      const { name: title } = splitTitle(item.title);
      const url = cellText_(item.url).trim();
      if (!url) {
        push('itemTitle', `${itemNumber}. ${buildItemLine({ ...item, title })}`);
        push('text', '[[アンカーリンク]]');
        return;
      }
      push('itemTitle', `${itemNumber}. ${title}`);
      // A flagged （整理対象） item gets the blog-wide
      // [[画像グレーアウト]] marker right before its link, so its
      // card grays out on the front-ends, like a cancelled
      // subscription's does.
      if (cellText_(item.flag).trim() === FLAG_VALUE) {
        push('text', '[[画像グレーアウト]]');
      }
      push('link', normalizeArticleUrl(url));
      const comment = cellText_(item.comment).trim();
      if (comment) {
        push('text', comment);
      }
      push('text', '[[アンカーリンク]]');
    });
  });

  push('heading2', 'サブスクリプション');
  push('toc');
  (subscriptions || []).forEach(({ name, purpose, url, flag }) => {
    push('itemTitle', cellText_(name));
    const note = cellText_(purpose).trim();
    if (note) {
      push('text', note);
    }
    const clean = cellText_(url).trim();
    if (clean) {
      // A subscription whose purpose says it was cancelled, or that
      // carries the tidy-up flag, gets the blog-wide
      // [[画像グレーアウト]] marker right before its link, so
      // blog-gas flags the card and the front-ends gray it out.
      if (note.includes('解約') || cellText_(flag).trim() === FLAG_VALUE) {
        push('text', '[[画像グレーアウト]]');
      }
      push('link', normalizeArticleUrl(clean));
    }
    push('text', '[[アンカーリンク]]');
  });
  return blocks;
};

// Services from the subscriptions sheet's raw values (header row
// included), as { name, purpose, url, flag } entries. Amounts are
// never read into the output. Header, blank, total and school rows
// are skipped: the draft is public-facing.
const extractSubscriptions = values => {
  const objects = sheetValuesToObjects(values.slice());
  return objects
    .map(obj => ({
      name: cellText_(obj.Service).trim(),
      purpose: cellText_(obj.Purpose).trim(),
      url: cellText_(obj.URL).trim(),
      flag: cellText_(obj.flag).trim(),
    }))
    .filter(
      ({ name }) =>
        name !== '' &&
        !/合計|total/i.test(name) &&
        !/学校|中学校|高等学校/.test(name)
    );
};

// Read every source sheet up front, so the Doc body is only replaced
// once the full content exists (fetch-before-clear, same pattern as
// crawlRoom).
const readListArticleSource = () => {
  const ss = SpreadsheetApp.openById(spreadsheetId_());
  const itemsByTab = {};
  LIST_ARTICLE_SHEETS.forEach(sheet => {
    const values = ss.getSheetByName(sheet).getDataRange().getValues();
    itemsByTab[sheet] = sheetValuesToObjects(values.slice());
  });
  // Room flags live in the flags sheet (the room sheet itself is
  // rewritten by the ROOM sync); merge them in by URL so a flagged
  // room item grays out in the article like any other flagged item.
  const flaggedUrls = flaggedRoomUrls_();
  (itemsByTab.room || []).forEach(item => {
    item.flag = flaggedUrls.has(String(item.url ?? '').trim()) ? FLAG_VALUE : '';
  });
  const subscriptionValues = ss
    .getSheetByName('subscriptions')
    .getDataRange()
    .getValues();
  return {
    itemsByTab,
    subscriptions: extractSubscriptions(subscriptionValues),
  };
};

// The one fixed dummy document the generator writes into. Created in
// the blog drafts folder on first use; its ID is stored in a script
// property so every later run reuses (and replaces) the same Doc.
const getListArticleDoc = () => {
  const properties = PropertiesService.getScriptProperties();
  const docId = properties.getProperty(LIST_ARTICLE_DUMMY_DOC_PROPERTY);
  if (docId) {
    return DocumentApp.openById(docId);
  }
  const doc = DocumentApp.create('持ち物リスト（試作）');
  DriveApp.getFileById(doc.getId()).moveTo(
    DriveApp.getFolderById(LIST_ARTICLE_DRAFT_FOLDER_ID)
  );
  properties.setProperty(LIST_ARTICLE_DUMMY_DOC_PROPERTY, doc.getId());
  return doc;
};

function generateListArticleDraft() {
  // Build the entire content in memory before touching the Doc, so a
  // failed read can never leave a half-written document behind.
  const { itemsByTab, subscriptions } = readListArticleSource();
  const blocks = buildListArticleBlocks(itemsByTab, subscriptions);

  const doc = getListArticleDoc();
  // The Doc's file name is the article's public title. A regeneration
  // rewrites only the body and must never change the name: the Doc
  // was created as 持ち物リスト（試作） and renamed for publication,
  // and a save once restored the creation name over the published
  // one (2026-11-07). Capture the current name and put it back if
  // the save changed it.
  const file = DriveApp.getFileById(doc.getId());
  const nameBefore = file.getName();
  const body = doc.getBody();
  body.clear();
  blocks.forEach(block => {
    const { type, text } = block;
    const paragraph = body.appendParagraph(type === 'toc' ? '[[目次]]' : text);
    if (type === 'heading1') {
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    } else if (type === 'heading2') {
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    }
    // Character formatting carries over to each newly appended
    // paragraph, so bold is set explicitly on every paragraph: item
    // titles bold (the converted markdown then holds them as
    // bold-only paragraphs — the title shape the [[アンカーリンク]]
    // marker binds to on the front-ends), everything else not.
    // Without this, the comment and the marker paragraph itself
    // would inherit the title's bold, and a bold marker no longer
    // matches as a marker.
    paragraph.editAsText().setBold(type === 'itemTitle');
  });
  doc.saveAndClose();
  if (file.getName() !== nameBefore) {
    file.setName(nameBefore);
  }
  Logger.log(`Updated list article draft: ${doc.getUrl()}`);
  return doc.getUrl();
}

// Install the every-6-hours trigger for generateListArticleDraft().
// Idempotent: all existing triggers for the handler are deleted first,
// so re-running this can never stack duplicates. Run manually, and only
// after the dummy document's output has been verified — nothing calls
// this automatically.
function setupListArticleTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'generateListArticleDraft')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger('generateListArticleDraft')
    .timeBased()
    .everyHours(6)
    .create();
}

export { buildItemLine, buildListArticleBlocks, extractSubscriptions, generateListArticleDraft, normalizeArticleUrl, setupListArticleTrigger, splitTitle };
