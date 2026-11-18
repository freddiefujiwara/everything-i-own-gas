import {
  createItem_,
  deleteItem_,
  listGenres_,
  readAdminItems_,
  readRoomItems_,
  setItemFlag_,
  setRoomFlag_,
  updateItem_,
} from './admin.js';
import {
  issueOwnerSession_,
  issueShareSession_,
  requireOwnerAuth_,
  revokeOwnerSession_,
  revokeShareSession_,
  validateOwnerToken_,
  validateShareToken_,
  verifyIdToken_,
} from './auth.js';
import { createJsonpOutput, fetchAllItems, readItemsCache, writeItemsCache } from './items.js';
import {
  SUBSCRIPTION_SHEET,
  createSubscription_,
  deleteSubscription_,
  readSubscriptions_,
  setSubscriptionFlag_,
  updateSubscription_,
} from './subscriptions.js';

// ---------------------------------------------------------------------------
// Admin API (doPost) for the list-of-stuff-admin site.
//
// The room sheet is rewritten by crawlRoom() from Rakuten ROOM, so the
// admin API manages the hand-edited sheets: the item sheets (workman,
// variety, others) and the subscriptions sheet. Authentication mirrors the other private sites (alarm,
// health, tweet): the owner's Google ID token, verified against the
// AVAILABLE_GMAILS allowlist, or a temporary share token issued by the
// owner (owner-equivalent; sessions are stored hashed in a
// share_sessions sheet in this spreadsheet). Actions:
//   { action: "list_items", sheet }             -> the sheet's items
//   { action: "create_item", sheet, item }      -> append one row
//   { action: "update_item", sheet, row, item } -> rewrite one data row
//   { action: "delete_item", sheet, row }       -> delete one data row
//   { action: "issue_share" }                   -> owner only
//   { action: "revoke_share" }                  -> owner only
// Every mutation clears the doGet response cache, so the public items
// API and everything built on it reflects the edit right away.
// ---------------------------------------------------------------------------

function doGet(e) {
  const cached = readItemsCache();
  if (cached !== null) {
    return createJsonpOutput(cached, e.parameter.callback);
  }

  const result = fetchAllItems();
  writeItemsCache(result);
  return createJsonpOutput(result, e.parameter.callback);
}

const jsonOut_ = payload => {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  output.setContent(JSON.stringify(payload));
  return output;
};

const authError_ = err => {
  const message = (err && err.message) || 'unauthorized';
  return jsonOut_({ ok: false, status: message === 'forbidden email' ? 403 : 401, error: message });
};

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '');
  } catch (err) {
    return jsonOut_({ ok: false, error: 'invalid JSON body' });
  }

  // Exchanging a Google ID token for a 30-day owner session is how a
  // sign-in starts; revoking the session is how it ends (logout).
  if (body.action === 'exchange_session') {
    try {
      verifyIdToken_(body.idToken);
    } catch (err) {
      return authError_(err);
    }
    try {
      return jsonOut_({ ok: true, ...issueOwnerSession_() });
    } catch (err) {
      return jsonOut_({ ok: false, error: err.message });
    }
  }
  if (body.action === 'revoke_session') {
    try {
      return jsonOut_({ ok: true, ...revokeOwnerSession_(body.session_token) });
    } catch (err) {
      return jsonOut_({ ok: false, error: err.message });
    }
  }

  // Issuing and revoking access links requires owner proof (the
  // Google ID token, or the owner session from it); a share token
  // cannot mint more share tokens.
  if (body.action === 'issue_share' || body.action === 'revoke_share') {
    try {
      requireOwnerAuth_(body);
    } catch (err) {
      return authError_(err);
    }
    try {
      const result = body.action === 'issue_share' ? issueShareSession_() : revokeShareSession_();
      return jsonOut_({ ok: true, ...result });
    } catch (err) {
      return jsonOut_({ ok: false, error: err.message });
    }
  }

  // The remaining actions accept the owner's ID token, the owner
  // session exchanged from it, or a share token from an issued access
  // link (owner-equivalent).
  let shareSession = null;
  if (body.session_token) {
    let ownerSession = null;
    try {
      ownerSession = validateOwnerToken_(body.session_token);
    } catch (err) {
      return jsonOut_({ ok: false, error: err.message });
    }
    if (!ownerSession) {
      return jsonOut_({ ok: false, status: 401, error: 'session expired or invalid' });
    }
  } else if (body.share_token) {
    try {
      shareSession = validateShareToken_(body.share_token);
    } catch (err) {
      return jsonOut_({ ok: false, error: err.message });
    }
    if (!shareSession) {
      return jsonOut_({ ok: false, status: 401, error: 'share session expired or invalid' });
    }
  } else {
    try {
      verifyIdToken_(body.idToken);
    } catch (err) {
      return authError_(err);
    }
  }
  const out = payload =>
    shareSession ? { ...payload, _share: { expires_at: shareSession.expires_at } } : payload;

  try {
    if (body.action === 'list_genres') {
      return jsonOut_(out({ ok: true, genres: listGenres_() }));
    }
    if (body.action === 'set_flag') {
      const flagged = body.flagged === true;
      if (body.sheet === 'room') {
        return jsonOut_(out({ ok: true, ...setRoomFlag_(body.url, flagged) }));
      }
      if (body.sheet === SUBSCRIPTION_SHEET) {
        return jsonOut_(out({ ok: true, ...setSubscriptionFlag_(body.row, flagged) }));
      }
      return jsonOut_(out({ ok: true, ...setItemFlag_(body.sheet, body.row, flagged) }));
    }
    if (body.action === 'list_items' && body.sheet === 'room') {
      return jsonOut_(out({ ok: true, sheet: 'room', items: readRoomItems_() }));
    }
    if (body.sheet === SUBSCRIPTION_SHEET) {
      if (body.action === 'list_items') {
        return jsonOut_(out({ ok: true, sheet: body.sheet, subscriptions: readSubscriptions_() }));
      }
      if (body.action === 'create_item') {
        return jsonOut_(out({ ok: true, ...createSubscription_(body.item) }));
      }
      if (body.action === 'update_item') {
        return jsonOut_(out({ ok: true, ...updateSubscription_(body.row, body.item) }));
      }
      if (body.action === 'delete_item') {
        return jsonOut_(out({ ok: true, ...deleteSubscription_(body.row) }));
      }
      return jsonOut_({ ok: false, error: 'unknown action' });
    }
    if (body.action === 'list_items') {
      return jsonOut_(out({ ok: true, sheet: body.sheet, items: readAdminItems_(body.sheet) }));
    }
    if (body.action === 'create_item') {
      return jsonOut_(out({ ok: true, ...createItem_(body.sheet, body.item) }));
    }
    if (body.action === 'update_item') {
      return jsonOut_(out({ ok: true, ...updateItem_(body.sheet, body.row, body.item) }));
    }
    if (body.action === 'delete_item') {
      return jsonOut_(out({ ok: true, ...deleteItem_(body.sheet, body.row) }));
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
  return jsonOut_({ ok: false, error: 'unknown action' });
}

export { doGet, doPost };
