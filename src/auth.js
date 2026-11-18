// Authentication for the admin API: the owner's Google ID token,
// verified against an allowlist, and the share / owner sessions
// exchanged from it.

import { scriptProperty_, spreadsheetId_ } from './common.js';

const allowedEmails_ = () =>
  (scriptProperty_('AVAILABLE_GMAILS') || '')
    .split(',')
    .map(v => v.trim().toLowerCase())
    .filter(v => v.length > 0);

// Throws an Error when the token is missing, invalid, or not allowed.
const verifyIdToken_ = idToken => {
  if (!idToken) throw new Error('missing id token');
  const clientId = scriptProperty_('GOOGLE_OAUTH_CLIENT_ID');
  const tokenInfoUrl = `https://oauth2.googleapis.com/tokeninfo?id_token=${encodeURIComponent(idToken)}`;
  let res;
  try {
    res = UrlFetchApp.fetch(tokenInfoUrl, { muteHttpExceptions: true });
  } catch (err) {
    // A transport-level failure (DNS, connection) is transient;
    // retry once before giving up. HTTP-level failures arrive as a
    // response (muteHttpExceptions) and are not retried.
    res = UrlFetchApp.fetch(tokenInfoUrl, { muteHttpExceptions: true });
  }
  if (res.getResponseCode() !== 200) throw new Error('token verification failed');
  const payload = JSON.parse(res.getContentText());
  if (payload.iss !== 'accounts.google.com' && payload.iss !== 'https://accounts.google.com') {
    throw new Error('invalid iss');
  }
  if (payload.aud !== clientId) throw new Error('invalid aud');
  const exp = Number(payload.exp || 0);
  if (!exp || exp * 1000 <= Date.now()) throw new Error('token expired');
  if (payload.email_verified !== true && payload.email_verified !== 'true') {
    throw new Error('email not verified');
  }
  const email = String(payload.email || '').toLowerCase();
  if (!email) throw new Error('missing email');
  if (!allowedEmails_().includes(email)) throw new Error('forbidden email');
  return email;
};

// --- Temporary share sessions (same scheme as the alarm site) ---
const SHARE_SHEET_NAME = 'share_sessions';
// One month, so the owner can hand the link over well in advance.
const SHARE_TTL_SECONDS = 30 * 24 * 60 * 60;
const SHARE_HEADERS = [
  'token_hash',
  'issued_at',
  'expires_at',
  'revoked_at',
  'access_count',
  'last_accessed_at',
];

const getSessionSheet_ = name => {
  const ss = SpreadsheetApp.openById(spreadsheetId_());
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(SHARE_HEADERS);
  }
  return sheet;
};

const hashShareToken_ = token =>
  Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token))
    .map(byte => ((byte % 256) + 256) % 256)
    .map(byte => byte.toString(16).padStart(2, '0'))
    .join('');

const readSessionRows_ = name => {
  const sheet = getSessionSheet_(name);
  const values = sheet.getDataRange().getValues();
  return { sheet, values };
};

const shareCellIso_ = value => {
  if (!value) {
    return '';
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
};

const revokeActiveSessions_ = (name, nowIso) => {
  const { sheet, values } = readSessionRows_(name);
  values.forEach((row, index) => {
    if (index === 0) {
      return;
    }
    const revokedAt = shareCellIso_(row[3]);
    const expiresAt = shareCellIso_(row[2]);
    if (!revokedAt && expiresAt && expiresAt > nowIso) {
      sheet.getRange(index + 1, 4).setValue(nowIso);
    }
  });
};

const revokeActiveShareSessions_ = nowIso => revokeActiveSessions_(SHARE_SHEET_NAME, nowIso);

const issueSession_ = (name, tokenKey, { revokePrevious = true } = {}) => {
  const now = new Date();
  const nowIso = now.toISOString();
  const expiresIso = new Date(now.getTime() + SHARE_TTL_SECONDS * 1000).toISOString();
  if (revokePrevious) {
    revokeActiveSessions_(name, nowIso);
  }
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  getSessionSheet_(name).appendRow([hashShareToken_(token), nowIso, expiresIso, '', 0, '']);
  return {
    [tokenKey]: token,
    issued_at: nowIso,
    expires_at: expiresIso,
    ttl_seconds: SHARE_TTL_SECONDS,
  };
};

const issueShareSession_ = () => issueSession_(SHARE_SHEET_NAME, 'share_token');

const revokeShareSession_ = () => {
  revokeActiveShareSessions_(new Date().toISOString());
  return { revoked: true };
};

const validateSessionToken_ = (name, token) => {
  if (!token) {
    return null;
  }
  const hash = hashShareToken_(token);
  const { sheet, values } = readSessionRows_(name);
  const nowIso = new Date().toISOString();
  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    if (row[0] !== hash) {
      continue;
    }
    if (shareCellIso_(row[3])) {
      return null;
    }
    const expiresAt = shareCellIso_(row[2]);
    if (!expiresAt || expiresAt <= nowIso) {
      return null;
    }
    sheet.getRange(index + 1, 5).setValue(Number(row[4] || 0) + 1);
    sheet.getRange(index + 1, 6).setValue(nowIso);
    return { expires_at: expiresAt };
  }
  return null;
};

const validateShareToken_ = token => validateSessionToken_(SHARE_SHEET_NAME, token);

// --- Owner sessions (long-lived sign-in) ---
// Google ID tokens expire after about an hour, which used to drop the
// owner back at the sign-in gate. After a Google sign-in the frontend
// exchanges the ID token for one of these 30-day owner sessions (same
// storage scheme as share sessions, separate sheet) and uses it for
// all later calls. Owner sessions coexist across devices: issuing a
// new one does not revoke the previous ones (unlike share sessions,
// where reissuing invalidates the old link). The frontend revokes
// its own session on logout.
const OWNER_SHEET_NAME = 'owner_sessions';

const issueOwnerSession_ = () =>
  issueSession_(OWNER_SHEET_NAME, 'session_token', { revokePrevious: false });

const validateOwnerToken_ = token => validateSessionToken_(OWNER_SHEET_NAME, token);

const revokeOwnerSession_ = token => {
  if (token) {
    const hash = hashShareToken_(token);
    const { sheet, values } = readSessionRows_(OWNER_SHEET_NAME);
    const nowIso = new Date().toISOString();
    values.forEach((row, index) => {
      if (index > 0 && row[0] === hash && !shareCellIso_(row[3])) {
        sheet.getRange(index + 1, 4).setValue(nowIso);
      }
    });
  }
  return { revoked: true };
};

// Owner proof for owner-only actions: a valid Google ID token, or a
// live owner session token (the post-sign-in steady state).
const requireOwnerAuth_ = body => {
  if (body.session_token && validateOwnerToken_(body.session_token)) {
    return;
  }
  verifyIdToken_(body.idToken);
};

export { issueOwnerSession_, issueShareSession_, requireOwnerAuth_, revokeOwnerSession_, revokeShareSession_, validateOwnerToken_, validateShareToken_, verifyIdToken_ };
