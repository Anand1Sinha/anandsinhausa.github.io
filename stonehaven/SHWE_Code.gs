// ============================================================
// SHWE Annual Social 2026 — Apps Script Backend
// Google Account: StonehavenwoodsEast@gmail.com
// RSVP Sheet: "SHWE Annual Social 2026" (SHEET_ID below)
// Directory Sheet: "StonehavenWoods East - Directory (2025-10-18)" (DIRECTORY_SHEET_ID below)
//
// SETUP INSTRUCTIONS:
// 1. Go to sheets.google.com — create new sheet named "SHWE Annual Social 2026"
// 2. Create these tabs: RSVPs | Reminders Log | Suggestions
// 3. Extensions → Apps Script
// 4. Paste this entire file, replacing any existing code
// 5. Save → Deploy → New deployment
//    - Type: Web app
//    - Execute as: Me (StonehavenwoodsEast@gmail.com)
//    - Who has access: Anyone
// 6. Copy the Web app URL
// 7. Paste it into 2026social.html where it says SCRIPT_URL = ''
// 8. Set up triggers (see TRIGGER SETUP below)
//
// NOTE: This script has NO resident data in it. Household names, emails, and
// addresses are read live from DIRECTORY_SHEET_ID / DIRECTORY_TAB below,
// cached briefly via CacheService. To add/remove/edit a household, edit that
// Sheet directly — no code change or redeploy needed. Both Apps Script
// deployments (SHWE Users + SHWE Admin) need edit access to that Sheet's ID,
// not just this project's own SHEET_ID — this account already owns both, so
// no extra sharing step is needed as long as it stays that way.
// ============================================================

const SHEET_ID     = '1U1Ee2Yhep56fcKI1JwN7fRPlJsSSC-8YABf6UIkooSU'; // the actual Google Sheet — see URL in Drive
const FROM_NAME    = 'Stonehaven Woods East HOA';

// ── RESIDENT DIRECTORY — read live from the mail-merge Sheet, NOT hardcoded ──
// This is the user's real household directory, maintained independently of the
// RSVP system. Read-only: this script never writes to it.
const DIRECTORY_SHEET_ID = '1xPx1ZkLyvSyj_QAGpZbRZVV9NwFCanJLElooA9UfvbU';
const DIRECTORY_TAB      = 'Live';
const RESIDENTS_CACHE_KEY = 'residents_v1';
const RESIDENTS_CACHE_TTL = 300; // 5 minutes — long enough to avoid re-reading the
                                  // Sheet on every autocomplete keystroke, short
                                  // enough that directory edits show up quickly.
// Admin access is now enforced by Google account identity, NOT a shared secret.
// Add every board member's Google account email that should have admin access.
// This ONLY works when the admin panel is deployed as its own deployment with:
//   Execute as: User accessing the web app
//   Who has access: Anyone with Google account
// (See deployment notes at the bottom of this file.)
const ADMIN_ALLOWLIST = [
  'StonehavenwoodsEast@gmail.com'
  // add board members' personal Google account emails here, e.g. 'sona_patil@hotmail.com'
  // NOTE: must be a Google account (Gmail or Google Workspace) to authenticate via Session.getActiveUser()
];

function isAuthorizedAdmin() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) return false;
    return ADMIN_ALLOWLIST.map(function(e){return e.toLowerCase();}).indexOf(email.toLowerCase()) !== -1;
  } catch (e) {
    return false; // no active user session available — deny by default
  }
}
const RSVP_LOCK     = new Date('2026-10-02T23:59:59-04:00'); // single authoritative deadline
const SITE_URL     = 'https://anandsinhausa.com/stonehaven/2026social.html';
const ADMIN_EMAIL  = 'StonehavenwoodsEast@gmail.com';
const EVENT_DATE   = 'Sunday, October 4, 2026';
const EVENT_TIME   = '3:00 – 5:00 PM';
const EVENT_ADDR   = 'Backyard of 3887 Fadi Drive, Troy MI';

// Social Coordinators — same names/numbers as the "Call Sona" / "Call Surya"
// buttons on 2026social.html. Keep these two in sync if either number changes.
const COORDINATOR_1_NAME    = 'Sona';
const COORDINATOR_1_PHONE   = '+12484625473';
const COORDINATOR_2_NAME    = 'Surya';
const COORDINATOR_2_PHONE   = '+12489358103';

// ── TAB HEADERS ──
const RSVP_HEADERS = [
  'Timestamp','ResidentID','House#','Address','Name1','Name2','Last',
  'Email','Email2','RSVP','Adults','Kids Under 5','Kids 6-15',
  'Event Suggestions','No Reason','No Reason Text',
  'Email Differs From Directory' // column 16 (0-indexed) — appended at the END, deliberately,
                                  // so every existing data[i][N] reference elsewhere in this file
                                  // keeps working unchanged. See §7 bug #4 in the project bible for
                                  // why a column insertion (vs. append) is the dangerous move here.
];

const SUGGESTION_HEADERS = [
  'Timestamp','Next Year Venue','Next Year Activities','Next Year Other'
];

const REMINDER_HEADERS = [
  'Timestamp','Type','House#','Address','Email','Status'
];

// ── RESIDENT DIRECTORY — read live from the mail-merge Sheet ──
// Replaces the old hardcoded 41-household array. No resident PII lives in
// this source file anymore; it's read fresh from DIRECTORY_SHEET_ID / "Live"
// on cache miss, and cached for RESIDENTS_CACHE_TTL seconds via CacheService.
//
// Returns the same shape the old hardcoded array used, so every downstream
// function (searchResidents, verifyPin, previewReminders, sendReminders,
// getStats) works unchanged: {house, id, name1, name2, last, email1, email2, addr}
//
// `id` is derived the same way it always was — first 3 letters of the street
// name (uppercased) + house number, e.g. "ROT-1228" / "PRO-1228" — which is
// what disambiguates the one real house-number collision (1228 Rothwell vs
// 1228 Provincial). This only works because the five streets in this sheet
// (Salma, Fadi, Jefferson, Rothwell, Provincial) all have distinct 3-letter
// prefixes; if a 6th street sharing a prefix is ever added, this needs revisiting.
function getResidents() {
  var cache = CacheService.getScriptCache();
  var cached;
  try { cached = cache.get(RESIDENTS_CACHE_KEY); } catch (e) { cached = null; }
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* corrupt cache entry, fall through to reload */ }
  }
  var residents = loadResidentsFromSheet();
  try {
    cache.put(RESIDENTS_CACHE_KEY, JSON.stringify(residents), RESIDENTS_CACHE_TTL);
  } catch (e) {
    // Non-fatal — e.g. payload larger than the ~100KB per-key cache limit.
    // Worst case we just re-read the Sheet next call instead of caching.
  }
  return residents;
}

function loadResidentsFromSheet() {
  var ss = SpreadsheetApp.openById(DIRECTORY_SHEET_ID);
  var sh = ss.getSheetByName(DIRECTORY_TAB);
  if (!sh) throw new Error('Directory sheet is missing the "' + DIRECTORY_TAB + '" tab.');

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function (h) { return String(h).trim(); });
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  var required = ['First Name', 'Spouse First Name', 'Last Name', 'Email Primary',
                   'Email Secondary', 'Address 1', 'House #', 'Street Name'];
  required.forEach(function (h) {
    if (!(h in col)) throw new Error('Directory sheet is missing expected column: "' + h + '"');
  });

  var residents = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var house     = String(row[col['House #']] || '').trim();
    var email1    = String(row[col['Email Primary']] || '').trim();
    var firstName = String(row[col['First Name']] || '').trim();

    // Skip rows that aren't real RSVP-eligible households: no house number
    // means either a footer/note row (e.g. "updated 10/18/2025") or a
    // household without a street address (Karen, P.O. Box — intentionally
    // excluded from RSVP eligibility, same as before this migration).
    if (!house || !email1 || !firstName) continue;

    var street = String(row[col['Street Name']] || '').trim();
    var streetPrefix = (street.substring(0, 3) || 'UNK').toUpperCase();

    residents.push({
      house:  house,
      id:     streetPrefix + '-' + house,
      name1:  firstName,
      name2:  String(row[col['Spouse First Name']] || '').trim(),
      last:   String(row[col['Last Name']] || '').trim(),
      email1: email1,
      // Defensive: if a cell ever has multiple emails jammed together again
      // (semicolon or comma separated), only take the first rather than
      // silently mailing whatever else got typed into that cell.
      email2: String(row[col['Email Secondary']] || '').trim().split(/[;,]/)[0].trim(),
      addr:   String(row[col['Address 1']] || '').trim()
    });
  }
  return residents;
}


// ── SANITIZE UNTRUSTED TEXT FOR SHEETS (prevents formula injection) ──
function sanitizeForSheet(value) {
  var s = String(value == null ? '' : value);
  if (/^[=+\-@]/.test(s)) {
    return "'" + s; // leading apostrophe forces Sheets to treat it as plain text
  }
  return s;
}

// ── HELPERS ──
function getSheet(tabName) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(tabName);
  if (!sh) {
    sh = ss.insertSheet(tabName);
    if (tabName === 'RSVPs')       sh.appendRow(RSVP_HEADERS);
    if (tabName === 'Suggestions') sh.appendRow(SUGGESTION_HEADERS);
    if (tabName === 'Reminders Log') sh.appendRow(REMINDER_HEADERS);
  }
  return sh;
}

function displayName(r) {
  // Display uses ONLY name1 (which already contains both spouse first names,
  // e.g. "Kevin Kim and Heeok Han") plus last name. name2 is a redundant
  // duplicate of the second spouse's name and is never used for display.
  var parts = [];
  if (r.name1) parts.push(r.name1);
  if (r.last)  parts.push(r.last);
  return parts.join(' ');
}

function getRsvpedHouses() {
  var sh = getSheet('RSVPs');
  var data = sh.getDataRange().getValues();
  var rsvps = {};
  for (var i = 1; i < data.length; i++) {
    rsvps[String(data[i][1])] = String(data[i][9]).toLowerCase(); // ResidentID → rsvp yes/no (col 9 = RSVP)
  }
  return rsvps;
}

// ── EMAIL TEMPLATES ──
function confirmationEmail(p) {
  var name = displayName(p);
  var isYes = p.rsvp === 'yes';
  var subject = isYes
    ? 'You\'re on the list! \u2014 SHWE Annual Social Oct 4'
    : 'Thanks for letting us know \u2014 SHWE Annual Social';

  var body = '<div style="font-family:Georgia,serif;max-width:560px;margin:0 auto;color:#2C1A0E;">'
    + '<div style="background:#1E3D2F;padding:20px 28px;border-radius:8px 8px 0 0;">'
    + '<p style="color:#F2C94C;font-size:17px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;margin:0">Stonehaven Woods East</p>'
    + '<h1 style="color:#fff;font-size:22px;margin:8px 0 0;">Annual Social 2026</h1>'
    + '</div>'
    + '<div style="background:#FFF;border:1px solid #EFE5CC;border-top:none;padding:28px;border-radius:0 0 8px 8px;">';

  if (isYes) {
    body += '<h2 style="color:#1E3D2F;margin:0 0 12px">You\'re on the list, ' + p.name1 + '!</h2>'
      + '<p>We\'ve received your RSVP and can\'t wait to see you on <strong>' + EVENT_DATE + '</strong>.</p>'
      + '<div style="background:#F7F0E3;border-radius:8px;padding:16px 20px;margin:16px 0;">'
      + '<p style="margin:4px 0"><strong>Date:</strong> ' + EVENT_DATE + '</p>'
      + '<p style="margin:4px 0"><strong>Time:</strong> ' + EVENT_TIME + '</p>'
      + '<p style="margin:4px 0"><strong>Location:</strong> ' + EVENT_ADDR + '</p>'
      + '<p style="margin:4px 0"><strong>Adults:</strong> ' + p.adults + '</p>'
      + (parseInt(p.kids5) > 0 ? '<p style="margin:4px 0"><strong>Under 5:</strong> ' + p.kids5 + '</p>' : '')
      + (parseInt(p.kids15) > 0 ? '<p style="margin:4px 0"><strong>Ages 6\u201315:</strong> ' + p.kids15 + '</p>' : '')
      + '</div>'
      + '<p>Need to make a change? You can update your RSVP anytime before October 2nd.</p>'
      + '<div style="text-align:center;margin:24px 0;">'
      + '<a href="' + SITE_URL + '" style="background:#6B1A1A;color:#fff;text-decoration:none;padding:14px 28px;border-radius:8px;font-weight:bold;display:inline-block;">Update My RSVP</a>'
      + '</div>';
  } else {
    body += '<h2 style="color:#1E3D2F;margin:0 0 12px">We\'ll miss you, ' + p.name1 + '!</h2>'
      + '<p>Thanks for letting us know. We hope to see you at next year\'s social!</p>'
      + '<p>If your plans change before October 2nd, you\'re always welcome to update your RSVP.</p>'
      + '<div style="text-align:center;margin:24px 0;">'
      + '<a href="' + SITE_URL + '" style="background:#1E3D2F;color:#fff;text-decoration:none;padding:14px 28px;border-radius:8px;font-weight:bold;display:inline-block;">Update My RSVP</a>'
      + '</div>';
  }

  body += '<hr style="border:none;border-top:1px solid #EFE5CC;margin:20px 0">'
    + '<p style="font-size:12px;color:#6B5744;text-align:center;margin:0 0 12px">Questions about the event? Reach your Social Coordinators directly:</p>'
    + '<div style="text-align:center;margin:0 0 14px">'
    + '<a href="tel:' + COORDINATOR_1_PHONE + '" style="background:#1E3D2F;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:bold;display:inline-block;margin:0 6px 8px;font-size:13px;">Call ' + COORDINATOR_1_NAME + '</a>'
    + '<a href="tel:' + COORDINATOR_2_PHONE + '" style="background:#1E3D2F;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:bold;display:inline-block;margin:0 6px 8px;font-size:13px;">Call ' + COORDINATOR_2_NAME + '</a>'
    + '</div>'
    + '<p style="font-size:11px;color:#6B5744;text-align:center;">Or email <a href="mailto:' + ADMIN_EMAIL + '">' + ADMIN_EMAIL + '</a></p>'
    + '</div></div>';

  return { subject: subject, body: body };
}

function reminderEmail(r, type) {
  var name = displayName(r);
  var isInvite  = (type === 'invite');
  var isUpdate  = (type === 'update');
  var isClosing = (type === 'closing');

  var subject = isInvite
    ? 'You\'re invited \u2014 SHWE Annual Social, Sunday October 4'
    : isUpdate
    ? 'See you Sunday! \u2014 SHWE Annual Social Oct 4'
    : isClosing
    ? 'Last chance to RSVP \u2014 SHWE Annual Social Oct 4'
    : 'Don\'t forget \u2014 SHWE Annual Social is one week away!';

  var headline = isInvite
    ? 'You\'re invited, ' + r.name1 + '!'
    : isUpdate
    ? 'See you Sunday, ' + r.name1 + '!'
    : isClosing
    ? 'Last chance to RSVP, ' + r.name1 + '!'
    : 'One week to go, ' + r.name1 + '!';

  var msg = isInvite
    ? 'Please join your neighbors at the Stonehaven Woods East Annual Social. We would love to see your household there \u2014 please RSVP below.'
    : isUpdate
    ? 'Just a friendly reminder that the Annual Social is coming up this Sunday. We\'re looking forward to seeing you!'
    : isClosing
    ? 'RSVPs close tonight at midnight. If you\'re planning to join us, please let us know now so we can plan accordingly.'
    : 'The Stonehaven Woods East Annual Social is just one week away. If you haven\'t RSVPed yet, we\'d love to know if you\'re coming!';

  var btnLabel = (isInvite || isClosing) ? 'RSVP Now' : (isUpdate ? 'View or Update My RSVP' : 'RSVP Now');
  var btnColor = isUpdate ? '#1E3D2F' : '#6B1A1A';

  var body = '<div style="font-family:Georgia,serif;max-width:560px;margin:0 auto;color:#2C1A0E;">'
    + '<div style="background:#1E3D2F;padding:20px 28px;border-radius:8px 8px 0 0;">'
    + '<p style="color:#F2C94C;font-size:17px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;margin:0">Stonehaven Woods East</p>'
    + '<h1 style="color:#fff;font-size:22px;margin:8px 0 0;">Annual Social 2026</h1>'
    + '</div>'
    + '<div style="background:#FFF;border:1px solid #EFE5CC;border-top:none;padding:28px;border-radius:0 0 8px 8px;">'
    + '<h2 style="color:#1E3D2F;margin:0 0 12px">' + headline + '</h2>'
    + '<p>' + msg + '</p>'
    + '<div style="background:#F7F0E3;border-radius:8px;padding:16px 20px;margin:16px 0;">'
    + '<p style="margin:4px 0"><strong>Date:</strong> ' + EVENT_DATE + '</p>'
    + '<p style="margin:4px 0"><strong>Time:</strong> ' + EVENT_TIME + '</p>'
    + '<p style="margin:4px 0"><strong>Location:</strong> ' + EVENT_ADDR + '</p>'
    + '</div>'
    + '<div style="text-align:center;margin:24px 0;">'
    + '<a href="' + SITE_URL + '" style="background:' + btnColor + ';color:#fff;text-decoration:none;padding:14px 28px;border-radius:8px;font-weight:bold;display:inline-block;">' + btnLabel + '</a>'
    + '</div>'
    + '<hr style="border:none;border-top:1px solid #EFE5CC;margin:20px 0">'
    + '<p style="font-size:12px;color:#6B5744;text-align:center;">Stonehaven Woods East HOA \u00B7 Troy, Michigan</p>'
    + '</div></div>';

  return { subject: subject, body: body };
}

// ── SEND EMAIL HELPER ──
function sendEmail(to, cc, subject, htmlBody) {
  var opts = {
    name: FROM_NAME,
    htmlBody: htmlBody,
    replyTo: ADMIN_EMAIL
  };
  if (cc) opts.cc = cc;
  GmailApp.sendEmail(to, subject, '', opts);
}


// ── SESSION TOKEN (uses CacheService — persists across separate requests) ──
function generateToken() {
  return Utilities.getUuid() + '-' + Date.now().toString(36);
}
function cachePutSession(token, session) {
  CacheService.getScriptCache().put('session:' + token, JSON.stringify(session), 7200); // 2 hours
}
function cacheGetSession(token) {
  var raw = CacheService.getScriptCache().get('session:' + token);
  return raw ? JSON.parse(raw) : null;
}
function cacheDeleteSession(token) {
  CacheService.getScriptCache().remove('session:' + token);
}

// ── ACTION: SEARCH RESIDENTS (returns display name + addr only, no emails) ──
function searchResidents(q) {
  q = (q || '').toLowerCase().trim();
  if (q.length < 2) return { results: [] };

  // Word-based matching: every word the resident typed must appear as a substring
  // somewhere in the household's name tokens, in ANY order. This correctly matches
  // "Anand Sinha" against a record where name1 is "Anand and Sonia" and last is
  // "Sinha" -- a naive whole-string substring search would miss this because
  // "and Sonia" sits between "Anand" and "Sinha" in the concatenated string.
  var queryWords = q.split(/\s+/).filter(Boolean);
  var residents = getResidents();

  var results = residents.filter(function(r) {
    var tokens = [r.name1, r.name2, r.last].filter(Boolean).join(' ').toLowerCase();
    return queryWords.every(function(word) {
      return tokens.indexOf(word) !== -1;
    });
  }).slice(0, 8).map(function(r) {
    // Display uses ONLY name1 + last. name1 already contains both spouse
    // first names for combined-style records; name2 is never appended.
    var display = r.last ? (r.name1 + ' ' + r.last) : r.name1;
    // Mask street: no house number before verification (house number is no longer the PIN,
    // but we still avoid handing out full addresses to an unauthenticated search).
    var streetOnly = r.addr.replace(/^\d+\s*/, '');
    return {
      display: display,
      street: streetOnly,
      residentId: r.id
      // deliberately NO house#, NO full address, NO email
    };
  });
  return { success: true, results: results };
}

// ── ACTION: VERIFY PIN (server-side — house# never sent to browser) ──
function verifyPin(residentId, pin) {
  var cache = CacheService.getScriptCache();

  // Rate limit is keyed by residentId + a coarse client signal (IP not available in Apps Script,
  // so we also cap globally per residentId but do NOT lock the household out —
  // failed attempts slow down, they never block a legitimate resident indefinitely).
  var attemptsKey = 'pinattempts:' + residentId;
  var attempts = parseInt(cache.get(attemptsKey) || '0', 10);
  if (attempts >= 10) {
    return { verified: false, error: 'Too many attempts. Please wait a few minutes and try again, or contact StonehavenwoodsEast@gmail.com.' };
  }
  cache.put(attemptsKey, String(attempts + 1), 300); // 5 min rolling window, not a hard lockout

  var resident = null;
  var residents = getResidents();
  for (var i = 0; i < residents.length; i++) {
    if (residents[i].id === residentId) { resident = residents[i]; break; }
  }
  if (!resident) return { verified: false };
  // PIN = house number. This is a lightweight collision check (e.g. avoiding two
  // 'Kevin' households mixing up RSVPs), not a real security control.
  if (String(pin) !== String(resident.house)) return { verified: false };

  var token = generateToken();
  cachePutSession(token, {
    id: resident.id,
    addr: resident.addr,
    name1: resident.name1,
    name2: resident.name2,
    last: resident.last,
    email1: resident.email1,
    email2: resident.email2,
    house: resident.house
  });
  cache.remove(attemptsKey); // clear rate-limit counter on success
  return {
    verified: true,
    token: token,
    addr: resident.addr,   // full address only revealed AFTER successful verification
    email: resident.email1 || '' // full email likewise — pre-fills the editable email field
  };
}

// ── ACTION: SUBMIT ANONYMOUS SUGGESTION ──
function submitSuggestion(p) {
  if (!p.nextVenue && !p.nextActivities && !p.nextOther) return { success: true };
  // Store date only (no time-of-day) to reduce correlation with a same-second RSVP submission
  var dateOnly = Utilities.formatDate(new Date(), 'America/Detroit', 'yyyy-MM-dd');
  getSheet('Suggestions').appendRow([
    dateOnly,
    sanitizeForSheet(String(p.nextVenue || '').substring(0,500)),
    sanitizeForSheet(String(p.nextActivities || '').substring(0,500)),
    sanitizeForSheet(String(p.nextOther || '').substring(0,500))
  ]);
  return { success: true };
}

// ── ACTION: SUBMIT RSVP ──
function submitRsvp(p) {
  // Enforce RSVP deadline server-side
  if (new Date() >= RSVP_LOCK) {
    return { success: false, error: 'RSVPs are now closed.' };
  }

  // Validate token (persisted via CacheService, survives separate Apps Script executions)
  var session = cacheGetSession(p.token);
  if (!session) {
    return { success: false, error: 'Session expired. Please refresh and try again.' };
  }

  // Server-side input validation — never trust the browser
  if (p.rsvp !== 'yes' && p.rsvp !== 'no') {
    return { success: false, error: 'Invalid RSVP value.' };
  }
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!p.email || !emailPattern.test(p.email) || p.email.length > 200) {
    return { success: false, error: 'Invalid email address.' };
  }
  function toSafeInt(v, max) {
    var n = parseInt(v, 10);
    if (isNaN(n) || n < 0 || n > max) return null;
    return n;
  }
  var adults = toSafeInt(p.adults, 20);
  var kids5  = toSafeInt(p.kids5, 20);
  var kids15 = toSafeInt(p.kids15, 20);
  if (adults === null || kids5 === null || kids15 === null) {
    return { success: false, error: 'Invalid attendance numbers.' };
  }
  if (p.rsvp === 'yes' && (adults + kids5 + kids15) < 1) {
    return { success: false, error: 'Please include at least one attendee.' };
  }
  p.adults = adults; p.kids5 = kids5; p.kids15 = kids15;
  p.eventSuggestions = String(p.eventSuggestions || '').substring(0, 500);
  p.noReasonText      = String(p.noReasonText || '').substring(0, 500);

  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // wait up to 10s for concurrent submissions to clear

  try {
    var sh = getSheet('RSVPs');
    var data = sh.getDataRange().getValues();

    // Overwrite if this household already RSVPed
    var existingRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(session.id)) {
        existingRow = i + 1; break;
      }
    }

    // Flag (not auto-sync) when the email they typed differs from what's on
    // file in the directory Sheet. The directory itself is never written to —
    // this only surfaces the mismatch in the RSVP sheet so a human can decide
    // whether to update the directory later.
    var directoryEmail = String(session.email1 || '').trim();
    var typedEmail      = String(p.email || '').trim();
    var emailDiffers = directoryEmail && directoryEmail.toLowerCase() !== typedEmail.toLowerCase();
    var emailDiffersNote = emailDiffers ? ('Yes \u2014 directory has: ' + directoryEmail) : '';

    var row = [
      new Date(), session.id, session.house, session.addr,
      session.name1, session.name2, session.last,
      sanitizeForSheet(p.email), session.email2 || '', p.rsvp,
      p.adults || 0, p.kids5 || 0, p.kids15 || 0,
      sanitizeForSheet(p.eventSuggestions || ''), p.noReason || '', sanitizeForSheet(p.noReasonText || ''),
      sanitizeForSheet(emailDiffersNote)
    ];

    if (existingRow > 0) {
      sh.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  } finally {
    lock.releaseLock();
  }

  // Invalidate token after use
  cacheDeleteSession(p.token);

  // Send confirmation email using session data (email from server, not browser)
  var emailPayload = {
    name1: session.name1, rsvp: p.rsvp,
    adults: p.adults, kids5: p.kids5, kids15: p.kids15
  };
  var tmpl = confirmationEmail(emailPayload);
  var emailSent = true;
  try {
    // Send to the email the resident actually entered.
    // CC the directory email on file, if different, so the household stays informed.
    var ccAddr = (session.email1 && session.email1 !== p.email) ? session.email1 : '';
    sendEmail(p.email, ccAddr, tmpl.subject, tmpl.body);
  } catch (err) {
    emailSent = false;
  }

  return { success: true, emailSent: emailSent };
}

// ── ACTION: SEND TEST CONFIRMATION EMAIL (admin only) ──
// Renders the real confirmationEmail() template with sample data and sends it
// to an admin-supplied address, so the actual confirmation email (fonts, colors,
// spacing, whether icons/emoji render correctly, etc.) can be visually checked
// before any real resident sees it. Does NOT touch the RSVPs sheet, does NOT
// use any real household's data, and is completely separate from the
// campaign test-send (sendReminders with testMode) which tests the invite/
// reminder/update/closing templates instead of this confirmation template.
function sendTestConfirmation(testEmail, rsvpType) {
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!testEmail || !emailPattern.test(testEmail)) {
    return { success: false, error: 'Valid testEmail is required.' };
  }
  var rsvp = (rsvpType === 'no') ? 'no' : 'yes';
  var sample = {
    name1: 'Test Household',
    rsvp: rsvp,
    adults: 2,
    kids5: 1,
    kids15: 1
  };
  var tmpl = confirmationEmail(sample);
  try {
    sendEmail(testEmail, '', '[TEST] ' + tmpl.subject, tmpl.body);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── ACTION: SEND REMINDERS (call manually from admin or triggered) ──
function previewReminders(type) {
  // Read-only preview: shows exactly who would receive this campaign and a
  // sample rendered email, without sending anything or writing any log entry.
  var allowedTypes = ['invite', 'reminder', 'update', 'closing'];
  if (allowedTypes.indexOf(type) === -1) {
    return { success: false, error: 'Invalid reminder type.' };
  }
  var rsvped = getRsvpedHouses();
  var residents = getResidents();
  var recipients = [];

  residents.forEach(function(r) {
    if (!r.email1) return;
    var hasRsvp = rsvped.hasOwnProperty(r.id);
    var shouldSend = false;
    if (type === 'invite') shouldSend = true; // everyone, regardless of RSVP status
    if ((type === 'reminder' || type === 'closing') && !hasRsvp) shouldSend = true;
    if (type === 'update' && hasRsvp && rsvped[r.id] === 'yes') shouldSend = true;
    if (shouldSend) {
      recipients.push({ name: r.name1, addr: r.addr, email: r.email1 });
    }
  });

  var sample = recipients.length ? reminderEmail(
    residents.filter(function(r) { return r.email1 === recipients[0].email; })[0],
    type
  ) : null;

  return {
    success: true,
    type: type,
    recipientCount: recipients.length,
    recipients: recipients,
    sampleSubject: sample ? sample.subject : null,
    sampleBody: sample ? sample.body : null
  };
}

function sendReminders(type, testMode, testEmail) {
  // type: 'invite'   (initial invitation, everyone regardless of RSVP status)
  //       'reminder' (Sept 27 non-RSVPed)
  //       'update'   (Sept 27 already RSVPed)
  //       'closing'  (Sept 30 non-RSVPed)
  // testMode: if true, every email is redirected to testEmail instead of the real
  //           recipient. The real recipient's name/address is shown in the subject
  //           line so you can verify targeting. Test sends do NOT count toward the
  //           one-time campaign lock and do NOT write 'campaign-complete', so the
  //           real send later is unaffected.
  var allowedTypes = ['invite', 'reminder', 'update', 'closing'];
  if (allowedTypes.indexOf(type) === -1) {
    return { success: false, error: 'Invalid reminder type.' };
  }

  if (testMode) {
    var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!testEmail || !emailPattern.test(testEmail)) {
      return { success: false, error: 'Valid testEmail is required when testMode is true.' };
    }
  } else if (campaignAlreadySent(type)) {
    return { success: false, error: 'This campaign (' + type + ') has already been sent. Contact the developer to force a resend.' };
  }

  var rsvped = getRsvpedHouses();
  var logSh  = getSheet('Reminders Log');
  var sent   = 0;

  getResidents().forEach(function(r) {
    if (!r.email1) return;
    var hasRsvp = rsvped.hasOwnProperty(r.id);

    // invite → everyone, regardless of RSVP status
    // reminder + closing → only non-RSVPed
    // update → only RSVPed with yes
    var shouldSend = false;
    if (type === 'invite') shouldSend = true;
    if ((type === 'reminder' || type === 'closing') && !hasRsvp) shouldSend = true;
    if (type === 'update' && hasRsvp && rsvped[r.id] === 'yes') shouldSend = true;

    if (!shouldSend) return;

    try {
      var tmpl = reminderEmail(r, type);
      if (testMode) {
        tmpl.subject = '[TEST - would go to ' + r.name1 + ', ' + r.addr + '] ' + tmpl.subject;
        sendEmail(testEmail, '', tmpl.subject, tmpl.body);
        logSh.appendRow([new Date(), type, r.house, r.addr, testEmail, 'test-sent (real recipient: ' + r.email1 + ')']);
      } else {
        sendEmail(r.email1, r.email2 || '', tmpl.subject, tmpl.body);
        logSh.appendRow([new Date(), type, r.house, r.addr, r.email1, 'sent']);
      }
      sent++;
    } catch(e) {
      logSh.appendRow([new Date(), type, r.house, r.addr, testMode ? testEmail : r.email1, 'error: ' + e.message]);
    }
  });

  if (!testMode) {
    getSheet('Reminders Log').appendRow([new Date(), type, '', '', '', 'campaign-complete']);
  }
  return { success: true, sent: sent, testMode: !!testMode };
}

function campaignAlreadySent(type) {
  var data = getSheet('Reminders Log').getDataRange().getValues();
  return data.slice(1).some(function(row) {
    return row[1] === type && row[5] === 'campaign-complete';
  });
}

// ── ACTION: GET STATS (for admin panel) ──
function getStats() {
  var sh   = getSheet('RSVPs');
  var data = sh.getDataRange().getValues();
  var eligibleCount = getResidents().length;
  if (data.length <= 1) {
    return { eligibleHouseholds: eligibleCount, totalHouseholds:0, yesHouseholds:0, noHouseholds:0,
             notRsvped: eligibleCount, totalAdults:0, totalKids5:0, totalKids15:0, totalAttending:0, rsvps:[] };
  }

  // Column layout: 0 Timestamp, 1 ResidentID, 2 House#, 3 Address, 4 Name1, 5 Name2,
  // 6 Last, 7 Email, 8 Email2, 9 RSVP, 10 Adults, 11 KidsU5, 12 Kids6-15,
  // 13 Event Suggestions, 14 No Reason, 15 No Reason Text, 16 Email Differs From Directory
  var rows = data.slice(1);
  var yesRows = rows.filter(function(r){ return r[9]==='yes'; });
  var noRows  = rows.filter(function(r){ return r[9]==='no'; });

  var totalAdults = yesRows.reduce(function(s,r){ return s + (parseInt(r[10])||0); }, 0);
  var totalKids5  = yesRows.reduce(function(s,r){ return s + (parseInt(r[11])||0); }, 0);
  var totalKids15 = yesRows.reduce(function(s,r){ return s + (parseInt(r[12])||0); }, 0);

  var rsvps = rows.map(function(r){
    return {
      addr: r[3], name: (r[4] || '') + (r[6] ? ' ' + r[6] : ''), // name1 + last only, name2 not used
      rsvp: r[9], adults: r[10], kids5: r[11], kids15: r[12]
    };
  });

  var respondedIds = {};
  rows.forEach(function(row) { respondedIds[String(row[1])] = true; });
  var notRsvped = Math.max(0, eligibleCount - Object.keys(respondedIds).length);

  return {
    eligibleHouseholds: eligibleCount,
    totalHouseholds: rows.length,
    yesHouseholds:   yesRows.length,
    noHouseholds:    noRows.length,
    notRsvped:       notRsvped,
    totalAdults:     totalAdults,
    totalKids5:      totalKids5,
    totalKids15:     totalKids15,
    totalAttending:  totalAdults + totalKids5 + totalKids15,
    rsvps:           rsvps
  };
}

// ── doPost ──
function doPost(e) {
  var cors = ContentService.createTextOutput('').setMimeType(ContentService.MimeType.JSON);
  try {
    var p = JSON.parse(e.postData.contents);
    var result;
    if      (p.action === 'submitRsvp')      result = submitRsvp(p);
    else if (p.action === 'submitSuggestion') result = submitSuggestion(p);
    else if (p.action === 'sendReminders') {
      if (!isAuthorizedAdmin()) { result = { error: 'Unauthorized' }; }
      else { result = sendReminders(p.type, p.testMode, p.testEmail); }
    }
    else if (p.action === 'previewReminders') {
      if (!isAuthorizedAdmin()) { result = { error: 'Unauthorized' }; }
      else { result = previewReminders(p.type); }
    }
    else if (p.action === 'sendTestConfirmation') {
      if (!isAuthorizedAdmin()) { result = { error: 'Unauthorized' }; }
      else { result = sendTestConfirmation(p.testEmail, p.rsvpType); }
    }
    else if (p.action === 'getStats') {
      if (!isAuthorizedAdmin()) { result = { error: 'Unauthorized' }; }
      else { result = getStats(); }
    }
    else result = { error: 'Unknown action' };
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── doGet (for admin stats) ──
function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action;
  var result;

  if (action === 'getStats') {
    if (!isAuthorizedAdmin()) { result = { error: 'Unauthorized' }; }
    else { result = getStats(); }
  } else if (action === 'searchResidents') {
    result = searchResidents(params.q);
  } else if (action === 'verifyPin') {
    result = verifyPin(params.residentId, params.pin);
  } else {
    result = { status: 'SHWE Social 2026 API' };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── TRIGGER SETUP INSTRUCTIONS ──
// Run these ONE TIME manually from the Apps Script editor:
//
// function setupTriggers() {
//   // Sept 27 at 9 AM — reminder to non-RSVPed + update to confirmed
//   ScriptApp.newTrigger('triggerSept27').timeBased()
//     .atDate(2026,9,27).atHour(9).create();
//   // Sept 30 at 9 AM — closing alert to non-RSVPed only
//   ScriptApp.newTrigger('triggerSept30').timeBased()
//     .atDate(2026,9,30).atHour(9).create();
// }
//
// function triggerSept27() {
//   sendReminders('reminder'); // non-RSVPed
//   sendReminders('update');   // already RSVPed
// }
// function triggerSept30() {
//   sendReminders('closing');  // non-RSVPed only
// }


// ══════════════════════════════════════════════════════════════════
// ADMIN DEPLOYMENT — READ BEFORE PUBLISHING
// ══════════════════════════════════════════════════════════════════
// This script must be deployed TWICE as two separate Web Apps from the
// same Apps Script project (Deploy → Manage deployments → New deployment):
//
// 1) PUBLIC (resident) deployment — unchanged from before:
//      Execute as: Me (StonehavenwoodsEast@gmail.com)
//      Who has access: Anyone
//    → this URL goes into 2026social.html (SCRIPT_URL)
//
// 2) ADMIN deployment — NEW, separate URL:
//      Execute as: User accessing the web app
//      Who has access: Anyone with Google account
//    → this URL goes into shwe-admin.html (SCRIPT_URL)
//    → Each admin's Google account must ALSO be shared as an Editor on
//      the "SHWE Annual Social 2026" Google Sheet, or their requests
//      will fail with a permissions error even after passing the
//      ADMIN_ALLOWLIST check above.
//
// No PIN or secret is stored in the admin HTML anymore. When an admin
// opens the admin URL, Google will prompt them to sign in (if not
// already) and the script checks their email against ADMIN_ALLOWLIST.
// Anyone not on the allowlist gets {error:'Unauthorized'} even if they
// are signed into a Google account.
