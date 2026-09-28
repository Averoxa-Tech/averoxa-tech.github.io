/**
 * Averoxa Mock Test — backend (Google Apps Script)
 *
 * HOW TO USE
 *  1. Open a new Google Sheet -> Extensions -> Apps Script.
 *  2. Delete the sample code, paste this whole file.
 *  3. Change ADMIN_PASSWORD below to your own strong password.
 *  4. Run the function "setup" once (it creates the sheets and asks for permission).
 *  5. Deploy -> New deployment -> type: Web app
 *        Execute as: Me
 *        Who has access: Anyone
 *     Copy the Web app URL and paste it into mock-test.html and mock-admin.html (API_URL).
 *
 * SECURITY: keep this file private. Do NOT upload it to a public GitHub repository,
 * because it contains your admin password.
 */

const ADMIN_PASSWORD = 'CHANGE-THIS-PASSWORD';

const EVENTS_SHEET = 'Events';
const ATTEMPTS_SHEET = 'Attempts';
const EVENTS_HEADERS = ['time', 'type', 'visitorId', 'roll', 'exam', 'state', 'userAgent'];
const ATTEMPT_HEADERS = ['submittedAt', 'attemptId', 'roll', 'name', 'dob', 'state', 'exam',
  'score', 'max', 'correct', 'incorrect', 'partial', 'unattempted', 'seconds', 'subjectsJson', 'answersJson'];

function setup() {
  sheet_(EVENTS_SHEET, EVENTS_HEADERS);
  sheet_(ATTEMPTS_SHEET, ATTEMPT_HEADERS);
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Neutralise spreadsheet formulas and limit length.
function clean_(v, max) {
  return String(v === undefined || v === null ? '' : v).slice(0, max).replace(/^[=+\-@]/, "'$&");
}

function doGet() {
  return json_({ ok: true, service: 'Averoxa Mock Test API' });
}

function doPost(e) {
  let b;
  try {
    b = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad request' });
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    switch (b.action) {
      case 'visit': return json_(logEvent_('visit', b));
      case 'start': return json_(logEvent_('start', b));
      case 'submit': return json_(saveAttempt_(b));
      case 'admin': return json_(adminData_(b));
      default: return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) { /* ignore */ }
  }
}

function logEvent_(type, b) {
  const sh = sheet_(EVENTS_SHEET, EVENTS_HEADERS);
  sh.appendRow([new Date(), type, clean_(b.visitorId, 40), clean_(b.roll, 40),
    clean_(b.exam, 10), clean_(b.state, 50), clean_(b.ua, 120)]);
  return { ok: true };
}

function saveAttempt_(b) {
  if (!b.attemptId || !b.roll || !b.name) return { ok: false, error: 'missing fields' };
  const exam = String(b.exam);
  if (exam !== 'main' && exam !== 'adv') return { ok: false, error: 'bad exam' };

  const sh = sheet_(ATTEMPTS_SHEET, ATTEMPT_HEADERS);
  const last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, ATTEMPT_HEADERS.length).getValues() : [];
  const id = String(b.attemptId);

  let mine = data.find(function (r) { return String(r[1]) === id; });
  if (!mine) {
    const num = function (v) { const n = Number(v); return isFinite(n) ? n : 0; };
    const score = num(b.score), max = num(b.max);
    if (max <= 0 || score > max) return { ok: false, error: 'invalid score' };
    mine = [new Date(), clean_(id, 40), clean_(b.roll, 40), clean_(b.name, 60), clean_(b.dob, 10),
      clean_(b.state, 50), exam, score, max, num(b.correct), num(b.incorrect), num(b.partial),
      num(b.unattempted), num(b.seconds),
      clean_(JSON.stringify(b.subjects || {}), 4000), clean_(JSON.stringify(b.answers || []), 30000)];
    sh.appendRow(mine);
    data.push(mine);
  }
  const myScore = Number(mine[7]);
  const scores = data.filter(function (r) { return r[6] === exam; }).map(function (r) { return Number(r[7]); });
  const rank = 1 + scores.filter(function (s) { return s > myScore; }).length;
  return { ok: true, rank: rank, total: scores.length };
}

function adminData_(b) {
  if (String(b.password || '') !== ADMIN_PASSWORD) {
    Utilities.sleep(1200); // slows down password guessing
    return { ok: false, error: 'unauthorized' };
  }
  const tz = 'UTC';

  const ev = sheet_(EVENTS_SHEET, EVENTS_HEADERS);
  const evLast = ev.getLastRow();
  const events = evLast > 1 ? ev.getRange(2, 1, evLast - 1, EVENTS_HEADERS.length).getValues() : [];
  let visits = 0, starts = 0;
  const unique = {};
  const perDay = {};
  events.forEach(function (r) {
    if (r[1] === 'visit') {
      visits++;
      unique[r[2]] = true;
      const d = Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd');
      perDay[d] = (perDay[d] || 0) + 1;
    } else if (r[1] === 'start') {
      starts++;
    }
  });

  const at = sheet_(ATTEMPTS_SHEET, ATTEMPT_HEADERS);
  const atLast = at.getLastRow();
  const rows = atLast > 1 ? at.getRange(2, 1, atLast - 1, ATTEMPT_HEADERS.length).getValues() : [];
  const attempts = rows.map(function (r) {
    return {
      submittedAt: new Date(r[0]).toISOString(), roll: r[2], name: r[3], dob: r[4], state: r[5],
      exam: r[6], score: Number(r[7]), max: Number(r[8]), correct: Number(r[9]),
      incorrect: Number(r[10]), partial: Number(r[11]), unattempted: Number(r[12]), seconds: Number(r[13])
    };
  });

  return {
    ok: true,
    totals: { visits: visits, uniqueVisitors: Object.keys(unique).length, starts: starts, submitted: attempts.length },
    perDay: perDay,
    attempts: attempts
  };
}
