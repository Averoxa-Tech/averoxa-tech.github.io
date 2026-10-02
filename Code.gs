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
const ATTEMPT_HEADERS = ['submittedAt', 'attemptId', 'roll', 'name', 'dob', 'state', 'exam', 'mode', 'subjectsSel', 'chapters', 'level', 'source',
  'score', 'max', 'correct', 'incorrect', 'partial', 'unattempted', 'pending', 'seconds', 'subjectsJson', 'answersJson'];
const VALID_EXAMS = ['jm', 'ja', 'neet', 'cbse', 'isc'];

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
  if (b.action === 'qparse') return json_(qParse_(b)); // slow AI call: run outside the lock
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    switch (b.action) {
      case 'visit': return json_(logEvent_('visit', b));
      case 'start': return json_(logEvent_('start', b));
      case 'submit': return json_(saveAttempt_(b));
      case 'admin': return json_(adminData_(b));
      case 'qcheck': return json_(qAuth_(b) ? { ok: true } : { ok: false, error: 'unauthorized' });
      case 'qsave': return json_(qSave_(b));
      case 'qget': return json_(qGet_());
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
  if (VALID_EXAMS.indexOf(exam) === -1) return { ok: false, error: 'bad exam' };

  const sh = sheet_(ATTEMPTS_SHEET, ATTEMPT_HEADERS);
  const last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, ATTEMPT_HEADERS.length).getValues() : [];
  const id = String(b.attemptId);
  const SCORE_COL = 12; // 0-based index of 'score' in ATTEMPT_HEADERS

  let mine = data.find(function (r) { return String(r[1]) === id; });
  const num = function (v) { const n = Number(v); return isFinite(n) ? n : 0; };
  const score = num(b.score), max = num(b.max);
  if (max <= 0 || score > max) return { ok: false, error: 'invalid score' };
  const row = [new Date(), clean_(id, 40), clean_(b.roll, 40), clean_(b.name, 60), clean_(b.dob, 10),
    clean_(b.state, 50), exam, clean_(b.mode, 10), clean_(b.subjectsSel, 200), clean_(b.chapters, 300),
    clean_(b.level, 10), clean_(b.source, 10), score, max, num(b.correct), num(b.incorrect), num(b.partial),
    num(b.unattempted), num(b.pending), num(b.seconds),
    clean_(JSON.stringify(b.subjects || {}), 4000), clean_(JSON.stringify(b.answers || []), 30000)];
  if (!mine) {
    sh.appendRow(row);
    data.push(row);
  } else {
    // re-submission after self-grading written answers: update the existing row in place
    const rowIndex = data.indexOf(mine) + 2; // +1 for header, +1 for 1-based
    sh.getRange(rowIndex, 1, 1, ATTEMPT_HEADERS.length).setValues([row]);
    data[data.indexOf(mine)] = row;
  }
  const myScore = score;
  const scores = data.filter(function (r) { return r[6] === exam; }).map(function (r) { return Number(r[SCORE_COL]); });
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
      exam: r[6], mode: r[7], subjectsSel: r[8], chapters: r[9], level: r[10], source: r[11],
      score: Number(r[12]), max: Number(r[13]), correct: Number(r[14]),
      incorrect: Number(r[15]), partial: Number(r[16]), unattempted: Number(r[17]),
      pending: Number(r[18]), seconds: Number(r[19])
    };
  });

  return {
    ok: true,
    totals: { visits: visits, uniqueVisitors: Object.keys(unique).length, starts: starts, submitted: attempts.length },
    perDay: perDay,
    attempts: attempts
  };
}

/* ===================== QUESTION UPLOADER (added) =====================
 * Extra setup: Project Settings -> Script properties -> add ANTHROPIC_KEY = your Anthropic API key.
 * Then re-deploy (Deploy -> Manage deployments -> edit -> New version).
 */
const QUESTIONS_SHEET = 'Questions';
const Q_TYPES = ['mcq', 'multi', 'num', 'vsa', 'sa', 'case', 'la', 'la4', 'la6', 'la10'];

function qAuth_(b) {
  if (String(b.password || '') === ADMIN_PASSWORD) return true;
  Utilities.sleep(1200);
  return false;
}

function qSafe_(s) { // strip dangerous HTML from AI/PDF text
  return String(s == null ? '' : s)
    .replace(/<\/?(script|iframe|object|embed|style|link|meta)[^>]*>/gi, '')
    .replace(/\son\w+\s*=\s*("[^"]*"|'[^']*')/gi, '');
}

function qParse_(b) {
  if (!qAuth_(b)) return { ok: false, error: 'unauthorized' };
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!key) return { ok: false, error: 'ANTHROPIC_KEY script property not set' };
  const h = b.hint || {};
  const content = [];
  (b.files || []).slice(0, 10).forEach(function (f) {
    if (f.mime === 'application/pdf') content.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: f.data } });
    else content.push({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: f.data } });
  });
  content.push({ type: 'text', text: 'Extract every question from the material above and below.\n' + String(b.text || '').slice(0, 30000) });
  let system = 'You convert exam questions into JSON for a mock-test question bank. Reply with ONLY a JSON array, no prose, no code fences. ' +
    'Each item: {"t": one of mcq|multi|num|vsa|sa|case|la, "s": subject, "ch": chapter, "q": question text (HTML allowed: <sup>, <sub>, <br>; write maths in plain text/unicode), ' +
    '"o": [option texts without A/B/C/D labels] (mcq/multi only), "a": correct answer (mcq: 0-based index; multi: array of 0-based indices; num: number; omit for written), ' +
    '"x": short worked solution or model answer, "m": marks (written only), "rb": [[point, marks],...] marking scheme (written only)}. ' +
    'Use t=mcq for single-correct options, multi if several can be correct, num for numerical answer, vsa(2 marks)/sa(3)/la(5) for written. ' +
    'If the answer is not given in the material, solve it yourself. Default subject: ' + (h.s || 'infer') + '; chapter: ' + (h.ch || 'infer') + '. Never invent questions that are not in the material.';
  if (b.mode !== 'generate') {
    system += ' ANSWERS: the uploaded questions usually come WITHOUT answers. For EVERY question you must work out the answer yourself: ' +
      'for mcq/multi/num solve it step by step, double-check, and set "a" and a full worked solution in "x"; ' +
      'for written questions (vsa/sa/case/la) write a complete model answer in "x" and a point-wise marking scheme in "rb" ' +
      '(e.g. [["Defines the term",1],["Gives example",1]]) with "m" equal to the sum of the marks. ' +
      'If the material already gives an answer, check it; if it looks wrong, use the correct one and mention the difference in "x". Never leave "a", "x" or "rb" empty.';
  }
  if (b.mode === 'generate') {
    const n = Math.max(1, Math.min(30, Number(b.count) || 10));
    system = system.replace('Never invent questions that are not in the material.', '') +
      ' MODE OVERRIDE - DO NOT EXTRACT. The material is only a SAMPLE. Detect its subject, language (if Hindi, write everything in Devanagari Hindi; if English, English) and style, and write ' + n +
      ' brand-new ORIGINAL questions of the SAME subject, language, difficulty and format, covering different chapters/topics of that subject. ' +
      'Never copy or lightly reword a sample question: use new scenarios, numbers, passages and concepts, and make the questions different from each other. ' +
      'For every written question (vsa/sa/case/la) "rb" MUST be a point-wise marking scheme, e.g. [["Defines the term correctly",1],["Gives one example",1],["Correct conclusion",1]], and "m" must equal the sum of the marks. ' +
      'For mcq/multi/num, carefully compute and double-check the correct answer, and put the full working in "x". Set "s" to the subject of each question.';
    content[content.length - 1] = { type: 'text', text: 'Sample material is above/below. Now write ' + n + ' new unique questions as specified.\n' + String(b.text || '').slice(0, 30000) };
  }
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: 'claude-sonnet-5-5', max_tokens: 16000, system: system, messages: [{ role: 'user', content: content }] })
  });
  if (res.getResponseCode() !== 200) return { ok: false, error: 'AI error ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200) };
  const txt = (JSON.parse(res.getContentText()).content || []).map(function (c) { return c.text || ''; }).join('');
  try {
    const arr = JSON.parse(txt.replace(/^```(?:json)?|```$/gm, '').trim());
    return { ok: true, questions: Array.isArray(arr) ? arr : [] };
  } catch (e) {
    return { ok: false, error: 'AI reply could not be read, try fewer pages at a time' };
  }
}

function qSave_(b) {
  if (!qAuth_(b)) return { ok: false, error: 'unauthorized' };
  const sh = sheet_(QUESTIONS_SHEET, ['time', 'json']);
  let n = 0;
  (b.questions || []).slice(0, 200).forEach(function (q) {
    if (!q || Q_TYPES.indexOf(q.t) === -1 || !q.q || !q.s || !q.ch) return;
    const ex = (q.ex || []).filter(function (e) { return VALID_EXAMS.indexOf(e) !== -1; });
    if (!ex.length) return;
    q.ex = ex; q.q = qSafe_(q.q); q.x = qSafe_(q.x);
    if (q.o) q.o = q.o.map(qSafe_);
    sh.appendRow([new Date(), JSON.stringify(q).slice(0, 45000)]);
    n++;
  });
  return { ok: true, saved: n };
}

function qGet_() {
  const sh = sheet_(QUESTIONS_SHEET, ['time', 'json']);
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 2, last - 1, 1).getValues() : [];
  const items = [];
  rows.forEach(function (r) { try { items.push(JSON.parse(r[0])); } catch (e) { /* skip bad row */ } });
  return { ok: true, items: items };
}
