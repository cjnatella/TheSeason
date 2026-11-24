/**
 * The Season: link between the app and the Engine sheet.
 * Paste into Extensions > Apps Script in "The Season: Engine", then deploy as a web app
 * (Execute as: Me, Who has access: Anyone). The app sends the token below with every request.
 */
const TOKEN = '__TOKEN__';

const DAILY_KEYS = {
  'Date': 'date', 'Week': 'week', 'Weight (kg)': 'weight', 'Trend (kg)': 'trend', 'Steps': 'steps',
  'Evening walk': 'walk', 'Protein target hit': 'protein', 'In bed on time': 'bed',
  'Cooked dinner at home': 'cooked', 'Mood (1–5)': 'mood', 'Energy (1–5)': 'energy',
  'Food spend (£)': 'spend', 'Calories (from League Two)': 'calories', 'Notes': 'notes',
  'Basics hit (of 4)': 'basics', 'Training-ground points': 'tg'
};
const DAY_WRITABLE = ['weight', 'steps', 'walk', 'protein', 'bed', 'cooked', 'mood', 'energy', 'spend', 'calories', 'notes'];
const DAY_BOOLEAN = ['walk', 'protein', 'bed', 'cooked'];

const FIXTURE_KEYS = {
  'Week': 'week', 'Date': 'date', 'Day': 'day', 'Fixture': 'type', 'Session': 'session',
  'Energy (1–5)': 'energy', 'Session size': 'size', 'Result (W/D/L)': 'result', 'Points': 'points',
  'Notes': 'notes', 'Day offset': 'offset', 'Unbeaten run': 'run'
};

function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== TOKEN) return json_({ ok: false, error: 'The token does not match. Check it in the app settings.' });
  return json_({ ok: true, data: readAll_() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'The request could not be read.' }); }
  if (body.token !== TOKEN) return json_({ ok: false, error: 'The token does not match. Check it in the app settings.' });
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (body.action === 'day') saveDay_(body);
    else if (body.action === 'fixture') saveFixture_(body);
    else if (body.action === 'lifts') saveLifts_(body);
    else return json_({ ok: false, error: 'Unknown action: ' + body.action });
    SpreadsheetApp.flush();
    return json_({ ok: true, data: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- reading ---------- */

function readAll_() {
  const ss = SpreadsheetApp.getActive();
  return {
    fetchedAt: new Date().toISOString(),
    settings: readSettings_(ss),
    daily: readTable_(ss.getSheetByName('Daily Log'), 1, DAILY_KEYS),
    fixtures: readTable_(ss.getSheetByName('Fixtures'), 1, FIXTURE_KEYS).filter(function (r) { return r.date; }),
    divisions: readDivisions_(ss),
    trophies: readTrophies_(ss),
    programme: readProgramme_(ss),
    lifts: readLifts_(ss)
  };
}

function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }

function cell_(v) {
  if (v instanceof Date) {
    if (v.getFullYear() < 1910) return Utilities.formatDate(v, tz_(), 'HH:mm');
    return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  }
  return v;
}

function readTable_(sh, headerRow, keys) {
  const values = sh.getDataRange().getValues();
  const header = values[headerRow - 1];
  const out = [];
  for (let i = headerRow; i < values.length; i++) {
    const row = { row: i + 1 };
    header.forEach(function (h, c) { if (keys[h]) row[keys[h]] = cell_(values[i][c]); });
    out.push(row);
  }
  return out;
}

function readSettings_(ss) {
  const sh = ss.getSheetByName('Settings');
  const v = sh.getRange('A5:B22').getValues();
  const byLabel = {};
  v.forEach(function (r) { if (r[0]) byLabel[r[0]] = cell_(r[1]); });
  return {
    kickoff: cell_(sh.getRange('B5').getValue()),
    heightCm: sh.getRange('B6').getValue(),
    startWeight: sh.getRange('B7').getValue(),
    goalWeight: sh.getRange('B8').getValue(),
    proteinTarget: sh.getRange('B10').getValue(),
    stepTarget: sh.getRange('B11').getValue(),
    bedBy: cell_(sh.getRange('B12').getValue()),
    foodBudget: sh.getRange('B13').getValue(),
    moodFloor: sh.getRange('B19').getValue(),
    all: byLabel
  };
}

function readDivisions_(ss) {
  const sh = ss.getSheetByName('Divisions');
  const v = sh.getRange('A5:M9').getValues();
  const rows = v.map(function (r) {
    return {
      name: r[0], start: cell_(r[1]), end: cell_(r[2]), weeks: r[3], fixturesPerWeek: r[4],
      available: r[5], share: r[6], line: r[7], points: r[8], milestone: r[9], status: r[10],
      milestoneDone: r[11] === true, current: r[12] === true
    };
  });
  return {
    rows: rows,
    currentIndex: Number(sh.getRange('B12').getValue()) - 1,
    daysIn: sh.getRange('B13').getValue(),
    ghost: sh.getRange('B14').getValue()
  };
}

function readTrophies_(ss) {
  const sh = ss.getSheetByName('Trophy Room');
  const last = sh.getLastRow();
  if (last < 5) return [];
  return sh.getRange(5, 1, last - 4, 3).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { name: r[0], how: r[1], earned: cell_(r[2]) || '' }; });
}

function readProgramme_(ss) {
  const sh = ss.getSheetByName('Programme');
  return sh.getRange('A5:F17').getValues()
    .filter(function (r) { return r[1]; })
    .map(function (r) { return { workout: r[0], exercise: r[1], full: r[2], standard: r[3], floor: r[4], progress: r[5] }; });
}

function readLifts_(ss) {
  const sh = ss.getSheetByName('Training Log');
  const colA = sh.getRange('A2:A').getValues();
  let n = 0;
  for (let i = 0; i < colA.length; i++) if (colA[i][0] !== '') n = i + 1;
  if (n === 0) return [];
  const start = Math.max(1, n - 400);
  return sh.getRange(start + 1, 1, n - start + 1, 9).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      return { date: cell_(r[0]), session: r[1], exercise: r[2], set: r[3], weight: r[4], reps: r[5], rir: r[6], e1rm: r[7], notes: r[8] };
    });
}

/* ---------- writing ---------- */

function parseDate_(s) { return Utilities.parseDate(String(s), tz_(), 'yyyy-MM-dd'); }

function saveDay_(body) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Daily Log');
  const values = sh.getDataRange().getValues();
  const header = values[0];
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] instanceof Date && Utilities.formatDate(values[i][0], tz_(), 'yyyy-MM-dd') === body.date) { rowIndex = i + 1; break; }
  }
  if (rowIndex < 0) throw new Error('No Daily Log row for ' + body.date);
  const fields = body.fields || {};
  header.forEach(function (h, c) {
    const key = DAILY_KEYS[h];
    if (!key || DAY_WRITABLE.indexOf(key) < 0 || !(key in fields)) return;
    let val = fields[key];
    if (DAY_BOOLEAN.indexOf(key) >= 0) val = val === true;
    else if (val === null || val === undefined) val = '';
    sh.getRange(rowIndex, c + 1).setValue(val);
  });
}

function saveFixture_(body) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Fixtures');
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = Number(body.row);
  if (!row || row < 2) throw new Error('Missing fixture row');
  const dateCol = header.indexOf('Date') + 1;
  const d = sh.getRange(row, dateCol).getValue();
  if (!(d instanceof Date) || Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') !== body.date) throw new Error('The fixture has moved in the sheet. Refresh and try again.');
  const set = function (name, v) { const c = header.indexOf(name) + 1; if (c > 0) sh.getRange(row, c).setValue(v); };
  if ('energy' in body) set('Energy (1–5)', body.energy === null ? '' : body.energy);
  if ('result' in body) set('Result (W/D/L)', body.result || '');
  if ('notes' in body) set('Notes', body.notes || '');
}

function saveLifts_(body) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Training Log');
  const sets = body.sets || [];
  if (!sets.length) return;
  const colA = sh.getRange('A2:A').getValues();
  let next = 2;
  for (let i = 0; i < colA.length; i++) if (colA[i][0] !== '') next = i + 3;
  const date = parseDate_(body.date);
  const ag = sets.map(function (s) { return [date, body.session || '', s.exercise, s.set, s.weight === '' ? '' : Number(s.weight), s.reps === '' ? '' : Number(s.reps), s.rir === '' ? '' : Number(s.rir)]; });
  sh.getRange(next, 1, ag.length, 7).setValues(ag);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
