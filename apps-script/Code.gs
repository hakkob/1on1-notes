// ============================================================
// 1-on-1 Notes — Apps Script Web App
// Deploy as: Execute as Me | Access: Anyone
// ============================================================

const SS_ID       = '1sfm8FVFLPDFnJZ3059q936jUAUwJ5p8aUddB23gWbzk';
const OWNER_EMAIL = 'felipe.jacob.g@gmail.com';

function doPost(e) {
  try {
    const req   = JSON.parse(e.postData.contents);
    const email = verifyToken(req.token);
    if (!email || email.toLowerCase() !== OWNER_EMAIL.toLowerCase())
      return jsonOut({ error: 'Unauthorized' });

    const ss = SpreadsheetApp.openById(SS_ID);

    switch (req.action) {
      case 'load':             return jsonOut(handleLoad(ss));
      case 'savePerson':       return jsonOut(handleSavePerson(ss, req.data));
      case 'deletePerson':     return jsonOut(handleDeletePerson(ss, req.id));
      case 'saveSession':      return jsonOut(handleSaveSession(ss, req.data));
      case 'deleteSession':    return jsonOut(handleDeleteSession(ss, req.id));
      case 'addActionItem':    return jsonOut(handleAddActionItem(ss, req.data));
      case 'toggleActionItem': return jsonOut(handleToggleActionItem(ss, req.id));
      case 'updateActionItem': return jsonOut(handleUpdateActionItem(ss, req.id, req.text));
      case 'deleteActionItem': return jsonOut(handleDeleteActionItem(ss, req.id));
      case 'search':           return jsonOut(handleSearch(ss, req.query));
      default:                 return jsonOut({ error: 'Unknown action' });
    }
  } catch (err) {
    return jsonOut({ error: err.toString() });
  }
}

// ── HELPERS ─────────────────────────────────────────────────
function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyToken(token) {
  if (!token) return null;
  try {
    const r = UrlFetchApp.fetch(
      'https://www.googleapis.com/oauth2/v3/userinfo',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (r.getResponseCode() !== 200) return null;
    const d = JSON.parse(r.getContentText());
    return d.email ? d.email.toLowerCase() : null;
  } catch (e) { return null; }
}

function getSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function findRow(sheet, id) {
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function sheetRows(sheet) {
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0]);
}

function uid(prefix) {
  return prefix + '_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
}

// ── LOAD ────────────────────────────────────────────────────
function handleLoad(ss) {
  const pSheet = getSheet(ss, 'People');
  const sSheet = getSheet(ss, 'Sessions');
  const aSheet = getSheet(ss, 'ActionItems');

  if (pSheet.getLastRow() === 0) pSheet.appendRow(['id','name','role','company','cadence','created_at','team']);
  if (sSheet.getLastRow() === 0) sSheet.appendRow(['id','person_id','date','went_well','to_improve','private_notes','next_agenda','created_at']);
  if (aSheet.getLastRow() === 0) aSheet.appendRow(['id','session_id','person_id','text','status','created_at']);

  const people = sheetRows(pSheet).map(r => ({
    id: String(r[0]), name: String(r[1]), role: String(r[2] || ''),
    company: String(r[3] || ''), cadence: String(r[4] || 'Weekly'), createdAt: String(r[5] || ''),
    team: String(r[6] || '')
  }));

  const sessions = sheetRows(sSheet).map(r => ({
    id: String(r[0]), personId: String(r[1]), date: String(r[2] || ''),
    wentWell: String(r[3] || ''), toImprove: String(r[4] || ''),
    privateNotes: String(r[5] || ''), nextAgenda: String(r[6] || ''),
    createdAt: String(r[7] || '')
  }));

  const actionItems = sheetRows(aSheet).map(r => ({
    id: String(r[0]), sessionId: String(r[1]), personId: String(r[2]),
    text: String(r[3] || ''), status: String(r[4] || 'open'), createdAt: String(r[5] || '')
  }));

  return { people, sessions, actionItems };
}

// ── PEOPLE ──────────────────────────────────────────────────
function handleSavePerson(ss, data) {
  const sheet  = getSheet(ss, 'People');
  if (!data.id) data.id = uid('p');
  const row    = [data.id, data.name, data.role || '', data.company || '', data.cadence || 'Weekly', new Date().toISOString(), data.team || ''];
  const rowNum = findRow(sheet, data.id);
  if (rowNum > 0) sheet.getRange(rowNum, 1, 1, 7).setValues([row]);
  else sheet.appendRow(row);
  return { ok: true, id: data.id };
}

function handleDeletePerson(ss, id) {
  const pSheet = getSheet(ss, 'People');
  const sSheet = getSheet(ss, 'Sessions');
  const aSheet = getSheet(ss, 'ActionItems');

  const pRow = findRow(pSheet, id);
  if (pRow > 0) pSheet.getRange(pRow, 1, 1, 7).clearContent();

  const sVals = sSheet.getDataRange().getValues();
  for (let i = sVals.length - 1; i >= 1; i--) {
    if (String(sVals[i][1]) === String(id))
      sSheet.getRange(i + 1, 1, 1, 8).clearContent();
  }

  const aVals = aSheet.getDataRange().getValues();
  for (let i = aVals.length - 1; i >= 1; i--) {
    if (String(aVals[i][2]) === String(id))
      aSheet.getRange(i + 1, 1, 1, 6).clearContent();
  }

  return { ok: true };
}

// ── SESSIONS ─────────────────────────────────────────────────
function handleSaveSession(ss, data) {
  const sSheet = getSheet(ss, 'Sessions');
  const aSheet = getSheet(ss, 'ActionItems');

  const isNew = !data.id;
  if (isNew) data.id = uid('s');

  const row    = [data.id, data.personId, data.date, data.wentWell || '', data.toImprove || '', data.privateNotes || '', data.nextAgenda || '', new Date().toISOString()];
  const rowNum = findRow(sSheet, data.id);
  if (rowNum > 0) sSheet.getRange(rowNum, 1, 1, 8).setValues([row]);
  else sSheet.appendRow(row);

  // Only parse action items for new sessions; edits manage items from detail view
  if (isNew && data.actionItemLines) {
    data.actionItemLines.filter(l => l.trim()).forEach(text => {
      aSheet.appendRow([uid('a'), data.id, data.personId, text.trim(), 'open', new Date().toISOString()]);
    });
  }

  return { ok: true, id: data.id };
}

function handleDeleteSession(ss, id) {
  const sSheet = getSheet(ss, 'Sessions');
  const aSheet = getSheet(ss, 'ActionItems');

  const sRow = findRow(sSheet, id);
  if (sRow > 0) sSheet.getRange(sRow, 1, 1, 8).clearContent();

  const aVals = aSheet.getDataRange().getValues();
  for (let i = aVals.length - 1; i >= 1; i--) {
    if (String(aVals[i][1]) === String(id))
      aSheet.getRange(i + 1, 1, 1, 6).clearContent();
  }

  return { ok: true };
}

// ── ACTION ITEMS ─────────────────────────────────────────────
function handleAddActionItem(ss, data) {
  const sheet = getSheet(ss, 'ActionItems');
  const id    = uid('a');
  sheet.appendRow([id, data.sessionId, data.personId, data.text, 'open', new Date().toISOString()]);
  return { ok: true, id };
}

function handleToggleActionItem(ss, id) {
  const sheet  = getSheet(ss, 'ActionItems');
  const rowNum = findRow(sheet, id);
  if (rowNum < 0) return { error: 'Not found' };
  const current = sheet.getRange(rowNum, 5).getValue();
  const next    = current === 'done' ? 'open' : 'done';
  sheet.getRange(rowNum, 5).setValue(next);
  return { ok: true, status: next };
}

function handleUpdateActionItem(ss, id, text) {
  const sheet  = getSheet(ss, 'ActionItems');
  const rowNum = findRow(sheet, id);
  if (rowNum < 0) return { error: 'Not found' };
  sheet.getRange(rowNum, 4).setValue(text);
  return { ok: true };
}

function handleDeleteActionItem(ss, id) {
  const sheet  = getSheet(ss, 'ActionItems');
  const rowNum = findRow(sheet, id);
  if (rowNum < 0) return { error: 'Not found' };
  sheet.getRange(rowNum, 1, 1, 6).clearContent();
  return { ok: true };
}

// ── SEARCH ───────────────────────────────────────────────────
function handleSearch(ss, query) {
  if (!query || query.trim().length < 2) return { results: [] };
  const q = query.toLowerCase().trim();

  const people   = sheetRows(getSheet(ss, 'People')).map(r => ({ id: String(r[0]), name: String(r[1]) }));
  const sessions = sheetRows(getSheet(ss, 'Sessions')).map(r => ({
    id: String(r[0]), personId: String(r[1]), date: String(r[2]),
    wentWell: String(r[3] || ''), toImprove: String(r[4] || ''), nextAgenda: String(r[6] || '')
  }));

  const nameMap = {};
  people.forEach(p => { nameMap[p.id] = p.name; });

  const results = [];
  const seen    = new Set();

  sessions.forEach(s => {
    const fields  = [s.wentWell, s.toImprove, s.nextAgenda];
    const matched = fields.find(f => f.toLowerCase().includes(q));
    if (matched && !seen.has(s.id)) {
      seen.add(s.id);
      results.push({
        sessionId:  s.id,
        personId:   s.personId,
        personName: nameMap[s.personId] || 'Unknown',
        date:       s.date,
        snippet:    matched.slice(0, 120)
      });
    }
    if (nameMap[s.personId] && nameMap[s.personId].toLowerCase().includes(q)) {
      const key = 'person_' + s.personId;
      if (!seen.has(key)) {
        seen.add(key);
        results.push({
          sessionId:  s.id,
          personId:   s.personId,
          personName: nameMap[s.personId],
          date:       s.date,
          snippet:    'Person match'
        });
      }
    }
  });

  people.filter(p => p.name.toLowerCase().includes(q)).forEach(p => {
    const key = 'person_' + p.id;
    if (!seen.has(key)) {
      seen.add(key);
      results.push({ sessionId: null, personId: p.id, personName: p.name, date: '', snippet: 'Person match' });
    }
  });

  return { results: results.slice(0, 20) };
}

// ── WEEKLY DIGEST ────────────────────────────────────────────
function sendWeeklyDigest() {
  const ss          = SpreadsheetApp.openById(SS_ID);
  const people      = sheetRows(getSheet(ss, 'People')).map(r => ({ id: String(r[0]), name: String(r[1]), cadence: String(r[4] || 'Weekly') }));
  const sessions    = sheetRows(getSheet(ss, 'Sessions')).map(r => ({ personId: String(r[1]), date: String(r[2]) }));
  const actionItems = sheetRows(getSheet(ss, 'ActionItems')).map(r => ({ personId: String(r[2]), text: String(r[3] || ''), status: String(r[4] || 'open') }));

  const nameMap       = {};
  people.forEach(p => { nameMap[p.id] = p.name; });
  const cadenceDaysMap = { 'Weekly': 7, 'Bi-weekly': 14, 'Monthly': 30 };
  const now            = new Date();

  const overdue = people.filter(p => {
    const days = cadenceDaysMap[p.cadence] || 7;
    const last = sessions.filter(s => s.personId === p.id).sort((a, b) => b.date.localeCompare(a.date))[0];
    if (!last) return true;
    return Math.floor((now - new Date(last.date)) / 86400000) >= days;
  });

  const openItems = actionItems.filter(a => a.status === 'open');

  if (!openItems.length && !overdue.length) return;

  let body = 'Weekly Digest — 1-on-1 Notes\n\n';
  if (overdue.length) {
    body += 'OVERDUE CHECK-INS\n';
    overdue.forEach(p => { body += '  • ' + p.name + '\n'; });
    body += '\n';
  }
  if (openItems.length) {
    body += 'OPEN ACTION ITEMS\n';
    openItems.forEach(a => { body += '  • [' + (nameMap[a.personId] || '?') + '] ' + a.text + '\n'; });
  }

  MailApp.sendEmail(OWNER_EMAIL, '[1-on-1 Notes] Weekly Digest', body);
}
