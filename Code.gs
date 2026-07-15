// ═══════════════════════════════════════════════
//  PádelTurnos — Google Apps Script Backend
//  Factory Padel Córdoba
// ═══════════════════════════════════════════════

const SHEET_ID = '17kXOZDi4dcuJhcd7WXM8V7u5Q5L_cWaQL80Egy8nTDY';

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

// ── CORS headers ──────────────────────────────
function cors(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    if (action === 'getAll')      result = getAll();
    else if (action === 'ping')   result = { ok: true };
    else result = { error: 'Acción no reconocida' };
  } catch(err) {
    result = { error: err.message };
  }
  return cors(ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON));
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch(err) { body = {}; }
  const action = body.action;
  let result;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    if      (action === 'saveReserva')   result = saveReserva(body.data);
    else if (action === 'deleteReserva') result = deleteReserva(body.key);
    else if (action === 'saveFijo')      result = saveFijo(body.data);
    else if (action === 'deleteFijo')    result = deleteFijo(body.key);
    else if (action === 'saveJugador')   result = saveJugador(body.data);
    else if (action === 'deleteJugador') result = deleteJugador(body.id);
    else if (action === 'importFijos')   result = importFijos(body.fijos);
    else if (action === 'clearFijos')    result = clearFijos();
    else result = { error: 'Acción no reconocida' };
  } catch(err) {
    result = { error: err.message };
  } finally {
    lock.releaseLock();
  }
  return cors(ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON));
}

// ── GET ALL ───────────────────────────────────
function getAll() {
  return {
    reservas:  sheetToObjects('reservas'),
    fijos:     sheetToObjects('fijos'),
    jugadores: sheetToObjects('jugadores'),
  };
}

// ── RESERVAS ──────────────────────────────────
function saveReserva(data) {
  const sheet = getSheet('reservas');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  // buscar si ya existe por key
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][headers.indexOf('key')] === data.key) {
      // actualizar fila
      updateRow(sheet, i + 1, headers, data);
      return { ok: true, action: 'updated' };
    }
  }
  // insertar nueva fila
  if (rows.length === 1 && rows[0].length === 1 && rows[0][0] === '') {
    initSheet(sheet, ['key','fecha','slot','cancha','nombre','tipo','esFijo','jugadorId','seña','ts']);
  }
  appendRow(sheet, ['key','fecha','slot','cancha','nombre','tipo','esFijo','jugadorId','seña','ts'], data);
  return { ok: true, action: 'created' };
}

function deleteReserva(key) {
  deleteRowByField('reservas', 'key', key);
  return { ok: true };
}

// ── FIJOS ─────────────────────────────────────
function saveFijo(data) {
  const sheet = getSheet('fijos');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][headers.indexOf('key')] === data.key) {
      updateRow(sheet, i + 1, headers, data);
      return { ok: true, action: 'updated' };
    }
  }
  appendRow(sheet, ['key','dow','slot','cancha','nombre','tipo','jugadorId','ts'], data);
  return { ok: true, action: 'created' };
}

function deleteFijo(key) {
  deleteRowByField('fijos', 'key', key);
  return { ok: true };
}

function clearFijos() {
  const sheet = getSheet('fijos');
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  return { ok: true };
}

function importFijos(fijos) {
  clearFijos();
  const sheet = getSheet('fijos');
  fijos.forEach(f => {
    appendRow(sheet, ['key','dow','slot','cancha','nombre','tipo','jugadorId','ts'], f);
  });
  return { ok: true, count: fijos.length };
}

// ── JUGADORES ─────────────────────────────────
function saveJugador(data) {
  const sheet = getSheet('jugadores');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][headers.indexOf('id')] === data.id) {
      updateRow(sheet, i + 1, headers, data);
      return { ok: true, action: 'updated' };
    }
  }
  appendRow(sheet, ['id','nombre','tel','email','turnos','ts'], data);
  return { ok: true, action: 'created' };
}

function deleteJugador(id) {
  deleteRowByField('jugadores', 'id', id);
  return { ok: true };
}

// ── HELPERS ───────────────────────────────────
function sheetToObjects(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function initSheet(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function appendRow(sheet, headers, data) {
  // asegurarse que la hoja tenga encabezados
  const existing = sheet.getDataRange().getValues();
  if (existing.length === 0 || (existing.length === 1 && existing[0][0] === '')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
}

function updateRow(sheet, rowIndex, headers, data) {
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function deleteRowByField(sheetName, field, value) {
  const sheet = getSheet(sheetName);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const colIdx = headers.indexOf(field);
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][colIdx]) === String(value)) {
      sheet.deleteRow(i + 1);
    }
  }
}
