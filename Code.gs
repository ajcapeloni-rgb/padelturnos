// ═══════════════════════════════════════════════
//  PádelTurnos — Google Apps Script Backend
//  Factory Padel Córdoba
// ═══════════════════════════════════════════════

const SHEET_ID = '17kXOZDi4dcuJhcd7WXM8V7u5Q5L_cWaQL80Egy8nTDY';

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// ── CORS ────────────────────────────────────────
// Nota: ContentService.TextOutput no soporta setHeader() en Apps Script,
// así que no hay headers CORS manuales que setear acá. Los deployments
// "Anyone" de Apps Script ya permiten fetch() desde cualquier origen.
function cors(output) {
  return output;
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
    if      (action === 'saveReserva')       result = saveReserva(body.data);
    else if (action === 'deleteReserva')     result = deleteReserva(body.key);
    else if (action === 'saveFijo')          result = saveFijo(body.data);
    else if (action === 'deleteFijo')        result = deleteFijo(body.key);
    else if (action === 'saveJugador')       result = saveJugador(body.data);
    else if (action === 'deleteJugador')     result = deleteJugador(body.id);
    else if (action === 'importFijos')       result = importFijos(body.fijos);
    else if (action === 'clearFijos')        result = clearFijos();
    else if (action === 'saveReservasBatch') result = saveReservasBatch(body);
    else if (action === 'saveExcepcion')     result = saveExcepcion(body.data);
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
    reservas:     sheetToObjects('reservas'),
    fijos:        sheetToObjects('fijos'),
    jugadores:    sheetToObjects('jugadores'),
    excepciones:  sheetToObjects('excepciones'),
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

// ── GUARDADO EN LOTE (reserva de varios slots en una sola llamada) ──
function saveReservasBatch(body) {
  const result = { ok: true };
  if (body.jugador) saveJugador(body.jugador);
  if (body.reservas && body.reservas.length) {
    result.reservas = batchUpsert('reservas', ['key','fecha','slot','cancha','nombre','tipo','esFijo','jugadorId','seña','ts'], 'key', body.reservas);
  }
  if (body.fijos && body.fijos.length) {
    result.fijos = batchUpsert('fijos', ['key','dow','slot','cancha','nombre','tipo','jugadorId','ts'], 'key', body.fijos);
  }
  return result;
}

function batchUpsert(sheetName, headers, keyField, items) {
  const sheet = getSheet(sheetName);
  let rows = sheet.getDataRange().getValues();
  if (rows.length === 1 && rows[0].length === 1 && rows[0][0] === '') {
    initSheet(sheet, headers);
    rows = [headers];
  }
  const hdrs = rows[0];
  const keyIdx = hdrs.indexOf(keyField);
  const keyToRow = {};
  for (let i = 1; i < rows.length; i++) keyToRow[rows[i][keyIdx]] = i + 1;
  const toAppend = [];
  items.forEach(item => {
    const rowIndex = keyToRow[item[keyField]];
    const rowValues = headers.map(h => item[h] !== undefined ? item[h] : '');
    if (rowIndex) {
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      toAppend.push(rowValues);
    }
  });
  if (toAppend.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toAppend.length, headers.length).setValues(toAppend);
  }
  return { updated: items.length - toAppend.length, created: toAppend.length };
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

// ── EXCEPCIONES (saltear una sola semana de un turno fijo) ──
function saveExcepcion(data) {
  const sheet = getSheet('excepciones');
  const headers = ['key','fecha','slot','cancha','dow','ts'];
  const rows = sheet.getDataRange().getValues();
  if (rows.length === 1 && rows[0].length === 1 && rows[0][0] === '') {
    initSheet(sheet, headers);
  } else {
    const hdrs = rows[0];
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][hdrs.indexOf('key')] === data.key) {
        return { ok: true, action: 'exists' };
      }
    }
  }
  appendRow(sheet, headers, data);
  return { ok: true, action: 'created' };
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

// ── REPARACIÓN ÚNICA: encabezado roto de la hoja "fijos" ──
// La fila 1 de la hoja "fijos" quedó con encabezados incorrectos
// ("fi" en la primera columna y el resto vacíos), por lo que la app
// nunca pudo leer bien los turnos fijos guardados. Ejecutar esta
// función UNA VEZ desde el editor (seleccionarla y tocar "Ejecutar")
// para corregir la fila 1. No borra ni toca los datos de las demás filas.
function fixFijosHeader() {
  const sheet = getSheet('fijos');
  sheet.getRange(1, 1, 1, 8).setValues([['key','dow','slot','cancha','nombre','tipo','jugadorId','ts']]);
  return { ok: true };
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
