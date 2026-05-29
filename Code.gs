// ═══════════════════════════════════════════════════════════════
// FINANZAS PORITOS — Google Apps Script Backend v3
// Pegar en: Extensiones > Apps Script > Código.gs
// Luego: Implementar > Administrar implementaciones > Nueva versión
// ═══════════════════════════════════════════════════════════════

const SHEET_TX         = 'Transacciones';
const SHEET_META       = 'Meta Vivienda';
const SHEET_CAT        = 'Categorias';
const SHEET_ACTIVOS    = 'Activos';
const SHEET_PASIVOS    = 'Pasivos';

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    if      (action === 'getTx')          result = getTx(e.parameter);
    else if (action === 'getMeta')         result = getMeta();
    else if (action === 'getSummary')      result = getSummary(e.parameter);
    else if (action === 'getPatrimonio')   result = getPatrimonio();
    else result = { error: 'Accion desconocida' };
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  let result;
  try {
    if      (action === 'addTx')           result = addTx(body.data);
    else if (action === 'deleteTx')        result = deleteTx(body.id);
    else if (action === 'updateMeta')      result = updateMeta(body.data);
    else if (action === 'savePatrimonio')  result = savePatrimonio(body.data);
    else result = { error: 'Accion desconocida' };
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Inicializar hojas ─────────────────────────────────────────
function inicializarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Transacciones
  let shTx = ss.getSheetByName(SHEET_TX);
  if (!shTx) {
    shTx = ss.insertSheet(SHEET_TX);
    shTx.getRange(1,1,1,9).setValues([[
      'ID','Fecha','Tipo','Categoria','Descripcion','Monto','Quien','Recurrente','Timestamp'
    ]]);
    shTx.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    shTx.setFrozenRows(1);
    shTx.setColumnWidths(1, 9, [140,110,80,160,220,110,80,100,180]);
  }

  // Meta Vivienda
  let shMeta = ss.getSheetByName(SHEET_META);
  if (!shMeta) {
    shMeta = ss.insertSheet(SHEET_META);
    shMeta.getRange(1,1,1,2).setValues([['Parametro','Valor']]);
    shMeta.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    shMeta.getRange(2,1,5,2).setValues([
      ['precio_objetivo', 120000000],
      ['pie_pct', 20],
      ['ahorro_acumulado', 1890000],
      ['ingreso_mensual', 5130000],
      ['notas', 'Meta: tercera vivienda']
    ]);
  }

  // Activos
  let shAct = ss.getSheetByName(SHEET_ACTIVOS);
  if (!shAct) {
    shAct = ss.insertSheet(SHEET_ACTIVOS);
    shAct.getRange(1,1,1,10).setValues([[
      'ID','Tipo','Nombre','Valor_Mercado','Descripcion','Anio_Compra','Arriendo_Mensual','Institucion','Aporte_Mensual','Updated'
    ]]);
    shAct.getRange(1,1,1,10).setFontWeight('bold').setBackground('#0F6E56').setFontColor('#FFFFFF');
    shAct.setFrozenRows(1);
    [120,100,180,130,200,90,120,130,110,180].forEach((w,i) => shAct.setColumnWidth(i+1, w));
  }

  // Pasivos
  let shPas = ss.getSheetByName(SHEET_PASIVOS);
  if (!shPas) {
    shPas = ss.insertSheet(SHEET_PASIVOS);
    shPas.getRange(1,1,1,12).setValues([[
      'ID','Tipo','Nombre','Banco','Saldo_Pendiente','Cuota_Mensual','Tasa_Anual','Cuotas_Pagadas','Cuotas_Totales','Fecha_Inicio','Activo_ID','Updated'
    ]]);
    shPas.getRange(1,1,1,12).setFontWeight('bold').setBackground('#7F1D1D').setFontColor('#FFFFFF');
    shPas.setFrozenRows(1);
    [120,100,180,120,130,110,90,100,100,110,120,180].forEach((w,i) => shPas.setColumnWidth(i+1, w));
  }

  return { ok: true, message: 'Hojas creadas correctamente' };
}

// ── Convierte fecha a string AAAA-MM-DD ──────────────────────
function fechaToString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(val).trim();
}

// ── GET Transacciones ─────────────────────────────────────────
function getTx(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_TX);
  if (!sh) return { rows: [] };
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return { rows: [] };
  const headers = data[0];
  let rows = data.slice(1).map((row, i) => {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = row[j]; });
    obj['Fecha'] = fechaToString(obj['Fecha']);
    obj['Monto'] = parseInt(obj['Monto'], 10) || 0;
    obj._row = i + 2;
    return obj;
  }).filter(r => r.ID && r.ID !== '');
  if (params.year && params.month) {
    const prefix = `${params.year}-${String(params.month).padStart(2, '0')}`;
    rows = rows.filter(r => String(r.Fecha).startsWith(prefix));
  }
  return { rows };
}

// ── POST Agregar transacción ──────────────────────────────────
function addTx(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_TX);
  if (!sh) { inicializarHojas(); sh = ss.getSheetByName(SHEET_TX); }
  const id = 'TX' + new Date().getTime();
  const monto = parseInt(data.monto, 10) || 0;
  sh.appendRow([id, data.fecha, data.tipo, data.categoria, data.descripcion, monto,
    data.quien, data.recurrente ? 'Sí' : 'No', new Date().toISOString()]);
  sh.getRange(sh.getLastRow(), 6).setNumberFormat('0');
  return { ok: true, id };
}

// ── POST Eliminar transacción ─────────────────────────────────
function deleteTx(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_TX);
  if (!sh) return { ok: false };
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, error: 'No encontrado' };
}

// ── GET Meta vivienda ─────────────────────────────────────────
function getMeta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_META);
  if (!sh) return {};
  const data = sh.getDataRange().getValues().slice(1);
  const meta = {};
  data.forEach(row => { if (row[0]) meta[row[0]] = row[1]; });
  return meta;
}

// ── POST Actualizar meta ──────────────────────────────────────
function updateMeta(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_META);
  if (!sh) return { ok: false };
  const rows = sh.getDataRange().getValues();
  Object.entries(data).forEach(([key, val]) => {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) { sh.getRange(i + 1, 2).setValue(val); return; }
    }
    sh.appendRow([key, val]);
  });
  return { ok: true };
}

// ── GET Patrimonio (activos + pasivos) ────────────────────────
function getPatrimonio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Activos
  const shAct = ss.getSheetByName(SHEET_ACTIVOS);
  let activos = [];
  if (shAct && shAct.getLastRow() > 1) {
    const data = shAct.getDataRange().getValues();
    const h = data[0];
    activos = data.slice(1).filter(r => r[0]).map(row => {
      const o = {}; h.forEach((k,i) => { o[k.toLowerCase()] = row[i]; }); return o;
    }).map(r => ({
      id: r['id'], tipo: r['tipo'], nombre: r['nombre'],
      valor_mercado: Number(r['valor_mercado']) || 0,
      descripcion: r['descripcion'] || '',
      anio_compra: r['anio_compra'] || '',
      arriendo: Number(r['arriendo_mensual']) || 0,
      institucion: r['institucion'] || '',
      aporte: Number(r['aporte_mensual']) || 0,
      updated: r['updated'] || ''
    }));
  }

  // Pasivos
  const shPas = ss.getSheetByName(SHEET_PASIVOS);
  let pasivos = [];
  if (shPas && shPas.getLastRow() > 1) {
    const data = shPas.getDataRange().getValues();
    const h = data[0];
    pasivos = data.slice(1).filter(r => r[0]).map(row => {
      const o = {}; h.forEach((k,i) => { o[k.toLowerCase()] = row[i]; }); return o;
    }).map(r => ({
      id: r['id'], tipo: r['tipo'], nombre: r['nombre'],
      banco: r['banco'] || '',
      saldo_pendiente: Number(r['saldo_pendiente']) || 0,
      cuota_mensual: Number(r['cuota_mensual']) || 0,
      tasa_anual: Number(r['tasa_anual']) || 0,
      cuotas_pagadas: Number(r['cuotas_pagadas']) || 0,
      cuotas_totales: Number(r['cuotas_totales']) || 0,
      fecha_inicio: fechaToString(r['fecha_inicio']),
      activo_id: r['activo_id'] || '',
      updated: r['updated'] || ''
    }));
  }

  return { activos, pasivos };
}

// ── POST Guardar patrimonio completo ──────────────────────────
function savePatrimonio(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Activos: rewrite completo
  let shAct = ss.getSheetByName(SHEET_ACTIVOS);
  if (!shAct) { inicializarHojas(); shAct = ss.getSheetByName(SHEET_ACTIVOS); }
  // Limpiar datos (mantener cabecera)
  if (shAct.getLastRow() > 1) shAct.getRange(2, 1, shAct.getLastRow()-1, 10).clearContent();
  (data.activos || []).forEach((a, i) => {
    shAct.getRange(i+2, 1, 1, 10).setValues([[
      a.id, a.tipo, a.nombre, Number(a.valor_mercado)||0,
      a.descripcion||'', a.anio_compra||'', Number(a.arriendo)||0,
      a.institucion||'', Number(a.aporte)||0, a.updated||new Date().toISOString()
    ]]);
    shAct.getRange(i+2, 4).setNumberFormat('0');
  });

  // Pasivos: rewrite completo
  let shPas = ss.getSheetByName(SHEET_PASIVOS);
  if (!shPas) { inicializarHojas(); shPas = ss.getSheetByName(SHEET_PASIVOS); }
  if (shPas.getLastRow() > 1) shPas.getRange(2, 1, shPas.getLastRow()-1, 12).clearContent();
  (data.pasivos || []).forEach((p, i) => {
    shPas.getRange(i+2, 1, 1, 12).setValues([[
      p.id, p.tipo, p.nombre, p.banco||'',
      Number(p.saldo_pendiente)||0, Number(p.cuota_mensual)||0,
      Number(p.tasa_anual)||0, Number(p.cuotas_pagadas)||0,
      Number(p.cuotas_totales)||0, p.fecha_inicio||'',
      p.activo_id||'', p.updated||new Date().toISOString()
    ]]);
    shPas.getRange(i+2, 5).setNumberFormat('0');
    shPas.getRange(i+2, 6).setNumberFormat('0');
  });

  return { ok: true };
}

// ── GET Resumen mensual ───────────────────────────────────────
function getSummary(params) {
  const txData = getTx(params);
  const rows   = txData.rows;
  const ingresos = rows.filter(r => r.Tipo === 'ingreso').reduce((s, r) => s + Number(r.Monto), 0);
  const gastos   = rows.filter(r => r.Tipo === 'gasto').reduce((s, r)  => s + Number(r.Monto), 0);
  const ahorros  = rows.filter(r => r.Tipo === 'ahorro').reduce((s, r)  => s + Number(r.Monto), 0);
  const catMap   = {};
  rows.filter(r => r.Tipo === 'gasto').forEach(r => { catMap[r.Categoria] = (catMap[r.Categoria] || 0) + Number(r.Monto); });
  return { ingresos, gastos, ahorros, libre: ingresos - gastos - ahorros, catMap, total: rows.length };
}
