/**
 * API de subsanaciones de lotes (Google Apps Script).
 * La "base de datos" es la Hoja de cálculo de Google donde se pega este script
 * (Extensiones > Apps Script). Ver LEEME_SUBSANACIONES.txt para instalarlo.
 *
 * Hojas que usa (se crean solas la primera vez):
 *   - Subsanaciones: una visita por fila, con las mismas columnas del Excel de interventoría
 *     (cada ítem SI / NO / N/A) y al final las columnas de seguimiento de la subsanación.
 *   - Items: ítems a revisar. Columna A = nombre, columna B = Activo (SI/NO).
 *   - Rutas: pre-rutas (lotes que mandan a revisar), una fila por lote. Se pueden escribir a mano:
 *     con CUADRILLA, LOTE y FECHA DE EJECUCIÓN basta; la página las ordena por cercanía al abrirse.
 *
 * Se pueden pegar registros del Excel directamente desde la fila 2 de "Subsanaciones":
 * al abrir la página se les asigna ID y ESTADO.
 */

// Opcional: si pone una clave aquí, la página la pedirá antes de leer o guardar.
var CLAVE = '';

// Opcional: ID de una hoja de Drive distinta a la que contiene el script
// (lo que va entre /d/ y /edit en el link). Vacío = la hoja donde está pegado el script.
var ID_HOJA = '';

var HOJA_DATOS = 'Subsanaciones';
var HOJA_ITEMS = 'Items';
var COLUMNAS_INICIO = ['MICRORRUTA', 'LOTE', 'ÁREA M2', 'ZONA', 'CUADRILLA', 'FECHA DE EJECUCIÓN', 'FECHA DE LA VISITA', 'INTERVENTOR'];
var ITEMS_INICIALES = ['LIMPIEZA PREVIA', 'AFECTACIÓN A ÁRBOLES', 'LIMPIEZA MANUAL ÁRBOLES', 'PLATEO',
  'CORTE DE CESPED: DISPAREJO/ALTURA', 'BORDEO', 'RECOLECCIÓN Y DISPOSICIÓN DE RESIDUOS',
  'BARRIDO: BARRIDO DEL LOTE Y LIMPIEZA DE ZONAS DURAS', 'SECTORES SIN INTERVENIR'];
var COLUMNAS_FIN = ['OBSERVACIONES NOVEDADES', 'ESTADO', 'FECHA SUBSANACIÓN', 'OBSERVACIÓN SUBSANACIÓN', 'EVIDENCIA',
  'ID', 'REGISTRADO POR', 'ÚLTIMA MODIFICACIÓN'];
var NO_EDITABLES = ['ID', 'ESTADO', 'ÚLTIMA MODIFICACIÓN'];
var HOJA_RUTAS = 'Rutas';
var COLUMNAS_RUTAS = ['CUADRILLA', 'LOTE', 'FECHA DE EJECUCIÓN', 'ORDEN', 'CARGADA'];

function doGet(e) {
  return responder(function () {
    verificarClave(e.parameter.clave);
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      completarPegados();
      completarRutas();
    } finally {
      lock.releaseLock();
    }
    var hoja = hojaDatos();
    return { items: leerItems(), columnas: encabezados(hoja), registros: listarRegistros(hoja), rutas: listarRutas() };
  });
}

function doPost(e) {
  return responder(function () {
    var cuerpo = JSON.parse(e.postData.contents || '{}');
    verificarClave(cuerpo.clave);
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (cuerpo.accion === 'crear') return { registro: guardarRegistro(null, cuerpo.registro || {}) };
      if (cuerpo.accion === 'actualizar') return { registro: guardarRegistro(cuerpo.id, cuerpo.registro || {}) };
      if (cuerpo.accion === 'eliminar') { eliminarRegistro(cuerpo.id); return { eliminado: cuerpo.id }; }
      if (cuerpo.accion === 'guardar_ruta') { guardarRuta(cuerpo.cuadrilla, cuerpo.paradas || []); return { rutas: listarRutas() }; }
      throw new Error('Acción no válida: ' + cuerpo.accion);
    } finally {
      lock.releaseLock();
    }
  });
}

function responder(fn) {
  var salida;
  try {
    salida = fn();
    salida.ok = true;
  } catch (err) {
    salida = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}

function verificarClave(clave) {
  if (CLAVE && String(clave || '') !== CLAVE) throw new Error('CLAVE_INVALIDA');
}

// Compara encabezados sin importar mayúsculas, tildes ni espacios sobrantes ("ZONA " = "Zona")
function norm(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function libro() {
  return ID_HOJA ? SpreadsheetApp.openById(ID_HOJA) : SpreadsheetApp.getActiveSpreadsheet();
}

function pintarEncabezado(rango) {
  rango.setFontWeight('bold').setBackground('#167a55').setFontColor('#ffffff').setWrap(true);
}

function hojaDatos() {
  var hoja = libro().getSheetByName(HOJA_DATOS);
  if (!hoja) {
    hoja = libro().insertSheet(HOJA_DATOS);
    var cab = COLUMNAS_INICIO.concat(leerItems(), COLUMNAS_FIN);
    pintarEncabezado(hoja.getRange(1, 1, 1, cab.length).setValues([cab]));
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function leerItems() {
  var hoja = libro().getSheetByName(HOJA_ITEMS);
  if (!hoja) {
    hoja = libro().insertSheet(HOJA_ITEMS);
    var filas = [['ÍTEM', 'ACTIVO']].concat(ITEMS_INICIALES.map(function (i) { return [i, 'SI']; }));
    hoja.getRange(1, 1, filas.length, 2).setValues(filas);
    hoja.getRange(1, 1, 1, 2).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja.getDataRange().getValues().slice(1).filter(function (f) {
    return String(f[0]).trim() !== '' && norm(f[1]) !== 'NO';
  }).map(function (f) { return String(f[0]).replace(/\s+/g, ' ').trim(); });
}

function encabezados(hoja) {
  var ultima = hoja.getLastColumn();
  return ultima ? hoja.getRange(1, 1, 1, ultima).getDisplayValues()[0].map(function (c) { return c.replace(/\s+/g, ' ').trim(); }) : [];
}

function columna(cab, nombre) {
  var n = norm(nombre);
  for (var i = 0; i < cab.length; i++) if (norm(cab[i]) === n) return i;
  return -1;
}

// Agrega al final las columnas que falten (ítem nuevo o columnas de seguimiento)
function asegurarColumnas(hoja, cab, nombres) {
  nombres.forEach(function (k) {
    if (columna(cab, k) === -1) {
      cab.push(k);
      pintarEncabezado(hoja.getRange(1, cab.length).setValue(k));
    }
  });
  return cab;
}

function filaARegistro(cab, f) {
  var r = {};
  cab.forEach(function (c, i) { if (c) r[c] = f[i]; });
  return r;
}

function listarRegistros(hoja) {
  var cab = encabezados(hoja), iId = columna(cab, 'ID');
  return hoja.getDataRange().getDisplayValues().slice(1)
    .filter(function (f) { return iId !== -1 && f[iId] !== ''; })
    .map(function (f) { return filaARegistro(cab, f); });
}

// SI = cumple. Algún NO = pendiente de subsanar. Todo SI después de un NO = subsanado.
function calcularEstado(valoresItems, estadoAnterior) {
  var hayNo = valoresItems.some(function (v) { return norm(v) === 'NO'; });
  if (hayNo) return 'PENDIENTE';
  var anterior = norm(estadoAnterior);
  return (anterior === 'PENDIENTE' || anterior === 'SUBSANADO') ? 'SUBSANADO' : 'SIN NOVEDAD';
}

function estadoDeFila(cab, fila, items, estadoAnterior) {
  return calcularEstado(items.map(function (it) {
    var i = columna(cab, it);
    return i === -1 ? '' : fila[i];
  }), estadoAnterior);
}

function siguienteNumero(hoja, cab) {
  var iId = columna(cab, 'ID'), max = 0;
  if (iId === -1 || hoja.getLastRow() < 2) return 1;
  hoja.getRange(2, iId + 1, hoja.getLastRow() - 1, 1).getDisplayValues().forEach(function (f) {
    var m = String(f[0]).match(/^SUB-(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return max + 1;
}

function formatoId(n) { return 'SUB-' + ('0000' + n).slice(-5); }

function ahora() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

// Filas pegadas a mano desde el Excel: les pone ID y ESTADO
function completarPegados() {
  var hoja = hojaDatos();
  var items = leerItems();
  var cab = asegurarColumnas(hoja, encabezados(hoja), COLUMNAS_FIN);
  if (hoja.getLastRow() < 2) return;
  var iId = columna(cab, 'ID'), iEstado = columna(cab, 'ESTADO');
  var iMr = columna(cab, 'MICRORRUTA'), iLote = columna(cab, 'LOTE');
  var valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, cab.length).getDisplayValues();
  var n = siguienteNumero(hoja, cab);
  valores.forEach(function (f, k) {
    if (f[iId] !== '' || (f[iMr] === '' && f[iLote] === '')) return;
    hoja.getRange(k + 2, iId + 1).setNumberFormat('@').setValue(formatoId(n++));
    if (f[iEstado] === '') hoja.getRange(k + 2, iEstado + 1).setValue(estadoDeFila(cab, f, items, ''));
  });
}

function buscarFila(hoja, cab, id) {
  var iId = columna(cab, 'ID');
  if (iId === -1 || hoja.getLastRow() < 2) return 0;
  var ids = hoja.getRange(2, iId + 1, hoja.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === String(id)) return i + 2;
  return 0;
}

function guardarRegistro(id, datos) {
  var hoja = hojaDatos();
  var items = leerItems();
  var cab = asegurarColumnas(hoja, encabezados(hoja), COLUMNAS_FIN.concat(Object.keys(datos)));
  var fila, actual;
  if (id) {
    fila = buscarFila(hoja, cab, id);
    if (!fila) throw new Error('No existe el registro ' + id);
    actual = hoja.getRange(fila, 1, 1, cab.length).getDisplayValues()[0];
  } else {
    fila = hoja.getLastRow() + 1;
    actual = cab.map(function () { return ''; });
    actual[columna(cab, 'ID')] = formatoId(siguienteNumero(hoja, cab));
  }
  Object.keys(datos).forEach(function (k) {
    if (NO_EDITABLES.some(function (c) { return norm(c) === norm(k); })) return;
    actual[columna(cab, k)] = datos[k] == null ? '' : String(datos[k]);
  });
  var iEstado = columna(cab, 'ESTADO');
  actual[iEstado] = estadoDeFila(cab, actual, items, actual[iEstado]);
  var iFechaSub = columna(cab, 'FECHA SUBSANACIÓN');
  if (actual[iEstado] === 'SUBSANADO' && !actual[iFechaSub]) actual[iFechaSub] = ahora().slice(0, 10);
  actual[columna(cab, 'ÚLTIMA MODIFICACIÓN')] = ahora();
  // Texto plano para que Sheets no cambie lotes, áreas, fechas ni códigos
  hoja.getRange(fila, 1, 1, cab.length).setNumberFormat('@').setValues([actual]);
  return filaARegistro(cab, actual);
}

function eliminarRegistro(id) {
  var hoja = hojaDatos();
  var fila = buscarFila(hoja, encabezados(hoja), id);
  if (!fila) throw new Error('No existe el registro ' + id);
  hoja.deleteRow(fila);
}

// ---------- Pre-rutas ----------
function hojaRutas() {
  var hoja = libro().getSheetByName(HOJA_RUTAS);
  if (!hoja) {
    hoja = libro().insertSheet(HOJA_RUTAS);
    pintarEncabezado(hoja.getRange(1, 1, 1, COLUMNAS_RUTAS.length).setValues([COLUMNAS_RUTAS]));
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, hoja.getMaxRows(), COLUMNAS_RUTAS.length).setNumberFormat('@');
  }
  asegurarColumnas(hoja, encabezados(hoja), COLUMNAS_RUTAS);
  return hoja;
}

function cuadrillaRuta(v) { return String(v == null ? '' : v).trim().replace(/^C\s*/i, ''); }

function filasRutas(hoja, cab) {
  return hoja.getLastRow() < 2 ? [] : hoja.getRange(2, 1, hoja.getLastRow() - 1, cab.length).getDisplayValues();
}

// Lotes escritos a mano: les pone la fecha de carga (sirve para saber si ya se revisaron)
function completarRutas() {
  var hoja = hojaRutas(), cab = encabezados(hoja);
  var iLote = columna(cab, 'LOTE'), iCargada = columna(cab, 'CARGADA'), hoy = ahora().slice(0, 10);
  filasRutas(hoja, cab).forEach(function (f, k) {
    if (String(f[iLote]).trim() !== '' && f[iCargada] === '') hoja.getRange(k + 2, iCargada + 1).setNumberFormat('@').setValue(hoy);
  });
}

function listarRutas() {
  var hoja = hojaRutas(), cab = encabezados(hoja);
  var i = {}; COLUMNAS_RUTAS.forEach(function (c) { i[c] = columna(cab, c); });
  return filasRutas(hoja, cab).filter(function (f) { return String(f[i.LOTE]).trim() !== ''; }).map(function (f) {
    return { cuadrilla: cuadrillaRuta(f[i.CUADRILLA]), lote: String(f[i.LOTE]).trim().toUpperCase(),
             fecha: f[i['FECHA DE EJECUCIÓN']], orden: f[i.ORDEN], cargada: f[i.CARGADA] };
  });
}

// Reemplaza todos los lotes de una cuadrilla por la lista recibida (vacía = quitar la cuadrilla)
function guardarRuta(cuadrilla, paradas) {
  var hoja = hojaRutas(), cab = encabezados(hoja), cuad = cuadrillaRuta(cuadrilla);
  var iCuad = columna(cab, 'CUADRILLA');
  var actuales = filasRutas(hoja, cab);
  var quedan = actuales.filter(function (f) { return cuadrillaRuta(f[iCuad]) !== cuad; });
  var nuevas = paradas.map(function (p, n) {
    var f = cab.map(function () { return ''; });
    f[iCuad] = cuad;
    f[columna(cab, 'LOTE')] = String(p.lote || '');
    f[columna(cab, 'FECHA DE EJECUCIÓN')] = String(p.fecha || '');
    f[columna(cab, 'ORDEN')] = String(n + 1);
    f[columna(cab, 'CARGADA')] = String(p.cargada || ahora().slice(0, 10));
    return f;
  });
  var todas = quedan.concat(nuevas);
  if (actuales.length) hoja.getRange(2, 1, actuales.length, cab.length).clearContent();
  if (todas.length) hoja.getRange(2, 1, todas.length, cab.length).setNumberFormat('@').setValues(todas);
}

// Ejecute esta función una vez desde el editor para crear las hojas y autorizar el script.
function configurar() {
  hojaDatos();
  hojaRutas();
  Logger.log('Listo. Ítems activos: ' + leerItems().join(', '));
}
