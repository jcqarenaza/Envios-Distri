/**
 * ════════════════════════════════════════════════════════
 *  DistriLogis V5 — Google Apps Script Backend
 *  Versión 5.0 · Abril 2026
 * ════════════════════════════════════════════════════════
 *
 *  CAMBIOS V5 respecto a V4:
 *  ─────────────────────────
 *  1. Nueva acción pública: 'authenticate' — el cliente ya NO
 *     tiene credenciales hardcodeadas. El login pasa usuario/pass
 *     y recibe rol + depot + label. Sin esta acción el login falla.
 *
 *  2. Nueva acción autenticada: 'revertirRecepcion' — antes el
 *     admin podía revertir localmente sin que el cambio llegara al Sheet.
 *     Ahora sincroniza correctamente.
 *
 *  3. obsRecv ya no fuerza 'OK' si está vacío — el cliente envía
 *     string vacío y se guarda tal cual.
 *
 *  INSTRUCCIONES DE DEPLOY:
 *  ─────────────────────────
 *  1. Extensiones → Apps Script → reemplazá todo con este archivo
 *  2. Guardá (Ctrl+S)
 *  3. Implementar → Administrar implementaciones → ✏️ editar
 *     → Versión: Nueva versión → Implementar
 *  4. Copiá la nueva URL y pegala en GAS_URL del HTML
 *
 * ════════════════════════════════════════════════════════
 *  ESTRUCTURA DE COLUMNAS — hoja "envios"
 *  A=id  B=origen  C=destino  D=fecha  E=remito
 *  F=tipo  G=modalidad  H=cantidad  I=transporte  J=obs
 *  K=estado  L=fechaLlegada  M=obsRecv
 *  N=creadoPor  O=creadoEn  P=modificadoPor  Q=modificadoEn
 * ════════════════════════════════════════════════════════
 */

// ── CONSTANTES ──────────────────────────────────────────
const SS        = SpreadsheetApp.getActiveSpreadsheet();
const SH_ENVIOS = 'envios';
const SH_ROUTES = 'rutas';
const SH_LOG    = 'log';
const SH_TRANS  = 'transportes';

// ════════════════════════════════════════════════════════
//  CONTROL DE VERSIÓN — administrado desde el Sheet
//
//  La versión mínima requerida NO está hardcodeada acá.
//  Se lee de la hoja "config", fila con clave "minVersion".
//
//  Para bloquear versiones viejas:
//  1. Abrí el Sheet → hoja "config"
//  2. Cambiá el valor de la fila "minVersion" al número nuevo
//  3. Listo — ningún deploy de GAS necesario.
//
//  El HTML se auto-registra en "config" la primera vez que
//  alguien se loguea (clave "lastSeenVersion"), para que el
//  admin sepa qué versiones hay en circulación.
// ════════════════════════════════════════════════════════

const SH_CONFIG = 'config';

const COL = {
  id:            1,   // A
  origen:        2,   // B
  destino:       3,   // C
  fecha:         4,   // D
  remito:        5,   // E
  tipo:          6,   // F
  modalidad:     7,   // G
  cantidad:      8,   // H
  transporte:    9,   // I
  obs:           10,  // J
  estado:        11,  // K
  fechaLlegada:  12,  // L
  obsRecv:       13,  // M
  creadoPor:     14,  // N
  creadoEn:      15,  // O
  modificadoPor: 16,  // P
  modificadoEn:  17,  // Q
};
const NCOLS = 17;

const COLS_HEADER = [
  'id','origen','destino','fecha','remito',
  'tipo','modalidad','cantidad','transporte','obs',
  'estado','fechaLlegada','obsRecv','creadoPor','creadoEn',
  'modificadoPor','modificadoEn'
];


// ════════════════════════════════════════════════════════
//  CONTROL DE VERSIÓN
// ════════════════════════════════════════════════════════

function getVersionInfo(clientVersion) {
  const minVersion = getConfig('minVersion') || '1.0';

  // Registrar la versión del cliente que se conectó (para auditoría)
  if(clientVersion){
    setConfig('lastSeenVersion', clientVersion);
    const hist = getConfig('versionHistory') || '';
    if(!hist.split(',').map(s=>s.trim()).includes(clientVersion)){
      setConfig('versionHistory', hist ? hist + ', ' + clientVersion : clientVersion);
    }
  }

  return { minVersion };
}

// ── Normalización de versiones ───────────────────────────────────────────
// Google Sheets puede guardar "20260404.1132" como número y devolverlo
// como "202604041132" (sin punto). Esta función lo restaura.
function normalizeVersion(v) {
  const s = String(v || '').trim().replace(/^'+/, ''); // quitar apóstrofes iniciales
  if(s.length === 12 && !s.includes('.')) {
    return s.slice(0, 8) + '.' + s.slice(8);
  }
  return s || '0';
}

function compareVersions(a, b) {
  return parseFloat(normalizeVersion(a)) - parseFloat(normalizeVersion(b));
}

// ── registerVersion: el HTML se auto-registra al login ──────────────────
// Solo actualiza si el número es MAYOR al deployedVersion actual.
function registerVersion(version, user) {
  if(!version) return { ok: false };
  const incoming = normalizeVersion(version);
  const current  = normalizeVersion(getConfig('deployedVersion') || '0');
  if(compareVersions(incoming, current) > 0){
    setConfig('deployedVersion', incoming);
    addLog(user.usuario, 'VERSION_DEPLOY', 'config', 'deployedVersion', incoming);
    Logger.log('[registerVersion] actualizado a v' + incoming + ' por ' + user.usuario);
    // Registrar en historial
    const hist = getConfig('versionHistory') || '';
    const versions = hist ? hist.split(',').map(s=>s.trim()) : [];
    if(!versions.includes(incoming)){
      setConfig('versionHistory', [...versions, incoming].join(', '));
    }
    return { updated: true, version: incoming };
  }
  return { updated: false, version: normalizeVersion(getConfig('deployedVersion')) };
}

// ── getSystemInfo: devuelve toda la info de versiones para el panel admin ─
function getSystemInfo() {
  return {
    minVersion:       normalizeVersion(getConfig('minVersion')      || '1.0'),
    deployedVersion:  normalizeVersion(getConfig('deployedVersion') || ''),
    lastSeenVersion:  normalizeVersion(getConfig('lastSeenVersion') || ''),
    versionHistory:   getConfig('versionHistory') || '',
  };
}

// ── setMinVersion: el admin cambia la versión mínima requerida desde la UI ──
function setMinVersion(minVersion, user) {
  if(!minVersion) throw new Error('Falta minVersion');
  const normalized = normalizeVersion(minVersion);
  if(normalized === '0') throw new Error('Versión inválida: ' + minVersion);
  setConfig('minVersion', normalized);
  addLog(user.usuario, 'SET_MIN_VERSION', 'config', 'minVersion', normalized);
  Logger.log('[setMinVersion] minVersion=' + normalized + ' por ' + user.usuario);
  return { minVersion: normalized };
}

// ── Config helpers — hoja "config" con columnas: clave | valor | ultimaModif ──

function getConfig(key) {
  const sh = _getOrCreateConfig();
  const data = sh.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0]).trim() === key) return String(data[i][1] || '').trim();
  }
  return null;
}

function setConfig(key, value) {
  const sh   = _getOrCreateConfig();
  const data = sh.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0]).trim() === key){
      sh.getRange(i + 1, 2).setValue(value);
      sh.getRange(i + 1, 3).setValue(new Date());
      return;
    }
  }
  // No existe la clave — agregar fila nueva
  sh.appendRow([key, value, new Date()]);
}

function _getOrCreateConfig() {
  let sh = SS.getSheetByName(SH_CONFIG);
  if(!sh){
    sh = SS.insertSheet(SH_CONFIG);
    sh.appendRow(['clave', 'valor', 'ultimaModif']);
    sh.getRange(1, 1, 1, 3).setBackground('#1E3A5F').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 200);
    sh.setColumnWidth(3, 160);
    // Valor por defecto: ninguna versión bloqueada
    sh.appendRow(['minVersion',       '1.0', new Date()]);
    sh.appendRow(['deployedVersion',  '',    new Date()]);
    sh.appendRow(['lastSeenVersion',  '',    new Date()]);
    sh.appendRow(['versionHistory',   '',    new Date()]);
    Logger.log('✅ Hoja config creada con minVersion=1.0');
  }
  return sh;
}


// ════════════════════════════════════════════════════════
//  AUTENTICACIÓN
// ════════════════════════════════════════════════════════

// ⚠️  ÚNICO lugar donde existen las credenciales.
//     El HTML del cliente NO las tiene — las valida aquí.
const USERS = {
  administrador: { pass: 'tinchohack', rol: 'admin',    depo: '',        label: 'Administrador' },
  pico:          { pass: 'pico',       rol: 'sucursal', depo: 'Pico',    label: 'Pico'          },
  mdp:           { pass: 'mdp',        rol: 'sucursal', depo: 'MDP',     label: 'MDP'           },
  bsas:          { pass: 'bsas',       rol: 'sucursal', depo: 'Bs As',   label: 'Bs As'         },
  rosario:       { pass: 'rosario',    rol: 'sucursal', depo: 'Rosario', label: 'Rosario'       },
  telev:         { pass: 'telev',      rol: 'readonly', depo: '',        label: 'Telev'         },
};

function authenticate(usuario, password) {
  if (!usuario || !password) return { ok: false, error: 'Faltan credenciales' };
  const u = USERS[String(usuario).trim()];
  if (!u) return { ok: false, error: 'Usuario no existe' };
  if (u.pass !== String(password).trim()) return { ok: false, error: 'Contraseña incorrecta' };
  // ✅ No devolvemos la contraseña — solo lo que el cliente necesita mostrar
  return { ok: true, user: { usuario, rol: u.rol, depo: u.depo, label: u.label } };
}


// ════════════════════════════════════════════════════════
//  WEB APP — doGet
// ════════════════════════════════════════════════════════

function doGet(e) {
  const p      = e.parameter;
  const action = (p.action || '').trim();

  Logger.log('[doGet] action=' + action + ' usuario=' + (p.usuario||'–'));

  try {
    // ── Acciones completamente públicas ─────────────────
    if (action === 'ping') return jsonOk({ status: 'ok', ts: new Date().toISOString() });

    // ── Control de versión — lee del Sheet, no del código ──
    // El admin controla qué versión mínima se requiere editando
    // una celda en la hoja "config". Sin deploy de GAS.
    if (action === 'getVersion'){
      const clientVersion = (p.v || '').trim(); // versión que reporta el cliente
      return jsonOk(getVersionInfo(clientVersion));
    }

    // ── Acción pública pero con credenciales (V5 NEW) ───
    // El cliente la llama durante el login para validar usuario/pass
    // y obtener rol + depot. No requiere pre-autenticación porque
    // ES la autenticación. Retorna datos mínimos (sin contraseña).
    if (action === 'authenticate') {
      const usuario  = (p.usuario  || '').trim();
      const password = (p.password || '').trim();
      const result   = authenticate(usuario, password);
      if (!result.ok) return jsonErr(result.error);
      return jsonOk(result.user);
    }

    // ── Acciones de solo lectura — no necesitan auth para simplificar ──
    // (la GAS URL no es pública, solo quien tiene el enlace puede usarla)
    if (action === 'getEnvios')      return jsonOk(getEnvios());
    if (action === 'getRutas')       return jsonOk(getRutas());
    if (action === 'getTransportes') return jsonOk(getTransportes());

  } catch(err) {
    Logger.log('[doGet] error en accion publica: ' + err.message);
    return jsonErr(err.message);
  }

  // ── Acciones con auth ────────────────────────────────
  const usuario  = (p.usuario  || '').trim();
  const password = (p.password || '').trim();
  let body = {};
  if (p.data) {
    try { body = JSON.parse(p.data); }
    catch(err) { Logger.log('[doGet] error parseando data: ' + err.message); }
  }

  const auth = authenticate(usuario, password);
  if (!auth.ok) {
    Logger.log('[doGet] auth FAIL: ' + auth.error);
    return jsonErr(auth.error || 'Credenciales incorrectas');
  }

  return dispatch(action, body, auth.user);
}


// ════════════════════════════════════════════════════════
//  WEB APP — doPost
// ════════════════════════════════════════════════════════

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch(err) { return jsonErr('JSON inválido: ' + err.message); }

  const action   = (body.action   || '').trim();
  const usuario  = (body.usuario  || '').trim();
  const password = (body.password || '').trim();

  Logger.log('[doPost] action=' + action + ' usuario=' + usuario);

  const auth = authenticate(usuario, password);
  if (!auth.ok) {
    Logger.log('[doPost] auth FAIL: ' + auth.error);
    return jsonErr(auth.error || 'Credenciales incorrectas');
  }

  return dispatch(action, body, auth.user);
}


// ════════════════════════════════════════════════════════
//  DISPATCH
// ════════════════════════════════════════════════════════

function dispatch(action, body, user) {
  try {
    switch (action) {

      case 'altaEnvio':
        if (user.rol === 'readonly') return jsonErr('Sin permiso');
        return jsonOk(altaEnvio(body.envio, user));

      // ✅ Auto-registro de versión al login — solo sube si el número es mayor
      case 'registerVersion':
        return jsonOk(registerVersion(body.version, user));

      // ✅ Info del sistema para el panel admin
      case 'getSystemInfo':
        if (user.rol !== 'admin') return jsonErr('Solo admin');
        return jsonOk(getSystemInfo());

      // ✅ El admin cambia la versión mínima requerida desde la UI
      case 'setMinVersion':
        if (user.rol !== 'admin') return jsonErr('Solo admin puede cambiar la versión mínima');
        return jsonOk(setMinVersion(body.minVersion, user));

      case 'confirmarRecepcion':
        if (user.rol === 'readonly') return jsonErr('Sin permiso');
        return jsonOk(confirmarRecepcion(body, user));

      // ✅ V5 NEW — revertir recepción (admin únicamente)
      case 'revertirRecepcion':
        if (user.rol !== 'admin') return jsonErr('Solo admin puede revertir');
        return jsonOk(revertirRecepcion(body, user));

      case 'eliminarEnvio':
        if (user.rol !== 'admin') return jsonErr('Solo admin puede eliminar');
        return jsonOk(eliminarEnvio(body, user));

      case 'editFechaDespacho':
        if (user.rol !== 'admin') return jsonErr('Solo admin puede editar');
        return jsonOk(editFechaDespacho(body, user));

      case 'updateRutas':
        if (user.rol !== 'admin') return jsonErr('Solo admin puede modificar rutas');
        return jsonOk(updateRutas(body.rutas, user));

      case 'updateTransportes':
        if (user.rol !== 'admin') return jsonErr('Solo admin puede modificar transportes');
        return jsonOk(updateTransportes(body.transportes, user));

      default:
        return jsonErr('Acción no reconocida: ' + action);
    }
  } catch(err) {
    Logger.log('[dispatch] ERROR en ' + action + ': ' + err.message + '\n' + err.stack);
    return jsonErr(err.message);
  }
}


// ════════════════════════════════════════════════════════
//  HELPER — buscar fila por id o remito
// ════════════════════════════════════════════════════════

function findRow(id, remito) {
  const sh = SS.getSheetByName(SH_ENVIOS);
  if (!sh || sh.getLastRow() < 2) return { sh: null, rowIdx: -1 };
  const lastRow = sh.getLastRow();
  const ids     = sh.getRange(2, COL.id,     lastRow - 1, 1).getValues();
  const remitos = sh.getRange(2, COL.remito, lastRow - 1, 1).getValues();

  const idStr = String(id || '').trim();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === idStr) {
      return { sh, rowIdx: i + 2 };
    }
  }

  const remStr = String(remito || '').trim();
  if (remStr) {
    for (let i = 0; i < remitos.length; i++) {
      if (String(remitos[i][0]).trim() === remStr) {
        return { sh, rowIdx: i + 2 };
      }
    }
  }

  Logger.log('[findRow] NO encontrado — id=' + idStr + ' remito=' + remStr);
  return { sh, rowIdx: -1 };
}


// ── fixConfigVersions: corregir valores de versión mal guardados en el Sheet ──
// Ejecutar manualmente UNA VEZ desde la consola de GAS si las celdas tienen
// valores como "202604041132" en lugar de "20260404.1132"
function fixConfigVersions() {
  const keys = ['minVersion', 'deployedVersion', 'lastSeenVersion'];
  keys.forEach(key => {
    const raw = getConfig(key);
    if(!raw) return;
    const fixed = normalizeVersion(raw);
    if(fixed !== raw.trim()){
      setConfig(key, fixed);
      Logger.log('[fixConfigVersions] ' + key + ': ' + raw + ' → ' + fixed);
    }
  });
  Logger.log('[fixConfigVersions] listo');
}


// ════════════════════════════════════════════════════════
//  ENVÍOS — CRUD
// ════════════════════════════════════════════════════════

function getEnvios() {
  const sh = SS.getSheetByName(SH_ENVIOS);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, NCOLS).getValues();
  return data
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => {
      const obj = {};
      COLS_HEADER.forEach((col, i) => {
        obj[col] = (r[i] instanceof Date && r[i].getTime() > 0)
          ? r[i].toISOString()
          : (r[i] === '' || r[i] === null ? null : r[i]);
      });
      return obj;
    });
}

function altaEnvio(envio, user) {
  if (!envio)                         throw new Error('Falta objeto envio');
  if (!envio.origen)                  throw new Error('Falta origen');
  if (!envio.destino)                 throw new Error('Falta destino');
  if (!envio.fecha)                   throw new Error('Falta fecha');
  if (!envio.remito)                  throw new Error('Falta remito');
  if (envio.origen === envio.destino) throw new Error('Origen y Destino iguales');

  const sh = SS.getSheetByName(SH_ENVIOS);
  if (sh && sh.getLastRow() > 1) {
    const remitos = sh.getRange(2, COL.remito, sh.getLastRow() - 1, 1).getValues();
    const remitoStr = String(envio.remito).trim();
    for (let i = 0; i < remitos.length; i++) {
      if (String(remitos[i][0]).trim() === remitoStr) {
        throw new Error('El remito ' + envio.remito + ' ya existe');
      }
    }
  }

  const now = new Date();
  const id  = envio.id || ('env-' + now.getTime());

  sh.appendRow([
    id,
    envio.origen,
    envio.destino,
    new Date(envio.fecha),
    String(envio.remito).trim(),
    envio.tipo       || '',
    envio.modalidad  || '',
    envio.cantidad   || '',
    envio.transporte || '',
    envio.obs        || '',
    'pendiente',
    '',
    '',
    user.usuario,
    now,
    '',
    '',
  ]);

  addLog(user.usuario, 'ALTA_ENVIO', 'envio', id, envio.remito + ' ' + envio.origen + '→' + envio.destino);
  Logger.log('[altaEnvio] OK id=' + id + ' remito=' + envio.remito);
  return { id, mensaje: 'Envío registrado correctamente' };
}

function confirmarRecepcion(body, user) {
  const { id, remito, fechaLlegada, obsRecv } = body;

  Logger.log('[confirmarRecepcion] id=' + id + ' remito=' + remito + ' fechaLlegada=' + fechaLlegada);

  if (!fechaLlegada) throw new Error('Falta fechaLlegada');

  const { sh, rowIdx } = findRow(id, remito);
  if (rowIdx === -1) throw new Error('Envío no encontrado — id: ' + id + ' / remito: ' + remito);

  const now = new Date();
  sh.getRange(rowIdx, COL.estado       ).setValue('recibido');
  sh.getRange(rowIdx, COL.fechaLlegada ).setValue(new Date(fechaLlegada));
  // ✅ V5 FIX: ya no fuerza 'OK' — guarda lo que manda el cliente (puede ser vacío)
  sh.getRange(rowIdx, COL.obsRecv      ).setValue(obsRecv || '');
  sh.getRange(rowIdx, COL.modificadoPor).setValue(user.usuario);
  sh.getRange(rowIdx, COL.modificadoEn ).setValue(now);

  addLog(user.usuario, 'RECEPCION', 'envio', id || remito, 'Llegada: ' + fechaLlegada);
  Logger.log('[confirmarRecepcion] OK fila=' + rowIdx);
  return { mensaje: 'Recepción confirmada' };
}

// ════════════════════════════════════════════════════════
//  V5 NEW — revertirRecepcion
//  Vuelve un envío de 'recibido' a 'pendiente'.
//  Solo admin. Antes este cambio ocurría solo en el cliente
//  y se perdía en el próximo refresh.
// ════════════════════════════════════════════════════════
function revertirRecepcion(body, user) {
  const { id } = body;
  Logger.log('[revertirRecepcion] id=' + id);

  const { sh, rowIdx } = findRow(id, null);
  if (rowIdx === -1) throw new Error('Envío no encontrado — id: ' + id);

  const now = new Date();
  sh.getRange(rowIdx, COL.estado       ).setValue('pendiente');
  sh.getRange(rowIdx, COL.fechaLlegada ).setValue('');
  sh.getRange(rowIdx, COL.obsRecv      ).setValue('');
  sh.getRange(rowIdx, COL.modificadoPor).setValue(user.usuario);
  sh.getRange(rowIdx, COL.modificadoEn ).setValue(now);

  addLog(user.usuario, 'REVERTIR_RECEPCION', 'envio', id, 'Revertido a pendiente');
  Logger.log('[revertirRecepcion] OK fila=' + rowIdx);
  return { mensaje: 'Recepción revertida' };
}

function eliminarEnvio(body, user) {
  const { id, remito } = body;
  Logger.log('[eliminarEnvio] id=' + id + ' remito=' + remito);

  const { sh, rowIdx } = findRow(id, remito);
  if (rowIdx === -1) throw new Error('Envío no encontrado — id: ' + id + ' / remito: ' + remito);

  sh.deleteRow(rowIdx);
  addLog(user.usuario, 'ELIMINAR_ENVIO', 'envio', id || remito, '');
  Logger.log('[eliminarEnvio] OK fila=' + rowIdx);
  return { mensaje: 'Envío eliminado' };
}

function editFechaDespacho(body, user) {
  const { id, fecha } = body;
  Logger.log('[editFechaDespacho] id=' + id + ' fecha=' + fecha);

  const { sh, rowIdx } = findRow(id, null);
  if (rowIdx === -1) return { mensaje: 'Envío no encontrado (puede ser seed local)' };

  sh.getRange(rowIdx, COL.fecha        ).setValue(new Date(fecha));
  sh.getRange(rowIdx, COL.modificadoPor).setValue(user.usuario);
  sh.getRange(rowIdx, COL.modificadoEn ).setValue(new Date());

  addLog(user.usuario, 'EDIT_FECHA', 'envio', id, fecha);
  return { mensaje: 'Fecha actualizada' };
}


// ════════════════════════════════════════════════════════
//  RUTAS
// ════════════════════════════════════════════════════════

function getRutas() {
  const sh = SS.getSheetByName(SH_ROUTES);
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const result = {};
  data.forEach(([key, label, normal, alerta]) => {
    if (key) result[key] = [Number(normal), Number(alerta)];
  });
  return result;
}

function updateRutas(rutas, user) {
  const sh   = SS.getSheetByName(SH_ROUTES);
  const data = sh.getDataRange().getValues();
  const now  = new Date();
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    if (key && rutas[key]) {
      sh.getRange(i + 1, 3).setValue(rutas[key][0]);
      sh.getRange(i + 1, 4).setValue(rutas[key][1]);
      sh.getRange(i + 1, 5).setValue(now);
      sh.getRange(i + 1, 6).setValue(user.usuario);
    }
  }
  addLog(user.usuario, 'UPDATE_RUTAS', 'rutas', '', JSON.stringify(rutas));
  return { mensaje: 'Tiempos actualizados' };
}


// ════════════════════════════════════════════════════════
//  TRANSPORTES
// ════════════════════════════════════════════════════════

function getTransportes() {
  const sh = SS.getSheetByName(SH_TRANS);
  if (!sh || sh.getLastRow() < 2) return { lista: [], depots: {} };
  const data   = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const result = { lista: [], depots: {} };
  data.forEach(([clave, valor]) => {
    try {
      const parsed = JSON.parse(valor);
      if (clave === 'lista') result.lista = parsed;
      else if (clave) result.depots[clave] = parsed;
    } catch(e) {}
  });
  return result;
}

function updateTransportes(transportes, user) {
  const sh = SS.getSheetByName(SH_TRANS);
  if (!sh) return { error: 'Hoja transportes no encontrada' };
  const data    = sh.getDataRange().getValues();
  const now     = new Date();
  const updates = {};
  if (transportes.lista)  updates['lista'] = JSON.stringify(transportes.lista);
  if (transportes.depots) {
    Object.entries(transportes.depots).forEach(([dep, val]) => {
      updates[dep] = JSON.stringify(val);
    });
  }
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    if (updates[key] !== undefined) {
      sh.getRange(i + 1, 2).setValue(updates[key]);
      sh.getRange(i + 1, 3).setValue(now);
      sh.getRange(i + 1, 4).setValue(user.usuario);
      delete updates[key];
    }
  }
  Object.entries(updates).forEach(([key, val]) => {
    sh.appendRow([key, val, now, user.usuario]);
  });
  addLog(user.usuario, 'UPDATE_TRANSPORTES', 'transportes', '', JSON.stringify(transportes));
  return { mensaje: 'Transportes actualizados' };
}


// ════════════════════════════════════════════════════════
//  LOG DE AUDITORÍA
// ════════════════════════════════════════════════════════

function addLog(usuario, accion, entidad, entidadId, detalle) {
  const sh = SS.getSheetByName(SH_LOG);
  if (!sh) return;
  sh.appendRow([new Date(), usuario, accion, entidad, String(entidadId || ''), detalle || '']);
}


// ════════════════════════════════════════════════════════
//  RESPUESTAS JSON
// ════════════════════════════════════════════════════════

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg, code: 400 }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════
//  SETUP — ejecutar una sola vez si creás el Sheet desde cero
// ════════════════════════════════════════════════════════

function setup() {
  _setupSheet(SH_ENVIOS,  COLS_HEADER, {1:140,2:90,3:90,4:160,5:160,6:80,7:90,8:70,9:120,10:200,11:90,12:160,13:200,14:100,15:160,16:100,17:160});
  _setupSheet(SH_ROUTES,  ['key','label','horasNormal','horasAlerta','ultimaModif','modificadoPor'], {});
  _setupSheet(SH_TRANS,   ['clave','valor','ultimaModif','modificadoPor'], {});
  _setupSheet(SH_LOG,     ['timestamp','usuario','accion','entidad','entidadId','detalle'], {});
  _getOrCreateConfig(); // crea hoja config con minVersion=1.0

  const shR = SS.getSheetByName(SH_ROUTES);
  if (shR.getLastRow() < 2) {
    shR.getRange(2,1,6,6).setValues([
      ['bsas-pico',    'Bs As ↔ Pico',    24, 48, new Date(), 'setup'],
      ['bsas-mdp',     'Bs As ↔ MDP',     24, 48, new Date(), 'setup'],
      ['bsas-rosario', 'Bs As ↔ Rosario', 24, 48, new Date(), 'setup'],
      ['pico-rosario', 'Pico ↔ Rosario',  48, 72, new Date(), 'setup'],
      ['pico-mdp',     'Pico ↔ MDP',      72, 96, new Date(), 'setup'],
      ['mdp-rosario',  'MDP ↔ Rosario',   48, 72, new Date(), 'setup'],
    ]);
  }

  const shT = SS.getSheetByName(SH_TRANS);
  if (shT.getLastRow() < 2) {
    const now = new Date();
    shT.getRange(2,1,5,4).setValues([
      ['lista',   JSON.stringify(['Brinatti','Cruz del Sur','El Directo','Santulli','Transban','Vía 2']), now, 'setup'],
      ['Pico',    JSON.stringify(['Brinatti','Cruz del Sur']),          now, 'setup'],
      ['MDP',     JSON.stringify(['Cruz del Sur','Vía 2']),             now, 'setup'],
      ['Bs As',   JSON.stringify(['Brinatti','Transban','Vía 2']),      now, 'setup'],
      ['Rosario', JSON.stringify(['Brinatti','Transban']),              now, 'setup'],
    ]);
  }

  Logger.log('✅ Setup DistriLogis V5 completo');
}

function _setupSheet(name, headers, widths) {
  let sh = SS.getSheetByName(name);
  if (sh) { Logger.log('Hoja ' + name + ' ya existe — saltando'); return; }
  sh = SS.insertSheet(name);
  sh.appendRow(headers);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#1E3A5F').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sh.setFrozenRows(1);
  Object.entries(widths).forEach(([col, w]) => sh.setColumnWidth(Number(col), w));
  Logger.log('✅ Hoja ' + name + ' creada');
}


// ════════════════════════════════════════════════════════
//  FORMATO — ejecutar manualmente para colorear el Sheet
// ════════════════════════════════════════════════════════

function aplicarFormato() {
  const sh = SS.getSheetByName(SH_ENVIOS);
  if (!sh) return;
  const lastRow = Math.max(sh.getLastRow(), 2);
  for (let i = 2; i <= lastRow; i++) {
    sh.getRange(i, 1, 1, NCOLS).setBackground(i % 2 === 0 ? '#F8FAFC' : '#FFFFFF');
  }
  const estados = sh.getRange(2, COL.estado, lastRow - 1, 1).getValues();
  estados.forEach(([val], i) => {
    sh.getRange(i + 2, COL.estado).setBackground(val === 'recibido' ? '#DCFCE7' : '#FEF3C7');
  });
  Logger.log('✅ Formato aplicado — ' + (lastRow - 1) + ' filas');
}

// ════════════════════════════════════════════════════════
//  DIAGNÓSTICO — helpers de consola
// ════════════════════════════════════════════════════════

function checkColumnas() {
  const sh = SS.getSheetByName('envios');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  headers.forEach((h, i) => {
    Logger.log('Col ' + (i+1) + ' (' + String.fromCharCode(65+i) + ') = ' + h);
  });
}

function testRecepcion() {
  const remito = '0010-10179'; // cambiar por uno real
  const sh   = SS.getSheetByName('envios');
  const data = sh.getDataRange().getValues();
  Logger.log('Total filas: ' + (data.length - 1));
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim() === String(remito).trim()) {
      Logger.log('ENCONTRADO en fila ' + (i+1) + ' estado=' + data[i][10]);
      return;
    }
  }
  Logger.log('NO ENCONTRADO: ' + remito);
}

// Test rápido del nuevo endpoint authenticate
function testAuthenticate() {
  const ok  = authenticate('administrador', 'tinchohack');
  const bad = authenticate('pico', 'wrong');
  Logger.log('OK:  ' + JSON.stringify(ok));
  Logger.log('BAD: ' + JSON.stringify(bad));
}
