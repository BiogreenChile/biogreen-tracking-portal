// ============================================
// BIOGREEN - Web App para consulta de pedidos
// Pega este código en Extensions → Apps Script
// ============================================

const SHEET_NAME    = 'Hoja 1';          // ← nombre exacto de tu hoja
const ALAS_BASE_URL = 'https://ws.alasxpress.com/api';
const ALAS_PARTNER  = 'Biogreen';

// ── Blue Express (Tracking Pull Corp) ──
const BLUE_TOKEN_URL = 'https://sso.blue.cl/oauth2/token';
const BLUE_BASE_URL  = 'https://cmkin.api.blue.cl/cmkin/bff/tracking-pull-corp/v1';

// ── SimpliRoute (operado por Globalship, courier "Global") ──
const SIMPLI_BASE_URL = 'https://api.simpliroute.com/v1';
const SIMPLI_DIAS_BUSQUEDA = 7; // ventana de días hacia atrás para buscar una visita

// ── Starken (Etracking + Consulta Imagen Entrega) ──
const STK_ETRACKING_URL = 'https://apiprd.starken.cl/etrackingRest/resumenTrackingCargaRedestinacion';
const STK_IMAGEN_URL    = 'https://restservices.starken.cl/apiprd/starkenservices/rest/consultarLinkImagenEntregayDevolucion';
const STK_TIPO_DOC      = '4'; // 4 = Boleta (formato confirmado con Starken)

// ============================================
// CREDENCIALES (Script Properties)
// Configúralas en: Configuración del proyecto (⚙️) → Propiedades del script
// Nunca las pongas como const en este archivo — este código es público.
// ============================================
function getSecret(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error('Falta configurar la propiedad "' + key + '" en Configuración del proyecto → Propiedades del script.');
  return value;
}

// ============================================
// HELPERS DE SEGURIDAD
// ============================================
// Normaliza un RUT chileno: quita puntos, guiones, espacios; mayúsculas.
// "12.345.678-K" → "12345678K"; "12345678" → "12345678"
function normalizarRut(v) {
  return String(v || '').replace(/[.\-\s]/g, '').toUpperCase();
}

// Whitelist de RUTs internos (acceso administrativo / pruebas). NUNCA quitar
// "12345678K" — es el ingreso interno de Iván / pruebas. Válido aunque su DV
// matemático no sea correcto.
const RUTS_INTERNOS = ['12345678K'];

// Valida el DV chileno con algoritmo mod 11. Reduce el espacio de fuerza bruta
// ~90% (solo 1 de cada 11 combinaciones tiene DV válido). Los RUTs internos
// pasan siempre aunque su DV no sea el matemáticamente correcto.
function validarDvRut(rutNormalizado) {
  if (!rutNormalizado || rutNormalizado.length < 2) return false;
  if (RUTS_INTERNOS.indexOf(rutNormalizado) !== -1) return true;
  const cuerpo = rutNormalizado.slice(0, -1);
  const dv     = rutNormalizado.slice(-1);
  if (!/^\d+$/.test(cuerpo)) return false;
  let suma = 0, factor = 2;
  for (let i = cuerpo.length - 1; i >= 0; i--) {
    suma += parseInt(cuerpo[i], 10) * factor;
    factor = factor === 7 ? 2 : factor + 1;
  }
  const resto = 11 - (suma % 11);
  const esperado = resto === 11 ? '0' : (resto === 10 ? 'K' : String(resto));
  return dv === esperado;
}

// Contador de fallos por pedido. Después de MAX_FALLOS fallos de RUT en la
// ventana de 15 min, el pedido queda bloqueado para nuevas consultas —
// bloquea fuerza bruta más allá del rate limit por request.
const MAX_FALLOS_POR_PEDIDO = 5;
function pedidoBloqueado(pedido) {
  try {
    const cache = CacheService.getScriptCache();
    const fails = parseInt(cache.get('fail_' + pedido) || '0', 10);
    return fails >= MAX_FALLOS_POR_PEDIDO;
  } catch (e) { return false; }
}
function registrarFalloRut(pedido) {
  try {
    const cache = CacheService.getScriptCache();
    const key   = 'fail_' + pedido;
    const fails = parseInt(cache.get(key) || '0', 10);
    cache.put(key, String(fails + 1), 900); // ventana de 15 min
  } catch (e) {}
}

// Neutraliza fórmulas de Google Sheets al escribir texto proveniente del usuario.
// Sin esto, `=IMPORTDATA(...)` se ejecuta al abrir la hoja.
function textoSeguroParaSheet(v) {
  const s = String(v || '');
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// Rate limit por clave (ej. pedido). Devuelve true si permite, false si se pasó.
// Usa CacheService; expira solo. Máximo 10 requests por clave cada 60s.
function checkRateLimit(clave) {
  try {
    const cache = CacheService.getScriptCache();
    const key = 'rl_' + String(clave || 'anon').substring(0, 50);
    const actual = parseInt(cache.get(key) || '0', 10);
    if (actual >= 10) return false;
    cache.put(key, String(actual + 1), 60);
    return true;
  } catch (e) {
    return true; // fail-open para no romper por un error del cache
  }
}

// Valida que la hoja tenga los 17 encabezados esperados y que las columnas
// críticas contengan las palabras clave. Evita "silent drift" si alguien
// reorganiza columnas (podría cambiar el significado de los datos expuestos).
function validarHeaders(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 17) return false;
    const h = sheet.getRange(1, 1, 1, 17).getValues()[0].map(function(v) {
      return String(v || '').toUpperCase();
    });
    // Firmas mínimas por columna crítica (case-insensitive, tolera variantes menores)
    const checks = [
      { col: 1,  contains: 'PEDIDO' },  // A
      { col: 3,  contains: 'RUT' },     // C
      { col: 9,  contains: 'RUT' },     // I (RUT SIN DV)
      { col: 15, contains: 'WMS' }      // O (NOTAS WMS)
    ];
    for (var i = 0; i < checks.length; i++) {
      if (h[checks[i].col - 1].indexOf(checks[i].contains) === -1) return false;
    }
    return true;
  } catch (e) {
    return false;
  }
}

// Columnas (basadas en tu estructura)
const COL = {
  pedido:       1,   // A — N PEDIDO
  nombre:       2,   // B — NOMBRE
  rut:          3,   // C — RUT
  fechaPedido:  4,   // D — FECHA PEDIDO
  categoria:    5,   // E — CAT
  puntos:       6,   // F — PTOS
  importe:      7,   // G — IMPORTE
  rutSoftland:  8,   // H — RUT SOFTLAND
  rutSinDv:     9,   // I — RUT SIN DV
  razonSocial:  10,  // J — RAZON SOCIAL
  comuna:       11,  // K — COMUNA
  comunaSoft:   12,  // L — COMUNA SOFT
  tipo:         13,  // M — TIPO
  importe2:     14,  // N — IMPORTE (segunda)
  notasWms:     15,  // O — NOTAS WMS  ← aquí va el courier
  formaPago:    16,  // P — FORMA DE PAGO
  estadoPedido: 17,  // Q — ESTADO PEDIDO
};

// Couriers a detectar en NOTAS WMS
const COURIERS = ['Alas', 'Bluexpress', 'Starken', 'Cacem', 'Mardam', 'Trapananda', 'Global'];

// Couriers con API integrada (se sincronizan al Tracking Cache)
// 'global' = Globalship, se rastrea vía API de SimpliRoute
const COURIERS_API = ['alas', 'bluexpress', 'global', 'starken'];

// ── Dashboard interno ──
const CACHE_SHEET_NAME = 'Tracking Cache';
const DASHBOARD_DOMAIN = 'biogreenchile.com';

// ── Log de consultas ──
// NO se guardan datos personales (nombre/comuna) — Ley 19.628. Solo lo mínimo
// necesario para auditoría y detección de abuso: timestamp, tipo, pedido,
// courier, resultado, detalle.
const LOG_SHEET_NAME = 'Log Consultas';
const LOG_MAX_FILAS  = 500;  // rotación agresiva: log de auditoría, no historial largo

// Registra una consulta en la hoja "Log Consultas". Se llama de forma
// asíncrona conceptual: si falla, NO rompe la consulta principal (try/catch).
// `info` es opcional: { nombre, comuna } — para consultas de pedidos.
// NO se guarda PII (nombre/comuna) — Ley 19.628. Parámetro `info` se ignora
// para mantener compatibilidad con llamadas existentes.
function registrarLog(tipo, pedido, courier, ok, extra, info) {
  // Serializa escrituras al log: evita races entre appendRow y deleteRows.
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(3000)) return; // si no obtiene lock en 3s, salta el log de esta request
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const HEADERS = ['Timestamp', 'Tipo', 'Pedido', 'Courier', 'Resultado', 'Detalle'];
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET_NAME);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    } else {
      // Migrar header si venimos de versiones anteriores (6 u 8 columnas con PII)
      const headerActual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headerActual.length !== HEADERS.length || headerActual[3] !== 'Courier') {
        // Limpia headers viejos y aplica los nuevos (las columnas extras quedan en blanco,
        // el operador puede borrar manualmente las columnas D/E antiguas de PII).
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      }
    }
    // Cada campo texto pasa por textoSeguroParaSheet para neutralizar fórmulas.
    // Sin esto, un input como "=IMPORTDATA(...)" se evaluaría al abrir la hoja.
    sheet.appendRow([
      new Date(),
      textoSeguroParaSheet(tipo),
      textoSeguroParaSheet(pedido),
      textoSeguroParaSheet(courier),
      ok ? 'OK' : 'ERROR',
      textoSeguroParaSheet(String(extra || '').substring(0, 300))
    ]);
    // Truncar si supera el límite (deja el header + últimas LOG_MAX_FILAS)
    const total = sheet.getLastRow();
    if (total > LOG_MAX_FILAS + 1) {
      sheet.deleteRows(2, total - LOG_MAX_FILAS - 1);
    }
  } catch (e) {
    // No propagar el error — no queremos que el log rompa la consulta
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ============================================
// FUNCIÓN PRINCIPAL
// ============================================
// ============================================
// PEDIDO DEMO (para presentaciones)
// Pedido 200999 → datos hardcodeados, nunca toca la hoja ni la API
// ============================================
function demoData(e) {
  // Respuesta del pedido
  if (e.parameter.pedido === '200999') {
    return jsonOut({
      ok: true,
      pedido:           '200999',
      nombre:           'Mariana',
      categoria:        'LID',
      puntos:           604,
      importe:          618000,
      tipo:             'Boleta',
      razonSocial:      'Biogreen Chile',
      comuna:           'Huechuraba',
      formaPago:        'Pagado',
      estadoPedido:     'Pagado',
      courier:          'Alas',
      fechaPedido:      'Miércoles, 24 de junio de 2026, 10:17',
      fechaObj:         '2026-06-24T10:17:00.000Z',
      despachoEstimado: 'Jueves, 25 de junio de 2026',
      despachoISO:      '2026-06-25T12:00:00.000Z',
      antesDeDoce:      true,
    });
  }
  // Respuesta del tracking Alas para el pedido demo
  if (e.parameter.courier === 'alas' && e.parameter.codigo === '200999') {
    return jsonOut({
      ok: true,
      data: {
        status:           'Entrega realizada',
        deliveryDate:     '2026-06-30T14:35:00',
        deliveryExpected: '2026-07-06T00:00:00',
        deliveryOrderId:  'DEMO200999',
        eventos: [
          { time: '2026-06-30T14:35:00', status: 'Entrega realizada',                 emoji: '✅', nuevo: true  },
          { time: '2026-06-30T08:10:00', status: 'Salió a Entregar',                  emoji: '🚀', nuevo: false },
          { time: '2026-06-29T18:45:00', status: 'Recibido en Centro de Distribución',emoji: '🏪', nuevo: false },
          { time: '2026-06-25T15:22:00', status: 'Pedido Ingresado',                  emoji: '📋', nuevo: false },
        ],
      },
    });
  }
  return null;
}

function doGet(e) {
  // Intercepción demo — siempre primero
  const demo = demoData(e);
  if (demo) return demo;

  // Dashboard interno (acceso restringido por dominio)
  if (e.parameter.dashboard) {
    return handleDashboardRequest();
  }
  // Consulta de tracking a un courier específico (usado por fetch() desde el frontend externo)
  if (e.parameter.courier && e.parameter.codigo) {
    return handleCourierRequest(e);
  }
  // Si viene con parámetro pedido → devuelve JSON
  if (e.parameter.pedido) {
    return handleRequest(e);
  }
  // Si no viene con parámetros, redirigimos al portal público oficial.
  // (El HTML antiguo "Seguimiento de pedido.html" fue removido; ahora todo el
  //  frontend vive en GitHub Pages / biogreenchile.com/seguimientodepedido.)
  return HtmlService.createHtmlOutput(
    '<!doctype html><meta charset="utf-8">' +
    '<meta http-equiv="refresh" content="0;url=https://biogreenchile.com/seguimientodepedido">' +
    '<p style="font-family:Arial;padding:40px;text-align:center;color:#555">' +
    'Redirigiendo al portal de seguimiento… ' +
    'Si no cargas automáticamente, ' +
    '<a href="https://biogreenchile.com/seguimientodepedido">haz clic aquí</a>.</p>'
  ).setTitle('Redirigiendo al portal · Biogreen');
}

function doPost(e) {
  return handleRequest(e);
}

// ============================================
// CONSULTA DE COURIER VÍA URL (para fetch() externo)
// GET /exec?pedido=97219&rut=12345678K&courier=alas
// SEGURIDAD:
//   - Requiere pedido + RUT (valida contra la hoja antes de llamar al courier).
//   - El courier se DERIVA de la fila validada, NO se acepta desde la URL como
//     autoridad. El parámetro `courier` de la URL se ignora si no coincide con
//     el detectado en la hoja.
//   - Rate limit por pedido para prevenir enumeración / relay masivo.
// ============================================
function handleCourierRequest(e) {
  const codigo = String(e.parameter.codigo || e.parameter.pedido || '').trim();
  const rut    = normalizarRut(e.parameter.rut);

  if (!codigo || !rut) {
    return jsonOut({ ok: false, error: 'Pedido y RUT son obligatorios.' });
  }

  // Validación matemática del DV (con whitelist de RUTs internos).
  if (!validarDvRut(rut)) {
    return jsonOut({ ok: false, error: 'Datos no coinciden.' });
  }

  // Bloqueo por fallos acumulados
  if (pedidoBloqueado(codigo)) {
    return jsonOut({ ok: false, error: 'Consulta temporalmente bloqueada. Intenta más tarde.' });
  }

  // Rate limit por pedido
  if (!checkRateLimit('cou_' + codigo)) {
    return jsonOut({ ok: false, error: 'Demasiadas consultas. Intenta en un minuto.' });
  }

  // Validar hoja + headers antes de exponer cualquier dato
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || !validarHeaders(sheet)) {
    return jsonOut({ ok: false, error: 'Servicio temporalmente no disponible.' });
  }

  // Buscar pedido + validar RUT en la fila. Si no coincide, error genérico
  // (no confirmar existencia del pedido — evita enumeración).
  const data = sheet.getDataRange().getValues();
  let matched = false;
  let filaCourier = null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[COL.pedido - 1]).trim() !== codigo) continue;
    const rutFila     = normalizarRut(row[COL.rut - 1]);
    const rutSinDvFila = normalizarRut(row[COL.rutSinDv - 1]);
    if (rut === rutFila || rut === rutSinDvFila) {
      matched = true;
      filaCourier = detectarCourier(String(row[COL.notasWms - 1] || ''));
      break;
    }
  }

  if (!matched) {
    registrarFalloRut(codigo);
    // Mensaje genérico (no diferenciar "pedido no existe" vs "RUT no coincide")
    return jsonOut({ ok: false, error: 'Datos no coinciden.' });
  }
  if (!filaCourier) {
    return jsonOut({ ok: false, error: 'Este pedido no tiene tracking en línea.' });
  }

  // Derivar courier de la fila, IGNORAR el parámetro de la URL como autoridad.
  const courier = String(filaCourier).toLowerCase();

  let resultado;
  if (courier === 'alas') {
    resultado = consultarAlas(codigo);
  } else if (courier === 'bluexpress') {
    resultado = consultarBlueExpress(codigo);
  } else if (courier === 'global') {
    resultado = consultarSimpliRoute(codigo);
  } else if (courier === 'starken') {
    resultado = consultarStarken(codigo);
  } else {
    resultado = { ok: false, error: 'Este pedido no tiene tracking en línea.' };
  }

  registrarLog('tracking', codigo, courier, resultado.ok, resultado.ok ? '' : (resultado.error || ''));
  return jsonOut(resultado);
}

// handleRequest(e, opts?):
//   opts._internal = true → salta la validación de RUT (usado por consultarPedido
//     desde el dashboard interno, que ya autentica por dominio).
// SEGURIDAD:
//   - Endpoint público → requiere pedido + RUT (evita enumeración por número
//     secuencial de pedido).
//   - Rate limit por pedido para prevenir fuerza bruta.
//   - Detecta duplicados: si hay >1 fila con el mismo pedido, se rechaza en
//     vez de devolver la primera al azar.
//   - Valida los 17 encabezados esperados; aborta si la estructura cambió.
function handleRequest(e, opts) {
  opts = opts || {};
  try {
    const pedido = (e.parameter.pedido || '').toString().trim();
    const rut    = normalizarRut(e.parameter.rut);

    if (!pedido) {
      return jsonOut({ ok: false, error: 'Falta el número de pedido.' });
    }
    if (!opts._internal && !rut) {
      return jsonOut({ ok: false, error: 'Pedido y RUT son obligatorios.' });
    }

    // Validación matemática del DV (con whitelist de RUTs internos).
    // Rechaza RUTs con DV inválido antes de tocar la hoja → corta ~90% del
    // espacio de fuerza bruta.
    if (!opts._internal && !validarDvRut(rut)) {
      return jsonOut({ ok: false, error: 'Datos no coinciden.' });
    }

    // Bloqueo si acumuló muchos fallos previos (fuerza bruta sostenida)
    if (!opts._internal && pedidoBloqueado(pedido)) {
      return jsonOut({ ok: false, error: 'Consulta temporalmente bloqueada. Intenta más tarde.' });
    }

    // Rate limit por pedido (protege contra fuerza bruta de RUT)
    if (!opts._internal && !checkRateLimit('req_' + pedido)) {
      return jsonOut({ ok: false, error: 'Demasiadas consultas. Intenta en un minuto.' });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonOut({ ok: false, error: 'Servicio temporalmente no disponible.' });
    }
    if (!validarHeaders(sheet)) {
      return jsonOut({ ok: false, error: 'Servicio temporalmente no disponible.' });
    }

    const data = sheet.getDataRange().getValues();

    // Recolectar TODAS las coincidencias (para detectar duplicados)
    const matches = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[COL.pedido - 1]).trim() === pedido) matches.push(row);
    }

    if (!matches.length) {
      registrarLog('consulta', pedido, '', false, 'No encontrado');
      return jsonOut({ ok: false, error: 'Datos no coinciden.' });
    }

    // Filtrar por RUT (salvo llamada interna del dashboard)
    let row;
    if (opts._internal) {
      // Confianza interna: si hay duplicados, se registra pero se toma la primera
      if (matches.length > 1) {
        registrarLog('consulta', pedido, '', false, 'DUPLICADO (' + matches.length + ' filas) — se usó la primera');
      }
      row = matches[0];
    } else {
      const validas = matches.filter(function(r) {
        const rutFila = normalizarRut(r[COL.rut - 1]);
        const rutSinDv = normalizarRut(r[COL.rutSinDv - 1]);
        return rut === rutFila || rut === rutSinDv;
      });
      if (!validas.length) {
        // Mensaje GENÉRICO: no diferenciar "pedido no existe" vs "RUT no coincide"
        registrarFalloRut(pedido);
        registrarLog('consulta', pedido, '', false, 'RUT no coincide');
        return jsonOut({ ok: false, error: 'Datos no coinciden.' });
      }
      if (validas.length > 1) {
        registrarLog('consulta', pedido, '', false, 'DUPLICADO validado (' + validas.length + ') — rechazado');
        return jsonOut({ ok: false, error: 'Hay una inconsistencia con este pedido. Contáctanos en Atención al Cliente.' });
      }
      row = validas[0];
    }

    // Extraer courier desde NOTAS WMS
    const notasWms  = String(row[COL.notasWms - 1] || '');
    const courier   = detectarCourier(notasWms);

    // Fecha y hora del pedido
    const fechaRaw  = row[COL.fechaPedido - 1];
    const fechaInfo = parsearFecha(fechaRaw);

    // Calcular despacho estimado
    const despachoInfo = calcularDespacho(fechaInfo.dateObj);

    const nombreLog = toTitleCase(String(row[COL.nombre - 1] || ''));
    const comunaLog = toTitleCase(String(row[COL.comuna - 1] || ''));
    // Log SIN PII: solo pedido + courier + resultado
    registrarLog('consulta', pedido, courier || '', true, '');

    return jsonOut({
      ok:           true,
      pedido:       pedido,
      nombre:       nombreLog,
      categoria:    String(row[COL.categoria - 1] || ''),
      puntos:       row[COL.puntos - 1] || 0,
      importe:      row[COL.importe - 1] || 0,
      tipo:         String(row[COL.tipo - 1] || ''),        // Boleta / Factura
      razonSocial:  String(row[COL.razonSocial - 1] || ''),
      comuna:       comunaLog,
      formaPago:    String(row[COL.formaPago - 1] || ''),
      estadoPedido: String(row[COL.estadoPedido - 1] || ''),
      courier:      courier,
      fechaPedido:  fechaInfo.texto,
      fechaObj:     fechaInfo.iso,
      despachoEstimado: despachoInfo.texto,
      despachoISO:      despachoInfo.iso,
      antesDeDoce:      despachoInfo.antesDeDoce,
    });

  } catch (err) {
    return jsonOut({ ok: false, error: 'Error interno.' }); // no exponer err.message
  }
}

// ============================================
// DETECTAR COURIER EN NOTAS WMS
// ============================================
function detectarCourier(texto) {
  if (!texto) return null;
  const upper = texto.toLowerCase();
  for (const c of COURIERS) {
    if (upper.includes(c.toLowerCase())) return c;
  }
  return null;
}

// ============================================
// PARSEAR FECHA (Google Sheets puede devolver Date o string)
// ============================================
function parsearFecha(raw) {
  if (!raw) return { texto: 'No disponible', iso: null, dateObj: null };

  let d;
  if (raw instanceof Date) {
    d = raw;
  } else {
    d = new Date(raw);
  }

  if (isNaN(d.getTime())) {
    return { texto: String(raw), iso: null, dateObj: null };
  }

  const texto = d.toLocaleDateString('es-CL', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric',
    hour: '2-digit', minute: '2-digit', hour12: false
  });

  return {
    texto:   texto.charAt(0).toUpperCase() + texto.slice(1),
    iso:     d.toISOString(),
    dateObj: d,
  };
}

// ============================================
// CALCULAR FECHA DE DESPACHO ESTIMADO
// Reglas:
//   - Antes de 12:00 día hábil  → despacha día hábil SIGUIENTE (cualquier hora)
//   - Después de 12:00 día hábil → despacha SUBSIGUIENTE día hábil (PM)
//   - Fin de semana              → se trata como lunes antes de 12:00 → despacha martes
// ============================================
function calcularDespacho(dateObj) {
  if (!dateObj) return { texto: 'No disponible', iso: null, antesDeDoce: null };

  const hora         = dateObj.getHours() + dateObj.getMinutes() / 60;
  const diaSemana    = dateObj.getDay(); // 0=Dom, 6=Sab
  const esFinSemana  = diaSemana === 0 || diaSemana === 6;
  const antesDeDoce  = hora < 12;

  let diasAgregar;

  if (esFinSemana) {
    // Fin de semana = lunes antes de 12 → despacha martes
    const diasHastaLunes = diaSemana === 6 ? 2 : 1; // Sab→2, Dom→1
    diasAgregar = diasHastaLunes + 1; // martes
  } else if (antesDeDoce) {
    diasAgregar = 1; // siguiente día hábil
  } else {
    diasAgregar = 2; // subsiguiente día hábil
  }

  // Calcular fecha destino saltando fines de semana
  let despacho = new Date(dateObj);
  let agregados = 0;
  while (agregados < diasAgregar) {
    despacho.setDate(despacho.getDate() + 1);
    const dow = despacho.getDay();
    if (dow !== 0 && dow !== 6) agregados++; // solo días hábiles
  }

  // Siempre PM para despachos post-12
  const horaTexto = (esFinSemana || antesDeDoce) ? '' : ' PM';

  const textoFecha = despacho.toLocaleDateString('es-CL', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
  });

  return {
    texto:        textoFecha.charAt(0).toUpperCase() + textoFecha.slice(1) + horaTexto,
    iso:          despacho.toISOString(),
    antesDeDoce:  antesDeDoce && !esFinSemana,
  };
}

// ============================================
// HELPERS
// ============================================
function toTitleCase(str) {
  // Divide por espacios para evitar problemas con letras acentuadas
  return str.toLowerCase().split(' ')
    .map(word => word.length > 0 ? word.charAt(0).toUpperCase() + word.slice(1) : '')
    .join(' ');
}

// ============================================
// FUNCIÓN PARA google.script.run (sin CORS)
// SOLO uso interno desde el Dashboard (autenticado por dominio).
// ============================================
function consultarPedido(pedido) {
  // Auth por dominio antes de saltar la validación de RUT
  const email = Session.getActiveUser().getEmail();
  if (!email || email.split('@')[1] !== DASHBOARD_DOMAIN) {
    return { ok: false, error: 'Acceso no autorizado.' };
  }
  var fakeEvent = { parameter: { pedido: pedido } };
  var output = handleRequest(fakeEvent, { _internal: true });
  return JSON.parse(output.getContent());
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// CONSULTA TRACKING ALAS (llamada server-side)
// ============================================
function consultarAlas(pedido) {
  try {
    const payload = {
      partner:           ALAS_PARTNER,
      senderCode:        getSecret('ALAS_SENDER'),
      deliveryOrderCode: String(pedido)
    };

    const options = {
      method:           'post',
      contentType:      'application/json',
      headers:          { 'x-alas-ce0-api-key': getSecret('ALAS_API_KEY') },
      payload:          JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(ALAS_BASE_URL + '/delivery-orders/status', options);
    const code     = response.getResponseCode();

    if (code !== 200) {
      return { ok: false, error: 'Alas HTTP ' + code };
    }

    const data = JSON.parse(response.getContentText());
    return { ok: true, data: data };

  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// ============================================
// CONSULTA TRACKING BLUE EXPRESS (llamada server-side)
// ============================================
function obtenerTokenBlue() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('blue_token');
  if (cached) return cached;

  const options = {
    method:      'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type:    'client_credentials',
      client_id:     getSecret('BLUE_CLIENT_ID'),
      client_secret: getSecret('BLUE_CLIENT_SECRET')
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(BLUE_TOKEN_URL, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('No se pudo obtener token Blue Express (HTTP ' + response.getResponseCode() + ')');
  }

  const data  = JSON.parse(response.getContentText());
  const token = data.access_token;
  const ttl   = Math.max((data.expires_in || 3600) - 60, 60);
  cache.put('blue_token', token, ttl);
  return token;
}

// La API tracking-pull-corp de Blue Express SOLO busca por referencia exacta:
//   - "references" es obligatorio en cada llamada (no hay "listar todo")
//   - No acepta wildcards ni búsqueda por prefijo/sufijo
// Desde la implementación con cuenta propia, el operador registra como
// referencia en Blue SOLO el número de pedido Biogreen (ej. "101216"), por lo
// que basta con pasar el número de pedido para encontrar el envío.
// (Los pedidos antiguos cargados como "{pedido}-MT{wms}-BIOGREEN" no se
//  encontrarán con esto — solo aplican los emitidos con el nuevo formato.)
// Acepta una referencia o varias separadas por coma (multi-ref, más eficiente).
function consultarBlueExpress(referencias) {
  try {
    const token   = obtenerTokenBlue();
    const account = encodeURIComponent(getSecret('BLUE_ACCOUNT'));
    const refs    = encodeURIComponent(String(referencias));

    const options = {
      method:  'get',
      headers: {
        'Authorization': 'Bearer ' + token,
        'x-api-key':     getSecret('BLUE_API_KEY')
      },
      muteHttpExceptions: true
    };

    const url      = BLUE_BASE_URL + '/search?accounts=' + account + '&references=' + refs;
    const response = UrlFetchApp.fetch(url, options);
    const code     = response.getResponseCode();

    if (code !== 200) {
      return { ok: false, error: 'Blue Express HTTP ' + code };
    }

    const data = JSON.parse(response.getContentText());
    if (!data.data || !data.data.length) {
      return { ok: false, error: 'No encontrado en Blue Express' };
    }

    // Una sola referencia → devuelve el objeto; varias → devuelve el array completo
    return { ok: true, data: data.data.length === 1 ? data.data[0] : data.data };

  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// ============================================
// CONSULTA TRACKING SIMPLIROUTE / GLOBALSHIP (courier "Global")
// ============================================
// La API lista visitas por fecha (planned_date), no por referencia. La
// referencia en SimpliRoute es "{pedido}BIOGREEN" (ej. "101432BIOGREEN").
// Buscamos el pedido escaneando los últimos SIMPLI_DIAS_BUSQUEDA días.
// Cada día se cachea 10 min para acelerar consultas repetidas.
function obtenerVisitasSimpliPorFecha(fechaStr) {
  const cache  = CacheService.getScriptCache();
  const clave  = 'simpli_' + fechaStr;
  const cached = cache.get(clave);
  if (cached) return JSON.parse(cached);

  const url = SIMPLI_BASE_URL + '/routes/visits/?planned_date=' + fechaStr;
  const resp = UrlFetchApp.fetch(url, {
    method:  'get',
    headers: { 'Authorization': 'Token ' + getSecret('SIMPLI_TOKEN') },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) return [];

  const visitas = JSON.parse(resp.getContentText()) || [];
  // Guardamos solo lo necesario para no exceder el límite de 100KB del cache
  const liviano = visitas.map(function(v) {
    return {
      reference: v.reference, status: v.status, title: v.title,
      address: v.address, tracking_id: v.tracking_id,
      checkout_time: v.checkout_time, checkout_comment: v.checkout_comment,
      planned_date: v.planned_date, estimated_time_arrival: v.estimated_time_arrival,
      signature: v.signature, pictures: v.pictures
    };
  });
  try { cache.put(clave, JSON.stringify(liviano), 600); } catch (e) {}
  return liviano;
}

function consultarSimpliRoute(pedido) {
  try {
    const objetivo = String(pedido).trim();
    const hoy = new Date();

    // Prioridad de estados: el más avanzado gana si un pedido aparece duplicado
    const PRIORIDAD = { 'completed': 5, 'failed': 4, 'partial': 3, 'canceled': 2, 'pending': 1 };
    const rankVisita = function(v) { return PRIORIDAD[String(v.status || '').toLowerCase()] || 0; };

    for (let i = 0; i < SIMPLI_DIAS_BUSQUEDA; i++) {
      const d = new Date(hoy.getTime() - i * 86400000);
      const fechaStr = Utilities.formatDate(d, 'America/Santiago', 'yyyy-MM-dd');
      const visitas  = obtenerVisitasSimpliPorFecha(fechaStr);

      const coincidencias = visitas.filter(function(v) {
        const ref = String(v.reference || '').toUpperCase();
        // SEGURIDAD: comparación EXACTA. Antes se aceptaba prefix match
        // (`indexOf === 0`), lo que permitía que `codigo=1` matchee "12345BIOGREEN"
        // → exposición de tracking/dirección/firma de otros clientes.
        return ref === (objetivo.toUpperCase() + 'BIOGREEN');
      });

      if (coincidencias.length) {
        // Si hay duplicados (mismo pedido, distinto estado), tomamos el estado final
        coincidencias.sort(function(a, b) { return rankVisita(b) - rankVisita(a); });
        return { ok: true, data: coincidencias[0] };
      }
    }

    return { ok: false, error: 'No encontrado en SimpliRoute' };

  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// ── Extrae estado normalizado desde una visita SimpliRoute (para el cache) ──
function extraerEstadoSimpli(visit) {
  if (!visit) return null;
  const status = String(visit.status || '').toLowerCase();
  const entregado = status === 'completed';
  const fechaFin = entregado && visit.checkout_time ? visit.checkout_time : null;
  const mapaEstado = {
    'pending':   'Pendiente de entrega',
    'partial':   'Entrega parcial',
    'completed': 'Entregado',
    'failed':    'Entrega fallida',
    'canceled':  'Anulado'
  };
  return { estado: mapaEstado[status] || visit.status || 'Desconocido', entregado: entregado, fechaFin: fechaFin };
}

// ============================================
// CONSULTA TRACKING STARKEN (Etracking + Imagen de entrega)
// ============================================
// numeroDocumento = número de pedido Biogreen (Starken lo registra como "Boleta"
// tipoDocumento=4). Devuelve estado, historial y — si está entregado — se
// consulta también la imagen de entrega para pasarla al portal.
function consultarStarken(pedido) {
  try {
    const payload = {
      tracking: [{
        numeroDocumento: String(pedido).trim(),
        numeroOrdenFlete: '',
        tipoDocumento: STK_TIPO_DOC
      }],
      rutEmpresa: getSecret('STK_RUT')
    };

    const resp = UrlFetchApp.fetch(STK_ETRACKING_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'api-key':  getSecret('STK_API_KEY'),
        'cli-rut':  getSecret('STK_RUT'),
        'password': getSecret('STK_PASSWORD')
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      return { ok: false, error: 'Starken HTTP ' + resp.getResponseCode() };
    }

    const data = JSON.parse(resp.getContentText());
    const lista = data.listaResumenRedestinacion && data.listaResumenRedestinacion.ordenFlete || [];
    const orden = lista[0];

    if (!orden || orden.codigoSalida !== 1) {
      return { ok: false, error: (orden && (orden.mensaje || orden.mensajeSalida)) || 'No encontrado en Starken' };
    }

    // Si ya está entregado, intentamos traer la imagen de prueba de entrega
    let imagen = null;
    if (/entreg/i.test(orden.estadoOrdenFlete || '') && orden.numeroOrdenFlete) {
      imagen = consultarImagenStarken(orden.numeroOrdenFlete);
    }

    return { ok: true, data: Object.assign({}, orden, { _imagen: imagen }) };

  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// Devuelve { linkImagen, latitud, longitud, descripcion } o null si no hay
function consultarImagenStarken(numeroOrdenFlete) {
  try {
    const resp = UrlFetchApp.fetch(STK_IMAGEN_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'rut':   getSecret('STK_RUT'),
        'clave': getSecret('STK_PASSWORD')
      },
      payload: JSON.stringify({
        codigoOrdenFlete: Number(numeroOrdenFlete),
        rutEmpresa:       Number(getSecret('STK_RUT')),
        rutUsuario:       10,
        password:         'SOPORTE'
      }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return null;
    const data = JSON.parse(resp.getContentText());
    if (data.codigoRespuesta !== 1 || !data.lista || !data.lista.length) return null;

    // Priorizamos evento de entrega efectiva; si no hay, tomamos el último con imagen
    const conImagen = data.lista.filter(function(e) { return e.linkImagen; });
    if (!conImagen.length) return null;
    const entrega = conImagen.find(function(e) { return /entreg/i.test(e.tipoEvento || ''); }) || conImagen[conImagen.length - 1];

    return {
      linkImagen: entrega.linkImagen,
      latitud:    entrega.latitud,
      longitud:   entrega.longitud,
      fecha:      entrega.fecha,
      descripcion: entrega.descripcionDevolucion || entrega.tipoEvento || null
    };
  } catch(e) {
    return null;
  }
}

// ── Extrae estado normalizado desde la respuesta de Starken (para el cache) ──
function extraerEstadoStarken(orden) {
  if (!orden) return null;
  const estado    = orden.estadoOrdenFlete || 'Desconocido';
  const entregado = /entreg/i.test(estado);
  const fechaFin  = entregado && orden.fechaHoraEntregaOrdenFlete ? orden.fechaHoraEntregaOrdenFlete : null;
  return { estado: estado, entregado: entregado, fechaFin: fechaFin };
}

// ============================================
// DASHBOARD INTERNO (acceso restringido por dominio)
// ============================================
function handleDashboardRequest() {
  const email = Session.getActiveUser().getEmail();
  if (!email || email.split('@')[1] !== DASHBOARD_DOMAIN) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:60px;text-align:center;">' +
      '<h2>Acceso restringido</h2>' +
      '<p>Debes iniciar sesión con tu cuenta @' + DASHBOARD_DOMAIN + ' para ver este dashboard.</p>' +
      '</div>'
    );
  }
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Dashboard de Despachos · Biogreen')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Crea o retorna la hoja de caché de tracking ──
const CACHE_COLS = 12; // columnas del Tracking Cache
function obtenerCacheSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CACHE_SHEET_NAME);
  const HEADERS = ['Pedido', 'Courier', 'Region', 'Comuna', 'Estado', 'Entregado', 'FechaDespacho', 'DiasEnTransito', 'Fuente', 'YaDespachado', 'AlertaBodega', 'FechaPedido'];
  if (!sheet) {
    sheet = ss.insertSheet(CACHE_SHEET_NAME);
  }
  // Asegura el header actual (12 columnas) — migra automáticamente del formato viejo de 11
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sheet;
}

// ── Mapa Comuna → Región, usando la hoja "Region y comuna" (A=Region, B=Comuna) ──
const MATRIZ_SHEET_NAME = 'Region y comuna';
let _mapaComunaRegion = null;
function obtenerRegionPorComuna(comuna) {
  if (!_mapaComunaRegion) {
    _mapaComunaRegion = {};
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MATRIZ_SHEET_NAME);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const region = String(data[i][0] || '').trim();
        const com    = String(data[i][1] || '').trim().toLowerCase();
        if (com && region) _mapaComunaRegion[com] = region;
      }
    }
  }
  return _mapaComunaRegion[String(comuna || '').trim().toLowerCase()] || 'Sin clasificar';
}

// ── Extrae estado normalizado desde la respuesta de Alas (sin asumir atraso) ──
function extraerEstadoAlas(order) {
  if (!order || !order.status) return null;
  const estado = order.description || order.status;
  const entregado = /entreg/i.test(estado);
  // deliveryDate = fecha real de entrega informada por Alas (si ya se entregó)
  const fechaFin = entregado && order.deliveryDate ? order.deliveryDate : null;
  return { estado: estado, entregado: entregado, fechaFin: fechaFin };
}

// ── Extrae estado normalizado desde la respuesta de Blue Express ──
// IMPORTANTE: order.stateDesc queda pegado en "IMPRESO" para siempre (es el
// estado inicial de emisión, NO se actualiza). El estado REAL vive en
// packages[0].latestStatus.statusCode. Nunca mostrar stateDesc.
function extraerEstadoBlue(order) {
  if (!order) return null;
  const pkg    = (order.packages && order.packages[0]) || {};
  const latest = pkg.latestStatus || order.latestStatus || {};
  const code   = latest.statusCode || '';

  // Códigos → { texto legible, si cierra el ciclo }
  // NOTA: DLV = "Devolución Entregada" (el paquete volvió al remitente, cierra ciclo pero NO se entregó al cliente).
  //       DR/RD = Devuelto/Rechazado (también cierran ciclo).
  const MAPA = {
    'DL':  {t:'Entregado',                 fin:true,  ok:true  }, // entregado al cliente ✅
    'DLV': {t:'Devolución entregada al remitente', fin:true,  ok:false }, // volvió al remitente
    'DR':  {t:'Devuelto al remitente',     fin:true,  ok:false },
    'RD':  {t:'Rechazado por destinatario',fin:true,  ok:false },
    'PU':  {t:'Retirado por Blue',         fin:false, ok:false },
    'SOB': {t:'En bodega Blue',            fin:false, ok:false },
    'PUH': {t:'En hub Blue',               fin:false, ok:false },
    'IC':  {t:'En camino',                 fin:false, ok:false },
    'AS':  {t:'En camino',                 fin:false, ok:false },
    'DA':  {t:'Arribado a sucursal',       fin:false, ok:false },
    'LD':  {t:'En reparto',                fin:false, ok:false },
    'NH':  {t:'Nadie en casa',             fin:false, ok:false },
    'TS':  {t:'En solución',               fin:false, ok:false },
    'MRC': {t:'Mal ruteo cliente',         fin:false, ok:false }
  };
  const m = MAPA[code] || {t: code || 'Desconocido', fin:false, ok:false};

  // "Entregado" a efectos del dashboard = ciclo cerrado (llegó a destino o volvió).
  // Esto evita que pedidos devueltos/rechazados queden como "en tránsito" para siempre.
  const entregado = m.fin;
  const fechaFin  = entregado && latest.statusDate ? latest.statusDate : null;
  return { estado: m.t, entregado: entregado, fechaFin: fechaFin };
}

// ============================================
// SINCRONIZAR TRACKING (ejecutar vía trigger por tiempo)
// ============================================
// No tiene sentido seguir consultando pedidos antiguos ya resueltos.
const SYNC_DIAS_MAXIMO       = 30; // ventana de pedidos a sincronizar (días desde la fecha del pedido)
const ALERTA_DIAS_MAX        = 5;  // solo alerta de bodega si el despacho estimado fue en los últimos N días
const GLOBAL_DIAS_VISIBILIDAD = 8; // SimpliRoute solo muestra ~7 días; los Global más viejos se dan por entregados
const BLUE_DIAS_IMPRESO_ANULADO = 8; // Blue en "IMPRESO" más de N días = etiqueta nunca retirada → anulado en la práctica
const MANUAL_DIAS_ENTREGADO  = 10; // Couriers sin API (Cacem, Mardam, etc.): >N días desde despacho → asumir entregado

// Optimizado para no exceder el límite de 6 min de Apps Script con >1000 filas:
//  - Un solo recorrido para recolectar pedidos activos
//  - Alas y Starken en PARALELO con UrlFetchApp.fetchAll() (por lotes)
//  - Blue en multi-ref (varias referencias por llamada)
//  - SimpliRoute/Global: prefetch de los últimos días UNA sola vez (mapa en memoria)
function sincronizarTracking() {
  // Lock a nivel script: si el trigger de 30 min se solapa con una ejecución
  // manual, la segunda se salta en vez de sobrescribir el cache a medio construir.
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;

  // Aborta si la estructura de la hoja cambió (evita sobrescribir cache con
  // datos leídos de columnas erróneas).
  if (!validarHeaders(sheet)) {
    Logger.log('sincronizarTracking abortada: headers no válidos');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const ahora = new Date();
  const limiteFecha = new Date(ahora.getTime() - SYNC_DIAS_MAXIMO * 86400000);

  // ── PASO 1: recolectar pedidos a procesar (sin llamar APIs todavía) ──
  const pendientes = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const pedido = String(row[COL.pedido - 1] || '').trim();
    if (!pedido) continue;

    const estadoPedidoSheet = String(row[COL.estadoPedido - 1] || '').trim();
    if (/entreg/i.test(estadoPedidoSheet)) continue; // ya entregado → no reconsultar

    const fechaPedidoRaw = row[COL.fechaPedido - 1];
    const fechaPedidoObj = fechaPedidoRaw instanceof Date ? fechaPedidoRaw : new Date(fechaPedidoRaw);
    if (!isNaN(fechaPedidoObj.getTime()) && fechaPedidoObj < limiteFecha) continue; // fuera de ventana

    const courier      = detectarCourier(String(row[COL.notasWms - 1] || ''));
    const courierLower = (courier || '').toLowerCase();
    const fechaInfo    = parsearFecha(fechaPedidoRaw);
    const despachoInfo = calcularDespacho(fechaInfo.dateObj);
    const fechaDespacho = despachoInfo.iso ? new Date(despachoInfo.iso) : null;

    pendientes.push({
      pedido: pedido,
      courier: courier,
      courierLower: courierLower,
      esAnulado: /anula/i.test(estadoPedidoSheet),
      estadoPedidoSheet: estadoPedidoSheet,
      fechaPedidoIso: fechaInfo.iso,
      fechaDespacho: fechaDespacho,
      yaDespachado: !!(fechaDespacho && ahora >= fechaDespacho),
      comuna: String(row[COL.comuna - 1] || '').trim()
    });
  }

  // ── PASO 2: consultar cada courier por lote / en paralelo ──
  const activos = pendientes.filter(function(p) { return !p.esAnulado; });
  const alasMap   = batchAlas(activos.filter(function(p){return p.courierLower==='alas';}).map(function(p){return p.pedido;}));
  const stkMap    = batchStarken(activos.filter(function(p){return p.courierLower==='starken';}).map(function(p){return p.pedido;}));
  const blueMap   = batchBlue(activos.filter(function(p){return p.courierLower==='bluexpress';}).map(function(p){return p.pedido;}));
  const simpliMap = activos.some(function(p){return p.courierLower==='global';}) ? prefetchSimpli() : {};

  // ── PASO 3: armar filas del cache ──
  const cacheData = [];
  for (let k = 0; k < pendientes.length; k++) {
    const p = pendientes[k];
    let info = null, fuente = 'Manual', alertaBodega = false;

    if (p.esAnulado) {
      info = { estado: 'Anulado', entregado: true, fechaFin: null };
    } else if (p.courierLower === 'alas' && alasMap[p.pedido]) {
      info = extraerEstadoAlas(alasMap[p.pedido]); fuente = 'API';
    } else if (p.courierLower === 'bluexpress' && blueMap[p.pedido]) {
      info = extraerEstadoBlue(blueMap[p.pedido]); fuente = 'API';
    } else if (p.courierLower === 'starken' && stkMap[p.pedido]) {
      info = extraerEstadoStarken(stkMap[p.pedido]); fuente = 'API';
    } else if (p.courierLower === 'global' && simpliMap[p.pedido]) {
      info = extraerEstadoSimpli(simpliMap[p.pedido]); fuente = 'API';
    }

    // Global (SimpliRoute) solo tiene visibilidad de ~7 días. Un pedido Global que
    // ya no aparece y cuyo despacho fue hace más de GLOBAL_DIAS_VISIBILIDAD días se
    // da por ENTREGADO (a esas alturas ya llegó; SimpliRoute simplemente lo perdió).
    if (!info && p.courierLower === 'global' && p.fechaDespacho &&
        (ahora - p.fechaDespacho) > GLOBAL_DIAS_VISIBILIDAD * 86400000) {
      info = { estado: 'Entregado', entregado: true, fechaFin: null };
      fuente = 'API';
    }

    // Blue sin ningún evento real en el paquete = etiqueta emitida pero nunca
    // retirada. Después de N días es en la práctica un pedido abandonado.
    if (info && p.courierLower === 'bluexpress' && !info.entregado &&
        /desconocido/i.test(info.estado) && p.fechaDespacho &&
        (ahora - p.fechaDespacho) > BLUE_DIAS_IMPRESO_ANULADO * 86400000) {
      info = { estado: 'Anulado (nunca retirado por courier)', entregado: true, fechaFin: null };
      fuente = 'API';
    }

    const esCourierApi = COURIERS_API.indexOf(p.courierLower) !== -1;
    // La alerta de bodega solo aplica a pedidos RECIENTES: si un pedido que debía
    // despacharse en los últimos ALERTA_DIAS_MAX días no aparece en el courier, es
    // un atraso real de bodega. Pedidos viejos sin registro = dato viejo (los
    // couriers ya no los devuelven, o SimpliRoute solo ve 7 días) → NO es alerta.
    const despachoReciente = p.fechaDespacho && (ahora - p.fechaDespacho) <= ALERTA_DIAS_MAX * 86400000;
    if (!info && esCourierApi && p.yaDespachado && despachoReciente) {
      alertaBodega = true;
      info = { estado: 'Sin guía en courier (posible atraso de bodega)', entregado: false, fechaFin: null };
    }

    // Couriers SIN API integrada (Cacem, Mardam, Trapananda, Sin courier): después
    // de MANUAL_DIAS_ENTREGADO días desde el despacho, se dan por entregados
    // (no podemos verificar, y a esas alturas ya llegó).
    if (!info && !esCourierApi && p.fechaDespacho &&
        (ahora - p.fechaDespacho) > MANUAL_DIAS_ENTREGADO * 86400000) {
      info = { estado: 'Entregado (asumido - sin API)', entregado: true, fechaFin: null };
    }

    if (!info) {
      info = { estado: 'Sin tracking disponible', entregado: /entreg/i.test(p.estadoPedidoSheet), fechaFin: null };
    }

    const region = obtenerRegionPorComuna(p.comuna);
    let diasEnTransito = null;
    if (p.fechaDespacho && !p.esAnulado && p.yaDespachado) {
      if (info.entregado) {
        if (info.fechaFin) {
          diasEnTransito = Math.max(0, Math.round((new Date(info.fechaFin) - p.fechaDespacho) / 86400000));
        }
      } else {
        diasEnTransito = Math.max(0, Math.round((ahora - p.fechaDespacho) / 86400000));
      }
    }

    cacheData.push([
      p.pedido, p.courier || 'Sin courier', region, p.comuna || 'Sin comuna', info.estado,
      info.entregado ? 'SI' : 'NO',
      p.fechaDespacho || '', diasEnTransito, fuente, p.yaDespachado ? 'SI' : 'NO',
      alertaBodega ? 'SI' : 'NO', p.fechaPedidoIso || ''
    ]);
  }

  const cacheSheet = obtenerCacheSheet();
  if (cacheSheet.getLastRow() > 1) {
    cacheSheet.getRange(2, 1, cacheSheet.getLastRow() - 1, CACHE_COLS).clearContent();
  }
  if (cacheData.length) {
    cacheSheet.getRange(2, 1, cacheData.length, CACHE_COLS).setValues(cacheData);
  }

  PropertiesService.getScriptProperties().setProperty('ULTIMA_SYNC', ahora.toISOString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ── Alas en PARALELO con fetchAll (lotes de 50). Devuelve { pedido: orderData } ──
function batchAlas(pedidos) {
  const out = {};
  if (!pedidos.length) return out;
  const apiKey = getSecret('ALAS_API_KEY');
  const sender = getSecret('ALAS_SENDER');
  const CHUNK = 50;
  for (let i = 0; i < pedidos.length; i += CHUNK) {
    const chunk = pedidos.slice(i, i + CHUNK);
    const requests = chunk.map(function(ped) {
      return {
        url: ALAS_BASE_URL + '/delivery-orders/status',
        method: 'post', contentType: 'application/json',
        headers: { 'x-alas-ce0-api-key': apiKey },
        payload: JSON.stringify({ partner: ALAS_PARTNER, senderCode: sender, deliveryOrderCode: String(ped) }),
        muteHttpExceptions: true
      };
    });
    let responses;
    try { responses = UrlFetchApp.fetchAll(requests); } catch (e) { continue; }
    for (let j = 0; j < responses.length; j++) {
      try {
        if (responses[j].getResponseCode() !== 200) continue;
        const d = JSON.parse(responses[j].getContentText());
        if (d && d.status) out[chunk[j]] = d;
      } catch (e) {}
    }
  }
  return out;
}

// ── Starken en PARALELO con fetchAll (solo Etracking, la imagen no se usa en el
//    cache). Lotes de 50. Devuelve { pedido: ordenFlete } ──
function batchStarken(pedidos) {
  const out = {};
  if (!pedidos.length) return out;
  const apiKey = getSecret('STK_API_KEY');
  const rut    = getSecret('STK_RUT');
  const pass   = getSecret('STK_PASSWORD');
  const CHUNK = 50;
  for (let i = 0; i < pedidos.length; i += CHUNK) {
    const chunk = pedidos.slice(i, i + CHUNK);
    const requests = chunk.map(function(ped) {
      return {
        url: STK_ETRACKING_URL, method: 'post', contentType: 'application/json',
        headers: { 'api-key': apiKey, 'cli-rut': rut, 'password': pass },
        payload: JSON.stringify({ tracking: [{ numeroDocumento: String(ped), numeroOrdenFlete: '', tipoDocumento: STK_TIPO_DOC }], rutEmpresa: rut }),
        muteHttpExceptions: true
      };
    });
    let responses;
    try { responses = UrlFetchApp.fetchAll(requests); } catch (e) { continue; }
    for (let j = 0; j < responses.length; j++) {
      try {
        if (responses[j].getResponseCode() !== 200) continue;
        const d = JSON.parse(responses[j].getContentText());
        const lista = d.listaResumenRedestinacion && d.listaResumenRedestinacion.ordenFlete || [];
        const orden = lista[0];
        if (orden && orden.codigoSalida === 1) out[chunk[j]] = orden;
      } catch (e) {}
    }
  }
  return out;
}

// ── Blue Express en multi-ref (lotes de 40 referencias por llamada).
//    Devuelve { pedido: order } ──
function batchBlue(pedidos) {
  const out = {};
  if (!pedidos.length) return out;
  let token;
  try { token = obtenerTokenBlue(); } catch (e) { return out; }
  const account = encodeURIComponent(getSecret('BLUE_ACCOUNT'));
  const apiKey  = getSecret('BLUE_API_KEY');
  // La API de Blue pagina a 10 resultados por llamada. Para NO perder ninguno,
  // enviamos lotes de 10 referencias (así la página 1 los trae todos).
  const CHUNK = 10;
  for (let i = 0; i < pedidos.length; i += CHUNK) {
    const chunk = pedidos.slice(i, i + CHUNK);
    const url = BLUE_BASE_URL + '/search?accounts=' + account +
                '&pageSize=50' +
                '&references=' + encodeURIComponent(chunk.join(','));
    try {
      const resp = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + token, 'x-api-key': apiKey },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() !== 200) continue;
      const d = JSON.parse(resp.getContentText());
      (d.data || []).forEach(function(order) {
        const ref = (order.references && order.references[0]) || order.reference || '';
        const m = String(ref).match(/^(\d+)/); // toma el número de pedido del inicio
        if (m) out[m[1]] = order;
      });
    } catch (e) {}
  }
  return out;
}

// ── SimpliRoute/Global: prefetch de los últimos días UNA sola vez.
//    Devuelve { pedido: visita } con el estado más avanzado si hay duplicados ──
function prefetchSimpli() {
  const map = {};
  const PRIORIDAD = { 'completed': 5, 'failed': 4, 'partial': 3, 'canceled': 2, 'pending': 1 };
  const rank = function(v) { return PRIORIDAD[String(v.status || '').toLowerCase()] || 0; };
  const hoy = new Date();
  for (let i = 0; i < SIMPLI_DIAS_BUSQUEDA; i++) {
    const d = new Date(hoy.getTime() - i * 86400000);
    const fechaStr = Utilities.formatDate(d, 'America/Santiago', 'yyyy-MM-dd');
    const visitas  = obtenerVisitasSimpliPorFecha(fechaStr);
    visitas.forEach(function(v) {
      const m = String(v.reference || '').toUpperCase().match(/^(\d+)/);
      if (!m) return;
      const key = m[1];
      if (!map[key] || rank(v) > rank(map[key])) map[key] = v;
    });
  }
  return map;
}

// ── Instala el trigger de sincronización automática (ejecutar UNA VEZ manualmente) ──
function instalarTriggerSync() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sincronizarTracking') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizarTracking')
    .timeBased()
    .everyMinutes(30)
    .create();
}

// ── Datos crudos para el dashboard (el cliente calcula todos los agregados,
//    así puede recalcular en vivo al cambiar el filtro de fechas sin volver
//    a llamar al servidor) ──
function obtenerDashboardData() {
  const email = Session.getActiveUser().getEmail();
  if (!email || email.split('@')[1] !== DASHBOARD_DOMAIN) {
    throw new Error('Acceso no autorizado.');
  }

  const cacheSheet = obtenerCacheSheet();
  const data = cacheSheet.getDataRange().getValues();
  const rows = data.slice(1).filter(function(r) { return r[0]; });
  const ultimaSync = PropertiesService.getScriptProperties().getProperty('ULTIMA_SYNC');

  const resultado = {
    detalle: rows.slice(0, 1000).map(function(r) {
      return {
        pedido: r[0], courier: r[1] || 'Sin courier', region: r[2] || 'Sin clasificar', comuna: r[3] || 'Sin comuna',
        estado: r[4] || 'Desconocido', entregado: r[5] === 'SI',
        fechaDespacho: r[6] instanceof Date ? r[6].toISOString() : (r[6] || null),
        diasEnTransito: typeof r[7] === 'number' ? r[7] : null,
        fuente: r[8] || 'Manual', yaDespachado: r[9] === 'SI',
        alertaBodega: r[10] === 'SI',
        fechaPedido: r[11] instanceof Date ? r[11].toISOString() : (r[11] || null)
      };
    }),
    ultimaSync: ultimaSync
  };

  // Se serializa a texto explícitamente: google.script.run a veces falla
  // silenciosamente (devuelve null al cliente) con objetos grandes o con
  // tipos mixtos (Date vs string) dentro del mismo array.
  return JSON.stringify(resultado);
}
