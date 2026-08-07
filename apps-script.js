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
const LOG_SHEET_NAME = 'Log Consultas';
const LOG_MAX_FILAS  = 5000; // límite: mantiene los últimos N registros

// Registra una consulta en la hoja "Log Consultas". Se llama de forma
// asíncrona conceptual: si falla, NO rompe la consulta principal (try/catch).
// `info` es opcional: { nombre, comuna } — para consultas de pedidos.
function registrarLog(tipo, pedido, courier, ok, extra, info) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const HEADERS = ['Timestamp', 'Tipo', 'Pedido', 'Nombre', 'Comuna', 'Courier', 'Resultado', 'Detalle'];
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET_NAME);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    } else {
      // Migrar header si venimos de la versión de 6 columnas
      const headerActual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headerActual.length < HEADERS.length || headerActual[3] !== 'Nombre') {
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      }
    }
    info = info || {};
    sheet.appendRow([
      new Date(),
      tipo || '',
      String(pedido || ''),
      info.nombre || '',
      info.comuna || '',
      courier || '',
      ok ? 'OK' : 'ERROR',
      String(extra || '').substring(0, 300)
    ]);
    // Truncar si supera el límite (deja el header + últimas LOG_MAX_FILAS)
    const total = sheet.getLastRow();
    if (total > LOG_MAX_FILAS + 1) {
      sheet.deleteRows(2, total - LOG_MAX_FILAS - 1);
    }
  } catch (e) {
    // No propagar el error — no queremos que el log rompa la consulta
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
// GET /exec?courier=alas&codigo=97219
// ============================================
function handleCourierRequest(e) {
  const courier = String(e.parameter.courier || '').toLowerCase();
  const codigo  = String(e.parameter.codigo || '').trim();

  if (!codigo) {
    return jsonOut({ ok: false, error: 'Falta el código de pedido.' });
  }

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
    resultado = { ok: false, error: 'Courier no soportado: ' + courier };
  }

  registrarLog('tracking', codigo, courier, resultado.ok, resultado.ok ? '' : (resultado.error || ''));
  return jsonOut(resultado);
}

function handleRequest(e) {
  try {
    const pedido = (e.parameter.pedido || '').toString().trim();

    if (!pedido) {
      return jsonOut({ ok: false, error: 'Falta el número de pedido.' });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonOut({ ok: false, error: 'Hoja no encontrada: ' + SHEET_NAME });
    }

    const data = sheet.getDataRange().getValues();

    // Buscar pedido (fila 0 = encabezados)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nPedido = String(row[COL.pedido - 1]).trim();

      if (nPedido === pedido) {
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
        registrarLog('consulta', pedido, courier || '', true, '', { nombre: nombreLog, comuna: comunaLog });

        return jsonOut({
          ok:           true,
          pedido:       nPedido,
          nombre:       toTitleCase(String(row[COL.nombre - 1] || '')),
          categoria:    String(row[COL.categoria - 1] || ''),
          puntos:       row[COL.puntos - 1] || 0,
          importe:      row[COL.importe - 1] || 0,
          tipo:         String(row[COL.tipo - 1] || ''),        // Boleta / Factura
          razonSocial:  String(row[COL.razonSocial - 1] || ''),
          comuna:       toTitleCase(String(row[COL.comuna - 1] || '')),
          formaPago:    String(row[COL.formaPago - 1] || ''),
          estadoPedido: String(row[COL.estadoPedido - 1] || ''),
          courier:      courier,
          fechaPedido:  fechaInfo.texto,
          fechaObj:     fechaInfo.iso,
          despachoEstimado: despachoInfo.texto,
          despachoISO:      despachoInfo.iso,
          antesDeDoce:      despachoInfo.antesDeDoce,
        });
      }
    }

    // No encontrado
    registrarLog('consulta', pedido, '', false, 'No encontrado');
    return jsonOut({ ok: false, error: 'No encontramos ese número de pedido. Verifica e intenta nuevamente.' });

  } catch (err) {
    return jsonOut({ ok: false, error: 'Error interno: ' + err.message });
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
    hour: '2-digit', minute: '2-digit'
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
// ============================================
function consultarPedido(pedido) {
  var fakeEvent = { parameter: { pedido: pedido } };
  var output = handleRequest(fakeEvent);
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
        // referencia = "{pedido}BIOGREEN"; también aceptamos coincidencia por prefijo
        return ref === (objetivo + 'BIOGREEN') || ref.indexOf(objetivo) === 0;
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

// ── Extrae estado normalizado desde la respuesta de Blue Express (sin asumir atraso) ──
function extraerEstadoBlue(order) {
  if (!order) return null;
  const pkg = (order.packages && order.packages[0]) || {};
  // latestStatus puede venir a NIVEL DE ORDEN (típico en entregados) o dentro del paquete
  const latest = order.latestStatus || pkg.latestStatus || {};
  const code   = latest.statusCode || '';
  const entregado = code === 'DL';
  // Mapa breve de códigos → texto legible; si no está, usa stateDesc
  const MAPA = {
    'DL': 'Entregado', 'PU': 'Retirado', 'SOB': 'En bodega', 'IC': 'En camino',
    'AS': 'En camino', 'DA': 'En sucursal', 'LD': 'En reparto', 'DR': 'Devuelto al remitente',
    'RD': 'Rechazado', 'NH': 'Nadie en casa'
  };
  const estado = MAPA[code] || order.stateDesc || code || 'Desconocido';
  const fechaFin = entregado && latest.statusDate ? latest.statusDate : null;
  return { estado: estado, entregado: entregado, fechaFin: fechaFin };
}

// ============================================
// SINCRONIZAR TRACKING (ejecutar vía trigger por tiempo)
// ============================================
// No tiene sentido seguir consultando pedidos antiguos ya resueltos.
const SYNC_DIAS_MAXIMO = 30; // ventana de pedidos a sincronizar (días desde la fecha del pedido)

// Optimizado para no exceder el límite de 6 min de Apps Script con >1000 filas:
//  - Un solo recorrido para recolectar pedidos activos
//  - Alas y Starken en PARALELO con UrlFetchApp.fetchAll() (por lotes)
//  - Blue en multi-ref (varias referencias por llamada)
//  - SimpliRoute/Global: prefetch de los últimos días UNA sola vez (mapa en memoria)
function sincronizarTracking() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;

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

    const esCourierApi = COURIERS_API.indexOf(p.courierLower) !== -1;
    if (!info && esCourierApi && p.yaDespachado) {
      // Ya debía estar despachado y el courier no lo tiene → posible atraso de bodega
      alertaBodega = true;
      info = { estado: 'Sin guía en courier (posible atraso de bodega)', entregado: false, fechaFin: null };
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
  // La API de Blue pagina a 10 resultados por llamada, así que pedimos pageSize
  // suficiente y limitamos el lote para no perder resultados.
  const CHUNK = 25;
  for (let i = 0; i < pedidos.length; i += CHUNK) {
    const chunk = pedidos.slice(i, i + CHUNK);
    const url = BLUE_BASE_URL + '/search?accounts=' + account +
                '&pageSize=' + CHUNK +
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
