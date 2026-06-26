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
const COURIERS_API = ['alas', 'bluexpress'];

// ── Dashboard interno ──
const CACHE_SHEET_NAME = 'Tracking Cache';
const DASHBOARD_DOMAIN = 'biogreenchile.com';

// ============================================
// FUNCIÓN PRINCIPAL
// ============================================
function doGet(e) {
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
  // Si no → sirve la página HTML (legado; el frontend público ahora vive en GitHub Pages)
  return HtmlService.createHtmlOutputFromFile('Seguimiento de pedido')
    .setTitle('Seguimiento de Pedido · Biogreen')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  } else {
    resultado = { ok: false, error: 'Courier no soportado: ' + courier };
  }

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

function consultarBlueExpress(pedido) {
  try {
    const token = obtenerTokenBlue();
    const url = BLUE_BASE_URL + '/search?accounts=' + encodeURIComponent(getSecret('BLUE_ACCOUNT')) +
                '&references=' + encodeURIComponent(String(pedido));

    const options = {
      method:  'get',
      headers: {
        'Authorization': 'Bearer ' + token,
        'x-api-key':     getSecret('BLUE_API_KEY')
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code !== 200) {
      return { ok: false, error: 'Blue Express HTTP ' + code };
    }

    const data = JSON.parse(response.getContentText());
    if (!data.data || !data.data.length) {
      return { ok: false, error: 'No encontrado en Blue Express' };
    }

    return { ok: true, data: data.data[0] };

  } catch(err) {
    return { ok: false, error: err.message };
  }
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
function obtenerCacheSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CACHE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CACHE_SHEET_NAME);
    sheet.getRange(1, 1, 1, 9).setValues([
      ['Pedido', 'Courier', 'Region', 'Comuna', 'Estado', 'Entregado', 'FechaDespacho', 'DiasEnTransito', 'Fuente']
    ]);
  }
  return sheet;
}

// ── Mapa Comuna → Región, usando la hoja "Matriz Alasxpress" (A=Region, C=Comuna) ──
const MATRIZ_SHEET_NAME = 'Matriz Alasxpress';
let _mapaComunaRegion = null;
function obtenerRegionPorComuna(comuna) {
  if (!_mapaComunaRegion) {
    _mapaComunaRegion = {};
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MATRIZ_SHEET_NAME);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const region = String(data[i][0] || '').trim();
        const com    = String(data[i][2] || '').trim().toLowerCase();
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
  const pkg    = (order.packages && order.packages[0]) || {};
  const latest = pkg.latestStatus || {};
  const estado = order.stateDesc || latest.statusCode || 'Desconocido';
  const entregado = latest.statusCode === 'DL';
  // statusDate = fecha del último evento; si está entregado, es la fecha de entrega
  const fechaFin = entregado && latest.statusDate ? latest.statusDate : null;
  return { estado: estado, entregado: entregado, fechaFin: fechaFin };
}

// ============================================
// SINCRONIZAR TRACKING (ejecutar vía trigger por tiempo)
// ============================================
// No tiene sentido seguir consultando pedidos antiguos ya resueltos.
const SYNC_DIAS_MAXIMO = 30; // ventana de pedidos a sincronizar (días desde la fecha del pedido)

function sincronizarTracking() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const cacheData = [];
  const ahora = new Date();
  const limiteFecha = new Date(ahora.getTime() - SYNC_DIAS_MAXIMO * 86400000);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const pedido = String(row[COL.pedido - 1] || '').trim();
    if (!pedido) continue;

    const estadoPedidoSheet = String(row[COL.estadoPedido - 1] || '').trim();
    const esAnulado = /anula/i.test(estadoPedidoSheet);

    // Omitir pedidos ya entregados (no hace falta seguir consultando su API)
    if (/entreg/i.test(estadoPedidoSheet)) continue;

    // Omitir pedidos fuera de la ventana de días configurada
    const fechaPedidoRaw = row[COL.fechaPedido - 1];
    const fechaPedidoObj = fechaPedidoRaw instanceof Date ? fechaPedidoRaw : new Date(fechaPedidoRaw);
    if (!isNaN(fechaPedidoObj.getTime()) && fechaPedidoObj < limiteFecha) continue;

    const notasWms = String(row[COL.notasWms - 1] || '');
    const courier  = detectarCourier(notasWms);
    const courierLower = (courier || '').toLowerCase();

    let info = null;
    let fuente = 'Manual';

    if (esAnulado) {
      // Pedido anulado: nunca se despachó, no aplica consultar ninguna API de courier
      info = { estado: estadoPedidoSheet, entregado: true, fechaFin: null };
      fuente = 'Manual';
    } else if (COURIERS_API.indexOf(courierLower) !== -1) {
      // Couriers con API: intenta tracking en vivo primero
      try {
        if (courierLower === 'alas') {
          const r = consultarAlas(pedido);
          if (r.ok) { info = extraerEstadoAlas(r.data); fuente = 'API'; }
        } else if (courierLower === 'bluexpress') {
          const r = consultarBlueExpress(pedido);
          if (r.ok) { info = extraerEstadoBlue(r.data); fuente = 'API'; }
        }
      } catch (e) {
        info = null;
      }
    }

    // Sin API o sin respuesta: usa el ESTADO PEDIDO de la hoja (Starken, Cacem, Mardam, etc.)
    if (!info) {
      if (!estadoPedidoSheet) continue; // nada que mostrar para este pedido
      info = {
        estado: estadoPedidoSheet,
        entregado: /entreg/i.test(estadoPedidoSheet),
        fechaFin: null
      };
      fuente = 'Manual';
    }

    const comuna = String(row[COL.comuna - 1] || '').trim();
    const region = obtenerRegionPorComuna(comuna);

    // Fecha de despacho: usamos la misma estimación interna que ve el cliente en el portal
    const fechaInfo    = parsearFecha(row[COL.fechaPedido - 1]);
    const despachoInfo = calcularDespacho(fechaInfo.dateObj);
    const fechaDespacho = despachoInfo.iso ? new Date(despachoInfo.iso) : null;

    // Los anulados nunca se despacharon: no corresponde contar días en tránsito
    let diasEnTransito = null;
    if (fechaDespacho && !esAnulado) {
      const fin = info.fechaFin ? new Date(info.fechaFin) : ahora;
      diasEnTransito = Math.max(0, Math.round((fin - fechaDespacho) / 86400000));
    }

    cacheData.push([
      pedido, courier || 'Sin courier', region, comuna || 'Sin comuna', info.estado,
      info.entregado ? 'SI' : 'NO',
      fechaDespacho || '', diasEnTransito, fuente
    ]);
  }

  const cacheSheet = obtenerCacheSheet();
  if (cacheSheet.getLastRow() > 1) {
    cacheSheet.getRange(2, 1, cacheSheet.getLastRow() - 1, 9).clearContent();
  }
  if (cacheData.length) {
    cacheSheet.getRange(2, 1, cacheData.length, 9).setValues(cacheData);
  }

  PropertiesService.getScriptProperties().setProperty('ULTIMA_SYNC', ahora.toISOString());
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

// ── Datos agregados para el dashboard (llamado desde el HTML) ──
function obtenerDashboardData() {
  const email = Session.getActiveUser().getEmail();
  if (!email || email.split('@')[1] !== DASHBOARD_DOMAIN) {
    throw new Error('Acceso no autorizado.');
  }

  const cacheSheet = obtenerCacheSheet();
  const data = cacheSheet.getDataRange().getValues();
  const rows = data.slice(1).filter(function(r) { return r[0]; });

  const porCourier = {};       // courier → {estado: cantidad}
  const porRegion   = {};      // region → {totalEnTransito, sumaDias, comunas:{comuna:{totalEnTransito, sumaDias}}}
  let conTrackingApi = 0;

  rows.forEach(function(r) {
    const courier  = r[1] || 'Sin courier';
    const region   = r[2] || 'Sin clasificar';
    const comuna   = r[3] || 'Sin comuna';
    const estado   = r[4] || 'Desconocido';
    const entregado = r[5] === 'SI';
    const dias     = typeof r[7] === 'number' ? r[7] : null;
    const fuente   = r[8] || 'Manual';

    if (fuente === 'API') conTrackingApi++;

    if (!porCourier[courier]) porCourier[courier] = {};
    porCourier[courier][estado] = (porCourier[courier][estado] || 0) + 1;

    if (!entregado && dias !== null) {
      if (!porRegion[region]) porRegion[region] = { totalEnTransito: 0, sumaDias: 0, comunas: {} };
      porRegion[region].totalEnTransito++;
      porRegion[region].sumaDias += dias;

      if (!porRegion[region].comunas[comuna]) porRegion[region].comunas[comuna] = { totalEnTransito: 0, sumaDias: 0 };
      porRegion[region].comunas[comuna].totalEnTransito++;
      porRegion[region].comunas[comuna].sumaDias += dias;
    }
  });

  // calcular promedios
  Object.keys(porRegion).forEach(function(region) {
    const r = porRegion[region];
    r.promedioDias = r.totalEnTransito ? +(r.sumaDias / r.totalEnTransito).toFixed(1) : 0;
    Object.keys(r.comunas).forEach(function(comuna) {
      const c = r.comunas[comuna];
      c.promedioDias = c.totalEnTransito ? +(c.sumaDias / c.totalEnTransito).toFixed(1) : 0;
    });
  });

  const ultimaSync = PropertiesService.getScriptProperties().getProperty('ULTIMA_SYNC');

  const resultado = {
    totalActivos: rows.length,
    enTransito: rows.filter(function(r) { return r[5] !== 'SI'; }).length,
    conTrackingApi: conTrackingApi,
    porCourier: porCourier,
    porRegion:  porRegion,
    detalle: rows.slice(0, 500).map(function(r) {
      return {
        pedido: r[0], courier: r[1], region: r[2], comuna: r[3],
        estado: r[4], entregado: r[5] === 'SI',
        fechaDespacho: r[6] instanceof Date ? r[6].toISOString() : (r[6] || null),
        diasEnTransito: r[7], fuente: r[8] || 'Manual'
      };
    }),
    ultimaSync: ultimaSync
  };

  // Se serializa a texto explícitamente: google.script.run a veces falla
  // silenciosamente (devuelve null al cliente) con objetos grandes o con
  // tipos mixtos (Date vs string) dentro del mismo array.
  return JSON.stringify(resultado);
}
