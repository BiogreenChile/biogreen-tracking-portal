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

// ============================================
// FUNCIÓN PRINCIPAL
// ============================================
function doGet(e) {
  // Si viene con parámetro pedido → devuelve JSON
  if (e.parameter.pedido) {
    return handleRequest(e);
  }
  // Si no → sirve la página HTML
  return HtmlService.createHtmlOutputFromFile('Seguimiento de pedido')
    .setTitle('Seguimiento de Pedido · Biogreen')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return handleRequest(e);
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
