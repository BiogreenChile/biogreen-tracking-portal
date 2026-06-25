# Biogreen — Portal de Seguimiento de Pedidos

Portal web (Google Apps Script + Google Sheets) para que los clientes de Biogreen Chile consulten el estado de su pedido en tiempo real, integrado con las APIs de **Alas Express** y **Blue Express**.

## Estructura

- `apps-script.js` → backend (Code.gs en el editor de Apps Script)
- `Seguimiento de pedido.html` → frontend servido por Apps Script

## Configuración

Este código **no contiene ninguna credencial**. Las claves se leen desde **Script Properties** del proyecto de Apps Script (`PropertiesService`), nunca desde el código fuente.

### 1. Crear el proyecto

1. En tu Google Sheet con los pedidos: **Extensiones → Apps Script**
2. Reemplaza el contenido de `Código.gs` con `apps-script.js`
3. Crea un archivo HTML llamado exactamente `Seguimiento de pedido` y pega el contenido de `Seguimiento de pedido.html`

### 2. Configurar credenciales (Script Properties)

En el editor de Apps Script: **⚙️ Configuración del proyecto → Propiedades del script → Añadir propiedad de script**

| Propiedad | Descripción |
|---|---|
| `ALAS_API_KEY` | API key entregada por Alas Express (header `x-alas-ce0-api-key`) |
| `ALAS_SENDER` | Código de remitente (sender code) asignado por Alas Express |
| `BLUE_CLIENT_ID` | Client ID OAuth2 de Blue Express (Tracking Pull Corp) |
| `BLUE_CLIENT_SECRET` | Client secret OAuth2 de Blue Express |
| `BLUE_API_KEY` | x-api-key de Blue Express |
| `BLUE_ACCOUNT` | Número de cuenta Blue Express (parámetro `accounts`) |

### 3. Permisos (`appsscript.json`)

Debe incluir:

```json
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.container.ui"
  ],
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE_ANONYMOUS"
  }
}
```

### 4. Desplegar

**Deploy → Manage deployments → New version → Deploy**

Cada cambio de código requiere una **nueva versión** del despliegue para que se refleje en la URL pública.

## Estructura de la hoja de cálculo

Columnas esperadas (hoja `Hoja 1`):

| Col | Campo |
|---|---|
| A | N° Pedido |
| B | Nombre |
| C | RUT |
| D | Fecha pedido |
| E | Categoría |
| F | Puntos |
| G | Importe |
| H | RUT Softland |
| I | RUT sin DV |
| J | Razón social |
| K | Comuna |
| L | Comuna Softland |
| M | Tipo (Boleta/Factura) |
| N | Importe (2) |
| O | Notas WMS (debe contener el nombre del courier: Alas, Bluexpress, etc.) |
| P | Forma de pago |
| Q | Estado pedido |

## Couriers soportados con tracking en línea

- **Alas Express** — vía API REST con `deliveryOrderCode`
- **Blue Express** — vía API REST (Tracking Pull Corp), búsqueda por `reference`

Otros couriers detectados en NOTAS WMS se muestran sin tracking en línea (mensaje genérico).
