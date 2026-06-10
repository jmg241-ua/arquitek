# Arquitek — Automatización de Certificaciones de Obra con n8n

Sistema para un arquitecto que automatiza la gestión de certificaciones mensuales de obra usando **n8n** + **DeepSeek API** + **Docker**.

## Flujo de negocio

```
MES 1: Presupuesto (PDF) ──→ Excel base con columnas fijas (A-F)
                                      │
MES 2: Certificación 1 (PDF) ────────→ + columnas Certif 01 (G-L) en azul
                                      │
MES 3: Certificación 2 (PDF) ────────→ + columnas Certif 02 (N-R) en azul
                                      │
...                                    → el arquitecto valida en obra
                                         y cambia azul → negro
```

### Estructura del Excel

El Excel maestro tiene esta estructura, donde cada certificación añade columnas a la derecha:

| Col | Contenido | Tipo |
|-----|-----------|------|
| A | Código de partida | Fija (presupuesto) |
| B | Unidad (Ud, M, M2...) | Fija |
| C | Resumen / descripción | Fija |
| D | Cantidad presupuestada | Fija |
| E | Precio presupuestado | Fija |
| F | Importe presupuestado | Fija |
| G-L | **Certif 01**: Can, Imp, Comprobación Can/Imp, %, exceso | Azul (nueva) |
| N-R | **Certif 02**: Can, Imp, Comprobación Can/Imp, % | Azul (nueva) |
| U-Y | **Certif 03**: Can, Imp, Comprobación Can/Imp, % | Azul (nueva) |
| AA-AE | **Certif 04**: Can, Imp, Comprobación Can/Imp, % | Azul (nueva) |
| AG | COMENTARIO | Acumulativa |
| AH | MEDIOS AUXILIARES | Acumulativa |

Las filas pueden ser:
- **Capítulos** (ej: `1`, `2`, `3`...) — solo contienen totales en columnas de importe.
- **Partidas** (ej: `1.1`, `1.2`, `2.1`...) — desglose con cantidades, precio e importe.

### Flujo de validación (azul → negro)

1. **Se genera el Excel** con los datos extraídos por DeepSeek. Las columnas de la nueva certificación aparecen en **azul**.
2. **El arquitecto revisa en obra** cada partida: verifica cantidades, precios y ejecución real.
3. **Al volver al estudio**, cambia el formato de **azul a negro** en las celdas ya validadas.
4. Este proceso se repite cada mes con cada nueva certificación.

> El azul indica "pendiente de validar en obra" — no implica error, solo que aún no se ha confirmado presencialmente.

## Implementación técnica

### Arquitectura

```
                    docker compose
┌──────────────────────────────────────────────────────────┐
│                                                           │
│  ┌──────────────┐    HTTP GET     ┌──────────────────┐   │
│  │    n8n       │ ──────────────> │    Proxy         │   │
│  │  (5678)      │   ?binaryId=..  │  Node.js 20      │   │
│  │              │ <────────────── │  (3456)          │   │
│  │  Webhook →   │     JSON resp   │                  │   │
│  │  Code →      │                 │  1. Lee PDF      │   │
│  │  HTTP Req    │                 │  2. Extrae texto │   │
│  └──────┬───────┘                 │  3. DeepSeek API │   │
│         │                         │  4. Genera Excel │   │
│         │ n8n_data volume         │  5. Guarda /output│  │
│         ▼                         └────────┬─────────┘   │
│  ┌──────────────┐                          │             │
│  │  n8n storage  │   ┌─────────────┐       │ ./output    │
│  │  (SQLite +    │   │ DeepSeek    │       ▼             │
│  │   binary)     │   │ API (cloud) │  ┌──────────────┐  │
│  └──────────────┘   └─────────────┘  │  ./output/    │  │
│                                        │  certificacion│  │
│                                        │  _*.xlsx      │  │
│                                        └──────────────┘  │
└──────────────────────────────────────────────────────────┘
```

```
├── docker-compose.yml              # n8n (5678) + proxy (3456) con Dockerfile
├── .env                            # N8N_HOST, DEEPSEEK_API_KEY
├── proxy-server/
│   ├── Dockerfile                  # Construye imagen con Tesseract OCR + xlsx
│   ├── server.js                   # Proxy con endpoints /budget y /cert
│   └── package.json                # xlsx
├── workflows/
│   ├── 01-importar-contrato.json   # Workflow contrato (pendiente)
│   └── certificacion-pdf-a-excel.json  # Workflow certificación activo
├── ejemplo-real/                   # PDFs y Excel de ejemplo
│   ├── GEANSAR DOC-20230525-WA0005..pdf          # Presupuesto inicial (escaneado)
│   ├── CERTIFICACION Nº4 LIQUIDACION NUEVA AVENIDA.pdf  # Certificación mensual
│   └── NUEVA AVENIDA COMPROBACION CERT 04 v02.xlsx      # Excel esperado
├── test-envio.sh
├── output/                         # Excel generados (volumen Docker)
└── README.md
```

### Workflow n8n (3 nodos)

```
[Webhook] ──→ [Code: Preparar] ──→ [HTTP Request: Proxy /cert]
```

1. **Webhook** — Recibe el PDF vía `POST /webhook/cert-XXXXXX` con `multipart/form-data`. Responde `{"message":"Workflow was started"}` inmediatamente.
2. **Code** — Extrae el `binaryId` del archivo recibido y construye la URL del proxy: `http://arquitek-proxy:3456/cert?binaryId=filesystem-v2:...`
3. **HTTP Request** — Hace GET al proxy `/cert`. El proxy (fuera del sandbox de n8n) lee el PDF del disco compartido, llama a DeepSeek, añade columnas al Excel del presupuesto y guarda en `/output`.

### Proxy server (`proxy-server/server.js`)

Servicio Node.js 20 que corre como contenedor separado. Expone dos endpoints que reciben un `binaryId` vía query param:

#### `GET /budget?binaryId=...`
Procesa un PDF de **presupuesto** (el punto de partida de la obra):
1. Lee el archivo binario del volumen compartido `n8n_data`
2. Extrae texto del PDF — dos métodos: raw parsing para PDFs con texto incrustado, OCR con `pdftoppm` + `tesseract` (español) para PDFs escaneados
3. Llama a DeepSeek con un prompt específico para extraer todas las partidas del presupuesto
4. Genera un Excel con columnas **A-F**: Código, Ud, Resumen, CanPres, PrPres, ImpPres
5. Responde `{ success, numPartidas, excelFile }`

#### `GET /cert?binaryId=...`
Procesa un PDF de **certificación mensual** y añade columnas al Excel del presupuesto existente:
1. Lee el PDF, extrae texto (igual que /budget)
2. Llama a DeepSeek con prompt específico para extraer datos de certificación y el array `partidas` con `{ codigo, can, imp }`
3. Busca el Excel de presupuesto más reciente en `/output/`
4. Añade columnas de certificación (Can, Imp, CompCan, CompImp, %) emparejando por código de partida
5. Guarda el Excel actualizado
6. Responde `{ success, certLabel, numPartidas, excelFile, appendedTo }`

## Estado actual

### ✅ Funcionando
- n8n 2.25.6 en Docker con SQLite
- Proxy server con dos endpoints:
  - **`/budget`**: extrae presupuesto desde PDF (con OCR), crea Excel base con columnas A-F
  - **`/cert`**: extrae certificación desde PDF, añade columnas al Excel del presupuesto existente
- Detección automática de PDFs escaneados → OCR con Tesseract (español)
- Workflow **Certificacion PDF a Excel**: 3 nodos ejecutándose correctamente
- Webhook en modo `onReceived` (respuesta inmediata)

### ⚠️ Limitaciones
| Limitación | Detalle |
|---|---|
| **OCR lento** | Para PDFs escaneados, la conversión pdftoppm + tesseract puede tardar varios minutos. |
| **HTTP Request node POST** | Bug en typeVersion 4.2: `jsonBody` con arrays los elimina. Solución: GET con query param. |
| **Code node sandbox** | No permite `http`, `https`, `fs`, `fetch`, `child_process`, `process.env`, `$env`, módulos npm externos. Solo `buffer`, `path`, `crypto`, `stream`, `url`, `util`, `zlib`. |
| **Execute Command node** | No disponible en esta versión de n8n. |

### 🔜 Próximos pasos
1. Workflow de n8n para presupuesto inicial (Webhook → Code → `/budget`)
2. Añadir formato azul a las columnas nuevas de certificación en el Excel
3. Mejorar el emparejamiento de partidas entre certificación y presupuesto (actualmente por código exacto)
4. Mejorar el header del Excel a dos filas (grupo de certificación + nombre de columna)

## Requisitos

- Docker y Docker Compose
- API Key de DeepSeek (https://platform.deepseek.com/api_keys)

## Configuración

1. Copia `.env.example` a `.env` y edita:
   ```bash
   cp .env.example .env
   ```
2. Asegúrate de que `DEEPSEEK_API_KEY` tenga tu clave válida.

## Uso

### Iniciar servicios
```bash
docker compose up -d
```
Accede a `http://localhost:5678` para la interfaz de n8n.

### Enviar una certificación
```bash
curl -X POST "http://localhost:5678/webhook/cert-080640" \
  -F "data=@ejemplo-real/CERTIFICACION Nº4 LIQUIDACION  NUEVA AVENIDA.pdf"
```
El Excel aparece en `./output/certificacion_*.xlsx` tras ~60s.

### Script de prueba
```bash
./test-envio.sh ejemplo-real/CERTIFICACION_N4.pdf
```

## Notas técnicas

- `executionOrder: "v1"` y `callerPolicy: "workflowsFromSameOwner"` requeridos en settings del workflow.
- El proxy usa GET para evitar el bug de arrays en HTTP Request node.
- Binary file path mapping: `filesystem-v2:workflows/...` → `/n8n-storage/storage/workflows/...`.
- La API key de DeepSeek se pasa al proxy via `DEEPSEEK_API_KEY` env var en docker-compose.
- El proxy se reconstruye automáticamente al cambiar `proxy-server/server.js` (volume mount + `npm install` en cada inicio).
