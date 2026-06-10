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

Cada certificación añade **6 columnas** a la derecha (etiqueta + 5 nombres columna). El bloque de la primera certificación empieza en la columna G:

| Col | Contenido | Tipo |
|-----|-----------|------|
| A | Código de partida | Fija (presupuesto) |
| B | Unidad (Ud, M, M2...) | Fija |
| C | Resumen / descripción | Fija |
| D | Cantidad presupuestada | Fija |
| E | Precio presupuestado | Fija |
| F | Importe presupuestado | Fija |
| G-L | **Certif 01** (6 cols): etiqueta + Can, Imp, CompCan, CompImp, % | Pendiente de formato azul |
| M-R | **Certif 02** (6 cols): etiqueta + Can, Imp, CompCan, CompImp, % | Pendiente de formato azul |
| S-X | **Certif 03** (6 cols): etiqueta + Can, Imp, CompCan, CompImp, % | Pendiente de formato azul |
| Y-AD | **Certif 04** (6 cols): etiqueta + Can, Imp, CompCan, CompImp, % | Pendiente de formato azul |
| AE | COMENTARIO | Acumulativa |
| AF | MEDIOS AUXILIARES | Acumulativa |

Las filas pueden ser:
- **Capítulos** (ej: `1`, `2`, `3`...) — solo contienen totales en columnas de importe.
- **Partidas** (ej: `1.1`, `1.2`, `2.1`...) — desglose con cantidades, precio e importe.

### Flujo de validación (pendiente de formato azul)

1. **Se genera el Excel** con los datos extraídos por DeepSeek. Las columnas de la nueva certificación se añaden a la derecha.
2. **El arquitecto revisa en obra** cada partida: verifica cantidades, precios y ejecución real.
3. **Al volver al estudio**, actualiza el Excel manualmente (pendiente de implementar formato azul automático).
4. Este proceso se repite cada mes con cada nueva certificación.

> ⚠️ El formato azul automático no está disponible actualmente — la librería `xlsx` Community Edition no soporta estilos. Pendiente migrar a `exceljs`.

## Implementación técnica

### Arquitectura

```
┌──────────────────────────────────────────────────────────────┐
│  Arquitecto (navegador)                                       │
│  http://localhost:3456                                        │
│  ┌────────────────┐   ┌──────────────────┐   ┌────────────┐ │
│  │  Subir PDF      │   │  Ver/Descargar   │   │  Estado    │ │
│  │  Presupuesto    │   │  Excel           │   │            │ │
│  └───────┬────────┘   └──────────────────┘   └────────────┘ │
└──────────┼───────────────────────────────────────────────────┘
           │ HTTP multipart upload
           ▼
┌──────────────────────────────────────────────────────────────┐
│  Proxy Node.js 20 (puerto 3456)                              │
│                                                              │
│  ┌──────────────┐   ┌─────────────┐   ┌──────────────────┐ │
│  │  /upload      │   │  /budget    │   │  /cert           │ │
│  │  (web UI)     │   │  ?binaryId  │   │  ?binaryId       │ │
│  └──────┬───────┘   └──────┬──────┘   └──────┬───────────┘ │
│         │                  │                  │              │
│         ▼                  ▼                  ▼              │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  1. Extraer texto (raw parsing o Tesseract OCR)      │   │
│  │  2. DeepSeek API (estructurar datos)                 │   │
│  │  3. Generar Excel (xlsx) ← o → Añadir columnas       │   │
│  └──────────────────────┬───────────────────────────────┘   │
│                         │                                    │
│                         ▼                                    │
│                  ┌──────────────┐   ┌─────────────┐         │
│                  │  ./output/   │   │ DeepSeek    │         │
│                  │  *.xlsx      │   │ API (cloud) │         │
│                  └──────────────┘   └─────────────┘         │
└──────────────────────────────────────────────────────────────┘
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
- **Interfaz web** en `http://localhost:3456` — subir PDFs, ver y descargar Excel generados
- Proxy server con endpoints `/budget` y `/cert`:
  - **`/upload/budget`**: sube PDF, extrae texto (OCR si escaneado), llama DeepSeek, crea Excel A-F
  - **`/upload/cert`**: sube PDF, extrae datos, añade columnas al Excel de presupuesto existente
- Detección automática de PDFs escaneados → OCR con Tesseract (español)
- n8n 2.25.6 con workflow alternativo para certificaciones
- Webhook en modo `onReceived`

### 🔧 Últimas mejoras
| Mejora | Descripción |
|---|---|
| **Parser multipart corregido** | El `endBoundary` se encontraba desde el inicio del buffer, consumiendo todas las partes en una sola. Ahora busca el siguiente delimitador con prefijo `\r\n`. |
| **Detección OCR mejorada** | `isLikelyScanned` comprueba si el PDF contiene objetos `BT`/`ET` (texto incrustado) antes de decidir aplicar OCR, evitando OCR innecesario en PDFs con streams comprimidos pero texto real. |
| **Emparejamiento difuso de partidas** | `matchPartida()` prueba coincidencia exacta primero, luego prefijo del código de presupuesto, luego prefijo inverso — para cuando los códigos de partida difieren ligeramente entre presupuesto y certificación. |
| **Layout de certificación** | Cada bloque de certificación ocupa 6 columnas: etiqueta `Certif XX` + 5 nombres de columna (Can, Imp, CompCan, CompImp, %). El dropdown de selección muestra todos los archivos (base y con certificaciones). |
| **Formato azul no disponible** | La librería `xlsx` (Community Edition) no soporta estilos de celda. Para formato azul habría que migrar a `exceljs`. |

### ⚠️ Limitaciones
| Limitación | Detalle |
|---|---|
| **OCR lento** | Para PDFs escaneados, la conversión pdftoppm + tesseract puede tardar varios minutos. |
| **HTTP Request node POST** | Bug en typeVersion 4.2: `jsonBody` con arrays los elimina. Solución: GET con query param. |
| **Code node sandbox** | No permite `http`, `https`, `fs`, `fetch`, `child_process`, `process.env`, `$env`, módulos npm externos. Solo `buffer`, `path`, `crypto`, `stream`, `url`, `util`, `zlib`. |
| **Execute Command node** | No disponible en esta versión de n8n. |
| **Sin formato azul** | La librería xlsx Community Edition no permite estilos. Pendiente migrar a exceljs. |

### 🔜 Próximos pasos
1. Migrar a `exceljs` para formato azul en columnas nuevas de certificación
2. Indicador de progreso en la interfaz web (barra de carga durante OCR/DeepSeek)
3. Mejorar manejo de errores (timeouts, reintentos DeepSeek)

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

### Interfaz web (recomendada)

Abre en el navegador: **http://localhost:3456**

Tres secciones simples:

1. **Presupuesto Inicial** — Sube el PDF del presupuesto (GEANSAR). El proxy extrae el texto (OCR si es necesario), llama a DeepSeek, y genera el Excel base con columnas A-F (Código, Ud, Resumen, CanPres, PrPres, ImpPres). Solo la primera vez.

2. **Certificación Mensual** — Sube el PDF de la certificación del mes. El proxy busca el Excel de presupuesto, extrae los datos, y **añade columnas nuevas** (Can, Imp, CompCan, CompImp, %) emparejando por código de partida.

3. **Archivos Generados** — Lista todos los Excel con enlaces de descarga.

La página se recarga automáticamente al terminar (1-3 min).

### API directa (curl)

```bash
# Presupuesto inicial
curl -F "file=@presupuesto.pdf" http://localhost:3456/upload/budget

# Certificación mensual
curl -F "file=@certificacion.pdf" http://localhost:3456/upload/cert
```

### Webhooks n8n (alternativa)

```bash
curl -X POST "http://localhost:5678/webhook/budget-084049" \
  -F "data=@presupuesto.pdf"
curl -X POST "http://localhost:5678/webhook/cert-080640" \
  -F "data=@certificacion.pdf"
```

## Notas técnicas

- La API key de DeepSeek se pasa al proxy via `DEEPSEEK_API_KEY` en docker-compose.
- El proxy se construye con Dockerfile (Tesseract OCR + xlsx pre-instalados).
- Al modificar `proxy-server/server.js`, reconstruir con `docker compose up -d --build proxy`.
- Los PDFs se procesan directamente en el proxy (sin pasar por n8n) cuando se usa la interfaz web.
- El endpoint `/cert` y `/upload/cert` buscan el Excel de presupuesto más reciente en `/output/` para añadir columnas.
- `executionOrder: "v1"` y `callerPolicy: "workflowsFromSameOwner"` requeridos en workflows n8n.
