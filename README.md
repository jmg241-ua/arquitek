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
├── docker-compose.yml              # n8n (5678) + proxy (3456)
├── .env                            # N8N_HOST, DEEPSEEK_API_KEY
├── proxy-server/
│   ├── server.js                   # Proxy PDF→texto→DeepSeek→Excel
│   └── package.json                # xlsx
├── workflows/
│   ├── 01-importar-contrato.json   # Workflow contrato (pendiente)
│   └── certificacion-pdf-a-excel.json
├── ejemplo-real/                   # PDFs y Excel de ejemplo
│   ├── GEANSAR DOC-20230525-WA0005..pdf          # Presupuesto inicial
│   ├── CERTIFICACION Nº4 LIQUIDACION NUEVA AVENIDA.pdf  # Certif mes
│   └── NUEVA AVENIDA COMPROBACION CERT 04 v02.xlsx      # Excel esperado
├── test-envio.sh
├── output/                         # Excel generados (volumen Docker)
└── README.md
```

### Workflow n8n (3 nodos)

```
[Webhook] ──→ [Code: Preparar] ──→ [HTTP Request: Proxy]
```

1. **Webhook** — Recibe el PDF vía `POST /webhook/cert-XXXXXX` con `multipart/form-data`. Responde `{"message":"Workflow was started"}` inmediatamente.
2. **Code** — Extrae el `binaryId` del archivo recibido y construye la URL del proxy: `http://arquitek-proxy:3456/?binaryId=filesystem-v2:...`
3. **HTTP Request** — Hace GET al proxy. El proxy (fuera del sandbox de n8n) lee el PDF del disco compartido, llama a DeepSeek, genera el Excel y lo guarda en `/output`.

### Proxy server (`proxy-server/server.js`)

Servicio Node.js 20 que corre como contenedor separado. Recibe un `binaryId` vía query param y:

1. **Lee el archivo binario** del volumen compartido `n8n_data` (montado en `/n8n-storage`)
   - `binaryId: "filesystem-v2:workflows/.../binary_data/..."`
   - Path real: `/n8n-storage/storage/workflows/.../binary_data/...`
2. **Extrae texto** del PDF mediante raw buffer parsing (expresiones regulares sobre contenido `(...)` en el PDF)
   - ⚠️ Solo funciona con PDFs de texto incrustado, no escaneados/comprimidos
3. **Llama a DeepSeek API** vía `https` (módulo Node estándar) con el texto extraído
   - Prompt: extraer datos estructurados de la certificación
   - `response_format: {type: "json_object"}`
4. **Genera Excel** con la librería `xlsx` y lo guarda en `/output/certificacion_*.xlsx`
5. **Responde** a n8n con `{ success, textLength, aiResponse, excelFile }`

## Estado actual

### ✅ Funcionando
- n8n 2.25.6 en Docker con SQLite
- Proxy server operativo: lee PDFs, llama DeepSeek, genera Excel
- Workflow **Certificacion PDF a Excel**: 3 nodos ejecutándose correctamente
- Excel se genera en `./output/`
- Webhook en modo `onReceived` (respuesta inmediata)

### ⚠️ Limitaciones
| Limitación | Detalle |
|---|---|
| **PDF escaneados** | El raw buffer parsing solo extrae texto de PDFs con texto incrustado. PDFs escaneados/comprimidos (como GEANSAR) no se pueden procesar sin OCR. |
| **HTTP Request node POST** | Bug en typeVersion 4.2: `jsonBody` con arrays los elimina. Solución: GET con query param. |
| **Code node sandbox** | No permite `http`, `https`, `fs`, `fetch`, `child_process`, `process.env`, `$env`, módulos npm externos. Solo `buffer`, `path`, `crypto`, `stream`, `url`, `util`, `zlib`. |
| **Execute Command node** | No disponible en esta versión de n8n. |
| **Proxy: flujo completo** | Pendiente: el proxy actual crea Excel desde cero. Falta implementar: leer Excel existente + añadir columnas de nueva certificación. |

### 🔜 Próximos pasos
1. Actualizar el proxy para que reciba el Excel existente (binaryId o file path) junto con el nuevo PDF de certificación
2. Leer el Excel maestro, añadir columnas de la nueva certificación (en azul)
3. Guardar el Excel actualizado
4. Workflow de contrato (presupuesto inicial desde PDF)

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
