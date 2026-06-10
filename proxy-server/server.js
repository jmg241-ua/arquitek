const http = require('http');
const fs = require('fs');
const https = require('https');
const path = require('path');
const url = require('url');
const { execSync } = require('child_process');
const XLSX = require('xlsx');

const PORT = process.env.PORT || 3456;
const N8N_STORAGE = process.env.N8N_STORAGE || '/home/node/.n8n/storage';
const DEEPSEEK_API_KEY = process.env.DEEPSEEK_API_KEY || '';
const OUTPUT_DIR = '/output';
const TMP_DIR = '/tmp/proxy-uploads';

if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR, { recursive: true });

// ─── PDF text extraction ──────────────────────────────────────────────────
function extractPdfText(pdfBuffer) {
  const raw = pdfBuffer.toString('binary');
  const textMatches = raw.match(/\(([^)]*)\)/g) || [];
  let extractedText = textMatches
    .map(t => t.slice(1, -1))
    .filter(t => t.length > 1 && !t.match(/^[\d\s.,]+$/))
    .join('\n');
  const btEtMatches = raw.match(/BT[\s\S]*?ET/g) || [];
  const btEtText = btEtMatches.map(block => {
    const tMatches = block.match(/\(([^)]*)\)/g) || [];
    return tMatches.map(t => t.slice(1, -1)).join(' ');
  }).join('\n');
  return (extractedText + '\n' + btEtText).trim();
}

function isLikelyScanned(text) {
  if (text.length < 100) return true;
  const alpha = (text.match(/[a-zA-ZáéíóúñüÁÉÍÓÚÑÜ]/g) || []).length;
  if (alpha / text.length < 0.4) return true;
  const lines = text.split('\n').filter(l => l.trim().length > 5);
  if (lines.length < 5) return true;
  const words = text.split(/\s+/).filter(w => w.length > 0);
  if (words.length === 0) return true;
  const avgWordLen = words.reduce((s, w) => s + w.length, 0) / words.length;
  return avgWordLen > 20;
}

function extractTextViaOCR(pdfPath) {
  const tmpDir = '/tmp/ocr-' + Date.now();
  fs.mkdirSync(tmpDir, { recursive: true });
  try {
    execSync(`pdftoppm -png -r 300 "${pdfPath}" "${tmpDir}/page"`, { timeout: 120000 });
    const pages = fs.readdirSync(tmpDir).filter(f => f.startsWith('page')).sort();
    if (pages.length === 0) return '';
    let fullText = '';
    for (const page of pages) {
      const pagePath = path.join(tmpDir, page);
      const baseName = pagePath.replace(/\.\w+$/, '');
      try {
        execSync(`tesseract "${pagePath}" "${baseName}" -l spa --psm 6`, { timeout: 60000 });
        const txtPath = baseName + '.txt';
        if (fs.existsSync(txtPath)) {
          fullText += fs.readFileSync(txtPath, 'utf-8') + '\n';
        }
      } catch (e) {
        console.error('OCR error on page', page, e.message);
      }
    }
    return fullText.trim();
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true }); } catch (e) {}
  }
}

function readPdfFromBuffer(buffer) {
  let text = extractPdfText(buffer);
  if (isLikelyScanned(text)) {
    const tmpPath = path.join(TMP_DIR, 'pdf-' + Date.now() + '.pdf');
    fs.writeFileSync(tmpPath, buffer);
    console.log('PDF parece escaneado, aplicando OCR...');
    text = extractTextViaOCR(tmpPath) || text;
    try { fs.unlinkSync(tmpPath); } catch (e) {}
  }
  return text;
}

function readPdfFromFile(storagePath) {
  return readPdfFromBuffer(fs.readFileSync(storagePath));
}

// ─── DeepSeek API ─────────────────────────────────────────────────────────
function callDeepSeek(messages) {
  return new Promise((resolve) => {
    const body = JSON.stringify({
      model: "deepseek-chat",
      messages: messages,
      response_format: { type: "json_object" },
      temperature: 0.1
    });
    const options = {
      hostname: 'api.deepseek.com',
      path: '/v1/chat/completions',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + DEEPSEEK_API_KEY,
      }
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          const content = parsed.choices?.[0]?.message?.content || '{}';
          resolve({ success: true, aiResponse: content });
        } catch (e) {
          resolve({ success: false, error: e.message, raw: data.substring(0, 500) });
        }
      });
    });
    req.on('error', (e) => resolve({ success: false, error: e.message }));
    req.write(body);
    req.end();
  });
}

function saveExcel(wb, prefix) {
  const ts = new Date().toISOString().replace(/[:.]/g, '-');
  const fileName = prefix + '_' + ts + '.xlsx';
  XLSX.writeFile(wb, path.join(OUTPUT_DIR, fileName));
  return fileName;
}

// ─── /budget logic ─────────────────────────────────────────────────────
async function processBudget(text) {
  const result = await callDeepSeek([
    {
      role: "system",
      content: "Eres un asistente que extrae presupuestos de obra. Devuelve SOLO JSON valido sin markdown."
    },
    {
      role: "user",
      content: "Extrae este presupuesto de obra en formato JSON. Debe incluir:\n" +
        "- datos generales (numero_presupuesto, fecha, contratista, obra, total)\n" +
        "- un array 'capitulos' donde cada capitulo tiene: codigo, nombre, total\n" +
        "- un array 'partidas' donde cada partida tiene: codigo, ud, resumen, canPres, prPres, impPres\n\n" +
        "Ejemplo partida: {\"codigo\":\"1.1\",\"ud\":\"UD\",\"resumen\":\"transporte y retirada\",\"canPres\":1,\"prPres\":990,\"impPres\":990}\n\n" +
        "Texto:\n" + text.substring(0, 50000)
    }
  ]);

  if (!result.success) return result;
  let data;
  try {
    data = JSON.parse(typeof result.aiResponse === 'string' ? result.aiResponse : JSON.stringify(result.aiResponse));
  } catch (e) {
    return { ...result, error: 'Error al parsear respuesta: ' + e.message };
  }

  const partidas = Array.isArray(data.partidas) ? data.partidas : [];
  const wb = XLSX.utils.book_new();
  const headers = ['Código', 'Ud', 'Resumen', 'CanPres', 'PrPres', 'ImpPres'];
  const wsData = [headers];
  for (const p of partidas) {
    wsData.push([String(p.codigo ?? ''), String(p.ud ?? ''), String(p.resumen ?? ''), p.canPres ?? 0, p.prPres ?? 0, p.impPres ?? 0]);
  }
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{ wch: 10 }, { wch: 8 }, { wch: 55 }, { wch: 12 }, { wch: 12 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Presupuesto');
  const fileName = saveExcel(wb, 'presupuesto');

  return {
    success: true,
    tipo: 'Presupuesto',
    texto_extraido: text.length + ' caracteres',
    num_partidas: partidas.length,
    total: data.total || data.total_presupuesto || null,
    archivo: fileName
  };
}

// ─── /cert logic ───────────────────────────────────────────────────────
function findLatestExcel() {
  if (!fs.existsSync(OUTPUT_DIR)) return null;
  return fs.readdirSync(OUTPUT_DIR)
    .filter(f => f.startsWith('presupuesto_') && f.endsWith('.xlsx'))
    .sort().reverse()[0] || null;
}

function getNextCertColumns(existingPath) {
  if (!existingPath) return { startCol: 7, label: 'Certif 01' };
  const wb = XLSX.readFile(existingPath);
  const ws = wb.Sheets['Presupuesto'] || wb.Sheets[wb.SheetNames[0]];
  if (!ws || !ws['!ref']) return { startCol: 7, label: 'Certif 01' };
  const range = XLSX.utils.decode_range(ws['!ref']);
  let certCount = 0;
  for (let c = 6; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c: c })];
    if (cell && String(cell.v).startsWith('Certif ')) certCount++;
  }
  const next = certCount + 1;
  const startCol = 6 + (next - 1) * 5 + (next > 1 ? 1 : 0);
  return { startCol, label: 'Certif ' + String(next).padStart(2, '0') };
}

async function processCert(text) {
  const result = await callDeepSeek([
    {
      role: "system",
      content: "Eres un asistente que extrae datos de certificaciones de obra en formato JSON. Devuelve SOLO JSON valido sin markdown."
    },
    {
      role: "user",
      content: "Extrae los datos de esta certificacion de obra en formato JSON.\n" +
        "Incluye: numero_certificacion, fecha, contratista, obra, importe_certificado, importe_acumulado.\n" +
        "Ademas incluye un array 'partidas' donde cada una tiene: codigo, can (cantidad), imp (importe).\n\n" +
        "Texto:\n" + text.substring(0, 30000)
    }
  ]);

  if (!result.success) return result;
  let data;
  try {
    data = JSON.parse(typeof result.aiResponse === 'string' ? result.aiResponse : JSON.stringify(result.aiResponse));
  } catch (e) {
    return { ...result, error: 'Error al parsear respuesta: ' + e.message };
  }

  const partidas = Array.isArray(data.partidas) ? data.partidas : [];
  const budgetFileName = findLatestExcel();

  if (budgetFileName) {
    const budgetPath = path.join(OUTPUT_DIR, budgetFileName);
    const wb = XLSX.readFile(budgetPath);
    const ws = wb.Sheets['Presupuesto'] || wb.Sheets[wb.SheetNames[0]];
    if (ws && ws['!ref']) {
      const certInfo = getNextCertColumns(budgetPath);
      const range = XLSX.utils.decode_range(ws['!ref']);
      const colLabels = [certInfo.label, 'Can', 'Imp', 'CompCan', 'CompImp', '%'];
      for (let i = 0; i < colLabels.length; i++) {
        ws[XLSX.utils.encode_cell({ r: 0, c: certInfo.startCol + i })] = { t: 's', v: colLabels[i] };
      }
      for (let r = 1; r <= range.e.r; r++) {
        const codigoCell = ws[XLSX.utils.encode_cell({ r: r, c: 0 })];
        if (!codigoCell) continue;
        const codigo = String(codigoCell.v).trim();
        const match = partidas.find(p => String(p.codigo ?? '').trim() === codigo);
        if (match) {
          ws[XLSX.utils.encode_cell({ r: r, c: certInfo.startCol })] = { t: 'n', v: match.can ?? 0 };
          ws[XLSX.utils.encode_cell({ r: r, c: certInfo.startCol + 1 })] = { t: 'n', v: match.imp ?? 0 };
        }
      }
      if (range.e.c < certInfo.startCol + 5) {
        range.e.c = certInfo.startCol + 5;
        ws['!ref'] = XLSX.utils.encode_range(range);
      }
      const baseName = path.basename(budgetFileName, '.xlsx');
      const newFile = saveExcel(wb, baseName);
      return {
        success: true,
        tipo: 'Certificación',
        certificacion: data.numero_certificacion || certInfo.label,
        num_partidas: partidas.length,
        texto_extraido: text.length + ' caracteres',
        archivo: newFile,
        basado_en: budgetFileName
      };
    }
  }

  // No budget Excel found
  const wb = XLSX.utils.book_new();
  const objHeaders = Object.keys(data).filter(k => k !== 'partidas');
  const values = objHeaders.map(h => typeof data[h] === 'object' ? JSON.stringify(data[h]) : String(data[h] ?? ''));
  const ws = XLSX.utils.aoa_to_sheet([objHeaders, values]);
  XLSX.utils.book_append_sheet(wb, ws, 'Certificacion');
  const fileName = saveExcel(wb, 'certificacion');
  return { success: true, tipo: 'Certificación', archivo: fileName, num_partidas: partidas.length };
}

// ─── Multipart form parser ─────────────────────────────────────────────
function parseMultipart(buffer, boundary) {
  const parts = [];
  const blocks = buffer.toString('binary').split('--' + boundary);
  for (const block of blocks) {
    if (block.trim() === '' || block.trim() === '--') continue;
    const headerEnd = block.indexOf('\r\n\r\n');
    if (headerEnd === -1) continue;
    const headerSection = block.substring(0, headerEnd);
    const contentStart = headerEnd + 4;
    const nameMatch = headerSection.match(/name="([^"]+)"/);
    const filenameMatch = headerSection.match(/filename="([^"]+)"/);
    const content = Buffer.from(block.substring(contentStart), 'binary');
    parts.push({ name: nameMatch ? nameMatch[1] : null, filename: filenameMatch ? filenameMatch[1] : null, content });
  }
  return parts;
}

// ─── HTML Interface ──────────────────────────────────────────────────────
function serveHTML(res, message) {
  const files = fs.existsSync(OUTPUT_DIR) ? fs.readdirSync(OUTPUT_DIR).filter(f => f.endsWith('.xlsx')).sort().reverse() : [];

  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.end(`<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Arquitek - Gestión de Certificaciones</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f5f5; color: #333; }
header { background: #1a1a2e; color: white; padding: 20px 0; text-align: center; }
header h1 { font-size: 24px; font-weight: 300; }
header p { font-size: 14px; opacity: 0.7; margin-top: 4px; }
.container { max-width: 900px; margin: 0 auto; padding: 20px; }
.cards { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 20px; }
.card { background: white; border-radius: 12px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
.card h2 { font-size: 18px; margin-bottom: 16px; color: #1a1a2e; }
.card p { font-size: 13px; color: #666; margin-bottom: 12px; }
.card form { display: flex; flex-direction: column; gap: 12px; }
.card input[type=file] { font-size: 14px; }
.card button { background: #1a1a2e; color: white; border: none; padding: 10px 20px; border-radius: 6px; font-size: 14px; cursor: pointer; }
.card button:hover { background: #16213e; }
.card button:disabled { opacity: 0.5; }
.msg { margin-top: 20px; padding: 16px 20px; border-radius: 8px; font-size: 14px; line-height: 1.5; display: ${message ? 'block' : 'none'}; }
.msg.ok { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
.msg.err { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
.msg.wait { background: #fff3cd; color: #856404; border: 1px solid #ffeeba; }
.msg pre { margin-top: 8px; font-size: 12px; white-space: pre-wrap; }
.files { margin-top: 20px; background: white; border-radius: 12px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
.files h2 { font-size: 18px; margin-bottom: 12px; color: #1a1a2e; }
.files table { width: 100%; border-collapse: collapse; font-size: 14px; }
.files th { text-align: left; padding: 8px 12px; border-bottom: 2px solid #eee; color: #666; font-weight: 600; }
.files td { padding: 8px 12px; border-bottom: 1px solid #f0f0f0; }
.files a { color: #1a73e8; text-decoration: none; }
.files a:hover { text-decoration: underline; }
@media (max-width: 600px) { .cards { grid-template-columns: 1fr; } }
.loading { display: none; }
</style>
</head>
<body>
<header>
  <h1>Arquitek</h1>
  <p>Gestión de Presupuestos y Certificaciones de Obra</p>
</header>
<div class="container">
  <div class="cards">
    <div class="card">
      <h2>📄 Presupuesto Inicial</h2>
      <p>Sube el PDF del presupuesto de obra (solo la primera vez).</p>
      <form action="/upload/budget" method="post" enctype="multipart/form-data" target="hidden-frame" onsubmit="uploading('budget')">
        <input type="file" name="file" accept=".pdf" required>
        <button type="submit" id="btn-budget">Procesar Presupuesto</button>
        <span id="spinner-budget" class="loading" style="font-size:13px;color:#666;">⏳ Procesando... puede tardar 2-3 min</span>
      </form>
    </div>
    <div class="card">
      <h2>📋 Certificación Mensual</h2>
      <p>Sube el PDF de la certificación del mes para añadirla al Excel.</p>
      <form action="/upload/cert" method="post" enctype="multipart/form-data" target="hidden-frame" onsubmit="uploading('cert')">
        <input type="file" name="file" accept=".pdf" required>
        <button type="submit" id="btn-cert">Procesar Certificación</button>
        <span id="spinner-cert" class="loading" style="font-size:13px;color:#666;">⏳ Procesando... puede tardar 1-2 min</span>
      </form>
    </div>
  </div>

  <div id="msg" class="msg ${message ? (message.startsWith('✅') || message.startsWith('✓') ? 'ok' : 'err') : ''}">${message || ''}</div>

  <div class="files">
    <h2>📁 Archivos Generados</h2>
    ${files.length === 0 ? '<p style="color:#999;font-size:14px;">Aún no hay archivos. Sube un presupuesto para empezar.</p>' : ''}
    ${files.length > 0 ? '<table><tr><th>Archivo</th><th>Tamaño</th><th></th></tr>' + files.map(f => {
      const stats = fs.statSync(path.join(OUTPUT_DIR, f));
      const size = stats.size < 1024 ? stats.size + ' B' : (stats.size / 1024).toFixed(1) + ' KB';
      const icon = f.startsWith('presupuesto') ? '📄' : '📊';
      return '<tr><td>' + icon + ' <a href="/download?file=' + encodeURIComponent(f) + '">' + f + '</a></td><td>' + size + '</td></tr>';
    }).join('') + '</table>' : ''}
  </div>
</div>
<iframe name="hidden-frame" style="display:none" id="hidden-frame"></iframe>
<script>
function uploading(type) {
  document.getElementById('btn-' + type).disabled = true;
  document.getElementById('spinner-' + type).style.display = 'inline';
  document.getElementById('msg').style.display = 'none';
}
document.getElementById('hidden-frame').onload = function() {
  document.getElementById('btn-budget').disabled = false;
  document.getElementById('btn-cert').disabled = false;
  document.getElementById('spinner-budget').style.display = 'none';
  document.getElementById('spinner-cert').style.display = 'none';
  location.reload();
};
</script>
</body>
</html>`);
}

// ─── HTTP Server ─────────────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);
  const pathname = parsed.pathname;

  // ── API endpoints (JSON) ──────────────────────────────────────────────
  if (pathname === '/budget' || pathname === '/cert') {
    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Access-Control-Allow-Origin', '*');
    const binaryId = parsed.query.binaryId;
    if (!binaryId) {
      res.statusCode = 400;
      res.end(JSON.stringify({ success: false, error: 'Falta binaryId' }));
      return;
    }
    const storagePath = path.join(N8N_STORAGE, binaryId.replace('filesystem-v2:', 'storage/'));
    if (!fs.existsSync(storagePath)) {
      res.statusCode = 404;
      res.end(JSON.stringify({ success: false, error: 'Archivo no encontrado' }));
      return;
    }
    const text = readPdfFromFile(storagePath);
    const result = pathname === '/budget' ? await processBudget(text) : await processCert(text);
    res.end(JSON.stringify(result));
    return;
  }

  // ── Upload endpoints ─────────────────────────────────────────────────
  if (pathname === '/upload/budget' || pathname === '/upload/cert') {
    const contentType = req.headers['content-type'] || '';
    const boundaryMatch = contentType.match(/boundary=(.+)/);
    if (!boundaryMatch) {
      serveHTML(res, '❌ Error: formato multipart esperado');
      return;
    }

    const chunks = [];
    req.on('data', c => chunks.push(c));
    req.on('end', async () => {
      const buffer = Buffer.concat(chunks);
      const parts = parseMultipart(buffer, boundaryMatch[1]);
      const filePart = parts.find(p => p.content.length > 0);
      if (!filePart) {
        serveHTML(res, '❌ Error: no se recibió ningún archivo');
        return;
      }

      let msg;
      try {
        const text = readPdfFromBuffer(filePart.content);
        if (pathname === '/upload/budget') {
          const result = await processBudget(text);
          msg = result.success
            ? '✅ Presupuesto procesado correctamente.<br>' + result.num_partidas + ' partidas extraídas. Archivo: ' + result.archivo
            : '❌ Error: ' + (result.error || 'desconocido');
        } else {
          const result = await processCert(text);
          msg = result.success
            ? '✅ Certificación ' + (result.certificacion || '') + ' añadida.<br>' + result.num_partidas + ' partidas. Archivo: ' + result.archivo
            : '❌ Error: ' + (result.error || 'desconocido');
        }
      } catch (e) {
        msg = '❌ Error: ' + e.message;
      }
      serveHTML(res, msg);
    });
    return;
  }

  // ── Download ──────────────────────────────────────────────────────────
  if (pathname === '/download') {
    const file = parsed.query.file;
    if (!file) { serveHTML(res, '❌ Error: falta parámetro file'); return; }
    const filePath = path.join(OUTPUT_DIR, path.basename(file));
    if (!fs.existsSync(filePath)) { serveHTML(res, '❌ Error: archivo no encontrado'); return; }
    const stat = fs.statSync(filePath);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + path.basename(filePath) + '"');
    res.setHeader('Content-Length', stat.size);
    fs.createReadStream(filePath).pipe(res);
    return;
  }

  // ── Web UI ────────────────────────────────────────────────────────────
  serveHTML(res, null);
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Proxy server listening on port ${PORT}`);
  console.log('Interfaz web: http://localhost:' + PORT);
  console.log('Endpoints API: GET /budget?binaryId=...  GET /cert?binaryId=...');
});
