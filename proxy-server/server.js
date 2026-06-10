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

function resolveBinaryPath(binaryId) {
  return path.join(N8N_STORAGE, binaryId.replace('filesystem-v2:', 'storage/'));
}

function readPdf(storagePath) {
  const pdfBuffer = fs.readFileSync(storagePath);
  let text = extractPdfText(pdfBuffer);
  if (isLikelyScanned(text)) {
    console.log('PDF parece escaneado, aplicando OCR...');
    text = extractTextViaOCR(storagePath) || text;
  }
  return text;
}

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
  const filePath = path.join(OUTPUT_DIR, fileName);
  XLSX.writeFile(wb, filePath);
  return fileName;
}

// ─── /budget endpoint ─────────────────────────────────────────────────────
async function handleBudget(storagePath) {
  const text = readPdf(storagePath);
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
    return { ...result, error: 'Parse failed: ' + e.message };
  }

  const wb = XLSX.utils.book_new();
  const partidas = Array.isArray(data.partidas) ? data.partidas : [];

  // Columnas A-F: Código, Ud, Resumen, CanPres, PrPres, ImpPres
  const headers = ['Código', 'Ud', 'Resumen', 'CanPres', 'PrPres', 'ImpPres'];
  const wsData = [headers];

  for (const p of partidas) {
    wsData.push([
      String(p.codigo ?? ''),
      String(p.ud ?? ''),
      String(p.resumen ?? ''),
      p.canPres ?? 0,
      p.prPres ?? 0,
      p.impPres ?? 0
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [
    { wch: 10 }, { wch: 8 }, { wch: 55 }, { wch: 12 }, { wch: 12 }, { wch: 14 }
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Presupuesto');

  const fileName = saveExcel(wb, 'presupuesto');
  return {
    success: true,
    textLength: text.length,
    numPartidas: partidas.length,
    total: data.total || data.total_presupuesto || null,
    excelFile: fileName
  };
}

// ─── /cert endpoint ────────────────────────────────────────────────────────
function findLatestExcel() {
  if (!fs.existsSync(OUTPUT_DIR)) return null;
  const files = fs.readdirSync(OUTPUT_DIR)
    .filter(f => f.startsWith('presupuesto_') && f.endsWith('.xlsx'))
    .sort()
    .reverse();
  return files.length > 0 ? path.join(OUTPUT_DIR, files[0]) : null;
}

function getNextCertColumns(existingPath) {
  // Columnas por certificación: Can, Imp, CompCan, CompImp, %
  // Buscar en el Excel existente cuántas certificaciones hay
  if (!existingPath) return { startCol: 7, label: 'Certif 01' };

  const wb = XLSX.readFile(existingPath);
  const ws = wb.Sheets['Presupuesto'] || wb.Sheets[wb.SheetNames[0]];
  if (!ws || !ws['!ref']) return { startCol: 7, label: 'Certif 01' };

  const range = XLSX.utils.decode_range(ws['!ref']);
  // Read header row to find existing certs
  let certCount = 0;
  for (let c = 6; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c: c })];
    if (cell && String(cell.v).startsWith('Certif ')) {
      certCount++;
    }
  }
  const next = certCount + 1;
  const startCol = 6 + (next - 1) * 5 + (next > 1 ? 1 : 0); // gap between certs
  const label = 'Certif ' + String(next).padStart(2, '0');
  return { startCol, label };
}

async function handleCert(storagePath) {
  const text = readPdf(storagePath);
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
    return { ...result, error: 'Parse failed: ' + e.message };
  }

  const partidas = Array.isArray(data.partidas) ? data.partidas : [];

  // Find existing budget Excel to append certification columns
  const budgetPath = findLatestExcel();
  let existingWb, existingWs;

  if (budgetPath) {
    existingWb = XLSX.readFile(budgetPath);
    existingWs = existingWb.Sheets['Presupuesto'] || existingWb.Sheets[existingWb.SheetNames[0]];
  }

  if (existingWs && existingWs['!ref']) {
    // Append certification columns to existing Excel
    const certInfo = getNextCertColumns(budgetPath);
    const range = XLSX.utils.decode_range(existingWs['!ref']);
    const colLabels = [certInfo.label, 'Can', 'Imp', 'CompCan', 'CompImp', '%'];

    // Write header row
    for (let i = 0; i < colLabels.length; i++) {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: certInfo.startCol + i });
      existingWs[cellRef] = { t: 's', v: colLabels[i] };
    }

    // Map partidas to existing rows by codigo
    for (let r = 1; r <= range.e.r; r++) {
      const codigoCell = existingWs[XLSX.utils.encode_cell({ r: r, c: 0 })];
      if (!codigoCell) continue;
      const codigo = String(codigoCell.v).trim();
      const match = partidas.find(p => String(p.codigo ?? '').trim() === codigo);
      if (match) {
        existingWs[XLSX.utils.encode_cell({ r: r, c: certInfo.startCol })] = { t: 'n', v: match.can ?? 0 };
        existingWs[XLSX.utils.encode_cell({ r: r, c: certInfo.startCol + 1 })] = { t: 'n', v: match.imp ?? 0 };
      }
    }

    if (range.e.c < certInfo.startCol + 5) {
      range.e.c = certInfo.startCol + 5;
      existingWs['!ref'] = XLSX.utils.encode_range(range);
    }

    const baseName = path.basename(budgetPath, '.xlsx');
    const fileName = saveExcel(existingWb, baseName);
    return {
      success: true,
      textLength: text.length,
      certLabel: certInfo.label,
      numPartidas: partidas.length,
      excelFile: fileName,
      appendedTo: path.basename(budgetPath)
    };
  } else {
    // No existing Excel, create a simple flat one
    const wb = XLSX.utils.book_new();
    const headers = Object.keys(data).filter(k => k !== 'partidas');
    const values = headers.map(h => {
      const v = data[h];
      return typeof v === 'object' ? JSON.stringify(v) : String(v ?? '');
    });
    const wsData = [headers, values];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Certificacion');
    const fileName = saveExcel(wb, 'certificacion');
    return {
      success: true,
      textLength: text.length,
      excelFile: fileName
    };
  }
}

// ─── HTTP Router ───────────────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  res.setHeader('Content-Type', 'application/json');
  res.setHeader('Access-Control-Allow-Origin', '*');

  const parsed = url.parse(req.url, true);
  const pathname = parsed.pathname;
  const binaryId = parsed.query.binaryId;

  if (!binaryId) {
    res.statusCode = 400;
    res.end(JSON.stringify({ success: false, error: 'Missing binaryId query param' }));
    return;
  }

  const storagePath = resolveBinaryPath(binaryId);
  if (!fs.existsSync(storagePath)) {
    res.statusCode = 404;
    res.end(JSON.stringify({ success: false, error: 'File not found: ' + storagePath }));
    return;
  }

  let result;
  try {
    if (pathname === '/budget') {
      result = await handleBudget(storagePath);
    } else {
      result = await handleCert(storagePath);
    }
  } catch (e) {
    result = { success: false, error: e.message };
  }

  res.end(JSON.stringify(result));
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Proxy server listening on port ${PORT}`);
  console.log('Endpoints: GET /budget?binaryId=...  GET /cert?binaryId=...');
});
