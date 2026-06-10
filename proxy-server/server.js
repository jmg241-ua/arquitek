const http = require('http');
const fs = require('fs');
const https = require('https');
const path = require('path');
const url = require('url');
const { execSync } = require('child_process');
const ExcelJS = require('exceljs');

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

function isLikelyScanned(text, rawPdf) {
  if (rawPdf) {
    const hasBTET = /BT[\s\S]*?ET/.test(rawPdf.toString('binary'));
    if (hasBTET) {
      const alpha = (text.match(/[a-zA-ZáéíóúñüÁÉÍÓÚÑÜ]/g) || []).length;
      const ratio = text.length > 0 ? alpha / text.length : 0;
      if (ratio >= 0.3) return false;
    }
    return true;
  }
  if (text.length < 100) return true;
  const alpha = (text.match(/[a-zA-ZáéíóúñüÁÉÍÓÚÑÜ]/g) || []).length;
  if (alpha / text.length < 0.4) return true;
  return false;
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
  if (isLikelyScanned(text, buffer)) {
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

// ─── Excel helpers ────────────────────────────────────────────────────────
function generateUniqueName(base, ext) {
  let name = base + ext;
  if (!fs.existsSync(path.join(OUTPUT_DIR, name))) return name;
  let v = 1;
  while (fs.existsSync(path.join(OUTPUT_DIR, base + '_v' + v + ext))) v++;
  return base + '_v' + v + ext;
}

function saveExcel(wb, baseName) {
  const filePath = path.join(OUTPUT_DIR, baseName);
  return wb.xlsx.writeFile(filePath).then(() => baseName);
}

const BLUE_FONT = { color: { argb: 'FF0070C0' } };
const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } };

function styleCertHeader(cell) {
  cell.font = { ...BLUE_FONT, bold: true };
  cell.fill = HEADER_FILL;
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
}

function styleCertData(cell) {
  cell.font = BLUE_FONT;
}

function getNextCertColumns(filePath) {
  // Returns { startCol, label } for the next certification block
  // Each block is 6 columns: merged label (row 1) + Can, Imp, CompCan, CompImp, % (row 2)
  return new Promise((resolve) => {
    if (!filePath || !fs.existsSync(filePath)) {
      return resolve({ startCol: 7, label: 'Certif 01' });
    }
    const wb = new ExcelJS.Workbook();
    wb.xlsx.readFile(filePath).then(() => {
      const ws = wb.getWorksheet('Presupuesto');
      let certCount = 0;
      if (ws) {
        // Check only the label columns (col 7, 13, 19, ...) to avoid merged cell duplication
        for (let c = 7; c <= 200; c += 6) {
          const cell = ws.getCell(1, c);
          const v = String(cell.value || '').trim();
          if (v.startsWith('Certif ')) certCount++;
          else break;
        }
      }
      const next = certCount + 1;
      const startCol = 7 + (next - 1) * 6;
      resolve({ startCol, label: 'Certif ' + String(next).padStart(2, '0') });
    }).catch(() => {
      resolve({ startCol: 7, label: 'Certif 01' });
    });
  });
}

function findLatestExcel() {
  if (!fs.existsSync(OUTPUT_DIR)) return null;
  return fs.readdirSync(OUTPUT_DIR)
    .filter(f => f.startsWith('presupuesto_') && f.endsWith('.xlsx'))
    .sort().reverse()[0] || null;
}

// ─── Fuzzy partida matching ──────────────────────────────────────────────
function matchPartida(budgetCode, partidas) {
  const bc = String(budgetCode ?? '').trim().toLowerCase();
  const exact = partidas.find(p => String(p.codigo ?? '').trim().toLowerCase() === bc);
  if (exact) return exact;
  const prefix = partidas.find(p => bc.startsWith(String(p.codigo ?? '').trim().toLowerCase()));
  if (prefix) return prefix;
  const revPrefix = partidas.find(p => String(p.codigo ?? '').trim().toLowerCase().startsWith(bc));
  if (revPrefix) return revPrefix;
  return null;
}

// ─── /budget logic ─────────────────────────────────────────────────────
async function processBudget(text, originalFilename) {
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
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Presupuesto');

  // Row 1: reserved for future cert merged labels (empty)
  // Row 2: budget headers
  const headers = ['Código', 'Ud', 'Resumen', 'CanPres', 'PrPres', 'ImpPres'];
  const headerRow = ws.getRow(2);
  headerRow.values = headers;
  headerRow.font = { bold: true };
  headerRow.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // Row 3+: data
  for (let i = 0; i < partidas.length; i++) {
    const p = partidas[i];
    const row = ws.getRow(i + 3);
    row.values = [String(p.codigo ?? ''), String(p.ud ?? ''), String(p.resumen ?? ''), p.canPres ?? 0, p.prPres ?? 0, p.impPres ?? 0];
    row.eachCell((cell, col) => {
      if (col >= 4) cell.numFmt = '#,##0.00';
    });
  }

  // Column widths
  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 8;
  ws.getColumn(3).width = 55;
  ws.getColumn(4).width = 12;
  ws.getColumn(5).width = 12;
  ws.getColumn(6).width = 14;

  // Freeze panes: freeze first 6 cols (A-F) and first 2 rows
  ws.views = [{ state: 'frozen', xSplit: 6, ySplit: 2 }];

  let pdfBase = 'presupuesto';
  if (originalFilename) {
    pdfBase = path.basename(originalFilename).replace(/\.pdf$/i, '').replace(/[^a-zA-Z0-9_ -]/g, '_').trim();
    if (!pdfBase) pdfBase = 'presupuesto';
  }
  pdfBase = 'presupuesto_' + pdfBase;
  const fileName = generateUniqueName(pdfBase, '.xlsx');
  await saveExcel(wb, fileName);
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
async function processCert(text, budgetFileName) {
  if (!budgetFileName) {
    return { success: false, error: 'No hay presupuesto. Sube primero el presupuesto inicial.' };
  }

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
  const budgetPath = path.join(OUTPUT_DIR, budgetFileName);

  if (!fs.existsSync(budgetPath)) {
    return { success: false, error: 'Archivo de presupuesto no encontrado: ' + budgetFileName };
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(budgetPath);
  const ws = wb.getWorksheet('Presupuesto');
  if (!ws) {
    return { success: false, error: 'El archivo de presupuesto no tiene hoja válida' };
  }

  const certInfo = await getNextCertColumns(budgetPath);
  const sc = certInfo.startCol;

  // Detect format: old format has data starting at row 2 (exceljs row 2),
  // new format has data starting at row 3 (exceljs row 3)
  // Check if row 1 (exceljs) has anything (new format) or is empty (old format)
  const row1CellA = ws.getCell(1, 1).value;
  const isNewFormat = row1CellA === null || row1CellA === undefined || row1CellA === '';

  let dataStartRow = 2; // old format: data starts at row 2
  if (isNewFormat) {
    dataStartRow = 3; // new format: data starts at row 3 (row 1=empty, row 2=headers)
  }

  // Row 1: merged cert label
  // In old format, existing data is at row 2; in new format, at row 3
  // We need to handle old format by shifting data down
  if (!isNewFormat) {
    // Old format: insert a row at position 2 (between header and data)
    // Actually, old format has: row 1 = headers, row 2+ = data
    // We need: row 1 = empty (for label), row 2 = headers, row 3+ = data
    // So shift everything from row 2 down by 1
    // But we need to preserve merged cells
    // Simpler: just handle new format by checking if row 1 exists and has merged cells for certs
    // For old format, just use the old approach (single row header)
    // Actually, let me just use the old layout for simplicity in old files

    // Old format: single row header at row 1
    sc; // sc is correct
    ws.getCell(1, sc).value = certInfo.label;
    styleCertHeader(ws.getCell(1, sc));
    // Merge label across all 6 cols  
    ws.mergeCells(1, sc, 1, sc + 5);
    ws.getCell(1, sc).alignment = { horizontal: 'center', vertical: 'middle' };
    const colLabels = ['Can', 'Imp', 'CompCan', 'CompImp', '%'];
    for (let i = 0; i < colLabels.length; i++) {
      const cell = ws.getCell(1, sc + i);
      cell.value = colLabels[i];
      styleCertHeader(cell);
    }
    // Data starts at row 2
    const lastRow = ws.rowCount;
    for (let r = 2; r <= lastRow; r++) {
      const codigoCell = ws.getCell(r, 1);
      if (!codigoCell.value) continue;
      const codigo = String(codigoCell.value).trim();
      const match = matchPartida(codigo, partidas);
      if (match) {
        const canCell = ws.getCell(r, sc);
        canCell.value = match.can ?? 0;
        styleCertData(canCell);
        canCell.numFmt = '#,##0.00';
        const impCell = ws.getCell(r, sc + 1);
        impCell.value = match.imp ?? 0;
        styleCertData(impCell);
        impCell.numFmt = '#,##0.00';
      }
    }
  } else {
    // New format: rows 1 = reserved, 2 = headers, 3+ = data

    // Row 1: merged cert label across 6 cols
    ws.mergeCells(1, sc, 1, sc + 5);
    const labelCell = ws.getCell(1, sc);
    labelCell.value = certInfo.label;
    styleCertHeader(labelCell);
    labelCell.alignment = { horizontal: 'center', vertical: 'middle' };

    // Row 2: column names (Can, Imp, CompCan, CompImp, %)
    const colLabels = ['Can', 'Imp', 'CompCan', 'CompImp', '%'];
    for (let i = 0; i < colLabels.length; i++) {
      const cell = ws.getCell(2, sc + i);
      cell.value = colLabels[i];
      styleCertHeader(cell);
    }

    // Rows 3+: cert data values
    const lastRow = ws.rowCount;
    for (let r = dataStartRow; r <= lastRow; r++) {
      const codigoCell = ws.getCell(r, 1);
      if (!codigoCell.value) continue;
      const codigo = String(codigoCell.value).trim();
      const match = matchPartida(codigo, partidas);
      if (match) {
        const canCell = ws.getCell(r, sc);
        canCell.value = match.can ?? 0;
        styleCertData(canCell);
        canCell.numFmt = '#,##0.00';
        const impCell = ws.getCell(r, sc + 1);
        impCell.value = match.imp ?? 0;
        styleCertData(impCell);
        impCell.numFmt = '#,##0.00';
      }
    }
  }

  // Set column widths for cert columns if not set
  for (let c = sc; c <= sc + 5; c++) {
    if (!ws.getColumn(c).width || ws.getColumn(c).width < 12) {
      ws.getColumn(c).width = 12;
    }
  }

  // Derive base name from budget file, strip _cXX and _vX suffixes
  let budgetBase = path.basename(budgetFileName, '.xlsx')
    .replace(/^presupuesto_/, '')
    .replace(/_(c\d+)/g, '')
    .replace(/_(v\d+)$/, '')
    .replace(/_+$/, '');
  const certNum = certInfo.label.replace('Certif ', '');
  const newFile = generateUniqueName('presupuesto_' + budgetBase + '_c' + certNum, '.xlsx');
  await saveExcel(wb, newFile);
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

// ─── Multipart form parser ─────────────────────────────────────────────
function parseMultipart(buffer, boundary) {
  const parts = [];
  const delimBuf = Buffer.from('\r\n--' + boundary);
  const endDelimBuf = Buffer.from('\r\n--' + boundary + '--');
  const doubleCRLF = Buffer.from('\r\n\r\n');
  let searchFrom = 0;
  while (searchFrom < buffer.length) {
    const bIdx = buffer.indexOf(Buffer.from('--' + boundary), searchFrom);
    if (bIdx === -1) break;
    const blockStart = bIdx + 2 + Buffer.byteLength(boundary);
    if (blockStart >= buffer.length) break;
    let dataStart = blockStart;
    if (buffer[dataStart] === 13) dataStart++;
    if (buffer[dataStart] === 10) dataStart++;
    const headerEnd = buffer.indexOf(doubleCRLF, dataStart);
    if (headerEnd === -1) break;
    const contentStart = headerEnd + 4;
    const nextDelim = buffer.indexOf(delimBuf, contentStart);
    const nextEndDelim = buffer.indexOf(endDelimBuf, contentStart);
    let blockEnd;
    if (nextDelim !== -1 && nextEndDelim !== -1) {
      blockEnd = Math.min(nextDelim, nextEndDelim);
    } else if (nextDelim !== -1) {
      blockEnd = nextDelim;
    } else if (nextEndDelim !== -1) {
      blockEnd = nextEndDelim;
    } else {
      blockEnd = buffer.length;
    }
    const headerSection = buffer.slice(dataStart, headerEnd).toString('utf-8');
    let content = buffer.slice(contentStart, blockEnd);
    if (content.length >= 2 && content[content.length - 2] === 13 && content[content.length - 1] === 10) {
      content = buffer.slice(contentStart, blockEnd - 2);
    }
    const nameMatch = headerSection.match(/name="([^"]+)"/);
    const filenameMatch = headerSection.match(/filename="([^"]+)"/);
    parts.push({ name: nameMatch ? nameMatch[1] : null, filename: filenameMatch ? filenameMatch[1] : null, content });
    searchFrom = blockEnd;
  }
  return parts;
}

// ─── HTML Interface ──────────────────────────────────────────────────────
function serveHTML(res, message) {
  const files = fs.existsSync(OUTPUT_DIR) ? fs.readdirSync(OUTPUT_DIR).filter(f => f.endsWith('.xlsx')).sort().reverse() : [];
  const baseFiles = files.filter(f => {
    if (!f.startsWith('presupuesto_')) return false;
    const base = f.replace('.xlsx', '');
    const parts = base.split('_');
    return parts.length === 2;
  });
  const hasBudget = baseFiles.length > 0;

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
      <p>Sube el PDF de la certificación del mes para añadirla al Excel del presupuesto.</p>
      ${!hasBudget ? '<p style="color:#c00;font-size:13px;">⚠️ Primero sube el presupuesto inicial.</p>' : '<p style="font-size:13px;color:#666;">Selecciona el archivo de presupuesto al que añadir esta certificación.</p>'}
      <form action="/upload/cert" method="post" enctype="multipart/form-data" target="hidden-frame" onsubmit="uploading('cert')">
        ${hasBudget ? '<select name="budgetFile" style="padding:8px;font-size:14px;border:1px solid #ccc;border-radius:6px;">' +
          files.filter(f => f.startsWith('presupuesto_')).map(f => {
            const isBase = f.split('_').length === 2;
            return '<option value="' + f + '">' + (isBase ? '📄 ' : '📊 ') + f + '</option>';
          }).join('') +
          '</select>' : ''}
        <input type="file" name="file" accept=".pdf" required>
        <button type="submit" id="btn-cert" ${!hasBudget ? 'disabled' : ''}>Procesar Certificación</button>
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
      const icon = f.split('_').length === 2 ? '📄' : '📊';
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
    if (pathname === '/budget') {
      const result = await processBudget(text, null);
      res.end(JSON.stringify(result));
    } else {
      const budgetFile = parsed.query.budgetFile || findLatestExcel();
      const result = await processCert(text, budgetFile);
      res.end(JSON.stringify(result));
    }
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
      const filePart = parts.find(p => p.name === 'file');
      if (!filePart) {
        serveHTML(res, '❌ Error: no se recibió ningún archivo');
        return;
      }

      let msg;
      try {
        const text = readPdfFromBuffer(filePart.content);
        if (pathname === '/upload/budget') {
          const result = await processBudget(text, filePart.filename);
          msg = result.success
            ? '✅ Presupuesto procesado correctamente.<br>' + result.num_partidas + ' partidas extraídas. Archivo: ' + result.archivo
            : '❌ Error: ' + (result.error || 'desconocido');
        } else {
          const budgetField = parts.find(p => p.name === 'budgetFile');
          let budgetFileName = budgetField ? budgetField.content.toString('utf-8').trim() : null;
          const result = await processCert(text, budgetFileName);
          msg = result.success
            ? '<strong>✅ Certificación ' + (result.certificacion || '') + ' añadida.</strong><br>' + result.num_partidas + ' partidas. Archivo: ' + result.archivo + '<br>Basado en: ' + result.basado_en
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
