const http = require('http');
const fs = require('fs');
const https = require('https');
const path = require('path');
const url = require('url');
const XLSX = require('xlsx');

const PORT = process.env.PORT || 3456;
const N8N_STORAGE = process.env.N8N_STORAGE || '/home/node/.n8n/storage';
const DEEPSEEK_API_KEY = process.env.DEEPSEEK_API_KEY || '';

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

function callDeepSeek(text) {
  return new Promise((resolve) => {
    const body = JSON.stringify({
      model: "deepseek-chat",
      messages: [
        {role: "system", content: "Eres un asistente que extrae datos de certificaciones de obra en formato JSON. Devuelve SOLO JSON valido sin markdown."},
        {role: "user", content: "Extrae los datos de esta certificacion de obra en formato JSON. Incluye: numero de certificacion, fecha, contratista, obra, importe certificado, importe acumulado, y cualquier otra informacion relevante.\n\nTexto:\n" + text.substring(0, 30000)}
      ],
      response_format: {type: "json_object"},
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
          resolve({success: true, textLength: text.length, aiResponse: content});
        } catch(e) {
          resolve({success: false, error: e.message, raw: data.substring(0, 500)});
        }
      });
    });
    req.on('error', (e) => resolve({success: false, error: e.message}));
    req.write(body);
    req.end();
  });
}

const server = http.createServer(async (req, res) => {
  res.setHeader('Content-Type', 'application/json');
  res.setHeader('Access-Control-Allow-Origin', '*');

  const parsed = url.parse(req.url, true);
  const binaryId = parsed.query.binaryId;

  if (!binaryId) {
    res.statusCode = 400;
    res.end(JSON.stringify({success: false, error: 'Missing binaryId query param'}));
    return;
  }

  const relativePath = binaryId.replace('filesystem-v2:', 'storage/');
const storagePath = path.join(N8N_STORAGE, relativePath);

  let pdfBuffer;
  try {
    pdfBuffer = fs.readFileSync(storagePath);
  } catch(e) {
    res.statusCode = 404;
    res.end(JSON.stringify({success: false, error: 'File not found: ' + storagePath, message: e.message}));
    return;
  }

  const text = extractPdfText(pdfBuffer);
  const result = await callDeepSeek(text);

  let aiData = {};
  try {
    const raw = result.aiResponse || '{}';
    aiData = JSON.parse(typeof raw === 'string' ? raw : JSON.stringify(raw));
  } catch(e) {
    aiData = {error: 'Parse failed', message: e.message};
  }

  const wb = XLSX.utils.book_new();
  const headers = Object.keys(aiData);
  const values = headers.map(h => {
    const v = aiData[h];
    return typeof v === 'object' ? JSON.stringify(v) : String(v ?? '');
  });
  if (headers.length > 0) {
    const ws = XLSX.utils.aoa_to_sheet([headers, values]);
    ws['!cols'] = headers.map((h, i) => ({
      wch: Math.min(Math.max(h.length, (values[i] || '').length) + 2, 80)
    }));
    XLSX.utils.book_append_sheet(wb, ws, 'Certificacion');
  }

  const ts = new Date().toISOString().replace(/[:.]/g, '-');
  const fileName = 'certificacion_' + ts + '.xlsx';
  const filePath = path.join('/output', fileName);
  try {
    XLSX.writeFile(wb, filePath);
    result.excelFile = fileName;
  } catch(e) {
    result.excelError = e.message;
  }

  res.end(JSON.stringify(result));
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Proxy server listening on port ${PORT}`);
});
