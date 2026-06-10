#!/usr/bin/env bash
set -euo pipefail

# Script de prueba: envia un PDF al webhook de certificacion y muestra el Excel generado
# Uso: ./test-envio.sh <ruta-al-pdf>
# Ejemplo: ./test-envio.sh ruta/al-certificado.pdf

N8N_URL="${N8N_URL:-http://localhost:5678}"
PDF_PATH="${1:-}"

if [ -z "$PDF_PATH" ]; then
  echo "Uso: $0 <ruta-al-pdf>"
  echo "  Ej: $0 ruta/al-certificado.pdf"
  exit 1
fi

if [ ! -f "$PDF_PATH" ]; then
  echo "No se encuentra el archivo: $PDF_PATH"
  exit 1
fi

# Obtener el path del webhook desde el workflow activo
# Buscar el workflow "Certificacion PDF a Excel" y extraer su webhook path
echo "Buscando workflow activo..."
WEBHOOK_PATH=$(curl -s -b /tmp/n8n-cookies "$N8N_URL/rest/workflows" 2>/dev/null | \
  python3 -c "
import sys, json
try:
    d = json.load(sys.stdin).get('data',{})
    for wf in d.values() if isinstance(d, dict) else (d.get('results',[]) if isinstance(d, dict) else []):
        if 'Certificacion' in str(wf.get('name','')):
            nodes = wf.get('nodes',[])
            for n in nodes:
                if n.get('type','') == 'n8n-nodes-base.webhook':
                    print(n.get('parameters',{}).get('path',''))
                    sys.exit(0)
except: pass
print('cert-080640')
")
WEBHOOK_PATH="${WEBHOOK_PATH:-cert-080640}"

echo "Enviando $(basename "$PDF_PATH") a $N8N_URL/webhook/$WEBHOOK_PATH ..."
echo ""

START=$(date +%s)
curl -s -w "\nHTTP %{http_code}" \
  -X POST "$N8N_URL/webhook/$WEBHOOK_PATH" \
  -F "data=@$PDF_PATH"
echo ""

echo ""
echo "Esperando 60s a que el proxy procese..."
sleep 60

echo ""
echo "Archivos en output/:"
ls -lt output/ 2>/dev/null | head -5

END=$(date +%s)
echo ""
echo "Tiempo total: $((END - START))s"
echo "Para ver el Excel: ls -la output/certificacion_*.xlsx"
