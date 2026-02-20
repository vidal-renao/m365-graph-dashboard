#!/bin/bash

echo "=========================================="
echo "  My Microsoft 365 Dashboard"
echo "=========================================="
echo ""
echo "Iniciando servidor en puerto 8080..."
echo ""
echo "Una vez iniciado, abre tu navegador en:"
echo "  👉 http://localhost:8080"
echo ""
echo "Presiona Ctrl+C para detener el servidor"
echo ""
echo "=========================================="
echo ""

# Intentar con Python 3
if command -v python3 &> /dev/null; then
    echo "✅ Usando Python 3..."
    python3 -m http.server 8080
# Si no, intentar con Python 2
elif command -v python &> /dev/null; then
    echo "✅ Usando Python..."
    python -m SimpleHTTPServer 8080
else
    echo "❌ Python no está instalado"
    echo ""
    echo "Instala Python o usa otro método del README.md"
    exit 1
fi
