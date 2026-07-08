#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

SCRIPT="exportar_salud_completo.py"

echo "============================================================"
echo " Exportador de Apple Salud a Excel"
echo "============================================================"
echo

if [[ ! -f "$SCRIPT" ]]; then
  echo "ERROR: No encuentro $SCRIPT en esta carpeta."
  echo "Copia este .sh en la misma carpeta que $SCRIPT."
  exit 1
fi

if command -v python3 >/dev/null 2>&1; then
  PY="python3"
elif command -v python >/dev/null 2>&1; then
  PY="python"
else
  echo "ERROR: No encuentro Python instalado."
  echo "Instala Python 3 y vuelve a ejecutar este archivo."
  exit 1
fi

echo "Comprobando dependencia openpyxl..."
if ! "$PY" -c "import openpyxl" >/dev/null 2>&1; then
  echo "openpyxl no está instalado. Instalando..."
  "$PY" -m pip install --user openpyxl
fi

echo
echo "Versión del script:"
"$PY" "$SCRIPT" --version
echo
echo "Deja en blanco cualquier campo para usar el valor por defecto."
echo

read -r -p "Ruta a export.xml o carpeta apple_health_export [Enter = buscar automáticamente]: " XML
read -r -p "Fecha inicio YYYY-MM-DD [Enter = sin límite]: " INICIO
read -r -p "Fecha fin YYYY-MM-DD [Enter = sin límite]: " FIN
read -r -p "Detalle dia/hora/min/todos [Enter = hora]: " DETALLE
DETALLE=${DETALLE:-hora}
read -r -p "Nombre Excel salida .xlsx [Enter = automático]: " SALIDA

echo
echo "Ejecutando exportación..."
echo

"$PY" "$SCRIPT"   --xml "$XML"   --inicio "$INICIO"   --fin "$FIN"   --detalle "$DETALLE"   --salida "$SALIDA"

echo
echo "Exportación finalizada correctamente."
