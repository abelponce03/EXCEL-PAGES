#!/usr/bin/env bash
# Construye el libro final:
#   1) genera el .xlsx con fórmulas (openpyxl)
#   2) lo recalcula con LibreOffice Calc (script recalc.py de la skill xlsx, ruta en $RECALC)
#   3) inyecta los valores calculados en el original (conserva protecciones y validaciones)
#   4) verifica los resultados contra un cálculo independiente en Python
set -euo pipefail
cd "$(dirname "$0")/.."
TMP=$(mktemp -d)
FINAL="Analisis_Rentabilidad_Paqueteria.xlsx"
python3 build/generar_excel.py "$TMP/base.xlsx"
if [[ -n "${RECALC:-}" && -f "$RECALC" ]]; then
  cp "$TMP/base.xlsx" "$TMP/recalc.xlsx"
  python3 "$RECALC" "$TMP/recalc.xlsx" 300 | tee "$TMP/recalc.json"
  grep -q '"total_errors": 0' "$TMP/recalc.json" || { echo "ERRORES DE FÓRMULA"; exit 1; }
  python3 build/inyectar_valores.py "$TMP/base.xlsx" "$TMP/recalc.xlsx" "$FINAL"
  python3 build/verificar.py "$TMP/recalc.xlsx" | tail -1
else
  echo "Sin RECALC: se entrega sin valores en caché (Excel calcula al abrir)."
  cp "$TMP/base.xlsx" "$FINAL"
fi
echo "Listo: $FINAL"
