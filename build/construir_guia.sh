#!/usr/bin/env bash
# Genera la guía en PDF:
#   1) capturas y valores de los ejercicios a partir del libro (build/capturas.py)
#   2) compila guia/guia_usuario.tex con latexmk (2 pasadas para índice y referencias)
#   3) copia el PDF a la raíz del repositorio
# Requiere RECALC (ruta de recalc.py), LibreOffice Calc, pdftoppm, Pillow y TeX Live.
set -euo pipefail
cd "$(dirname "$0")/.."
python3 build/capturas.py
cd guia
latexmk -pdf -interaction=nonstopmode -halt-on-error guia_usuario.tex > latexmk.log 2>&1 || { tail -40 latexmk.log; exit 1; }
cd ..
cp guia/guia_usuario.pdf Guia_de_Uso_Rentabilidad_Paqueteria.pdf
echo "Listo: Guia_de_Uso_Rentabilidad_Paqueteria.pdf ($(pdfinfo Guia_de_Uso_Rentabilidad_Paqueteria.pdf | grep Pages))"
