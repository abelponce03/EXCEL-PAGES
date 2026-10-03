# -*- coding: utf-8 -*-
"""Genera GUIA_DE_USO.md y ANALISIS_Y_PROPUESTA.md desde build/contenido.py."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from contenido import ANALISIS, GUIA

ROOT = Path(__file__).resolve().parent.parent


def md(blocks):
    out, num = [], 0
    for kind, val in blocks:
        if kind == "h1":
            out += [f"# {val.capitalize()}", ""]
        elif kind == "h2":
            num = 0
            out += [f"## {val}", ""]
        elif kind == "p":
            out += [val, ""]
        elif kind == "b":
            out.append(f"- {val}")
        elif kind == "n":
            num += 1
            out.append(f"{num}. {val}")
        elif kind == "t":
            if out and out[-1] != "":
                out.append("")
            out.append("| " + " | ".join(val[0]) + " |")
            out.append("|" + "---|" * len(val[0]))
            out += ["| " + " | ".join(str(x) for x in row) + " |" for row in val[1:]]
            out.append("")
    text = "\n".join(out)
    # línea en blanco tras cada lista
    lines, res = text.split("\n"), []
    for i, l in enumerate(lines):
        res.append(l)
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        is_item = l.startswith("- ") or l[:3].rstrip(". ").isdigit()
        nxt_item = nxt.startswith("- ") or nxt[:3].rstrip(". ").isdigit()
        if is_item and nxt and not nxt_item:
            res.append("")
    return "\n".join(res).rstrip() + "\n"


(ROOT / "GUIA_DE_USO.md").write_text(md(GUIA), encoding="utf-8")
(ROOT / "ANALISIS_Y_PROPUESTA.md").write_text(md(ANALISIS), encoding="utf-8")
print("OK")
