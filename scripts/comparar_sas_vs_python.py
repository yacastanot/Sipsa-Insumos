"""Comparación genérica de salidas Python vs SAS para una carpeta de módulo.

Criterios (por archivo, por hoja):
  1. Mismos archivos en ambas carpetas (nombre sin distinguir mayúsculas/extensión).
  2. Mismas hojas.
  3. Mismas columnas y en el MISMO ORDEN.
  4. Mismo número de filas y mismas filas (multiconjunto: cada fila de SAS
     tiene su par idéntico en Python y viceversa). También se informa si el
     orden de las filas coincide.

Los valores se normalizan antes de comparar: texto sin espacios extremos y
con espacios internos colapsados, números a 9 cifras significativas (1.0 == 1),
vacíos/NaN como "". Los .XLSX de SAS se leen con python-calamine porque
openpyxl falla con su docProps/core.xml.

Uso:
    python scripts/comparar_sas_vs_python.py --py <carpeta_python> --sas <carpeta_sas> [--out reporte.txt]
"""
from __future__ import annotations

import argparse
import math
import re
from collections import Counter
from pathlib import Path

from python_calamine import CalamineWorkbook


def _norm(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(int(v))
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return ""
        r = float(f"{float(v):.12g}")
        return str(int(r)) if r == int(r) else repr(r)
    s = re.sub(r"\s+", " ", str(v)).strip()
    # Números guardados como texto ("3466302", "12.50") se comparan como número
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        return _norm(float(s))
    return s


TOL_REL = 1e-6
TOL_ABS = 1e-6


def _num(s: str) -> float | None:
    try:
        return float(s)
    except ValueError:
        return None


def _iguales_tol(a: str, b: str) -> bool:
    if a == b:
        return True
    x, y = _num(a), _num(b)
    if x is None or y is None:
        return False
    return abs(x - y) <= max(TOL_ABS, TOL_REL * max(abs(x), abs(y)))


def _emparejar_tolerante(solo_s: Counter, solo_p: Counter) -> tuple[Counter, Counter]:
    """Empareja filas sobrantes que solo difieren en números dentro de tolerancia.

    Agrupa por la "firma de texto" (campos no numéricos) para no comparar todo
    contra todo.
    """
    def firma(r):
        return tuple(c if _num(c) is None else "#" for c in r)

    grupos: dict[tuple, list] = {}
    for r in solo_p.elements():
        grupos.setdefault(firma(r), []).append(r)
    rest_s = Counter()
    for r in solo_s.elements():
        cands = grupos.get(firma(r), [])
        hit = next((i for i, c in enumerate(cands)
                    if all(_iguales_tol(a, b) for a, b in zip(r, c))), None)
        if hit is None:
            rest_s[r] += 1
        else:
            cands.pop(hit)
    rest_p = Counter(r for lst in grupos.values() for r in lst)
    return rest_s, rest_p


def _leer(path: Path) -> dict[str, list[list[str]]]:
    wb = CalamineWorkbook.from_path(str(path))
    hojas = {}
    for name in wb.sheet_names:
        rows = wb.get_sheet_by_name(name).to_python(skip_empty_area=False)
        rows = [[_norm(c) for c in r] for r in rows]
        # recortar filas y columnas vacías al final
        while rows and not any(rows[-1]):
            rows.pop()
        ncols = max((max((i + 1 for i, c in enumerate(r) if c), default=0) for r in rows), default=0)
        hojas[name] = [r[:ncols] + [""] * (ncols - len(r)) for r in rows]
    return hojas


def _clave_archivo(p: Path) -> str:
    return p.stem.upper()


def comparar(py_dir: Path, sas_dir: Path) -> list[str]:
    out: list[str] = []
    py = {_clave_archivo(p): p for p in py_dir.glob("*.xls*")}
    sas = {_clave_archivo(p): p for p in sas_dir.glob("*.xls*")}

    out.append("=" * 100)
    out.append("1. ARCHIVOS")
    out.append("=" * 100)
    solo_sas = sorted(set(sas) - set(py))
    solo_py = sorted(set(py) - set(sas))
    comunes = sorted(set(sas) & set(py))
    out.append(f"SAS={len(sas)} | Python={len(py)} | comunes={len(comunes)}")
    for k in solo_sas:
        out.append(f"  ✖ Solo en SAS:    {sas[k].name}")
    for k in solo_py:
        out.append(f"  ✖ Solo en Python: {py[k].name}")

    resumen: list[tuple[str, str, str]] = []
    for k in comunes:
        out.append("")
        out.append("=" * 100)
        out.append(f"ARCHIVO: {sas[k].name}")
        out.append("=" * 100)
        try:
            hs, hp = _leer(sas[k]), _leer(py[k])
        except Exception as exc:  # noqa: BLE001
            out.append(f"  ERROR leyendo: {exc}")
            resumen.append((sas[k].name, "-", f"ERROR lectura: {exc}"))
            continue

        nombres_ok = list(hs) == list(hp)
        if not nombres_ok:
            out.append(f"  ✖ Nombres de hoja distintos | SAS={list(hs)} | Python={list(hp)}")
        pares = ([(h, h) for h in hs] if nombres_ok
                 else list(zip(hs, hp)) if len(hs) == len(hp)
                 else [(h, h if h in hp else None) for h in hs])
        for hoja, hoja_py in pares:
            if hoja_py is None:
                resumen.append((sas[k].name, hoja, "✖ hoja falta en Python"))
                continue
            s_rows, p_rows = hs[hoja], hp[hoja_py]
            s_cols = s_rows[0] if s_rows else []
            p_cols = p_rows[0] if p_rows else []
            s_data, p_data = s_rows[1:], p_rows[1:]
            problemas = [] if hoja == hoja_py else [f"nombre de hoja: SAS '{hoja}' vs Python '{hoja_py}'"]

            if s_cols != p_cols:
                if set(s_cols) == set(p_cols):
                    problemas.append("orden de columnas distinto")
                    out.append(f"  [{hoja}] ✖ Mismas columnas, ORDEN distinto")
                    out.append(f"      SAS:    {s_cols}")
                    out.append(f"      Python: {p_cols}")
                else:
                    problemas.append("columnas distintas")
                    out.append(f"  [{hoja}] ✖ Columnas distintas ({len(s_cols)} SAS vs {len(p_cols)} Python)")
                    out.append(f"      Solo SAS:    {[c for c in s_cols if c not in p_cols]}")
                    out.append(f"      Solo Python: {[c for c in p_cols if c not in s_cols]}")
                    out.append(f"      SAS:    {s_cols}")
                    out.append(f"      Python: {p_cols}")
            else:
                out.append(f"  [{hoja}] ✔ Columnas iguales y en el mismo orden ({len(s_cols)})")

            if len(s_data) != len(p_data):
                problemas.append(f"filas {len(s_data)} SAS vs {len(p_data)} Python")
            # Comparar filas alineando columnas por nombre (si el conjunto coincide)
            if set(s_cols) == set(p_cols) and len(set(s_cols)) == len(s_cols):
                idx = [p_cols.index(c) for c in s_cols]
                p_al = [[r[i] if i < len(r) else "" for i in idx] for r in p_data]
            else:
                p_al = p_data
            cs, cp = Counter(map(tuple, s_data)), Counter(map(tuple, p_al))
            solo_s, solo_p = _emparejar_tolerante(cs - cp, cp - cs)
            n_s, n_p = sum(solo_s.values()), sum(solo_p.values())
            out.append(f"  [{hoja}] Filas: SAS={len(s_data)} | Python={len(p_data)} | "
                       f"solo SAS={n_s} | solo Python={n_p}")
            if n_s or n_p:
                problemas.append(f"{n_s} filas solo SAS / {n_p} solo Python")
                # Columnas que explican las diferencias (por posición)
                if set(s_cols) == set(p_cols):
                    cols_dif = Counter()
                    for fs in list(solo_s.elements())[:200]:
                        # buscar la fila Python más parecida entre las no emparejadas
                        best = min(list(solo_p.elements())[:3000], key=lambda fp: sum(a != b for a, b in zip(fs, fp)),
                                   default=None)
                        if best:
                            for c, a, b in zip(s_cols, fs, best):
                                if a != b:
                                    cols_dif[c] += 1
                    if cols_dif:
                        out.append(f"      Columnas con diferencias (filas no emparejadas): {dict(cols_dif.most_common(10))}")
                for fs in list(solo_s.elements())[:3]:
                    out.append(f"      solo SAS:    {dict(zip(s_cols, fs))}")
                for fp in list(solo_p.elements())[:3]:
                    out.append(f"      solo Python: {dict(zip(s_cols, fp))}")
            elif s_data != p_al and not all(
                    all(_iguales_tol(a, b) for a, b in zip(x, y)) for x, y in zip(s_data, p_al)):
                problemas.append("mismas filas, orden distinto")
                out.append(f"  [{hoja}] ⚠ Mismas filas pero en distinto orden")
            resumen.append((sas[k].name, hoja, "✔ IGUAL" if not problemas else "✖ " + "; ".join(problemas)))

    out.insert(0, "")
    cab = ["=" * 100, "RESUMEN", "=" * 100]
    cab += [f"{a:<48} {h[:30]:<30} {r}" for a, h, r in resumen]
    return cab + out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--py", required=True, type=Path)
    ap.add_argument("--sas", required=True, type=Path)
    ap.add_argument("--out", type=Path)
    a = ap.parse_args()
    texto = "\n".join(comparar(a.py, a.sas))
    if a.out:
        a.out.write_text(texto, encoding="utf-8")
    print(texto)


if __name__ == "__main__":
    main()
