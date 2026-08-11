"""Validación Python vs SAS — SIPSA Insumos JUL2026.

Compara cada archivo de salida Python (data/08_reporting/JUL2026/) con su
contraparte SAS (7. JULIO 2026/REVISIÓN JUL2026/) hoja por hoja, columna por
columna, fila por fila. Basado en scripts/validar_vs_sas.py (MAY2026), con
rutas y el set de archivos ajustados a JUL2026 (solo agricolas/pecuarios/
elementos/empaques están activos este período — sin arriendos/servicios).

Uso:
    python scripts/validar_vs_sas_jul2026.py
    python scripts/validar_vs_sas_jul2026.py --modulo agricolas
"""
from __future__ import annotations

import argparse
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

SAS_BASE = Path(
    r"C:\Users\Jeferson\OneDrive - Cloud Integration Hub\Documentos"
    r"\DANE Automatización\SIPSA Insumos\SIPSA Insumos\7. JULIO 2026"
    r"\REVISIÓN JUL2026"
)
PY_BASE = Path(r"C:\Users\Jeferson\Kedro\Sipsa-Insumos\data\08_reporting\JUL2026")
PERIODO = "JUL2026"


def _pares() -> list[tuple[Path, Path, str]]:
    sas_a = SAS_BASE / "AGRICOLAS_JUL2026"
    sas_p = SAS_BASE / "PECUARIOS_JUL2026"
    sas_e = SAS_BASE / "ELEMENTOS_JUL2026"
    sas_em = SAS_BASE / "EMPAQUES_JUL2026"

    py_a = PY_BASE / "agricolas"
    py_p = PY_BASE / "pecuarios"
    py_e = PY_BASE / "elementos"
    py_em = PY_BASE / "empaques"
    py_sp = PY_BASE / "sin_precio_ant"

    return [
        # ── AGRÍCOLAS ──────────────────────────────────────────────────────
        (py_a / f"VAR_ATIPICO_AGRICOLA_{PERIODO}.xlsx",
         sas_a / f"VAR_ATIPICO_AGRICOLA_{PERIODO}.XLSX", "AGRICOLAS/VAR_ATIPICO"),
        (py_a / f"BASES INSUMOS AGRICOLAS {PERIODO}.xlsx",
         sas_a / f"BASES INSUMOS AGRICOLAS {PERIODO}.XLSX", "AGRICOLAS/BASES"),
        (py_a / f"BASE_INSUMOS_AGRICOLAS_{PERIODO}.xlsx",
         sas_a / f"BASE_INSUMOS_AGRICOLAS_{PERIODO}.XLSX", "AGRICOLAS/BASE_INSUMOS"),
        (py_a / f"CUADROS INSUMOS AGRICOLAS {PERIODO}.xlsx",
         sas_a / f"CUADROS INSUMOS AGRICOLAS {PERIODO}.XLSX", "AGRICOLAS/CUADROS"),
        (py_a / f"TABLAS INSUMOS AGRICOLAS {PERIODO}.xlsx",
         sas_a / f"TABLAS INSUMOS AGRICOLAS {PERIODO}.xlsx", "AGRICOLAS/TABLAS"),
        (py_a / f"INSUMOS_AGRICOLAS_DUPLI_{PERIODO}.xlsx",
         sas_a / f"INSUMOS_AGRICOLAS_DUPLI_{PERIODO}.XLSX", "AGRICOLAS/DUPLI"),
        (py_a / f"CV's INSUMOS AGRICOLAS {PERIODO}.xlsx",
         sas_a / f"CV's INSUMOS AGRICOLAS {PERIODO}.XLSX", "AGRICOLAS/CVs"),
        (py_a / f"FALTAN_GRUPO_AGRICOLA_{PERIODO}.xlsx",
         sas_a / f"FALTAN_GRUPO_AGRICOLA_{PERIODO}.XLSX", "AGRICOLAS/FALTAN_GRUPO"),
        (py_a / f"FALTAN_PUBLICA_AGRICOLA_{PERIODO}.xlsx",
         sas_a / f"FALTAN_PUBLICA_AGRICOLA_{PERIODO}.XLSX", "AGRICOLAS/FALTAN_PUBLICA"),
        (py_a / f"ANEXO INSUMOS AGRICOLAS {PERIODO}.xlsx",
         sas_a / f"ANEXO INSUMOS AGRICOLAS {PERIODO}.XLSX", "AGRICOLAS/ANEXO"),
        (py_a / f"INSU_AGRIC_MAYORESQUE2_{PERIODO}.xlsx",
         sas_a / f"INSU_AGRIC_MAYORESQUE2_{PERIODO}.XLSX", "AGRICOLAS/MAYORESQUE2"),
        (py_a / f"INSU_AGRIC_MENORESQUE2_{PERIODO}.xlsx",
         sas_a / f"INSU_AGRIC_MENORESQUE2_{PERIODO}.XLSX", "AGRICOLAS/MENORESQUE2"),
        (py_a / f"MAY_MEN3_AGRIC_{PERIODO}.xlsx",
         sas_a / f"MAY_MEN3_AGRIC_{PERIODO}.XLSX", "AGRICOLAS/MAYMEN3"),
        # ── PECUARIOS ─────────────────────────────────────────────────────
        (py_p / f"VAR_ATIPICO_PECUARIO_{PERIODO}.xlsx",
         sas_p / f"VAR_ATIPICO_PECUARIO_{PERIODO}.XLSX", "PECUARIOS/VAR_ATIPICO"),
        (py_p / f"BASES INSUMOS PECUARIOS {PERIODO}.xlsx",
         sas_p / f"BASES INSUMOS PECUARIOS {PERIODO}.xlsx", "PECUARIOS/BASES"),
        (py_p / f"CUADROS INSUMOS PECUARIOS {PERIODO}.xlsx",
         sas_p / f"CUADROS INSUMOS PECUARIOS {PERIODO}.xlsx", "PECUARIOS/CUADROS"),
        (py_p / f"TABLAS INSUMOS PECUARIOS {PERIODO}.xlsx",
         sas_p / f"TABLAS INSUMOS PECUARIOS {PERIODO}.xlsx", "PECUARIOS/TABLAS"),
        (py_p / f"INSUMOS_PECUARIOS_DUPLI_{PERIODO}.xlsx",
         sas_p / f"INSUMOS_PECUARIOS_DUPLI_{PERIODO}.XLSX", "PECUARIOS/DUPLI"),
        (py_p / f"CV's INSUMOS PECUARIOS {PERIODO}.xlsx",
         sas_p / f"CV's INSUMOS PECUARIOS {PERIODO}.XLSX", "PECUARIOS/CVs"),
        (py_p / f"FALTAN_GRUPO_PECUARIO_{PERIODO}.xlsx",
         sas_p / f"FALTAN_GRUPO_PECUARIO_{PERIODO}.XLSX", "PECUARIOS/FALTAN_GRUPO"),
        (py_p / f"FALTAN_PUBLICA_PECUARIO_{PERIODO}.xlsx",
         sas_p / f"FALTAN_PUBLICA_PECUARIO_{PERIODO}.XLSX", "PECUARIOS/FALTAN_PUBLICA"),
        (py_p / f"ANEXO INSUMOS PECUARIOS {PERIODO}.xlsx",
         sas_p / f"ANEXO INSUMOS PECUARIOS {PERIODO}.xlsx", "PECUARIOS/ANEXO"),
        (py_p / f"INSU_PECUA_MAYORESQUE2_{PERIODO}.xlsx",
         sas_p / f"INSU_PECUA_MAYORESQUE2_{PERIODO}.XLSX", "PECUARIOS/MAYORESQUE2"),
        (py_p / f"INSU_PECUA_MENORESQUE2_{PERIODO}.xlsx",
         sas_p / f"INSU_PECUA_MENORESQUE2_{PERIODO}.XLSX", "PECUARIOS/MENORESQUE2"),
        (py_p / f"MAY_MEN2_PECUA_{PERIODO}.xlsx",
         sas_p / f"MAY_MEN2_PECUA_{PERIODO}.XLSX", "PECUARIOS/MAYMEN2"),
        # ── ELEMENTOS ─────────────────────────────────────────────────────
        (py_e / f"VAR_ATIPICO_ELEMENTOS_{PERIODO}.xlsx",
         sas_e / f"VAR_ATIPICO_ELEMENTOS_{PERIODO}.XLSX", "ELEMENTOS/VAR_ATIPICO"),
        (py_e / f"BASES ELEMENTOS AGROPECUARIOS {PERIODO}.xlsx",
         sas_e / f"BASES ELEMENTOS AGROPECUARIOS {PERIODO}.xlsx", "ELEMENTOS/BASES"),
        (py_e / f"CUADROS ELEMENTOS AGROPECUARIOS {PERIODO}.xlsx",
         sas_e / f"CUADROS ELEMENTOS AGROPECUARIOS {PERIODO}.xlsx", "ELEMENTOS/CUADROS"),
        (py_e / f"TABLAS ELEMENTOS {PERIODO}.xlsx",
         sas_e / f"TABLAS ELEMENTOS {PERIODO}.xlsx", "ELEMENTOS/TABLAS"),
        (py_e / f"ELEM_AGROPE_DUPLI_{PERIODO}.xlsx",
         sas_e / f"ELEM_AGROPE_DUPLI_{PERIODO}.XLSX", "ELEMENTOS/DUPLI"),
        (py_e / f"CV's_ELEMENTOS_{PERIODO}.xlsx",
         sas_e / f"CV's_ELEMENTOS_{PERIODO}.XLSX", "ELEMENTOS/CVs"),
        (py_e / f"FALTAN_PUBLICA_ELEMENTOS_{PERIODO}.xlsx",
         sas_e / f"FALTAN_PUBLICA_ELEMENTOS_{PERIODO}.XLSX", "ELEMENTOS/FALTAN_PUBLICA"),
        (py_e / f"ANEXO ELEMENTOS AGROPECUARIOS {PERIODO}.xlsx",
         sas_e / f"ANEXO ELEMENTOS AGROPECUARIOS {PERIODO}.XLSX", "ELEMENTOS/ANEXO"),
        (py_e / f"ELEM_AGROPE_MAYORESQUE2_{PERIODO}.xlsx",
         sas_e / f"ELEM_AGROPE_MAYORESQUE2_{PERIODO}.XLSX", "ELEMENTOS/MAYORESQUE2"),
        (py_e / f"ELEM_AGROPE_MENORESQUE2_{PERIODO}.xlsx",
         sas_e / f"ELEM_AGROPE_MENORESQUE2_{PERIODO}.XLSX", "ELEMENTOS/MENORESQUE2"),
        (py_e / f"ELEM_MAY_MEN2_{PERIODO}.xlsx",
         sas_e / f"ELEM_MAY_MEN2_{PERIODO}.XLSX", "ELEMENTOS/MAYMEN2"),
        # ── EMPAQUES ──────────────────────────────────────────────────────
        (py_em / f"VAR_ATIPICO_EMPAQUES_{PERIODO}.xlsx",
         sas_em / f"VAR_ATIPICO_EMPAQUES_{PERIODO}.XLSX", "EMPAQUES/VAR_ATIPICO"),
        (py_em / f"BASES EMPAQUES AGROPECUARIOS {PERIODO}.xlsx",
         sas_em / f"BASES EMPAQUES AGROPECUARIOS {PERIODO}.xlsx", "EMPAQUES/BASES"),
        (py_em / f"CUADROS EMPAQUES AGROPECUARIOS {PERIODO}.xlsx",
         sas_em / f"CUADROS EMPAQUES AGROPECUARIOS {PERIODO}.xlsx", "EMPAQUES/CUADROS"),
        (py_em / f"TABLAS EMPAQUES {PERIODO}.xlsx",
         sas_em / f"TABLAS EMPAQUES {PERIODO}.xlsx", "EMPAQUES/TABLAS"),
        (py_em / f"EMPA_AGROPE_DUPLI_{PERIODO}.xlsx",
         sas_em / f"EMPA_AGROPE_DUPLI_{PERIODO}.XLSX", "EMPAQUES/DUPLI"),
        (py_em / f"CV's_EMPAQUES_{PERIODO}.xlsx",
         sas_em / f"CV's_EMPAQUES_{PERIODO}.XLSX", "EMPAQUES/CVs"),
        (py_em / f"FALTAN_PUBLICA_EMPAQUES_{PERIODO}.xlsx",
         sas_em / f"FALTAN_PUBLICA_EMPAQUES_{PERIODO}.XLSX", "EMPAQUES/FALTAN_PUBLICA"),
        (py_em / f"ANEXO EMPAQUES AGROPECUARIOS {PERIODO}.xlsx",
         sas_em / f"ANEXO EMPAQUES AGROPECUARIOS {PERIODO}.XLSX", "EMPAQUES/ANEXO"),
        (py_em / f"EMPA_AGROPE_MAYORESQUE2_{PERIODO}.xlsx",
         sas_em / f"EMPA_AGROPE_MAYORESQUE2_{PERIODO}.XLSX", "EMPAQUES/MAYORESQUE2"),
        (py_em / f"EMPA_AGROPE_MENORESQUE2_{PERIODO}.xlsx",
         sas_em / f"EMPA_AGROPE_MENORESQUE2_{PERIODO}.XLSX", "EMPAQUES/MENORESQUE2"),
        (py_em / f"EMPA_MAY_MEN2_{PERIODO}.xlsx",
         sas_em / f"EMPA_MAY_MEN2_{PERIODO}.XLSX", "EMPAQUES/MAYMEN2"),
    ]


def _read_excel_safe(path: Path) -> dict[str, pd.DataFrame] | None:
    try:
        xl = pd.ExcelFile(str(path), engine="openpyxl")
        return {sh: xl.parse(sh) for sh in xl.sheet_names}
    except Exception:
        pass
    try:
        xl = pd.ExcelFile(str(path), engine="calamine")
        return {sh: xl.parse(sh) for sh in xl.sheet_names}
    except Exception:
        return None


def _norm_col(c: str) -> str:
    return " ".join(str(c).strip().lower().split())


def _norm_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all").copy()
    df.columns = [_norm_col(c) for c in df.columns]
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": "", "None": ""})
        elif pd.api.types.is_float_dtype(df[col]):
            df[col] = df[col].round(6)
    return df


def _sort_df(df: pd.DataFrame) -> pd.DataFrame:
    sort_cols = list(df.columns)
    try:
        return df.sort_values(by=sort_cols, na_position="first").reset_index(drop=True)
    except Exception:
        return df.reset_index(drop=True)


def comparar_hojas(df_py: pd.DataFrame, df_sas: pd.DataFrame, etiqueta: str) -> dict:
    result: dict = {
        "etiqueta": etiqueta, "py_filas": len(df_py), "sas_filas": len(df_sas),
        "py_cols": len(df_py.columns), "sas_cols": len(df_sas.columns),
        "ok": True, "diferencias": [],
    }
    df_py = _norm_df(df_py)
    df_sas = _norm_df(df_sas)

    cols_py = set(df_py.columns)
    cols_sas = set(df_sas.columns)
    solo_py = cols_py - cols_sas
    solo_sas = cols_sas - cols_py
    if solo_py or solo_sas:
        result["ok"] = False
        if solo_py:
            result["diferencias"].append(f"Cols solo en Python: {sorted(solo_py)}")
        if solo_sas:
            result["diferencias"].append(f"Cols solo en SAS: {sorted(solo_sas)}")

    if len(df_py) != len(df_sas):
        result["ok"] = False
        result["diferencias"].append(
            f"Filas Python={len(df_py)} vs SAS={len(df_sas)} (diff={len(df_py) - len(df_sas):+d})"
        )

    cols_comunes = sorted(cols_py & cols_sas)
    if not cols_comunes:
        return result

    df_py_c = _sort_df(df_py[cols_comunes])
    df_sas_c = _sort_df(df_sas[cols_comunes])

    if len(df_py_c) == len(df_sas_c):
        diffs_por_col: dict[str, int] = {}
        for col in cols_comunes:
            s_py = df_py_c[col]
            s_sas = df_sas_c[col]
            if pd.api.types.is_numeric_dtype(s_py) and pd.api.types.is_numeric_dtype(s_sas):
                mask = ~np.isclose(
                    s_py.fillna(np.nan), s_sas.fillna(np.nan),
                    rtol=1e-5, atol=1e-5, equal_nan=True,
                )
            else:
                mask = s_py.astype(str) != s_sas.astype(str)
            n_diff = int(mask.sum())
            if n_diff > 0:
                diffs_por_col[col] = n_diff
        if diffs_por_col:
            result["ok"] = False
            total = sum(diffs_por_col.values())
            top = sorted(diffs_por_col.items(), key=lambda x: -x[1])[:5]
            result["diferencias"].append(
                f"Valores distintos: {total} celdas en {len(diffs_por_col)} cols — top5: {top}"
            )
            result["diffs_por_col"] = diffs_por_col

    return result


def comparar_par(py_path: Path, sas_path: Path, etiqueta: str) -> list[dict]:
    resultados = []
    if not py_path.exists():
        resultados.append({"etiqueta": etiqueta, "ok": False,
                            "diferencias": [f"Archivo Python no existe: {py_path.name}"]})
        return resultados
    if not sas_path.exists():
        resultados.append({"etiqueta": etiqueta, "ok": False,
                            "diferencias": [f"Archivo SAS no existe: {sas_path.name}"]})
        return resultados

    hojas_py = _read_excel_safe(py_path)
    hojas_sas = _read_excel_safe(sas_path)
    if hojas_py is None:
        resultados.append({"etiqueta": etiqueta, "ok": False, "diferencias": ["Error leyendo archivo Python"]})
        return resultados
    if hojas_sas is None:
        resultados.append({"etiqueta": etiqueta, "ok": False, "diferencias": ["Error leyendo archivo SAS"]})
        return resultados

    hojas_py_norm = {_norm_col(k): (k, v) for k, v in hojas_py.items()}
    hojas_sas_norm = {_norm_col(k): (k, v) for k, v in hojas_sas.items()}
    keys_py = set(hojas_py_norm)
    keys_sas = set(hojas_sas_norm)
    solo_py_sh = keys_py - keys_sas
    solo_sas_sh = keys_sas - keys_py
    comunes = keys_py & keys_sas

    if len(hojas_py) == 1 and len(hojas_sas) == 1 and not comunes:
        # Un único par de hojas con nombres distintos (convención SAS vs
        # Python difiere) — comparar por posición en vez de por nombre.
        (nombre_py, df_py), = hojas_py.items()
        (nombre_sas, df_sas), = hojas_sas.items()
        label = f"{etiqueta} [{nombre_py} vs {nombre_sas}]"
        resultados.append(comparar_hojas(df_py, df_sas, label))
        return resultados

    if solo_py_sh or solo_sas_sh:
        resultados.append({
            "etiqueta": etiqueta, "ok": False,
            "diferencias": [
                f"Hojas solo en Python: {sorted(solo_py_sh)}" if solo_py_sh else "",
                f"Hojas solo en SAS: {sorted(solo_sas_sh)}" if solo_sas_sh else "",
            ],
        })

    for key in sorted(comunes):
        _, df_py = hojas_py_norm[key]
        _, df_sas = hojas_sas_norm[key]
        label = f"{etiqueta} [{hojas_py_norm[key][0]}]"
        resultados.append(comparar_hojas(df_py, df_sas, label))

    return resultados


def _separador(texto: str) -> None:
    print(f"\n{'=' * 70}")
    print(f"  {texto}")
    print("=" * 70)


def main(filtro_modulo: str = "") -> None:
    pares = _pares()
    if filtro_modulo:
        pares = [p for p in pares if filtro_modulo.upper() in p[2].upper()]

    totales = {"ok": 0, "fail": 0}
    resumen: list[dict] = []

    for py_path, sas_path, etiqueta in pares:
        resultados = comparar_par(py_path, sas_path, etiqueta)
        for r in resultados:
            status = "OK" if r["ok"] else "FAIL"
            totales["ok" if r["ok"] else "fail"] += 1
            resumen.append({**r, "status": status})

    _separador(f"VALIDACIÓN PYTHON vs SAS — SIPSA Insumos {PERIODO}")
    print(f"\n{'Etiqueta':<55} {'Filas PY':>9} {'Filas SAS':>9}  Estado")
    print("-" * 90)
    for r in resumen:
        py_f = r.get("py_filas", "?")
        sas_f = r.get("sas_filas", "?")
        status = r["status"]
        marker = "  " if status == "OK" else "! "
        print(f"{marker}{r['etiqueta']:<53} {str(py_f):>9} {str(sas_f):>9}  {status}")

    fallos = [r for r in resumen if r["status"] == "FAIL"]
    if fallos:
        _separador(f"DETALLE DE DIFERENCIAS ({len(fallos)} archivos/hojas con diff)")
        for r in fallos:
            print(f"\n  [{r['etiqueta']}]")
            for d in r.get("diferencias", []):
                if d:
                    print(f"    - {d}")
    else:
        _separador("RESULTADO: TODOS LOS ARCHIVOS SON EXACTAMENTE IGUALES")

    _separador("RESUMEN FINAL")
    print(f"  OK:    {totales['ok']}")
    print(f"  FAIL:  {totales['fail']}")
    total = totales["ok"] + totales["fail"]
    pct = 100 * totales["ok"] / total if total else 0
    print(f"  TOTAL: {total}  ({pct:.1f}% OK)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--modulo", default="", help="Filtrar por módulo (ej: agricolas)")
    args = parser.parse_args()
    main(filtro_modulo=args.modulo)
