"""Nodos del pipeline Sin Precio Anterior — SIPSA Insumos.

Réplica del programa SAS "PROGRAMA REVISIÓN SIN PRECIO <MÓDULO> <MES>.sas"
(OneDrive SIPSA Insumos/sin_precio_ant/<año>/03Programas). Por módulo:

  1. Leer la serie histórica "para revisiones" (formato largo) y borrar filas vacías.
  2. ID = compress(CodigoMpio||Fuente||Articulo||CasaCom||RegICA||UnMed) y
     'Mes-año1' = mes en formato ESPDFMY. (p.ej. "AGO2026").
  3. PROC TRANSPOSE por ID → una columna Precio_<MMMAAAA> por mes, en el orden
     en que aparecen los meses al recorrer los ID ordenados.
  4. Filtrar VAR_ATIPICO: REVISA=2, Nov. no en (IA, IN) y precio del mes presente;
     el precio del mes pasa a llamarse Precio_<PERIODO>.
  5. MERGE por ID (if a): cada registro de VAR_ATIPICO con su historial.
  6. DATA step de salida: RETAIN de las columnas base y de la lista FIJA de meses
     del programa, luego el resto de columnas, y al final las variaciones
     Var_<PERIODO><sufijo> = (Precio_<PERIODO>/Precio_<mes>-1)*100 solo contra
     los meses que lista el programa.

La lista de meses y de variaciones está escrita a mano en cada programa SAS
(no se actualizó en 2026, así que no hay variaciones contra meses de 2026);
se replica tal cual desde parameters_sin_precio_ant.yml (formato_sas) para que
el archivo sea idéntico al de SAS.

Módulos y periodicidad:
  AGRICOLAS   — mensual (todos los meses)
  PECUARIOS   — mensual (todos los meses)
  ELEMENTOS   — bimestral impar (ENE, MAR, MAY, JUL, SEP, NOV)
  MATERIAL    — bimestral par   (FEB, ABR, JUN, AGO, OCT, DIC)
"""
from __future__ import annotations

import logging
from pathlib import Path

import numpy as np
import pandas as pd

log = logging.getLogger(__name__)

# Equivalente al formato SAS ESPDFMY.: 3 letras español + 4 dígitos año → ej. "MAY2026"
_MESES_ABBR: dict[int, str] = {
    1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR",
    5: "MAY", 6: "JUN", 7: "JUL", 8: "AGO",
    9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC",
}


def _fecha_a_mesc(dt) -> str:
    """datetime → 'MAY2026'  (equiv. SAS compress(upcase(put(date, ESPDFMY.))))."""
    if pd.isna(dt):
        return ""
    try:
        return f"{_MESES_ABBR[dt.month]}{dt.year}"
    except (AttributeError, KeyError):
        return ""


def _es_numerica(serie: pd.Series) -> bool:
    """PROC IMPORT de SAS crea la columna numérica si todas sus celdas con valor
    son números; si alguna es texto, la columna es de caracteres."""
    valores = serie.dropna()
    return len(valores) > 0 and all(
        isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) for v in valores
    )


def _num_texto(v) -> str:
    """Número como lo deja SAS al concatenarlo (BEST12. sin espacios)."""
    f = float(v)
    return str(int(f)) if f.is_integer() else repr(f)


def _a_texto_sas(serie: pd.Series) -> pd.Series:
    """Valor de la columna tal como entra en la concatenación del ID en SAS:
    numérica → número sin decimales sobrantes, faltante '.'; texto → faltante ''."""
    if _es_numerica(serie):
        return serie.map(lambda v: "." if pd.isna(v) else _num_texto(v))
    return serie.map(lambda v: "" if pd.isna(v) else (_num_texto(v) if isinstance(v, (int, float)) else str(v)))


def _construir_id(df: pd.DataFrame, cols: list[str], largo: int | None = None) -> pd.Series:
    """compress(c1||c2||...): concatena y quita todos los espacios; con `largo`
    se trunca como una variable SAS declarada con LENGTH ID $largo."""
    partes = [_a_texto_sas(df[c]) for c in cols]
    ids = partes[0]
    for p in partes[1:]:
        ids = ids + p
    ids = ids.str.replace(" ", "", regex=False)
    return ids.str.slice(0, largo) if largo else ids


def _detectar_col_fecha(columns: list[str]) -> str | None:
    """Detecta la columna 'Mes-año' tolerando distintas codificaciones de la ñ."""
    for c in columns:
        if str(c).lower().startswith("mes") and "anterio" not in str(c).lower():
            return c
    return None


def _col(df: pd.DataFrame, *candidatos: str) -> str:
    return next(c for c in candidatos if c in df.columns)


def revisar_sin_precio(
    ruta_historico: str,
    hoja_historico: str,
    ruta_var_atipico: str,
    hoja_var_atipico: str,
    casacom_col_hist: str,
    unmed_col_hist: str,
    periodo: str,
    modulo: str,
    ruta_reporting: str,
    mes_actual: str = "",
    activo: bool = True,
    formato_sas: dict | None = None,
) -> pd.DataFrame:
    """Genera REV_SIN_PRECIO_ANTE_<MÓDULO>_<PERIODO>.xlsx idéntico al de SAS.

    Args:
        ruta_historico:   Excel "para revisiones" (serie histórica larga).
        hoja_historico:   Hoja de ese archivo (ej: 'Agrícolas').
        ruta_var_atipico: VAR_ATIPICO_<MÓDULO>_<PERIODO>.xlsx del período.
        hoja_var_atipico: Hoja de ese archivo.
        casacom_col_hist: Columna CasaCom en el histórico (ej: 'CasaCom.').
        unmed_col_hist:   Columna UnMed en el histórico (ej: 'UnMed.').
        periodo:          Código del período MMMAAAA (ej: 'AGO2026').
        modulo:           Nombre del módulo (ej: 'AGRICOLAS').
        ruta_reporting:   Directorio raíz de reportes Kedro.
        mes_actual:       Nombre del mes actual (columna de precio en VAR_ATIPICO).
        activo:           False = el módulo no aplica este período.
        formato_sas:      nombre_salida, largo_id, columnas_base, meses_fijos y
                          variaciones, tal como están en el programa SAS.

    Returns:
        DataFrame de metadatos: archivo, módulo, período, filas generadas.
    """
    if not activo:
        log.info("[%s] Módulo no activo en %s — omitido.", modulo, periodo)
        return pd.DataFrame([{"modulo": modulo, "periodo": periodo, "filas": 0, "archivo": ""}])

    fmt = formato_sas or {}
    nombre_salida = fmt.get("nombre_salida", modulo)
    precio_actual = f"Precio_{periodo}"

    # ── 1-2. Histórico: ID y mes ───────────────────────────────────────────────
    log.info("[%s] Leyendo histórico: %s | hoja='%s'", modulo, ruta_historico, hoja_historico)
    insumos = pd.read_excel(ruta_historico, sheet_name=hoja_historico, engine="calamine", dtype=object,
                            keep_default_na=False, na_values=[""])
    insumos = insumos.rename(columns={"Codigo": "CodigoMpio"})
    col_fecha = _detectar_col_fecha(list(insumos.columns))
    if col_fecha is None:
        raise ValueError(f"[{modulo}] No se encontró columna Mes-año en '{hoja_historico}'.")
    vacia = (
        insumos["CodigoMpio"].isna()
        & insumos["Articulo"].fillna("").eq("")
        & insumos["RegICA"].isna()
        & insumos[unmed_col_hist].fillna("").eq("")
    )
    if "Grupo" in insumos.columns:
        vacia &= insumos["Grupo"].fillna("").eq("")
    insumos = insumos[~vacia].copy()
    insumos["_mesc"] = pd.to_datetime(insumos[col_fecha], errors="coerce").map(_fecha_a_mesc)
    insumos["_ID"] = _construir_id(
        insumos, ["CodigoMpio", "Fuente", "Articulo", casacom_col_hist, "RegICA", unmed_col_hist])
    log.info("[%s] Histórico | filas=%d", modulo, len(insumos))

    # ── 3. PROC TRANSPOSE by ID; id Mes-año1 ──────────────────────────────────
    hist = insumos[["_ID", "_mesc", "Precio_mes"]].sort_values("_ID", kind="mergesort")
    meses_hist = list(dict.fromkeys(m for m in hist["_mesc"] if m))
    hist["Precio_mes"] = pd.to_numeric(hist["Precio_mes"], errors="coerce")
    ancho = (hist[hist["_mesc"] != ""]
             .drop_duplicates(["_ID", "_mesc"], keep="last")
             .pivot(index="_ID", columns="_mesc", values="Precio_mes"))
    ancho = ancho.reindex(columns=meses_hist)
    ancho.columns = [f"Precio_{m}" for m in meses_hist]
    ancho = ancho.reset_index()

    # ── 4. VAR_ATIPICO filtrado ────────────────────────────────────────────────
    log.info("[%s] Leyendo VAR_ATIPICO: %s | hoja='%s'", modulo, ruta_var_atipico, hoja_var_atipico)
    var_at = pd.read_excel(ruta_var_atipico, sheet_name=hoja_var_atipico, engine="calamine", dtype=object,
                            keep_default_na=False, na_values=[""])
    col_nov = _col(var_at, "Nov.", "Nov#")
    col_casa = _col(var_at, "CasaCom.", "CasaCom#")
    col_unmed = _col(var_at, "UnMed.", "UnMed#")
    revisa = pd.to_numeric(var_at["REVISA"], errors="coerce")
    precio_mes = pd.to_numeric(var_at[mes_actual], errors="coerce")
    var_at1 = var_at[(revisa == 2) & ~var_at[col_nov].isin(["IA", "IN"]) & precio_mes.notna()].copy()
    var_at1.insert(0, "ID", _construir_id(
        var_at1, ["CodigoMpio", "Fuente", "Articulo", col_casa, "RegICA", col_unmed], fmt.get("largo_id")))
    var_at1 = var_at1.rename(columns={mes_actual: precio_actual})
    var_at1 = var_at1.sort_values("ID", kind="mergesort")
    log.info("[%s] VAR_ATIPICO filtrado (REVISA=2, Nov. no IA/IN, con precio) | filas=%d", modulo, len(var_at1))

    # ── 5. MERGE var_atipico1 (in=a) insumos3; by ID; if a ─────────────────────
    res = var_at1.merge(ancho.rename(columns={"_ID": "ID"}), on="ID", how="left", suffixes=("", "__h"))
    if f"{precio_actual}__h" in res.columns:
        # Variable en ambas tablas: si el ID cruzó, SAS deja el valor del histórico.
        cruzo = res["ID"].isin(ancho["_ID"])
        res.loc[cruzo, precio_actual] = res.loc[cruzo, f"{precio_actual}__h"]
        res = res.drop(columns=f"{precio_actual}__h")
    log.info("[%s] Merge OK | filas=%d", modulo, len(res))

    # ── 6. DATA step de salida (RETAIN + variaciones) ─────────────────────────
    por_mes = {c[7:].upper(): c for c in res.columns if c.startswith("Precio_")}
    variaciones = list(dict.fromkeys(fmt.get("variaciones", [])))
    meses_var = {s.lstrip("_").upper() for s in variaciones}
    salida: dict[str, pd.Series] = {}
    usadas: set[str] = set()

    for c in fmt.get("columnas_base", []):
        if c in res.columns:
            salida[c] = res[c]
            usadas.add(c)
    for m in fmt.get("meses_fijos", []):
        # Una variable del RETAIN solo queda en la salida si existe en los datos
        # o si alguna variación la usa (en ese caso, vacía).
        origen = por_mes.get(m.upper())
        if origen is not None:
            salida[f"Precio_{m}"] = res[origen]
            usadas.add(origen)
        elif m.upper() in meses_var:
            salida[f"Precio_{m}"] = pd.Series(np.nan, index=res.index)
    for c in res.columns:
        if c not in usadas and c not in salida:
            salida[c] = res[c]
    actual = pd.to_numeric(res[precio_actual], errors="coerce")
    for s in variaciones:
        ref = por_mes.get(s.lstrip("_").upper())
        base = pd.to_numeric(res[ref], errors="coerce") if ref else pd.Series(np.nan, index=res.index)
        salida[f"Var_{periodo}{s}"] = (actual / base.where(base != 0) - 1) * 100
    result = pd.DataFrame(salida)

    # ── Exportar (PROC EXPORT: la hoja toma el nombre del archivo, máx. 31) ───
    nombre = f"REV_SIN_PRECIO_ANTE_{nombre_salida}_{periodo}"
    ruta_out = Path(ruta_reporting) / "sin_precio_ant" / f"{nombre}.xlsx"
    ruta_out.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(str(ruta_out), index=False, sheet_name=nombre[:31])
    log.info("[%s] Exportado | archivo=%s | filas=%d | columnas=%d",
             modulo, ruta_out.name, len(result), result.shape[1])

    return pd.DataFrame([{"modulo": modulo, "periodo": periodo, "filas": len(result), "archivo": str(ruta_out)}])
