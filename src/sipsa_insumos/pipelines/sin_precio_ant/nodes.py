"""Nodos del pipeline Sin Precio Anterior — SIPSA Insumos.

SAS equivalente:
  PROGRAMA REVISIÓN SIN PRECIO INSUMOS AGRÍCOLAS [PERIODO].sas
  PROGRAMA REVISIÓN SIN PRECIO INSUMOS PECUARIOS [PERIODO].sas
  PROGRAMA REVISIÓN SIN PRECIO ELEMENTOS [PERIODO].sas
  PROGRAMA REVISIÓN SIN PRECIO MATERIAL PROPAGA [PERIODO].sas

Flujo por módulo:
  1. Leer serie histórica "para revisiones" (formato largo)
  2. Construir clave ID compuesta y convertir Mes-año a código MMMYYYY
  3. Pivotar largo → ancho: una columna Precio_MMMYYYY por período
  4. Leer VAR_ATIPICO del período (ya revisado por los analistas)
  5. Filtrar: REVISA=2, Nov. no en {IA, IN}, precio mes actual no nulo
  6. Construir misma clave ID en VAR_ATIPICO
  7. Left-join por ID: registros VAR_ATIPICO + columnas históricas de precio
  8. Calcular variaciones % vs cada período de referencia histórico
  9. Exportar REV_SIN_PRECIO_ANTE_[MODULO]_[PERIODO].xlsx

Módulos y periodicidad:
  AGRICOLAS   — mensual (todos los meses)
  PECUARIOS   — mensual (todos los meses)
  ELEMENTOS   — bimestral impar (ENE, MAR, MAY, JUL, SEP, NOV)
  MATERIAL    — bimestral par   (FEB, ABR, JUN, AGO, OCT, DIC)
"""
from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

log = logging.getLogger(__name__)

# Equivalente al formato SAS ESPDFMY.: 3 letras español + 4 dígitos año → ej. "MAY2026"
_MESES_ABBR: dict[int, str] = {
    1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR",
    5: "MAY", 6: "JUN", 7: "JUL", 8: "AGO",
    9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC",
}
_ABBR_MES: dict[str, int] = {v: k for k, v in _MESES_ABBR.items()}


def _fecha_a_mesc(dt) -> str:
    """datetime → 'MAY2026'  (equiv. SAS put(date, ESPDFMY.) + upcase + compress)."""
    if pd.isna(dt):
        return ""
    try:
        return f"{_MESES_ABBR[dt.month]}{dt.year}"
    except (AttributeError, KeyError):
        return ""


def _num_str(val) -> str:
    """Numeric/NaN → string equiv. SAS put(num, BEST12.) + compress().
    Valor faltante → '.'  (como SAS cuando concatena un numérico faltante).
    Valor presente → entero sin espacios ni decimales.
    """
    if pd.isna(val):
        return "."
    try:
        return str(int(float(val)))
    except (ValueError, OverflowError, TypeError):
        return str(val).strip()


def _construir_id(
    df: pd.DataFrame,
    codigo_col: str,
    casacom_col: str,
    unmed_col: str,
) -> pd.Series:
    """Clave compuesta equiv. SAS compress(CodigoMpio||Fuente||Articulo||CasaCom.||RegICA||UnMed.).

    SAS compress() elimina todos los espacios de la concatenación.
    """
    return (
        df[codigo_col].apply(_num_str)
        + df["Fuente"].fillna("")
        + df["Articulo"].fillna("")
        + df[casacom_col].fillna("")
        + df["RegICA"].apply(_num_str)
        + df[unmed_col].fillna("")
    ).str.replace(" ", "", regex=False)


def _detectar_col_fecha(columns: list[str]) -> str | None:
    """Detecta la columna 'Mes-año' tolerando distintas codificaciones de la ñ."""
    for c in columns:
        low = c.lower()
        if "mes" in low and any(x in low for x in ("a", "ñ", "\xf1", "ao")):
            if low.startswith("mes"):
                return c
    return None


def _ordenar_precio_cols(cols: list[str]) -> list[str]:
    """Ordena columnas Precio_MMMYYYY de más reciente a más antigua."""
    def _key(c: str) -> int:
        p = c.removeprefix("Precio_")
        mes_abbr = p[:3]
        try:
            anio = int(p[3:7])
            mes_num = _ABBR_MES.get(mes_abbr, 0)
            return -(anio * 12 + mes_num)
        except (ValueError, IndexError):
            return 0

    return sorted(cols, key=_key)


def revisar_sin_precio(
    ruta_historico: str,
    hoja_historico: str,
    ruta_var_atipico: str,
    hoja_var_atipico: str,
    casacom_col_hist: str,
    unmed_col_hist: str,
    casacom_col_var: str,
    unmed_col_var: str,
    nov_col: str,
    mes_actual: str,
    periodo: str,
    modulo: str,
    ruta_reporting: str,
    activo: bool = True,
) -> pd.DataFrame:
    """Genera REV_SIN_PRECIO_ANTE_[MODULO]_[PERIODO].xlsx para un módulo SIPSA.

    Replica fielmente los programas SAS de revisión sin precio anterior.

    Args:
        ruta_historico:   Ruta al Excel "para revisiones" (serie histórica larga).
        hoja_historico:   Hoja dentro de ese archivo (ej: 'Agrícolas').
        ruta_var_atipico: Ruta al Excel VAR_ATIPICO del período (revisado por analistas).
        hoja_var_atipico: Hoja dentro de ese archivo (ej: 'VAR_ATIPICO_AGRICOLA').
        casacom_col_hist: Columna CasaCom en el histórico (ej: 'CasaCom.').
        unmed_col_hist:   Columna UnMed en el histórico (ej: 'UnMed.').
        casacom_col_var:  Columna CasaCom en VAR_ATIPICO (ej: 'CasaCom.' o 'CasaCom#').
        unmed_col_var:    Columna UnMed en VAR_ATIPICO (ej: 'UnMed.' o 'UnMed#').
        nov_col:          Columna novedad en VAR_ATIPICO (ej: 'Nov.' o 'Nov#').
        mes_actual:       Nombre del mes actual en español (ej: 'Mayo').
        periodo:          Código del período MMMYYYY (ej: 'MAY2026').
        modulo:           Nombre del módulo en mayúsculas (ej: 'AGRICOLAS').
        ruta_reporting:   Directorio raíz de reportes Kedro.
        activo:           False = módulo no aplica este período (omite procesamiento).

    Returns:
        DataFrame de metadatos: archivo, módulo, período, filas generadas.
    """
    if not activo:
        log.info("[%s] Módulo no activo en %s — omitido.", modulo, periodo)
        return pd.DataFrame([{
            "modulo": modulo, "periodo": periodo, "filas": 0, "archivo": "",
        }])

    # ── 1. Cargar serie histórica "para revisiones" ────────────────────────────
    log.info("[%s] Leyendo histórico: %s | hoja='%s'", modulo, ruta_historico, hoja_historico)
    insumos = pd.read_excel(ruta_historico, sheet_name=hoja_historico)
    log.info("[%s] Histórico cargado | filas_raw=%d", modulo, len(insumos))

    # ── 2. Preparar histórico ──────────────────────────────────────────────────
    # Renombrar Codigo → CodigoMpio (equiv. SAS: CodigoMpio = Codigo)
    if "Codigo" in insumos.columns and "CodigoMpio" not in insumos.columns:
        insumos = insumos.rename(columns={"Codigo": "CodigoMpio"})

    # Detectar columna Mes-año y convertir a código MMMYYYY (equiv. SAS ESPDFMY.)
    col_fecha = _detectar_col_fecha(list(insumos.columns))
    if col_fecha is None:
        raise ValueError(
            f"[{modulo}] No se encontró columna Mes-año en '{hoja_historico}'. "
            f"Columnas disponibles: {list(insumos.columns)}"
        )
    insumos["_mesc"] = pd.to_datetime(insumos[col_fecha], errors="coerce").apply(_fecha_a_mesc)

    # Construir clave ID (equiv. SAS compress(CodigoMpio||Fuente||...))
    insumos["_ID"] = _construir_id(insumos, "CodigoMpio", casacom_col_hist, unmed_col_hist)

    # Eliminar filas vacías
    # SAS: if Codigo=. and Articulo='' and RegICA=. and 'UnMed.'n='' then delete;
    mask_vacia = (
        insumos["CodigoMpio"].isna()
        & (insumos["Articulo"].fillna("") == "")
        & insumos["RegICA"].isna()
        & (insumos[unmed_col_hist].fillna("") == "")
    )
    insumos = insumos[~mask_vacia].copy()
    log.info("[%s] Histórico filtrado (sin filas vacías) | filas=%d", modulo, len(insumos))

    # ── 3. Pivotar largo → ancho (equiv. SAS PROC TRANSPOSE) ──────────────────
    # SAS: by ID; id 'Mes-año1'n; var Precio_mes; prefix=Precio_
    insumos2 = (
        insumos[["_ID", "_mesc", "Precio_mes"]]
        .query("_mesc != ''")
        .dropna(subset=["Precio_mes"])
        .copy()
    )
    # Agregar: mismo ID+período puede tener varias filas → usar último precio
    wide = (
        insumos2.pivot_table(
            index="_ID", columns="_mesc", values="Precio_mes", aggfunc="last"
        )
        .reset_index()
    )
    wide.columns = ["_ID"] + [f"Precio_{c}" for c in wide.columns[1:]]
    log.info(
        "[%s] Pivot OK | IDs únicos=%d | períodos históricos=%d",
        modulo, len(wide), len(wide.columns) - 1,
    )

    # ── 4. Cargar VAR_ATIPICO ──────────────────────────────────────────────────
    log.info(
        "[%s] Leyendo VAR_ATIPICO: %s | hoja='%s'", modulo, ruta_var_atipico, hoja_var_atipico
    )
    var_at = pd.read_excel(ruta_var_atipico, sheet_name=hoja_var_atipico)
    log.info("[%s] VAR_ATIPICO cargado | filas_raw=%d", modulo, len(var_at))

    # ── 5. Filtrar VAR_ATIPICO ─────────────────────────────────────────────────
    # SAS: where REVISA = 2 and 'Nov.'n not in ('IA','IN') and &Mes. NE .
    if mes_actual not in var_at.columns:
        raise KeyError(
            f"[{modulo}] Columna '{mes_actual}' no encontrada en VAR_ATIPICO. "
            f"Columnas disponibles: {list(var_at.columns)}"
        )
    mask = (
        (var_at["REVISA"] == 2)
        & (~var_at[nov_col].fillna("").isin(["IA", "IN"]))
        & var_at[mes_actual].notna()
    )
    var_at1 = var_at[mask].copy()
    log.info("[%s] VAR_ATIPICO filtrado (REVISA=2, Nov valida, precio no nulo) | filas=%d", modulo, len(var_at1))

    if var_at1.empty:
        log.warning("[%s] Sin registros con REVISA=2 para el período %s.", modulo, periodo)
        return pd.DataFrame([{
            "modulo": modulo, "periodo": periodo, "filas": 0,
            "archivo": f"REV_SIN_PRECIO_ANTE_{modulo}_{periodo}.xlsx (vacío)",
        }])

    # Renombrar columna del mes actual → Precio_{periodo}
    # SAS: rename &Mes.=Precio_&Mesc.
    precio_actual_col = f"Precio_{periodo}"
    var_at1 = var_at1.rename(columns={mes_actual: precio_actual_col})

    # Construir ID en VAR_ATIPICO
    var_at1["_ID"] = _construir_id(var_at1, "CodigoMpio", casacom_col_var, unmed_col_var)

    # ── 6. Left-join por ID (equiv. SAS MERGE con if a) ───────────────────────
    # El Precio_{periodo} del VAR_ATIPICO prevalece; si el histórico también lo tiene se descarta
    wide_sin_actual = wide.drop(
        columns=[c for c in wide.columns if c == precio_actual_col],
        errors="ignore",
    )
    result = var_at1.merge(wide_sin_actual, on="_ID", how="left")
    log.info(
        "[%s] Merge OK | filas=%d | cols_hist=%d",
        modulo, len(result), len([c for c in result.columns if c.startswith("Precio_")]) - 1,
    )

    # ── 7. Calcular variaciones % (equiv. SAS Var_&Mesc._REFPER=((Precio_&Mesc./Precio_REF)-1)*100)
    hist_cols = _ordenar_precio_cols(
        [c for c in result.columns if c.startswith("Precio_") and c != precio_actual_col]
    )
    for hist_col in hist_cols:
        ref = hist_col.removeprefix("Precio_")
        result[f"Var_{periodo}_{ref}"] = (
            (result[precio_actual_col] / result[hist_col]) - 1
        ) * 100

    # ── 8. Ordenar columnas (equiv. SAS retain) ────────────────────────────────
    var_cols = [
        f"Var_{periodo}_{h.removeprefix('Precio_')}"
        for h in hist_cols
        if f"Var_{periodo}_{h.removeprefix('Precio_')}" in result.columns
    ]
    context_cols = [
        c for c in result.columns
        if not c.startswith("Precio_")
        and not c.startswith("Var_")
        and c not in ("_ID", "_mesc")
    ]
    col_order = ["_ID"] + context_cols + [precio_actual_col] + hist_cols + var_cols
    col_order = [c for c in col_order if c in result.columns]
    result = result[col_order].rename(columns={"_ID": "ID"})

    # ── 9. Exportar Excel ──────────────────────────────────────────────────────
    nombre_archivo = f"REV_SIN_PRECIO_ANTE_{modulo}_{periodo}.xlsx"
    ruta_out = Path(ruta_reporting) / "sin_precio_ant" / nombre_archivo
    ruta_out.parent.mkdir(parents=True, exist_ok=True)

    sheet_name = f"REV_{modulo}"[:31]  # Excel sheet names ≤31 chars
    result.to_excel(str(ruta_out), index=False, sheet_name=sheet_name)
    log.info("[%s] Exportado | archivo=%s | filas=%d", modulo, nombre_archivo, len(result))

    return pd.DataFrame([{
        "modulo": modulo,
        "periodo": periodo,
        "filas": len(result),
        "archivo": str(ruta_out),
    }])
