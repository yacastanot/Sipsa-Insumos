"""Nodos del pipeline de calidad — SIPSA Insumos.

SAS equivalente (pasos 6-9):
  - Calcular variaciones atípicas
  - Calcular CV por municipio-producto
  - Identificar duplicados

Flujo:
  base_enriquecida ──► [detectar_duplicados] ──► base_sin_dupli + duplicados
  base_sin_dupli ──► [calcular_cv] ──► base_con_cv + cvs_reporte
  base_con_cv + mayor2_anterior ──► [detectar_var_atipica] ──► base_calidad + var_atipico
"""
from __future__ import annotations

import logging

import pandas as pd

log = logging.getLogger(__name__)

# Llave compuesta para detección de duplicados (equivale a SAS PROC SORT nodup)
# Incluye PRECIO para solo eliminar registros completamente idénticos.
# Registros del mismo vendedor con diferente precio son observaciones válidas.
_LLAVE_DUPLICADOS = [
    "CÓDIGO DIVIPOLA",
    "CÓDIGO CPC",
    "ARTÍCULO",
    "CASA COMERCIAL",
    "REGISTRO ICA",
    "UNIDAD DE MEDIDA",
    "PRECIO",
]

# Columnas de agrupación para el CV por municipio-producto (nivel Nombre_Publica)
# Equivale al nivel de agregación SAS: (municipio, Nombre_productos_agr_publ)
_LLAVE_CV = ["CÓDIGO DIVIPOLA", "Nombre_Publica"]

# Columnas adicionales de contexto que se incluyen en el reporte de CVs
_COLS_CV_CONTEXTO = [
    "NombreDepartamento", "NombreMunicipio", "CÓDIGO CPC", "Grupo",
]


def detectar_duplicados(
    base_enriquecida: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Identifica registros duplicados por llave compuesta.

    SAS: PROC SORT nodup; BY CodigoMpio CPC ARTICULO CASACOM ICA UnMed;

    Args:
        base_enriquecida: DataFrame con todas las columnas de enriquecimiento.

    Returns:
        (base_sin_dupli, duplicados): sin duplicados + los registros duplicados.
    """
    llave_existente = [c for c in _LLAVE_DUPLICADOS if c in base_enriquecida.columns]

    es_dupli = base_enriquecida.duplicated(subset=llave_existente, keep="first")
    duplicados = base_enriquecida[es_dupli].copy()
    base_sin_dupli = base_enriquecida[~es_dupli].copy()

    if len(duplicados) > 0:
        log.warning(
            "detectar_duplicados | %d registros duplicados encontrados",
            len(duplicados),
        )

    log.info(
        "detectar_duplicados OK | únicos=%d | duplicados=%d",
        len(base_sin_dupli), len(duplicados),
    )
    return base_sin_dupli, duplicados


def calcular_cv(
    base_sin_dupli: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Calcula el Coeficiente de Variación por municipio-producto.

    SAS: PROC SQL — CV = STD/MEAN*100 por (CodigoMpio, Art_Unmed_Casacomer_ICA)

    Solo se calcula cuando N >= 2 (un solo precio no tiene CV).

    Args:
        base_sin_dupli: DataFrame sin duplicados con columna 'PRECIO'.

    Returns:
        (base_con_cv, cvs_reporte): base original con columna 'CV' agregada +
            DataFrame del reporte de CVs (N, MIN, MAX, MEAN, CV) por grupo.
    """
    llave_cv = [c for c in _LLAVE_CV if c in base_sin_dupli.columns]
    cols_contexto = [c for c in _COLS_CV_CONTEXTO if c in base_sin_dupli.columns]
    # Incluir columnas de contexto en el groupby para que aparezcan en el reporte
    llave_completa = llave_cv + [c for c in cols_contexto if c not in llave_cv]

    agg = (
        base_sin_dupli.groupby(llave_completa, dropna=False)["PRECIO"]
        .agg(N="count", MIN="min", MAX="max", MEAN="mean", STD="std")
        .reset_index()
    )
    agg["CV"] = (agg["STD"] / agg["MEAN"] * 100).where(agg["N"] >= 2)
    agg = agg.rename(columns={"MEAN": "PRECIO_PROMEDIO_CV"})

    # Unir CV de vuelta al DataFrame original usando solo la llave principal
    base_con_cv = base_sin_dupli.merge(
        agg[llave_cv + ["N", "CV"]],
        on=llave_cv,
        how="left",
    )

    cvs_altos = agg[agg["CV"].notna() & (agg["CV"] > 30)]
    if len(cvs_altos) > 0:
        log.warning(
            "calcular_cv | %d grupos con CV > 30%%",
            len(cvs_altos),
        )

    log.info(
        "calcular_cv OK | grupos=%d | con_cv=%d | cv_max=%.1f",
        len(agg),
        agg["CV"].notna().sum(),
        agg["CV"].max() if agg["CV"].notna().any() else 0,
    )
    return base_con_cv, agg


def detectar_var_atipica(
    base_con_cv: pd.DataFrame,
    mayor2_anterior: pd.DataFrame | None,
    umbral_var_alta: float,
    umbral_var_baja: float,
    umbral_var_extrema: float,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Calcula la variación respecto al período anterior e identifica atípicas.

    SAS:
      REVISA = 1 if VAR >= 25 OR VAR <= -25
      REVISA = 2 if VAR = .  (sin precio anterior)
      REVISA = 3 if VAR >= 100

    La variación se calcula a nivel de precio promedio por municipio-producto.
    Si mayor2_anterior es None o vacío (primer período), todos quedan con REVISA=2.

    Args:
        base_con_cv: DataFrame con columnas PRECIO y LLAVE_ARTICULO.
        mayor2_anterior: mayor2 del período anterior (puede ser None si es el primero).
        umbral_var_alta: Umbral de variación alta (SAS: 25).
        umbral_var_baja: Umbral de variación baja (SAS: -25).
        umbral_var_extrema: Umbral de variación extrema (SAS: 100).

    Returns:
        (base_calidad, var_atipico): base con columna 'REVISA' + reporte de atípicos.
    """
    llave = [
        c for c in ["CÓDIGO DIVIPOLA", "Nombre_Publica"]
        if c in base_con_cv.columns and (
            mayor2_anterior is None or c in mayor2_anterior.columns
        )
    ]
    if not llave:
        llave = [c for c in ["CÓDIGO DIVIPOLA", "LLAVE_ARTICULO"] if c in base_con_cv.columns]

    df = base_con_cv.copy()

    if mayor2_anterior is None or len(mayor2_anterior) == 0:
        log.warning(
            "detectar_var_atipica | No hay datos del período anterior — "
            "REVISA=2 para todos los registros."
        )
        df["PRECIO_ANTERIOR"] = float("nan")
        df["VAR"] = float("nan")
        df["REVISA"] = 2
    else:
        # Calcular precio promedio actual por municipio-producto para la comparativa
        precio_actual = (
            df.groupby(llave)["PRECIO"]
            .mean()
            .reset_index()
            .rename(columns={"PRECIO": "PRECIO_ACTUAL"})
        )

        anterior = mayor2_anterior[llave + ["PRECIO_PROMEDIO"]].rename(
            columns={"PRECIO_PROMEDIO": "PRECIO_ANTERIOR"}
        )

        comparativa = precio_actual.merge(anterior, on=llave, how="left")
        comparativa["VAR"] = (
            (comparativa["PRECIO_ACTUAL"] - comparativa["PRECIO_ANTERIOR"])
            / comparativa["PRECIO_ANTERIOR"]
            * 100
        )

        df = df.merge(comparativa[llave + ["PRECIO_ANTERIOR", "VAR"]], on=llave, how="left")

        # Clasificar variaciones (SAS: REVISA)
        df["REVISA"] = 0
        df.loc[df["VAR"].isna(), "REVISA"] = 2
        df.loc[
            (df["VAR"].notna()) & ((df["VAR"] >= umbral_var_alta) | (df["VAR"] <= umbral_var_baja)),
            "REVISA"
        ] = 1
        df.loc[df["VAR"].notna() & (df["VAR"] >= umbral_var_extrema), "REVISA"] = 3

    var_atipico = df[df["REVISA"] > 0].copy()

    n_revisa1 = (df["REVISA"] == 1).sum()
    n_revisa2 = (df["REVISA"] == 2).sum()
    n_revisa3 = (df["REVISA"] == 3).sum()
    log.info(
        "detectar_var_atipica OK | REVISA=1: %d | REVISA=2: %d | REVISA=3: %d",
        n_revisa1, n_revisa2, n_revisa3,
    )
    return df, var_atipico
