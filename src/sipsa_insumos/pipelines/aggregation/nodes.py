"""Nodos del pipeline de agregación — SIPSA Insumos.

SAS equivalente (pasos 10-12):
  PROC SQL: precio promedio por municipio
  PROC SQL: filtro secreto estadístico (N >= 2)

Flujo:
  base_calidad ──► [calcular_precio_promedio] ──► precio_promedio
  precio_promedio ──► [aplicar_secreto_estadistico] ──► mayor2 + menor2
"""
from __future__ import annotations

import logging

import pandas as pd

log = logging.getLogger(__name__)

# Columnas de agrupación para el precio promedio (nivel Nombre_Publica).
# Equivale al nivel SAS: (municipio, Nombre_productos_agr_publ).
# NOTA: CÓDIGO CPC se excluye del agrupamiento porque distintas marcas del
# mismo Nombre_Publica pueden estar codificadas con CPCs distintos en la base
# liviana. Incluirlo fragmentaría los conteos y haría fallar el secreto
# estadístico N>=2 para esos productos. El CPC se reasigna post-agregación
# tomando el código más frecuente por Nombre_Publica.
_LLAVE_PROMEDIO = [
    "CÓDIGO DIVIPOLA",
    "Nombre_Publica",
    "Grupo",
]

# Columnas DIVIPOLA que se incluyen en la agregación si existen.
# El nodo filtra automáticamente las que no están presentes.
_COLS_DIVIPOLA = ["CodigoDepto", "Departamento", "NombreDepartamento", "NombreMunicipio"]


def calcular_precio_promedio(base_calidad: pd.DataFrame) -> pd.DataFrame:
    """Calcula el precio promedio por municipio-producto.

    SAS:
      PROC SQL;
        SELECT CodigoMpio, Art_Unmed_Casacomer_ICA,
               COUNT(*) as N_FUENTE,
               MEAN(Precio) as PRECIO_PROMEDIO
        FROM base GROUP BY CodigoMpio, Art_Unmed_Casacomer_ICA;
      QUIT;

    Args:
        base_calidad: DataFrame con precios y columnas de enriquecimiento.

    Returns:
        DataFrame con una fila por (municipio, artículo) con:
        N_FUENTE, N_ARTICULOS, PRECIO_PROMEDIO, PRECIO_MIN, PRECIO_MAX,
        CÓDIGO CPC (modo por Nombre_Publica).
    """
    llave = [c for c in _LLAVE_PROMEDIO if c in base_calidad.columns]
    cols_divipola = [c for c in _COLS_DIVIPOLA if c in base_calidad.columns]

    # Para mantener las columnas DIVIPOLA en el resultado agrupado
    llave_completa = llave + [c for c in cols_divipola if c not in llave]

    # SAS: N_FUENTE = COUNT(DISTINCT Fuente) y, en módulos caracte,
    # N_INFORMANTE = COUNT(DISTINCT Informante) — no el número de precios
    # (eso es N_ARTICULOS). Verificado contra MAYORESQUE2/MAY_MEN de SAS.
    con_precio = base_calidad[base_calidad["PRECIO"].notna()]
    aggs = {
        "N_ARTICULOS": ("PRECIO", "count"),
        "PRECIO_PROMEDIO": ("PRECIO", "mean"),
        "PRECIO_MIN": ("PRECIO", "min"),
        "PRECIO_MAX": ("PRECIO", "max"),
    }
    if "FUENTE" in con_precio.columns:
        aggs["N_FUENTE"] = ("FUENTE", "nunique")
    if "INFORMANTE" in con_precio.columns:
        aggs["N_INFORMANTE"] = ("INFORMANTE", "nunique")
    agg = con_precio.groupby(llave_completa, dropna=False).agg(**aggs).reset_index()
    if "N_FUENTE" not in agg.columns:
        agg["N_FUENTE"] = agg["N_ARTICULOS"]
    orden = ["N_FUENTE", *(["N_INFORMANTE"] if "N_INFORMANTE" in agg.columns else []),
             "N_ARTICULOS", "PRECIO_PROMEDIO", "PRECIO_MIN", "PRECIO_MAX"]
    agg = agg[llave_completa + orden]
    # SAS: PRECIO_PROMEDIO=round(SUM/N_ARTICULOS, 0.0000001). Redondear evita
    # variaciones de ~1e-10 que cambian la tendencia de "Estable" a +/-.
    agg["PRECIO_PROMEDIO"] = agg["PRECIO_PROMEDIO"].round(7)

    # Reasignar CÓDIGO CPC: CPC más frecuente por Nombre_Publica
    if "CÓDIGO CPC" in base_calidad.columns and "Nombre_Publica" in base_calidad.columns:
        cpc_mode = (
            base_calidad.groupby("Nombre_Publica", dropna=False)["CÓDIGO CPC"]
            .agg(lambda x: x.mode().iloc[0] if not x.mode().empty else x.iloc[0])
            .reset_index()
        )
        agg = agg.merge(cpc_mode, on="Nombre_Publica", how="left")

    log.info(
        "calcular_precio_promedio OK | grupos=%d | precio_promedio_medio=%.0f",
        len(agg),
        agg["PRECIO_PROMEDIO"].mean() if len(agg) > 0 else 0,
    )
    return agg


def aplicar_secreto_estadistico(
    precio_promedio: pd.DataFrame,
    min_n: int = 2,
    cpc_publicar_siempre: list | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Aplica el criterio de secreto estadístico (N_ARTICULOS >= min_n).

    SAS:
      IF N_ARTICULOS >= 2 THEN OUTPUT MAYOR2; (agricolas/pecuarios/elementos)
      IF N_ARTICULOS >= 1 THEN OUTPUT MAYOR2; (arriendos/servicios/empaques)

    Los productos cuyo Código CPC está en cpc_publicar_siempre se publican
    aunque tengan una sola observación (Material AGO2026: semilla de arroz
    Fedearroz, CPC 113101). Quedan marcados en PUBLICA_SIEMPRE para que el
    reporte los escriba primero, como el SET de SAS.

    Args:
        precio_promedio: DataFrame del nodo calcular_precio_promedio.
        min_n: Mínimo de fuentes para publicar (2 para secreto estadístico, 1 sin secreto).
        cpc_publicar_siempre: Códigos CPC exentos del secreto estadístico.

    Returns:
        (mayor2, menor2): publicables + bajo umbral.
    """
    # Comparación numérica: la base guarda el CPC con ceros a la izquierda ("0113101").
    cpcs = {int(c) for c in (cpc_publicar_siempre or [])}
    siempre = (
        pd.to_numeric(precio_promedio["CÓDIGO CPC"], errors="coerce").isin(cpcs)
        if cpcs and "CÓDIGO CPC" in precio_promedio.columns
        else pd.Series(False, index=precio_promedio.index)
    )
    publicable = (precio_promedio["N_ARTICULOS"] >= min_n) | siempre
    mayor2 = precio_promedio[publicable].copy()
    menor2 = precio_promedio[~publicable].copy()
    if cpcs:
        mayor2["PUBLICA_SIEMPRE"] = siempre[publicable].values

    log.info(
        "aplicar_secreto_estadistico OK | publicables(N>=%d)=%d | secreto=%d",
        min_n, len(mayor2), len(menor2),
    )
    return mayor2, menor2
