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

    agg = (
        base_calidad.groupby(llave_completa, dropna=False)["PRECIO"]
        .agg(
            N_FUENTE=lambda x: x.notna().sum(),
            N_ARTICULOS="count",
            PRECIO_PROMEDIO="mean",
            PRECIO_MIN="min",
            PRECIO_MAX="max",
        )
        .reset_index()
    )

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
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Aplica el criterio de secreto estadístico (N_ARTICULOS >= min_n).

    SAS:
      IF N_ARTICULOS >= 2 THEN OUTPUT MAYOR2; (agricolas/pecuarios/elementos)
      IF N_ARTICULOS >= 1 THEN OUTPUT MAYOR2; (arriendos/servicios/empaques)

    Args:
        precio_promedio: DataFrame del nodo calcular_precio_promedio.
        min_n: Mínimo de fuentes para publicar (2 para secreto estadístico, 1 sin secreto).

    Returns:
        (mayor2, menor2): publicables + bajo umbral.
    """
    mayor2 = precio_promedio[precio_promedio["N_ARTICULOS"] >= min_n].copy()
    menor2 = precio_promedio[precio_promedio["N_ARTICULOS"] < min_n].copy()

    log.info(
        "aplicar_secreto_estadistico OK | publicables(N>=%d)=%d | secreto=%d",
        min_n, len(mayor2), len(menor2),
    )
    return mayor2, menor2
