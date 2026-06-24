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

# Columnas de agrupación para el precio promedio
_LLAVE_PROMEDIO = [
    "CÓDIGO DIVIPOLA",
    "LLAVE_ARTICULO",
    "Grupo",
    "Nombre_Publica",
    "CÓDIGO CPC",
    "ARTÍCULO",
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
        N_FUENTE, N_ARTICULOS, PRECIO_PROMEDIO, PRECIO_MIN, PRECIO_MAX.
    """
    llave = [c for c in _LLAVE_PROMEDIO if c in base_calidad.columns]
    cols_divipola = [c for c in _COLS_DIVIPOLA if c in base_calidad.columns]

    # Para mantener las columnas DIVIPOLA en el resultado agrupado
    if cols_divipola:
        llave_completa = llave + cols_divipola
    else:
        llave_completa = llave

    # Deduplicar llave_completa manteniendo orden
    seen: set[str] = set()
    llave_completa_unica = []
    for c in llave_completa:
        if c not in seen:
            seen.add(c)
            llave_completa_unica.append(c)

    agg = (
        base_calidad.groupby(llave_completa_unica, dropna=False)["PRECIO"]
        .agg(
            N_FUENTE=lambda x: x.notna().sum(),
            N_ARTICULOS="count",
            PRECIO_PROMEDIO="mean",
            PRECIO_MIN="min",
            PRECIO_MAX="max",
        )
        .reset_index()
    )

    log.info(
        "calcular_precio_promedio OK | grupos=%d | precio_promedio_medio=%.0f",
        len(agg),
        agg["PRECIO_PROMEDIO"].mean() if len(agg) > 0 else 0,
    )
    return agg


def aplicar_secreto_estadistico(
    precio_promedio: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Aplica el criterio de secreto estadístico (N_ARTICULOS >= 2).

    SAS:
      IF N_ARTICULOS >= 2 THEN OUTPUT MAYOR2;
      ELSE OUTPUT MENOR2;

    Los registros con N < 2 no se publican en los cuadros finales.

    Args:
        precio_promedio: DataFrame del nodo calcular_precio_promedio.

    Returns:
        (mayor2, menor2): publicables + bajo secreto estadístico.
    """
    mayor2 = precio_promedio[precio_promedio["N_ARTICULOS"] >= 2].copy()
    menor2 = precio_promedio[precio_promedio["N_ARTICULOS"] < 2].copy()

    log.info(
        "aplicar_secreto_estadistico OK | publicables(N>=2)=%d | secreto(N<2)=%d",
        len(mayor2), len(menor2),
    )
    return mayor2, menor2
