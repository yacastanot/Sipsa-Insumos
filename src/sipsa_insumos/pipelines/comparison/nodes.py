"""Nodos del pipeline de comparación interperiódica — SIPSA Insumos.

SAS equivalente (paso 12):
  MERGE mayor2_actual mayor2_anterior; by CodigoMpio Art_Unmed_Casacomer_ICA;
  VARIACION = (PRECIO_ACTUAL/PRECIO_ANTERIOR - 1) * 100;
  IF VARIACION > 0 THEN TENDENCIA = 'Positiva';
  ELSE IF VARIACION < 0 THEN TENDENCIA = 'Negativa';
  ELSE IF VARIACION = 0 THEN TENDENCIA = 'Estable';
  ELSE TENDENCIA = 'n.d.';

Flujo:
  mayor2 + mayor2_anterior ──► [calcular_variacion_tendencia] ──► base_comparada
"""
from __future__ import annotations

import logging

import pandas as pd

log = logging.getLogger(__name__)

_LLAVE_JOIN = ["CÓDIGO DIVIPOLA", "Nombre_Publica"]


def calcular_variacion_tendencia(
    mayor2: pd.DataFrame,
    mayor2_anterior: pd.DataFrame | None,
    mes_actual: str,
    mes_anterior: str,
) -> pd.DataFrame:
    """Calcula la variación interperiódica y clasifica la tendencia.

    SAS: MERGE mayor2 mayor2_anterior; by CodigoMpio Art_Unmed_Casacomer_ICA;
    VARIACION = (PRECIO_actual - PRECIO_anterior) / PRECIO_anterior * 100
    TENDENCIA: Positiva (>0) | Negativa (<0) | Estable (=0) | n.d. (sin dato anterior)

    Args:
        mayor2: DataFrame del período actual con PRECIO_PROMEDIO.
        mayor2_anterior: mayor2 del período anterior (None si es el primer período).
        mes_actual: Nombre del mes actual en español (ej: "Mayo").
        mes_anterior: Nombre del mes anterior en español (ej: "Abril").

    Returns:
        DataFrame con columnas adicionales:
        PRECIO_{mes_anterior}, VARIACION, TENDENCIA.
    """
    df = mayor2.copy()
    df = df.rename(columns={"PRECIO_PROMEDIO": f"PRECIO_{mes_actual}"})

    if mayor2_anterior is None or len(mayor2_anterior) == 0:
        log.warning(
            "calcular_variacion_tendencia | Sin datos del período anterior — "
            "TENDENCIA='n.d.' para todos los registros."
        )
        df[f"PRECIO_{mes_anterior}"] = float("nan")
        df["VARIACION"] = float("nan")
        df["TENDENCIA"] = "n.d."
        return df

    llave = [c for c in _LLAVE_JOIN if c in df.columns and c in mayor2_anterior.columns]

    anterior = mayor2_anterior[llave + ["PRECIO_PROMEDIO"]].rename(
        columns={"PRECIO_PROMEDIO": f"PRECIO_{mes_anterior}"}
    )

    df = df.merge(anterior, on=llave, how="left")

    df["VARIACION"] = (
        (df[f"PRECIO_{mes_actual}"] - df[f"PRECIO_{mes_anterior}"])
        / df[f"PRECIO_{mes_anterior}"]
        * 100
    )

    # Clasificar tendencia
    condiciones = [
        df["VARIACION"] > 0,
        df["VARIACION"] < 0,
        df["VARIACION"] == 0,
    ]
    opciones = ["Positiva", "Negativa", "Estable"]
    df["TENDENCIA"] = pd.Series(
        pd.Categorical(
            pd.array(
                [
                    "Positiva" if c1 else "Negativa" if c2 else "Estable" if c3 else "n.d."
                    for c1, c2, c3 in zip(*condiciones)
                ]
            )
        )
    )

    n_positiva = (df["TENDENCIA"] == "Positiva").sum()
    n_negativa = (df["TENDENCIA"] == "Negativa").sum()
    n_estable  = (df["TENDENCIA"] == "Estable").sum()
    n_nd       = (df["TENDENCIA"] == "n.d.").sum()
    log.info(
        "calcular_variacion_tendencia OK | Positiva=%d | Negativa=%d | Estable=%d | n.d.=%d",
        n_positiva, n_negativa, n_estable, n_nd,
    )
    return df
