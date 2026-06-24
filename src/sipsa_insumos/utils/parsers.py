"""Utilidades de parseo para SIPSA Insumos.

La columna 'UNIDAD DE MEDIDA' del Excel llega en formato pipe:
    NOMBRE|UNIDAD|CANTIDAD|0|0
Ejemplo:
    BULTO|KILOGRAMO|50|0|0
    FRASCO|LITRO|1|0|0
    JORNAL|DIA|1|0|0

La llave de lookup (LLAVE_ARTICULO) se construye a partir de las
mismas partes normalizadas (UPPER, strip) para coincidir con las
claves de mappings_grupos.yml y mappings_articulos.yml.
"""
from __future__ import annotations

import pandas as pd


def parsear_unidad_medida(valor: str) -> dict[str, str]:
    """Divide el campo pipe-delimitado en sus componentes.

    Args:
        valor: Cadena con formato "NOMBRE|UNIDAD|CANTIDAD|0|0".
               Si el formato es incompleto, rellena con cadenas vacías.

    Returns:
        Dict con claves: nombre_um, unidad, cantidad.
    """
    if not isinstance(valor, str):
        return {"nombre_um": "", "unidad": "", "cantidad": ""}

    partes = valor.split("|")
    return {
        "nombre_um": partes[0].strip() if len(partes) > 0 else "",
        "unidad":    partes[1].strip() if len(partes) > 1 else "",
        "cantidad":  partes[2].strip() if len(partes) > 2 else "",
    }


def construir_llave_articulo(articulo: str, unidad_medida: str) -> str:
    """Construye la llave normalizada para los mappings.

    Formato: ARTÍCULO.upper() + "_" + UNIDAD_DE_MEDIDA.upper()
    Equivale a la columna Art_Unmed_Casacomer_ICA del SAS:
        UPCASE(ARTICULO)||'_'||UPCASE('Unidad de medida'n)
    Coincide con las claves de mappings_grupos.yml y mappings_articulos.yml.

    Args:
        articulo: Nombre del artículo (ej: "Glifosato 480 SL").
        unidad_medida: Unidad de medida en formato pipe (ej: "FRASCO|LITRO|1|0|0").

    Returns:
        Clave normalizada lista para lookup en los diccionarios YAML.
        Ejemplo: "GLIFOSATO 480 SL_FRASCO|LITRO|1|0|0"
    """
    return f"{articulo.strip().upper()}_{unidad_medida.strip().upper()}"


def agregar_columnas_unidad_medida(df: pd.DataFrame) -> pd.DataFrame:
    """Parsea 'UNIDAD DE MEDIDA' y agrega NOMBRE_UM, UNIDAD, CANTIDAD y LLAVE_ARTICULO.

    LLAVE_ARTICULO = ARTÍCULO.upper() + "_" + "UNIDAD DE MEDIDA".upper()
    Equivale a Art_Unmed_Casacomer_ICA del SAS (sin CASA COMERCIAL ni REGISTRO ICA,
    que el DIVIPOLA no usa para la clasificación de grupos y artículos de publicación).

    Args:
        df: DataFrame con las columnas 'ARTÍCULO' y 'UNIDAD DE MEDIDA'.

    Returns:
        DataFrame con 4 columnas adicionales: NOMBRE_UM, UNIDAD, CANTIDAD, LLAVE_ARTICULO.
    """
    parsed = df["UNIDAD DE MEDIDA"].apply(parsear_unidad_medida).apply(pd.Series)
    df = df.copy()
    df["NOMBRE_UM"] = parsed["nombre_um"]
    df["UNIDAD"]    = parsed["unidad"]
    df["CANTIDAD"]  = parsed["cantidad"]
    df["LLAVE_ARTICULO"] = df.apply(
        lambda r: construir_llave_articulo(r["ARTÍCULO"], r["UNIDAD DE MEDIDA"]),
        axis=1,
    )
    return df
