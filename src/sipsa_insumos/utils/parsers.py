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

Modos de llave:
  "unmed"            ARTÍCULO_UNMED  (Agricolas, Pecuarios — marcas distintas
                     del mismo producto se agregan bajo la misma Nombre_Publica)
  "casacom_ica_unmed" ARTÍCULO_CASACOM_ICA_UNMED  (Elementos — la CasaCom y el
                     RegICA son especificaciones que diferencian productos)
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
    """Construye la llave normalizada para los mappings (modo 'unmed').

    Formato: ARTÍCULO.upper() + "_" + UNIDAD_DE_MEDIDA.upper()
    Equivale a la columna Art_Unmed_Casacomer_ICA del SAS (Agricolas/Pecuarios).

    Args:
        articulo: Nombre del artículo (ej: "Glifosato 480 SL").
        unidad_medida: Unidad de medida en formato pipe (ej: "FRASCO|LITRO|1|0|0").

    Returns:
        Clave normalizada lista para lookup en los diccionarios YAML.
        Ejemplo: "GLIFOSATO 480 SL_FRASCO|LITRO|1|0|0"
    """
    return f"{articulo.strip().upper()}_{unidad_medida.strip().upper()}"


def _clean(value: str) -> str:
    """Normaliza un componente de la llave: uppercase, sin NaN."""
    s = str(value).strip().upper()
    return "NA" if s in ("NAN", "NONE", "", "N/A") else s


def construir_llave_casacom_ica_unmed(
    articulo: str,
    casa_comercial: str,
    registro_ica: str,
    unidad_medida: str,
) -> str:
    """Construye la llave compuesta para Elementos (modo 'casacom_ica_unmed').

    Equivale a Art_Casacomer_ICA_Unmed del SAS:
        CATX('_', Articulo, 'CasaCom.'n, RegICA, 'UnMed.'n)

    Trata nulos como 'NA' (igual que CATX en SAS con missing string).
    """
    parts = [
        _clean(articulo),
        _clean(casa_comercial),
        _clean(registro_ica),
        _clean(unidad_medida),
    ]
    return "_".join(parts)


def parsear_llave_divipola(valor: str, tipo_llave: str = "unmed") -> str:
    """Reconstruye la LLAVE_ARTICULO canónica a partir de la clave compuesta
    de un archivo DIVIPOLA de módulo (columna Art_Unmed_Casacomer_ICA o
    Art_Casacomer_ICA_Unmed — el orden varía por módulo).

    No asume una posición fija: localiza el segmento con '|' (Unidad de
    Medida, siempre pipe-delimitado) y usa los segmentos restantes en su
    orden relativo — el primero es siempre el Artículo; si sigue Casa
    Comercial y Registro ICA, van en ese orden en ambas variantes observadas.

    Args:
        valor: Valor de la columna compuesta del DIVIPOLA.
        tipo_llave: "unmed" (Agricolas, Propagación) o "casacom_ica_unmed"
                    (Pecuarios, Elementos) — igual que agregar_columnas_unidad_medida.

    Returns:
        LLAVE_ARTICULO en el mismo formato que produce el pipeline al leer
        la base liviana, lista para usarse como clave en los mappings.
    """
    partes = str(valor).strip().split("_")
    idx_unidad = next((i for i, p in enumerate(partes) if "|" in p), None)
    if idx_unidad is None:
        return str(valor).strip().upper()

    unidad = partes[idx_unidad]
    resto = partes[:idx_unidad] + partes[idx_unidad + 1:]
    articulo = resto[0] if resto else ""
    extra = resto[1:]

    if tipo_llave == "casacom_ica_unmed":
        casacom = extra[0] if len(extra) > 0 else ""
        regica = extra[1] if len(extra) > 1 else ""
        return construir_llave_casacom_ica_unmed(articulo, casacom, regica, unidad)
    return construir_llave_articulo(articulo, unidad)


def agregar_columnas_unidad_medida(
    df: pd.DataFrame,
    tipo_llave: str = "unmed",
) -> pd.DataFrame:
    """Parsea 'UNIDAD DE MEDIDA' y agrega NOMBRE_UM, UNIDAD, CANTIDAD y LLAVE_ARTICULO.

    Args:
        df: DataFrame con las columnas 'ARTÍCULO' y 'UNIDAD DE MEDIDA'.
            Para tipo_llave='casacom_ica_unmed' también necesita
            'CASA COMERCIAL' y 'REGISTRO ICA'.
        tipo_llave: "unmed" (default) o "casacom_ica_unmed" (Elementos).

    Returns:
        DataFrame con 4 columnas adicionales: NOMBRE_UM, UNIDAD, CANTIDAD, LLAVE_ARTICULO.
    """
    parsed = df["UNIDAD DE MEDIDA"].apply(parsear_unidad_medida).apply(pd.Series)
    df = df.copy()
    df["NOMBRE_UM"] = parsed["nombre_um"]
    df["UNIDAD"]    = parsed["unidad"]
    df["CANTIDAD"]  = parsed["cantidad"]

    if tipo_llave == "casacom_ica_unmed":
        df["LLAVE_ARTICULO"] = df.apply(
            lambda r: construir_llave_casacom_ica_unmed(
                r["ARTÍCULO"],
                r.get("CASA COMERCIAL", ""),
                r.get("REGISTRO ICA", ""),
                r["UNIDAD DE MEDIDA"],
            ),
            axis=1,
        )
    else:
        df["LLAVE_ARTICULO"] = df.apply(
            lambda r: construir_llave_articulo(r["ARTÍCULO"], r["UNIDAD DE MEDIDA"]),
            axis=1,
        )
    return df
