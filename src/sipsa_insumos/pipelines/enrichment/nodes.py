"""Nodos del pipeline de enriquecimiento — SIPSA Insumos.

SAS equivalente (pasos 3-5):
  MERGE base_liviana DIVIPOLA; by CodigoMpio;
  MERGE base divipola_grupos; by Art_Unmed_Casacomer_ICA;
  MERGE base divipola_articulos; by Art_Unmed_Casacomer_ICA;

Flujo:
  base_bronze + divipola_raw ──► [merge_divipola] ──► base_con_mpio + faltan_divipola
  base_con_mpio + mappings_grupos ──► [asignar_grupo] ──► base_con_grupo + faltan_grupo
  base_con_grupo + mappings_articulos ──► [asignar_articulo_publica] ──► base_enriquecida + faltan_publica
"""
from __future__ import annotations

import logging

import pandas as pd

log = logging.getLogger(__name__)


def merge_divipola(
    base_bronze: pd.DataFrame,
    divipola_raw: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Une la base liviana con DIVIPOLA por código de municipio.

    SAS: MERGE base DIVIPOLA; by CodigoMpio;
    Soporta tanto el DIVIPOLA real (columna 'CódigoMunicipio', códigos sin
    padding) como fixtures de test (columna 'CodigoMpio').

    Args:
        base_bronze: DataFrame con columna 'CÓDIGO DIVIPOLA' (str 5 dígitos).
        divipola_raw: DataFrame con columna de código municipio.

    Returns:
        (base_con_mpio, faltan_divipola): base enriquecida + filas sin match.
    """
    divipola = divipola_raw.copy()

    # Detectar columna de código: DIVIPOLA real usa 'CódigoMunicipio', tests 'CodigoMpio'
    if "CódigoMunicipio" in divipola.columns:
        divipola = divipola.rename(columns={"CódigoMunicipio": "CÓDIGO DIVIPOLA"})
        # El DIVIPOLA real tiene códigos sin leading zero (ej: '5001' en vez de '05001')
        divipola["CÓDIGO DIVIPOLA"] = divipola["CÓDIGO DIVIPOLA"].str.zfill(5)
    elif "CodigoMpio" in divipola.columns:
        divipola = divipola.rename(columns={"CodigoMpio": "CÓDIGO DIVIPOLA"})
    else:
        # Fallback: primera columna como clave
        divipola = divipola.rename(columns={divipola.columns[0]: "CÓDIGO DIVIPOLA"})

    # Left join — conserva todas las filas de la base
    merged = base_bronze.merge(divipola, on="CÓDIGO DIVIPOLA", how="left")

    # Columna indicadora de match: cualquier columna que vino del DIVIPOLA
    col_match = next(
        (c for c in ["NombreDepartamento", "CodigoDepto", "NombreMunicipio", "Departamento"]
         if c in merged.columns),
        None,
    )
    sin_match = merged[col_match].isna() if col_match else pd.Series(False, index=merged.index)
    faltan_divipola = merged[sin_match].copy()
    base_con_mpio = merged[~sin_match].copy()

    if len(faltan_divipola) > 0:
        codigos = faltan_divipola["CÓDIGO DIVIPOLA"].unique().tolist()
        log.warning(
            "merge_divipola | %d filas sin match en DIVIPOLA | códigos: %s",
            len(faltan_divipola), codigos[:20],
        )

    log.info(
        "merge_divipola OK | filas_totales=%d | con_match=%d | sin_match=%d",
        len(base_bronze), len(base_con_mpio), len(faltan_divipola),
    )
    return base_con_mpio, faltan_divipola


def asignar_grupo(
    base_con_mpio: pd.DataFrame,
    mappings_grupos: dict,
    modulo: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Asigna el grupo de insumo a cada artículo usando el mapping YAML.

    SAS: MERGE base divipola_grupos; by Art_Unmed_Casacomer_ICA;
    Equivalente: lookup LLAVE_ARTICULO → Grupo en mappings_grupos.yml.

    Args:
        base_con_mpio: DataFrame con columna 'LLAVE_ARTICULO'.
        mappings_grupos: Dict {llave_articulo: grupo} cargado desde YAML.
        modulo: Nombre del módulo para logging.

    Returns:
        (base_con_grupo, faltan_grupo): base con columna 'Grupo' + filas sin match.
    """
    grupos_dict: dict[str, str] = mappings_grupos.get("grupos", mappings_grupos)
    df = base_con_mpio.copy()
    df["Grupo"] = df["LLAVE_ARTICULO"].map(grupos_dict)

    sin_grupo = df["Grupo"].isna()
    faltan_grupo = df[sin_grupo].copy()
    base_con_grupo = df[~sin_grupo].copy()

    if len(faltan_grupo) > 0:
        llaves = faltan_grupo["LLAVE_ARTICULO"].unique().tolist()
        log.warning(
            "asignar_grupo [%s] | %d filas sin grupo | llaves: %s",
            modulo, len(faltan_grupo), llaves[:10],
        )

    log.info(
        "asignar_grupo [%s] OK | con_grupo=%d | sin_grupo=%d",
        modulo, len(base_con_grupo), len(faltan_grupo),
    )
    return base_con_grupo, faltan_grupo


def asignar_articulo_publica(
    base_con_grupo: pd.DataFrame,
    mappings_articulos: dict,
    modulo: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Asigna el nombre de publicación a cada artículo usando el mapping YAML.

    SAS: MERGE base divipola_articulos; by Art_Unmed_Casacomer_ICA;
    Equivalente: lookup LLAVE_ARTICULO → Nombre_Publica en mappings_articulos.yml.

    Args:
        base_con_grupo: DataFrame con columna 'LLAVE_ARTICULO'.
        mappings_articulos: Dict {llave_articulo: nombre_publicacion} desde YAML.
        modulo: Nombre del módulo para logging.

    Returns:
        (base_enriquecida, faltan_publica): base con 'Nombre_Publica' + sin match.
    """
    articulos_dict: dict[str, str] = mappings_articulos.get("articulos_publicacion", mappings_articulos)
    df = base_con_grupo.copy()
    df["Nombre_Publica"] = df["LLAVE_ARTICULO"].map(articulos_dict)

    sin_publica = df["Nombre_Publica"].isna()
    faltan_publica = df[sin_publica].copy()
    base_enriquecida = df[~sin_publica].copy()

    if len(faltan_publica) > 0:
        llaves = faltan_publica["LLAVE_ARTICULO"].unique().tolist()
        log.warning(
            "asignar_articulo_publica [%s] | %d filas sin nombre publicación | llaves: %s",
            modulo, len(faltan_publica), llaves[:10],
        )

    log.info(
        "asignar_articulo_publica [%s] OK | con_nombre=%d | sin_nombre=%d",
        modulo, len(base_enriquecida), len(faltan_publica),
    )
    return base_enriquecida, faltan_publica
