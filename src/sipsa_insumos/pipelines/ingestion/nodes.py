"""Nodos del pipeline de ingesta — SIPSA Insumos.

SAS equivalente:
  proc import datafile="&FECHA./&ENTRADA." dbms=xlsx
    out=INSUMOS_BASE replace;
    sheet="Información Insumos"; getnames=yes;
  run;
  /* Filtrado: Estado='ANALIZADO CENTRAL' AND Precio > 0 */

Flujo:
  params:<modulo>.archivo_liviana ──► [leer_base_liviana] ──► <modulo>.base_bronze

Notas sobre el Excel de entrada:
  - La hoja "Información Insumos" contiene la base de supervisión COMPLETA,
    no solo los registros publicables. Se filtra a:
      Estado = 'ANALIZADO CENTRAL'  (centralmente verificados)
      Precio Actual > 0             (precio efectivo reportado)
  - Las columnas tienen nombres abreviados (Municipio, Articulo, etc.).
    Se renombran a los nombres canónicos del proyecto.
  - Municipio viene sin padding (ej: 5001 → 05001 con zfill(5)).
  - MES_AÑO no está en el Excel; se agrega desde el parámetro periodo.
"""
from __future__ import annotations

import logging

import pandas as pd

from sipsa_insumos.utils.parsers import agregar_columnas_unidad_medida
from sipsa_insumos.validations.schemas import SCHEMA_BASE_LIVIANA

log = logging.getLogger(__name__)

# Mapeo de columnas abreviadas del Excel → nombres canónicos del proyecto
_RENAME_COLS: dict[str, str] = {
    "Municipio":      "CÓDIGO DIVIPOLA",
    "Codigo CPC":     "CÓDIGO CPC",
    "Articulo":       "ARTÍCULO",
    "CasaCom.":       "CASA COMERCIAL",
    "RegICA":         "REGISTRO ICA",
    "Precio Actual":  "PRECIO",
    "UnMed.":         "UNIDAD DE MEDIDA",
}

# Columnas a leer como string (para evitar conversión numérica de códigos)
_DTYPE_EXCEL: dict[str, type] = {
    "Municipio":  str,
    "Codigo CPC": str,
    "Articulo":   str,
    "CasaCom.":   str,
    "RegICA":     str,
    "UnMed.":     str,
    "Estado":     str,
}


def leer_base_liviana(
    archivo_liviana: str,
    hoja_liviana: str,
    periodo: str,
) -> pd.DataFrame:
    """Lee el Excel mensual de supervisión, filtra y estandariza para el pipeline.

    Pasos:
    1. Leer Excel con nombres de columna abreviados.
    2. Filtrar a registros ANALIZADO CENTRAL con Precio Actual > 0.
    3. Renombrar columnas al estándar del proyecto.
    4. Zero-pad CÓDIGO DIVIPOLA a 5 dígitos (ej: '5001' → '05001').
    5. Agregar MES_AÑO = periodo.
    6. Validar con schema Pandera.
    7. Parsear UNIDAD DE MEDIDA → NOMBRE_UM, UNIDAD, CANTIDAD, LLAVE_ARTICULO.

    Args:
        archivo_liviana: Ruta relativa al Excel de entrada.
        hoja_liviana: Nombre de la hoja ("Información Insumos").
        periodo: Período mensual (ej: "MAY2026"). Se usa como valor de MES_AÑO.

    Returns:
        DataFrame listo para el pipeline de enriquecimiento.
    """
    log.info("Leyendo base liviana: %s | hoja: %s", archivo_liviana, hoja_liviana)

    df = pd.read_excel(
        archivo_liviana,
        sheet_name=hoja_liviana,
        header=0,
        dtype=_DTYPE_EXCEL,
    )
    log.info("Excel leído | filas_raw=%d | columnas=%d", len(df), len(df.columns))

    # 1. Filtrar: solo registros ANALIZADO CENTRAL con precio efectivo
    if "Estado" in df.columns:
        df = df[df["Estado"].str.strip().str.upper() == "ANALIZADO CENTRAL"].copy()
        log.info("Filtrado por Estado=ANALIZADO CENTRAL | filas=%d", len(df))

    df["Precio Actual"] = pd.to_numeric(df["Precio Actual"], errors="coerce")
    df = df[df["Precio Actual"] > 0].copy()
    log.info("Filtrado por Precio Actual > 0 | filas=%d", len(df))

    # 2. Renombrar columnas abreviadas
    df = df.rename(columns=_RENAME_COLS)

    # 3. Zero-pad CÓDIGO DIVIPOLA a 5 dígitos (ej: '5001' → '05001')
    df["CÓDIGO DIVIPOLA"] = df["CÓDIGO DIVIPOLA"].str.strip().str.zfill(5)

    # 4. Limpiar CÓDIGO CPC (puede tener espacios o decimal si se leyó como número)
    df["CÓDIGO CPC"] = df["CÓDIGO CPC"].astype(str).str.strip().str.split(".").str[0]

    # 5. Agregar MES_AÑO desde el parámetro periodo
    df["MES_AÑO"] = periodo

    # 6. Limpiar REGISTRO ICA
    df["REGISTRO ICA"] = df["REGISTRO ICA"].astype(str).str.strip().str.split(".").str[0]

    # 7. Conservar solo las columnas canónicas (descartar columnas auxiliares del Excel)
    _COLS_CANON = list(_RENAME_COLS.values()) + ["MES_AÑO"]
    df = df[[c for c in _COLS_CANON if c in df.columns]].copy()

    # 8. Validar schema Pandera (lazy=True → expone todas las violaciones de una vez)
    df = SCHEMA_BASE_LIVIANA.validate(df, lazy=True)

    # 9. Parsear UNIDAD DE MEDIDA → NOMBRE_UM, UNIDAD, CANTIDAD, LLAVE_ARTICULO
    df = agregar_columnas_unidad_medida(df)

    log.info(
        "leer_base_liviana OK | filas=%d | municipios_únicos=%d",
        len(df),
        df["CÓDIGO DIVIPOLA"].nunique(),
    )
    return df
