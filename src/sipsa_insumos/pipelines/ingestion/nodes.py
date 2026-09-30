"""Nodos del pipeline de ingesta — SIPSA Insumos.

SAS equivalente:
  proc import datafile="&FECHA./&ENTRADA." dbms=xlsx
    out=INSUMOS_BASE replace;
    sheet="Información Insumos"; getnames=yes;
  run;
  /* Filtrado: Estado IN ('ANALIZADO CENTRAL','APROBADO') AND Precio > 0 */

Flujo:
  params:<modulo>.archivo_liviana ──► [leer_base_liviana] ──► <modulo>.base_bronze

Tipos de módulo (tipo_modulo):
  "estandar" — Agricolas, Pecuarios, Elementos: usa columnas UnMed./CasaCom./RegICA.
  "caracte"  — Arriendos, Servicios, Empaques: usa columna Caracte. (sin UnMed).

Modos de llave (tipo_llave):
  "unmed"            — ARTÍCULO_UNMED (Agricolas, Pecuarios: marcas distintas ≡ misma publi)
  "casacom_ica_unmed" — ARTÍCULO_CASACOM_ICA_UNMED (Elementos: especificaciones diferencian)
"""
from __future__ import annotations

import logging
import re
from datetime import date

import pandas as pd

from sipsa_insumos.utils.parsers import agregar_columnas_unidad_medida
from sipsa_insumos.validations.schemas import SCHEMA_BASE_LIVIANA

log = logging.getLogger(__name__)

_ESTADOS_VALIDOS = {"ANALIZADO CENTRAL", "APROBADO"}

_RENAME_ESTANDAR: dict[str, str] = {
    "Municipio":      "CÓDIGO DIVIPOLA",
    "Fuente":         "FUENTE",
    "Codigo CPC":     "CÓDIGO CPC",
    "Articulo":       "ARTÍCULO",
    "CasaCom.":       "CASA COMERCIAL",
    "RegICA":         "REGISTRO ICA",
    "Precio Actual":  "PRECIO",
    "UnMed.":         "UNIDAD DE MEDIDA",
}

_RENAME_CARACTE: dict[str, str] = {
    "Municipio":      "CÓDIGO DIVIPOLA",
    "Fuente":         "FUENTE",
    "Informante":     "INFORMANTE",
    "Codigo CPC":     "CÓDIGO CPC",
    "Articulo":       "ARTÍCULO",
    "Precio Actual":  "PRECIO",
    "Caracte.":       "CARACTERÍSTICA",
}

_DTYPE_EXCEL: dict[str, type] = {
    "Municipio":  str,
    "Codigo CPC": str,
    "Articulo":   str,
    "CasaCom.":   str,
    "RegICA":     str,
    "UnMed.":     str,
    "Caracte.":   str,
    "Estado":     str,
}


def leer_base_liviana(
    archivo_liviana: str,
    hoja_liviana: str,
    periodo: str,
    tipo_modulo: str = "estandar",
    tipo_llave: str = "unmed",
) -> pd.DataFrame:
    """Lee el Excel mensual de supervisión, filtra y estandariza para el pipeline.

    Args:
        archivo_liviana: Ruta relativa al Excel de entrada.
        hoja_liviana: Nombre de la hoja ("Información Insumos").
        periodo: Período mensual (ej: "MAY2026"). Se usa como valor de MES_AÑO.
        tipo_modulo: "estandar" (UnMed/CasaCom/ICA) o "caracte" (Caracte.).
        tipo_llave: "unmed" (ARTÍCULO+UNMED) o "casacom_ica_unmed" (Elementos).
    """
    log.info(
        "Leyendo base liviana: %s | tipo_modulo=%s | tipo_llave=%s",
        archivo_liviana, tipo_modulo, tipo_llave,
    )

    df = pd.read_excel(
        archivo_liviana,
        sheet_name=hoja_liviana,
        header=0,
        dtype=_DTYPE_EXCEL,
    )
    log.info("Excel leído | filas_raw=%d | columnas=%d", len(df), len(df.columns))

    if "Estado" in df.columns:
        mask = df["Estado"].str.strip().str.upper().isin(_ESTADOS_VALIDOS)
        df = df[mask].copy()
        log.info("Filtrado por Estado válido | filas=%d", len(df))

    df["Precio Actual"] = pd.to_numeric(df["Precio Actual"], errors="coerce")
    df = df[df["Precio Actual"] > 0].copy()
    log.info("Filtrado por Precio Actual > 0 | filas=%d", len(df))

    if tipo_modulo == "caracte":
        return _procesar_caracte(df, periodo)
    return _procesar_estandar(df, periodo, tipo_llave)


_EXCEL_EPOCH = date(1899, 12, 30)
_RE_DATE_ICA = re.compile(r"^\d{4}-\d{2}-\d{2}")


def _ica_fix(v: str) -> str:
    """Convierte ICA leído como fecha de Excel de vuelta al serial numérico.

    pandas lee celdas de Excel con formato fecha aunque se pida dtype=str.
    Ejemplo: ICA 3119 → '1908-07-14 00:00:00' → se devuelve '3119'.
    Valores ya en formato correcto (p.ej. '3167-3', 'NA') se devuelven tal cual.
    """
    if _RE_DATE_ICA.match(v):
        try:
            d = pd.to_datetime(v).date()
            return str((d - _EXCEL_EPOCH).days)
        except Exception:
            pass
    return v


def _limpiar_registro_ica(serie: pd.Series) -> pd.Series:
    """Limpia REGISTRO ICA quitando el ".0" que agrega Excel a códigos
    puramente numéricos (ej. '3119.0' → '3119'), SIN truncar códigos
    alfanuméricos que sí incluyen un decimal legítimo (ej. '3167-3.5').

    OJO: no usar str.split(".").str[0] aquí — corta cualquier código con
    punto, incluyendo esos sufijos legítimos, y rompe el cruce contra
    Art_Casacomer_ICA_Unmed en mappings_grupos/mappings_articulos.
    """
    return serie.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


_RENAME_QA_EXTRA: dict[str, str] = {
    "Precio Ante.":  "PRECIO_ANTERIOR_RAW",
    "Nov.":          "NOV",
    "Estado":        "ESTADO",
    "Observación":   "OBSERVACION",
}


def leer_base_completa(
    archivo_liviana: str,
    hoja_liviana: str,
    periodo: str,
    tipo_modulo: str = "estandar",
    tipo_llave: str = "unmed",
) -> pd.DataFrame:
    """Lee la base liviana SIN filtrar por Estado/Precio, conservando las
    columnas de control de calidad (Precio Ante., Nov., Estado, Observación)
    que el filtro estándar de leer_base_liviana descarta.

    Necesario para reproducir BASE_INSUMOS_{MODULO}_{PERIODO}.xlsx y los
    archivos de diagnóstico (FALTAN_GRUPO/FALTAN_PUBLICA/DUPLI/VAR_ATIPICO)
    tal como los genera el programa SAS: ahí "Precio Ante." (precio del mes
    anterior) ya viene embebido por fila en el archivo del mes actual — no
    se cruza con un archivo de período anterior aparte.

    Args:
        archivo_liviana: Ruta relativa al Excel de entrada.
        hoja_liviana: Nombre de la hoja ("Información Insumos").
        periodo: Período mensual (ej: "MAY2026"). Se usa como valor de MES_AÑO.
        tipo_modulo: "estandar" (UnMed./CasaCom./RegICA) o "caracte" (Caracte.,
                     Informante — Arriendos/Servicios/Empaques).
        tipo_llave: "unmed" o "casacom_ica_unmed" — igual que leer_base_liviana,
                    para construir LLAVE_ARTICULO y poder cruzar Grupo/Nombre_Publica.
    """
    # Solo las celdas vacías son faltantes: textos como "NA" o "N/A" (p.ej. en
    # Registro ICA de Elementos) se conservan tal cual, igual que PROC IMPORT de SAS.
    df = pd.read_excel(
        archivo_liviana, sheet_name=hoja_liviana, header=0, dtype=_DTYPE_EXCEL,
        keep_default_na=False, na_values=[""],
    )
    rename_map = _RENAME_CARACTE if tipo_modulo == "caracte" else _RENAME_ESTANDAR
    df = df.rename(columns={**rename_map, **_RENAME_QA_EXTRA})
    df["CÓDIGO DIVIPOLA"] = df["CÓDIGO DIVIPOLA"].astype(str).str.strip().str.zfill(5)
    df["CÓDIGO CPC"] = df["CÓDIGO CPC"].astype(str).str.strip().str.split(".").str[0]
    df["MES_AÑO"] = periodo
    df["PRECIO"] = pd.to_numeric(df["PRECIO"], errors="coerce")
    df["PRECIO_ANTERIOR_RAW"] = pd.to_numeric(df["PRECIO_ANTERIOR_RAW"], errors="coerce")

    if tipo_modulo == "caracte":
        df["CARACTERÍSTICA"] = df["CARACTERÍSTICA"].astype(str).str.strip()
        df["LLAVE_ARTICULO"] = (
            df["ARTÍCULO"].str.strip().str.upper() + "_" + df["CARACTERÍSTICA"].str.upper()
        )
        df["UNIDAD DE MEDIDA"] = df["CARACTERÍSTICA"]
    else:
        df["REGISTRO ICA"] = _limpiar_registro_ica(df["REGISTRO ICA"]).map(_ica_fix)
        df = agregar_columnas_unidad_medida(df, tipo_llave=tipo_llave)

    log.info(
        "leer_base_completa OK | filas=%d (sin filtrar Estado/Precio)", len(df),
    )
    return df


def _procesar_estandar(df: pd.DataFrame, periodo: str, tipo_llave: str) -> pd.DataFrame:
    """Procesa módulos con columnas UnMed./CasaCom./RegICA."""
    df = df.rename(columns=_RENAME_ESTANDAR)
    df["CÓDIGO DIVIPOLA"] = df["CÓDIGO DIVIPOLA"].str.strip().str.zfill(5)
    df["CÓDIGO CPC"] = df["CÓDIGO CPC"].astype(str).str.strip().str.split(".").str[0]
    df["MES_AÑO"] = periodo
    df["REGISTRO ICA"] = _limpiar_registro_ica(df["REGISTRO ICA"]).map(_ica_fix)

    cols_canon = list(_RENAME_ESTANDAR.values()) + ["MES_AÑO"]
    df = df[[c for c in cols_canon if c in df.columns]].copy()

    df = SCHEMA_BASE_LIVIANA.validate(df, lazy=True)
    df = agregar_columnas_unidad_medida(df, tipo_llave=tipo_llave)

    log.info(
        "leer_base_liviana (estandar/%s) OK | filas=%d | municipios=%d",
        tipo_llave, len(df), df["CÓDIGO DIVIPOLA"].nunique(),
    )
    return df


def _procesar_caracte(df: pd.DataFrame, periodo: str) -> pd.DataFrame:
    """Procesa módulos con columna Caracte. (Arriendos, Servicios, Empaques).

    LLAVE_ARTICULO = ARTÍCULO.upper() + '_' + CARACTERÍSTICA.upper()
    """
    df = df.rename(columns=_RENAME_CARACTE)
    df["CÓDIGO DIVIPOLA"] = df["CÓDIGO DIVIPOLA"].str.strip().str.zfill(5)
    df["CÓDIGO CPC"] = df["CÓDIGO CPC"].astype(str).str.strip().str.split(".").str[0]
    df["MES_AÑO"] = periodo
    df["CARACTERÍSTICA"] = df["CARACTERÍSTICA"].astype(str).str.strip()

    df["LLAVE_ARTICULO"] = (
        df["ARTÍCULO"].str.strip().str.upper() + "_" + df["CARACTERÍSTICA"].str.upper()
    )
    df["UNIDAD DE MEDIDA"] = df["CARACTERÍSTICA"]
    df["NOMBRE_UM"] = df["CARACTERÍSTICA"]
    df["UNIDAD"] = ""
    df["CANTIDAD"] = ""

    cols_canon = [
        "CÓDIGO DIVIPOLA", "FUENTE", "INFORMANTE", "CÓDIGO CPC", "ARTÍCULO", "PRECIO",
        "CARACTERÍSTICA", "UNIDAD DE MEDIDA", "LLAVE_ARTICULO",
        "NOMBRE_UM", "UNIDAD", "CANTIDAD", "MES_AÑO",
    ]
    df = df[[c for c in cols_canon if c in df.columns]].copy()

    log.info(
        "leer_base_liviana (caracte) OK | filas=%d | municipios=%d",
        len(df), df["CÓDIGO DIVIPOLA"].nunique(),
    )
    return df
