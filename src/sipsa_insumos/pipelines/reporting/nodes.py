"""Nodos del pipeline de reportes — SIPSA Insumos.

SAS equivalente (pasos 13-16):
  - BASES   → "BASES {LABEL} {PERIODO}.xlsx"
  - TABLAS  → "TABLAS {LABEL} {PERIODO}.xlsx"
  - ANEXOS  → "ANEXO {LABEL} {PERIODO}.xlsx"
  - CUADROS → "CUADROS {LABEL} {PERIODO}.xlsx"

Nodos adicionales (paso 17 — revisión):
  - MAYORESQUE2 / MENORESQUE2 / MAY_MEN3 → secreto estadístico y revisión inter-período
  - TABREV + REVIS_4MESES → historial de precios 4 períodos (con hoja rodante)
  - REVISIÓN TEMÁTICA → tabla de revisión con variaciones por nivel geográfico

Todos los archivos son multi-hoja (una hoja por grupo) o single-sheet.
Se escriben con pd.ExcelWriter(engine='openpyxl') directamente.
El nodo retorna un DataFrame de metadatos.
"""
from __future__ import annotations

import logging
import unicodedata
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook

from sipsa_insumos.pipelines.quality.nodes import _LLAVE_DUPLICADOS
from sipsa_insumos.utils.excel_writer import (
    escribir_excel_multisheet,
    escribir_hoja_freq,
    escribir_hoja_pivot_cpc,
)

log = logging.getLogger(__name__)

# Etiquetas SAS por módulo para CUADROS y ANEXO (nombres de archivo SAS oficiales)
_LABEL_SAS: dict[str, str] = {
    "AGRICOLAS":   "INSUMOS AGRICOLAS",
    "PECUARIOS":   "INSUMOS PECUARIOS",
    "ELEMENTOS":   "ELEMENTOS AGROPECUARIOS",
    "EMPAQUES":    "EMPAQUES AGROPECUARIOS",
    "PROPAGACION": "MATERIAL PROPAGACION",
    "ARRIENDOS":   "ARRIENDOS",
    "SERVICIOS":   "SERVICIOS",
    "JORNALES":    "JORNALES",
    "ESPECIES":    "ESPECIES PRODUCTIVAS",
}

# BASES usa "MATERIAL PROPAGACIÓN" (con tilde) según archivos SAS de referencia
_LABEL_SAS_BASES: dict[str, str] = {
    **_LABEL_SAS,
    "PROPAGACION": "MATERIAL PROPAGACIÓN",
}

# TABLAS: etiquetas más cortas para algunos módulos (tal como genera SAS)
_LABEL_SAS_TABLAS: dict[str, str] = {
    **_LABEL_SAS,
    "ELEMENTOS":   "ELEMENTOS",
    "EMPAQUES":    "EMPAQUES",
    "PROPAGACION": "MATERIAL",
    "ESPECIES":    "ESPECIES PRODUC",
}


# Hoja única de módulos de un solo grupo (SAS no la nombra por grupo).
_HOJA_UNICA: dict[str, dict[str, str]] = {
    "BASES": {"ELEMENTOS": "Elementos", "EMPAQUES": "Hoja1", "ARRIENDOS": "Hoja1",
              "SERVICIOS": "Hoja1", "PROPAGACION": "Material", "JORNALES": "Hoja1",
              "ESPECIES": "Espe"},
    "ANEXO": {"ELEMENTOS": "Hoja1", "EMPAQUES": "Hoja1", "ARRIENDOS": "Hoja1",
              "SERVICIOS": "Hoja1", "PROPAGACION": "Material", "JORNALES": "Hoja1",
              "ESPECIES": "Hoja1"},
    "CUADROS": {m: "Tabulate 1 - Tabla 1" for m in
                ("ELEMENTOS", "EMPAQUES", "ARRIENDOS", "SERVICIOS", "PROPAGACION",
                 "JORNALES", "ESPECIES")},
}

# Columnas que SAS ordena como número (el resto como texto, orden de bytes).
_COLS_ORDEN_NUMERICO = {"CodigoMpio", "Codigo CPC", "CÓDIGO CPC", "CÓDIGO DIVIPOLA"}


# Módulos cuyo programa SAS guarda el Código CPC como texto: se ordena
# alfabéticamente ("1121" < "113101" < "1171"), no como número (Material AGO2026).
_CPC_TEXTO = {"PROPAGACION"}
_COLS_CPC = {"Codigo CPC", "CÓDIGO CPC"}


def _ordenar_sas(df: pd.DataFrame, cols: list[str], modulo: str = "") -> pd.DataFrame:
    """Ordena como PROC SORT de SAS: estable (empates conservan el orden de
    entrada), códigos como número, texto por orden de bytes y faltantes primero."""
    cols = [c for c in cols if c in df.columns]
    if df.empty or not cols:
        return df
    cpc_texto = modulo.upper() in _CPC_TEXTO

    def _clave(serie: pd.Series) -> pd.Series:
        if cpc_texto and serie.name in _COLS_CPC:
            txt = serie.astype(str).str.replace(r"\.0$", "", regex=True).str.lstrip("0")
            return txt.where(serie.notna())
        if serie.name in _COLS_ORDEN_NUMERICO or pd.api.types.is_numeric_dtype(serie):
            return pd.to_numeric(serie, errors="coerce")
        return serie.where(serie.isna(), serie.astype(str))

    return df.sort_values(cols, key=_clave, kind="mergesort", na_position="first")


def _hoja(tipo: str, modulo: str, grupo: str) -> str:
    return _HOJA_UNICA.get(tipo, {}).get(modulo.upper(), grupo)


def _nombre_sas(tipo: str, modulo: str, periodo: str, label_dict: dict[str, str] | None = None) -> str:
    """Construye el nombre de archivo siguiendo la convención SAS."""
    d = label_dict if label_dict is not None else _LABEL_SAS
    label = d.get(modulo.upper(), modulo.upper())
    return f"{tipo} {label} {periodo}.xlsx"


def exportar_bases(
    base_comparada: pd.DataFrame,
    grupos: list[str],
    modulo: str,
    periodo: str,
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta la base de precios por municipio, una hoja por grupo.

    Formato SAS: una fila por (municipio × Nombre_Publica) con precio actual,
    precio anterior, variación, cuenta de fuentes y tendencia.
    Columnas: CodigoDepto, NombreDepartamento, CodigoMpio, NombreMunicipio,
              Mercado, Codigo_CPC, Nombre_Publica, PRECIO_ABR, PRECIO_ACT,
              Variacion(%), Cuenta, Tendencia.

    Args:
        base_comparada: DataFrame con columnas de precio, variación y tendencia
                        a nivel (municipio × Nombre_Publica).
        grupos: Lista de nombres de grupo del módulo.
        modulo: Nombre del módulo (ej: "AGRICOLAS").
        periodo: ID del período (ej: "MAY2026").
        ruta_reporting: Directorio raíz de reportes.

    Returns:
        DataFrame de metadatos (archivo, hojas, filas_totales).
    """
    nombre = _nombre_sas("BASES", modulo, periodo, _LABEL_SAS_BASES)
    ruta = Path(ruta_reporting) / modulo.lower() / nombre
    ruta.parent.mkdir(parents=True, exist_ok=True)

    periodo_ant = _periodo_anterior_modulo(periodo, modulo)
    m = modulo.upper()
    col_pub_sas = _NOMBRE_PUBLICA_SAS.get(m, "Nombre_Publica")

    # Columnas de precio actual y anterior generadas en calcular_variacion_tendencia
    col_precio_actual = next(
        (c for c in base_comparada.columns if c.startswith("PRECIO_")
         and c not in ("PRECIO_MIN", "PRECIO_MAX") and periodo_ant not in c
         and not any(mes in c for mes in ["Abril","Enero","Febrero","Marzo","Mayo",
                                           "Junio","Julio","Agosto","Septiembre",
                                           "Octubre","Noviembre","Diciembre"])),
        next((c for c in base_comparada.columns if c.startswith("PRECIO_")
              and c not in ("PRECIO_MIN", "PRECIO_MAX")), None)
    )
    col_precio_anterior = next(
        (c for c in base_comparada.columns if c.startswith("PRECIO_")
         and c != col_precio_actual and c not in ("PRECIO_MIN", "PRECIO_MAX")), None
    )

    hojas: dict[str, pd.DataFrame] = {}
    for grupo in grupos:
        df_grupo = base_comparada[base_comparada["Grupo"] == grupo].copy()
        if df_grupo.empty:
            hojas[grupo] = pd.DataFrame()
            continue

        if "NombreMunicipio" in df_grupo.columns and "NombreDepartamento" in df_grupo.columns:
            df_grupo["Mercado"] = (
                df_grupo["NombreMunicipio"] + " (" + df_grupo["NombreDepartamento"] + ")"
            )

        # Derivar CodigoDepto y CodigoMpio desde CÓDIGO DIVIPOLA
        col_div = next((c for c in ["CÓDIGO DIVIPOLA", "CodigoMpio"] if c in df_grupo.columns), None)
        if col_div and "CodigoDepto" not in df_grupo.columns:
            df_grupo["CodigoDepto"] = df_grupo[col_div].astype(str).str[:2]

        # Cuenta SAS = número de municipios (mercados) a nivel nacional que
        # reportan este Nombre_Publica en el período — NO es N_FUENTE
        # (fuentes de precio de un solo municipio). Se repite el mismo valor
        # en todas las filas del mismo producto.
        if "Nombre_Publica" in df_grupo.columns:
            df_grupo["N_MUNICIPIOS_PUBLICA"] = df_grupo.groupby("Nombre_Publica")["Nombre_Publica"].transform("size")

        col_map = {
            "CodigoDepto":                   "CodigoDepto" if "CodigoDepto" in df_grupo.columns else None,
            "NombreDepartamento":            "NombreDepartamento" if "NombreDepartamento" in df_grupo.columns else None,
            "CodigoMpio":                    col_div,
            "NombreMunicipio":               "NombreMunicipio" if "NombreMunicipio" in df_grupo.columns else None,
            "Mercado":                       "Mercado" if "Mercado" in df_grupo.columns else None,
            "Codigo CPC":                    "CÓDIGO CPC" if "CÓDIGO CPC" in df_grupo.columns and m not in _SIN_CODIGO_CPC else None,
            col_pub_sas:                     "Nombre_Publica" if "Nombre_Publica" in df_grupo.columns else None,
            f"PRECIO_PROMEDIO_{periodo_ant}": col_precio_anterior,
            f"PRECIO_PROMEDIO_{periodo}":     col_precio_actual,
            "Variacion(%)":                  "VARIACION" if "VARIACION" in df_grupo.columns else None,
            "Cuenta":                        "N_MUNICIPIOS_PUBLICA" if "N_MUNICIPIOS_PUBLICA" in df_grupo.columns else None,
            "Tendencia":                     "TENDENCIA" if "TENDENCIA" in df_grupo.columns else None,
        }
        cols_sel = [v for v in col_map.values() if v and v in df_grupo.columns]
        rename_inv = {v: k for k, v in col_map.items() if v and v in df_grupo.columns}
        df_out = df_grupo[cols_sel].rename(columns=rename_inv)
        df_out = _ordenar_sas(df_out, ["CodigoMpio", "Codigo CPC", col_pub_sas], m)
        hojas[_hoja("BASES", m, grupo)] = df_out

    escribir_excel_multisheet(ruta, hojas)

    filas_totales = sum(len(df) for df in hojas.values())
    log.info("exportar_bases [%s] OK | archivo=%s | hojas=%d | filas=%d",
             modulo, nombre, len(hojas), filas_totales)
    return pd.DataFrame([{
        "archivo": str(ruta),
        "modulo": modulo,
        "periodo": periodo,
        "tipo": "BASES",
        "hojas": len(hojas),
        "filas_totales": filas_totales,
    }])


# Nombre de hoja SAS por grupo en TABLAS (PROC TABULATE → hojas TD_*),
# capturado de referencias JUL/AGO2026. JORNALES/ESPECIES sin referencia:
# conservan el nombre del grupo.
_TABLAS_SHEETS: dict[str, dict[str, str]] = {
    "AGRICOLAS": {"COADYUDANTES": "TD_Coady", "FERTILIZANTES": "TD_Ferti", "FUNGICIDAS": "TD_Fungi",
                  "HERBICIDAS": "TD_Herbic", "INSECTICIDAS": "TD_InseAgr", "BIOINSUMOS": "TD_BioInsu"},
    "PECUARIOS": {"ALIMENTOS": "TD_Alim", "ANTIBIOTICOS": "TD_Antibio", "ANTISEPTICOS": "TD_Antisep",
                  "HORMONALES": "TD_Hormon", "INSECTICIDAS": "TD_InsePec", "MEDICAMENTOS": "TD_Medic",
                  "VITAMINAS": "TD_Vitam"},
    "ELEMENTOS":   {"ELEMENTOS": "TD_Elementos"},
    "EMPAQUES":    {"EMPAQUES": "TD_Empaques"},
    "ARRIENDOS":   {"ARRIENDOS": "TD_Arriendo"},
    "SERVICIOS":   {"SERVICIOS": "TD_Servicios"},
    "PROPAGACION": {"MATERIAL DE PROPAGACION": "TD_Material"},
    "JORNALES":    {"JORNALES": "TD_Jornales"},
    "ESPECIES":    {"ESPECIES PRODUCTIVAS": "TD_Espec"},
}


def exportar_tablas(
    base_comparada: pd.DataFrame,
    grupos: list[str],
    modulo: str,
    periodo: str,
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta tablas dinámicas Producto × Tendencia, una hoja por grupo.

    SAS: PROC TABULATE: Nombre_Publica × (Positiva, Negativa, Estable, n.d.) → conteo

    Args:
        base_comparada: DataFrame con columnas Nombre_Publica, Grupo, TENDENCIA.
        grupos: Lista de nombres de grupo del módulo.
        modulo: Nombre del módulo.
        periodo: ID del período.
        ruta_reporting: Directorio raíz de reportes.

    Returns:
        DataFrame de metadatos.
    """
    nombre = _nombre_sas("TABLAS", modulo, periodo, _LABEL_SAS_TABLAS)
    ruta = Path(ruta_reporting) / modulo.lower() / nombre
    ruta.parent.mkdir(parents=True, exist_ok=True)

    hojas: dict[str, pd.DataFrame] = {}
    for grupo in grupos:
        df_grupo = base_comparada[base_comparada["Grupo"] == grupo].copy()
        if df_grupo.empty:
            hojas[grupo] = pd.DataFrame()
            continue

        pivot = (
            df_grupo.groupby(["CÓDIGO CPC", "Nombre_Publica", "TENDENCIA"], observed=True)
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        for tend in ["Positiva", "Negativa", "Estable", "n.d."]:
            if tend not in pivot.columns:
                pivot[tend] = 0
        pivot["Total"] = pivot[["Positiva", "Negativa", "Estable", "n.d."]].sum(axis=1)
        pivot = _ordenar_sas(pivot, ["CÓDIGO CPC", "Nombre_Publica"], modulo)
        hojas[grupo] = pivot

    wb = Workbook()
    wb.remove(wb.active)
    hojas_sas = _TABLAS_SHEETS.get(modulo.upper(), {})
    for nombre_hoja, df_hoja in hojas.items():
        ws = wb.create_sheet(title=hojas_sas.get(nombre_hoja, nombre_hoja)[:31])
        if df_hoja.empty:
            continue
        if modulo.upper() in _SIN_CODIGO_CPC:
            # Arriendos/Servicios: SAS usa PROC FREQ (sin CPC, filas por nombre).
            escribir_hoja_freq(
                ws,
                _ordenar_sas(df_hoja, ["Nombre_Publica"]),
                prod_col="Nombre_Publica",
                var_label=_NOMBRE_PUBLICA_SAS[modulo.upper()],
                metric_cols=["Positiva", "Negativa", "Estable", "n.d.", "Total"],
                sub_headers=["1.Positiva", "2.Negativa", "3.Estable", "4.n.d.", "Total"],
            )
            continue
        escribir_hoja_pivot_cpc(
            ws,
            df_hoja,
            cpc_col="CÓDIGO CPC",
            prod_col="Nombre_Publica",
            metric_cols=["Positiva", "Negativa", "Estable", "n.d.", "Total"],
            sub_headers=["1.Positiva", "2.Negativa", "3.Estable", "4.n.d.", "Total"],
            grupos_header=[("Tendencia_mod", 4)],
            cols_punto_si_cero={"Positiva", "Negativa", "Estable", "n.d."},
            cols_numero={"Total"},
            total_label="Total",
            prod_header=_NOMBRE_PUBLICA_SAS[modulo],
        )
    wb.save(str(ruta))

    filas_totales = sum(len(df) for df in hojas.values())
    log.info("exportar_tablas [%s] OK | archivo=%s | hojas=%d | filas=%d",
             modulo, nombre, len(hojas), filas_totales)
    return pd.DataFrame([{
        "archivo": str(ruta),
        "modulo": modulo,
        "periodo": periodo,
        "tipo": "TABLAS",
        "hojas": len(hojas),
        "filas_totales": filas_totales,
    }])


def _partir_nombre_publica(nombre, m: str) -> tuple:
    """Nombre_insumo / Presentación_insumo como los arma cada programa SAS con
    SCAN(nombre, k, ','), que omite partes vacías y quita espacios:
      - estándar: si hay 3+ partes, insumo = "p1,p2" ("p1, p2" en Material); presentación = última.
      - Jornales: insumo = p1; presentación = "p-2,p-1" si hay exactamente 2 comas.
      - Especies: si hay 3+ partes, insumo = "p1,p2,p3"; presentación = última.
    """
    if not isinstance(nombre, str):
        return (nombre, nombre)
    partes = [p.strip() for p in nombre.split(",") if p.strip()]
    if not partes:
        return ("", "")
    if m == "JORNALES":
        pres = f"{partes[-2]},{partes[-1]}" if nombre.count(",") == 2 and len(partes) >= 2 else partes[-1]
        return (partes[0], pres)
    if len(partes) < 3:
        return (partes[0], partes[-1])
    n = 3 if m == "ESPECIES" else 2
    sep = ", " if m == "PROPAGACION" else ","   # Material concatena con ", "
    return (sep.join(partes[:n]), partes[-1])


# ANEXO con orden de columnas propio (ARRIENDOS AGO2026: el nombre de
# publicación va antes de los precios y el insumo/presentación al final).
_ANEXO_ORDEN: dict[str, list[str]] = {
    "JORNALES": ["CodigoDepto", "NombreDepartamento", "CodigoMpio", "NombreMunicipio", "Codigo CPC", "__PUBLICA__",
                 "__PRECIO_ANT__", "__PRECIO_ACT__", "Variacion(%)", "Cuenta", "Tendencia", "Mercado",
                 "Nombre_insumo", "Presentación_insumo"],
    "ARRIENDOS": ["CodigoDepto", "NombreDepartamento", "CodigoMpio", "NombreMunicipio", "__PUBLICA__",
                  "__PRECIO_ANT__", "__PRECIO_ACT__", "Variacion(%)", "Cuenta", "Tendencia",
                  "Nombre_insumo", "Presentación_insumo", "Mercado"],
}


def exportar_anexos(
    base_comparada: pd.DataFrame,
    grupos: list[str],
    modulo: str,
    periodo: str,
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta el anexo por municipio con nombre e insumo por separado, una hoja por grupo.

    Formato SAS ANEXO: igual que BASES pero con Nombre_Publica desglosado en
    Nombre_insumo (texto antes de la coma) y Presentación_insumo (texto después).
    Columnas: CodigoDepto, NombreDepartamento, CodigoMpio/DIVIPOLA, NombreMunicipio,
              Codigo_CPC, Nombre_insumo, Presentación_insumo,
              PRECIO_ABR, PRECIO_ACT, Variacion(%), Cuenta, Tendencia, Mercado,
              Nombre_Publica.

    Args:
        base_comparada: DataFrame con columnas de precio, variación y tendencia
                        a nivel (municipio × Nombre_Publica).
        grupos: Lista de nombres de grupo.
        modulo: Nombre del módulo.
        periodo: ID del período.
        ruta_reporting: Directorio raíz de reportes.

    Returns:
        DataFrame de metadatos.
    """
    nombre = _nombre_sas("ANEXO", modulo, periodo)
    ruta = Path(ruta_reporting) / modulo.lower() / nombre
    ruta.parent.mkdir(parents=True, exist_ok=True)

    periodo_ant = _periodo_anterior_modulo(periodo, modulo)
    m = modulo.upper()
    col_pub_sas = _NOMBRE_PUBLICA_SAS.get(m, "Nombre_Publica")

    col_precio_actual = next(
        (c for c in base_comparada.columns if c.startswith("PRECIO_")
         and c not in ("PRECIO_MIN", "PRECIO_MAX") and periodo_ant not in c
         and not any(mes in c for mes in ["Abril","Enero","Febrero","Marzo","Mayo",
                                           "Junio","Julio","Agosto","Septiembre",
                                           "Octubre","Noviembre","Diciembre"])),
        next((c for c in base_comparada.columns if c.startswith("PRECIO_")
              and c not in ("PRECIO_MIN", "PRECIO_MAX")), None)
    )
    col_precio_anterior = next(
        (c for c in base_comparada.columns if c.startswith("PRECIO_")
         and c != col_precio_actual and c not in ("PRECIO_MIN", "PRECIO_MAX")), None
    )

    hojas: dict[str, pd.DataFrame] = {}
    for grupo in grupos:
        df_grupo = base_comparada[base_comparada["Grupo"] == grupo].copy()
        if df_grupo.empty:
            hojas[grupo] = pd.DataFrame()
            continue

        if "NombreMunicipio" in df_grupo.columns and "NombreDepartamento" in df_grupo.columns:
            df_grupo["Mercado"] = (
                df_grupo["NombreMunicipio"] + " (" + df_grupo["NombreDepartamento"] + ")"
            )

        if "Nombre_Publica" in df_grupo.columns:
            partidos = df_grupo["Nombre_Publica"].map(lambda x: _partir_nombre_publica(x, m))
            df_grupo["Nombre_insumo"] = partidos.str[0]
            df_grupo["Presentación_insumo"] = partidos.str[1]

        col_div = next((c for c in ["CÓDIGO DIVIPOLA", "CodigoMpio"] if c in df_grupo.columns), None)
        if col_div and "CodigoDepto" not in df_grupo.columns:
            df_grupo["CodigoDepto"] = df_grupo[col_div].astype(str).str[:2]

        # Cuenta SAS = número de municipios (mercados) a nivel nacional que
        # reportan este Nombre_Publica en el período — NO es N_FUENTE
        # (fuentes de precio de un solo municipio). Se repite el mismo valor
        # en todas las filas del mismo producto.
        if "Nombre_Publica" in df_grupo.columns:
            df_grupo["N_MUNICIPIOS_PUBLICA"] = df_grupo.groupby("Nombre_Publica")["Nombre_Publica"].transform("size")

        col_map = {
            "CodigoDepto":                   "CodigoDepto" if "CodigoDepto" in df_grupo.columns else None,
            "NombreDepartamento":            "NombreDepartamento" if "NombreDepartamento" in df_grupo.columns else None,
            "CodigoMpio":                    col_div,
            "NombreMunicipio":               "NombreMunicipio" if "NombreMunicipio" in df_grupo.columns else None,
            "Codigo CPC":                    "CÓDIGO CPC" if "CÓDIGO CPC" in df_grupo.columns and m not in _SIN_CODIGO_CPC else None,
            "Nombre_insumo":                 "Nombre_insumo" if "Nombre_insumo" in df_grupo.columns else None,
            "Presentación_insumo":           "Presentación_insumo" if "Presentación_insumo" in df_grupo.columns else None,
            f"PRECIO_PROMEDIO_{periodo_ant}": col_precio_anterior,
            f"PRECIO_PROMEDIO_{periodo}":     col_precio_actual,
            "Variacion(%)":                  "VARIACION" if "VARIACION" in df_grupo.columns else None,
            "Cuenta":                        "N_MUNICIPIOS_PUBLICA" if "N_MUNICIPIOS_PUBLICA" in df_grupo.columns else None,
            "Tendencia":                     "TENDENCIA" if "TENDENCIA" in df_grupo.columns else None,
            "Mercado":                       "Mercado" if "Mercado" in df_grupo.columns else None,
            col_pub_sas:                     "Nombre_Publica" if "Nombre_Publica" in df_grupo.columns else None,
        }
        cols_sel = [v for v in col_map.values() if v and v in df_grupo.columns]
        rename_inv = {v: k for k, v in col_map.items() if v and v in df_grupo.columns}
        df_out = df_grupo[cols_sel].rename(columns=rename_inv)
        if m in _ANEXO_ORDEN:
            orden = [col_pub_sas if c == "__PUBLICA__" else
                     f"PRECIO_PROMEDIO_{periodo_ant}" if c == "__PRECIO_ANT__" else
                     f"PRECIO_PROMEDIO_{periodo}" if c == "__PRECIO_ACT__" else c
                     for c in _ANEXO_ORDEN[m]]
            df_out = df_out[[c for c in orden if c in df_out.columns]]
        df_out = _ordenar_sas(df_out, ["CodigoMpio", "Codigo CPC", col_pub_sas], m)
        hojas[_hoja("ANEXO", m, grupo)] = df_out

    escribir_excel_multisheet(ruta, hojas)

    filas_totales = sum(len(df) for df in hojas.values())
    log.info("exportar_anexos [%s] OK | archivo=%s | hojas=%d | filas=%d",
             modulo, nombre, len(hojas), filas_totales)
    return pd.DataFrame([{
        "archivo": str(ruta),
        "modulo": modulo,
        "periodo": periodo,
        "tipo": "ANEXO",
        "hojas": len(hojas),
        "filas_totales": filas_totales,
    }])


def _construir_base_diagnostico(
    base_completa: pd.DataFrame,
    divipola_raw: pd.DataFrame,
    mappings_grupos: dict,
    mappings_articulos: dict,
) -> pd.DataFrame:
    """Enriquece base_completa (fila por fila, sin filtrar) con DIVIPOLA,
    Grupo y Nombre_Publica, y calcula VAR con el mismo criterio SAS: precio
    en 0 → nulo, y si Nov. IN ('IA','IN') el precio anterior también se anula.

    Insumo común para BASE_INSUMOS y los archivos de diagnóstico
    (FALTAN_GRUPO/FALTAN_PUBLICA/DUPLI/VAR_ATIPICO).
    """
    divipola = divipola_raw.copy()
    if "CódigoMunicipio" in divipola.columns:
        divipola = divipola.rename(columns={"CódigoMunicipio": "CÓDIGO DIVIPOLA"})
        divipola["CÓDIGO DIVIPOLA"] = divipola["CÓDIGO DIVIPOLA"].str.zfill(5)
    elif "CodigoMpio" in divipola.columns:
        divipola = divipola.rename(columns={"CodigoMpio": "CÓDIGO DIVIPOLA"})

    df = base_completa.merge(divipola, on="CÓDIGO DIVIPOLA", how="left")

    grupos_dict: dict[str, str] = mappings_grupos.get("grupos", mappings_grupos)
    articulos_dict: dict[str, str] = mappings_articulos.get("articulos_publicacion", mappings_articulos)
    df["Grupo"] = df["LLAVE_ARTICULO"].map(grupos_dict)
    df["Nombre_Publica"] = df["LLAVE_ARTICULO"].map(articulos_dict)

    # Recalcular variación fila por fila (equiv. SAS): precio en 0 → nulo, y
    # si Nov. in (IA, IN) el precio anterior también se anula. Se sobreescriben
    # las columnas de precio (no solo un cálculo auxiliar) para que Abril/Mayo
    # en la salida coincidan exactamente con SAS.
    df["PRECIO"] = df["PRECIO"].where(df["PRECIO"] != 0)
    df["PRECIO_ANTERIOR_RAW"] = df["PRECIO_ANTERIOR_RAW"].where(df["PRECIO_ANTERIOR_RAW"] != 0)
    df["PRECIO_ANTERIOR_RAW"] = df["PRECIO_ANTERIOR_RAW"].where(~df["NOV"].isin(["IA", "IN"]))
    df["VAR"] = (df["PRECIO"] - df["PRECIO_ANTERIOR_RAW"]) / df["PRECIO_ANTERIOR_RAW"] * 100

    df["CodigoDepto"] = df["CÓDIGO DIVIPOLA"].str[:2]
    return df


def _catx_display(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    """CATX case-preservado (a diferencia de LLAVE_ARTICULO, que va en
    mayúsculas para el cruce de Grupo/Nombre_Publica): concatena columnas
    con "_" tratando nulos como "NA" — igual que SAS CATX con missing string,
    pero sin tocar el case original de cada parte."""
    def _limpiar(s: pd.Series) -> pd.Series:
        s = s.astype(str).str.strip()
        return s.mask(s.isin(["nan", "None", "NaT", ""]), "NA")
    partes = [_limpiar(df[c]) for c in cols]
    out = partes[0]
    for p in partes[1:]:
        out = out + "_" + p
    return out


def _agregar_llave_display(df: pd.DataFrame, m: str) -> None:
    """Agrega _LLAVE_DISPLAY (y _LLAVE_DISPLAY_2 en Propagación): la llave que
    SAS muestra en BASE_INSUMOS y en los diagnósticos, con el orden de partes
    propio de cada programa SAS."""
    if m in ("AGRICOLAS", "PECUARIOS"):
        df["_LLAVE_DISPLAY"] = _catx_display(df, ["ARTÍCULO", "UNIDAD DE MEDIDA", "CASA COMERCIAL", "REGISTRO ICA"])
    elif m == "ELEMENTOS":
        df["_LLAVE_DISPLAY"] = _catx_display(df, ["ARTÍCULO", "CASA COMERCIAL", "REGISTRO ICA", "UNIDAD DE MEDIDA"])
    elif m in ("EMPAQUES", "ARRIENDOS", "SERVICIOS", "JORNALES", "ESPECIES"):
        df["_LLAVE_DISPLAY"] = _catx_display(df, ["ARTÍCULO", "CARACTERÍSTICA"])
    elif m == "PROPAGACION":
        df["_LLAVE_DISPLAY"] = _catx_display(df, ["ARTÍCULO", "UNIDAD DE MEDIDA"])
        df["_LLAVE_DISPLAY_2"] = _catx_display(df, ["ARTÍCULO", "CASA COMERCIAL", "REGISTRO ICA", "UNIDAD DE MEDIDA"])


# Módulos cuyo BASE_INSUMOS usa el mes calendario genérico (mes_anterior) en
# vez del mes según periodicidad del módulo — ver nota en exportar_base_insumos.
_BASE_INSUMOS_MES_ANT_CALENDARIO: set[str] = set()   # AGO2026: MATERIAL usa "Junio" (bimestral), ya no el mes calendario

# Nombre de archivo BASE_INSUMOS por módulo — cada programa SAS usa su propio
# nombre (capturado de archivos de referencia reales); Agrícolas/Pecuarios
# son los únicos que siguen el patrón por defecto BASE_INSUMOS_{MODULO}.
_BASE_INSUMOS_NOMBRES: dict[str, str] = {
    "ELEMENTOS":   "BASE_ELEMENTOS_AGROPECUARIOS",
    "EMPAQUES":    "BASE_EMPAQUES_AGROPECUARIOS",
    "ARRIENDOS":   "BASE_ARRIENDOS",
    "SERVICIOS":   "BASE_SERVICIOS",
    "PROPAGACION": "BASE_MATERIAL_PROPAGACION",
    "JORNALES":    "BASE_JORNALES",
    "ESPECIES":    "BASE_ESPECIE_PRODUCTIVA",
}

# Orden exacto de columnas SAS por módulo — capturado de archivos de
# referencia reales (JUL2026 para Agrícolas/Pecuarios/Elementos/Empaques,
# FEB2026 para Arriendos/Servicios/Propagación). Cada programa SAS ordena
# distinto — no sigue una fórmula única (mismo patrón que _DIAG_ORDEN_*).
# JORNALES/ESPECIES no tienen archivo de referencia disponible todavía —
# se construyen por analogía con Arriendos/Servicios (mismo tipo_modulo
# "caracte" y periodicidad trimestral) — SIN VALIDAR contra SAS real.
_BASE_INSUMOS_ORDEN: dict[str, list[str]] = {
    "AGRICOLAS":   ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "Grupo", "UnMed.", "CasaCom.", "RegICA", "__MES_ANT__", "__MES_ACT__", "VAR", "Estado", "Art_Unmed_Casacomer_ICA", "Observación", "Nov."],
    "PECUARIOS":   ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "Grupo", "UnMed.", "CasaCom.", "RegICA", "__MES_ANT__", "__MES_ACT__", "VAR", "Nov.", "Estado", "Art_Unmed_Casacomer_ICA", "Observación"],
    "ELEMENTOS":   ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "UnMed.", "CasaCom.", "RegICA", "Estado", "Art_Casacomer_ICA_Unmed", "Nov."],
    "EMPAQUES":    ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Estado", "Articulo_Caracte", "Informante", "Caracte.", "Nov."],
    "ARRIENDOS":   ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Codigo CPC", "Caracte#"],
    "SERVICIOS":   ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Codigo CPC", "Caracte#"],
    "PROPAGACION": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "UnMed#", "CasaCom#", "RegICA", "Nov#", "Estado", "Articulo_Unmed", "Art_Casacomer_ICA_Unmed"],
    "JORNALES":    ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Caracte#"],
    "ESPECIES":    ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Caracte#"],
}


def exportar_base_insumos(
    base_completa: pd.DataFrame,
    divipola_raw: pd.DataFrame,
    mappings_grupos: dict,
    mappings_articulos: dict,
    modulo: str,
    periodo: str,
    mes_actual: str,
    mes_anterior: str,
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta BASE_{...}_{PERIODO}.xlsx — réplica exacta del equivalente
    SAS: la base completa (sin filtrar por Estado ni Precio, a diferencia
    del flujo principal) enriquecida con DIVIPOLA/Grupo/Nombre_Publica, con
    la variación calculada fila por fila usando el precio anterior embebido
    en el propio archivo del mes (columna "Precio Ante."), no el promedio
    agregado del período anterior. Nombre de archivo y orden de columnas
    varían por módulo — ver _BASE_INSUMOS_NOMBRES / _BASE_INSUMOS_ORDEN.

    SAS (paso final del programa de CUADROS):
      MERGE DIVIPOLA base; BY CodigoMpio; IF B;             (conserva sin cruce)
      MERGE base grupos/articulos; BY LLAVE; IF A;          (conserva sin cruce)
      IF Precio Ante./Actual = 0 THEN falta;
      IF Nov. IN ('IA','IN') THEN Precio Ante. = falta;
      VAR = (Actual - Ante) / Ante * 100;

    La columna llave que SAS muestra (Art_Unmed_Casacomer_ICA / Articulo_Caracte
    / etc.) es un CATX case-preservado de sus partes — NO reutiliza
    LLAVE_ARTICULO, que va en mayúsculas y solo sirve para el cruce interno
    de Grupo/Nombre_Publica.

    Args:
        base_completa: Salida de leer_base_completa (sin filtrar).
        divipola_raw: DIVIPOLA maestro (municipios/departamentos).
        mappings_grupos: Dict {"grupos": {...}}.
        mappings_articulos: Dict {"articulos_publicacion": {...}}.
        modulo: Nombre del módulo en mayúsculas.
        periodo: Código del período (ej: "MAY2026").
        mes_actual: Nombre del mes actual en español (ej: "Mayo").
        mes_anterior: Nombre del mes anterior en español (ej: "Abril").
        ruta_reporting: Directorio raíz de reportes.

    Returns:
        DataFrame de metadatos (archivo, filas).
    """
    m = modulo.upper()
    df_out, _ = _base_insumos_sas(base_completa, divipola_raw, mappings_grupos, mappings_articulos,
                                  modulo, periodo, mes_actual, mes_anterior)
    nombre = f"{_BASE_INSUMOS_NOMBRES.get(m, f'BASE_INSUMOS_{m}')}_{periodo}.xlsx"
    ruta = Path(ruta_reporting) / modulo.lower() / nombre
    escribir_excel_multisheet(ruta, {"INFORMACIÓN INSUMOS": df_out})
    n = len(df_out)

    log.info("exportar_base_insumos [%s] OK | archivo=%s | filas=%d", modulo, nombre, n)
    return pd.DataFrame([{
        "archivo": str(ruta),
        "modulo": modulo,
        "periodo": periodo,
        "tipo": "BASE_INSUMOS",
        "filas_totales": n,
    }])


def _base_insumos_sas(
    base_completa: pd.DataFrame,
    divipola_raw: pd.DataFrame,
    mappings_grupos: dict,
    mappings_articulos: dict,
    modulo: str,
    periodo: str,
    mes_actual: str,
    mes_anterior: str,
) -> tuple[pd.DataFrame, str]:
    """Base fila por fila con etiquetas y orden SAS (tabla X6A de los programas
    SAS). Devuelve (base, nombre del mes anterior usado como columna)."""
    m = modulo.upper()
    df = _construir_base_diagnostico(base_completa, divipola_raw, mappings_grupos, mappings_articulos)

    # "Precio Ante." usa el mes anterior SEGÚN LA PERIODICIDAD DEL MÓDULO en
    # la mayoría de programas SAS (ej. "Mayo" para Empaques bimestral en
    # JUL2026, confirmado contra referencia real) — excepto Propagación, cuyo
    # BASE_MATERIAL_PROPAGACION real usa el mes calendario genérico ("Enero"
    # para FEB2026, no "Diciembre") — confirmado contra referencia FEB2026.
    if m in _BASE_INSUMOS_MES_ANT_CALENDARIO:
        mes_anterior_mod = mes_anterior
    else:
        periodo_ant_mod = _periodo_anterior_modulo(periodo, m)
        mes_anterior_mod = _ABBR_A_MES_LARGO.get(periodo_ant_mod[:3], mes_anterior)

    _agregar_llave_display(df, m)
    _truncar_observacion(df, m)

    col_map: dict[str, str | None] = {
        "CodigoDepto":        "CodigoDepto",
        "CodigoMpio":         "CÓDIGO DIVIPOLA",
        "NombreDepartamento": "NombreDepartamento" if "NombreDepartamento" in df.columns else None,
        "NombreMunicipio":    "NombreMunicipio" if "NombreMunicipio" in df.columns else None,
        "Fuente":             "FUENTE",
        "Codigo CPC":         "CÓDIGO CPC",
        "Articulo":           "ARTÍCULO",
        "__PUBLICA__":        "Nombre_Publica",
        "Grupo":              "Grupo",
        "__MES_ANT__":        "PRECIO_ANTERIOR_RAW",
        "__MES_ACT__":        "PRECIO",
        "VAR":                "VAR",
        "Observación":        "OBSERVACION",
        "Estado":             "ESTADO",
        "Informante":         "INFORMANTE" if "INFORMANTE" in df.columns else None,
    }
    if m in ("AGRICOLAS", "PECUARIOS"):
        col_map.update({
            "UnMed.": "UNIDAD DE MEDIDA", "CasaCom.": "CASA COMERCIAL", "RegICA": "REGISTRO ICA",
            "Art_Unmed_Casacomer_ICA": "_LLAVE_DISPLAY", "Nov.": "NOV",
        })
    elif m == "ELEMENTOS":
        col_map.update({
            "UnMed.": "UNIDAD DE MEDIDA", "CasaCom.": "CASA COMERCIAL", "RegICA": "REGISTRO ICA",
            "Art_Casacomer_ICA_Unmed": "_LLAVE_DISPLAY", "Nov.": "NOV",
        })
    elif m == "EMPAQUES":
        col_map.update({"Articulo_Caracte": "_LLAVE_DISPLAY", "Caracte.": "CARACTERÍSTICA", "Nov.": "NOV"})
    elif m in ("ARRIENDOS", "SERVICIOS", "JORNALES", "ESPECIES"):
        col_map.update({"Articulo_Caracte": "_LLAVE_DISPLAY", "Caracte#": "CARACTERÍSTICA", "Nov#": "NOV"})
    elif m == "PROPAGACION":
        col_map.update({
            "UnMed#": "UNIDAD DE MEDIDA", "CasaCom#": "CASA COMERCIAL", "RegICA": "REGISTRO ICA", "Nov#": "NOV",
            "Articulo_Unmed": "_LLAVE_DISPLAY", "Art_Casacomer_ICA_Unmed": "_LLAVE_DISPLAY_2",
        })

    orden = _BASE_INSUMOS_ORDEN.get(m)
    if orden is None:
        raise ValueError(f"exportar_base_insumos: módulo sin orden de columnas definido: {modulo}")
    return _armar_diag_hoja(df, orden, col_map, m, mes_anterior_mod, mes_actual), mes_anterior_mod


# =============================================================================
# Archivos de diagnóstico (FALTAN_GRUPO / FALTAN_PUBLICA / DUPLI / VAR_ATIPICO)
# =============================================================================
# Cada programa SAS de módulo usa etiquetas ligeramente distintas para las
# mismas columnas (p.ej. "Nov." vs "Nov#", "Var." vs "Var#") — se listan aquí
# tal como aparecen en los archivos SAS de referencia.
_DIAG_LABELS: dict[str, dict[str, str | None]] = {
    "AGRICOLAS":  {"llave": "Art_Unmed_Casacomer_ICA", "nov": "Nov.", "var_faltan": "Var.",  "precio_ante_faltan": "Precio Ante.",  "caracte": None},
    "PECUARIOS":  {"llave": "Art_Unmed_Casacomer_ICA", "nov": "Nov.", "var_faltan": "Var.",  "precio_ante_faltan": "Precio Ante.",  "caracte": None},
    "ELEMENTOS":  {"llave": "Art_Casacomer_ICA_Unmed",  "nov": "Nov.", "var_faltan": "Var.",  "precio_ante_faltan": "Precio Ante.",  "caracte": None},
    "EMPAQUES":   {"llave": "Articulo_Caracte",         "nov": "Nov.", "var_faltan": "Var.",  "precio_ante_faltan": "Precio Ante.",  "caracte": "Caracte."},
    "ARRIENDOS":  {"llave": "Articulo_Caracte",         "nov": "Nov#", "var_faltan": "Var#",  "precio_ante_faltan": "Precio Ante#",  "caracte": "Caracte#"},
    "SERVICIOS":  {"llave": "Articulo_Caracte",         "nov": "Nov#", "var_faltan": "Var#",  "precio_ante_faltan": "Precio Ante#",  "caracte": "Caracte#"},
    "JORNALES":   {"llave": "Articulo_Caracte",         "nov": "Nov#", "var_faltan": "Var#",  "precio_ante_faltan": "Precio Ante#",  "caracte": "Caracte#"},
    "ESPECIES":   {"llave": "Articulo_Caracte",         "nov": "Nov#", "var_faltan": "Var#",  "precio_ante_faltan": "Precio Ante#",  "caracte": "Caracte#"},
    "PROPAGACION": {"llave": "Articulo_Unmed", "llave2": "Art_Casacomer_ICA_Unmed", "nov": "Nov#", "var_faltan": "Var#",
                    "precio_ante_faltan": "Precio Ante#", "caracte": None, "unmed": "UnMed#", "casacom": "CasaCom#"},
}

# Orden de columnas SAS por módulo (capturado de los archivos de referencia;
# el orden varía entre programas SAS y no sigue una fórmula única).
_DIAG_ORDEN_FALTAN_GRUPO: dict[str, list[str]] = {
    "AGRICOLAS": ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Codigo CPC", "Articulo", "CasaCom.", "RegICA", "Precio Ante.", "Precio Actual", "Var.", "UnMed.", "Nov.", "Estado", "Observación", "Art_Unmed_Casacomer_ICA", "Grupo"],
    "PECUARIOS": ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Codigo CPC", "Articulo", "CasaCom.", "RegICA", "Precio Ante.", "Precio Actual", "Var.", "UnMed.", "Nov.", "Estado", "Observación", "Art_Unmed_Casacomer_ICA", "Grupo"],
}

_DIAG_ORDEN_FALTAN_PUBLICA: dict[str, list[str]] = {
    "AGRICOLAS": ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Codigo CPC", "Articulo", "CasaCom.", "RegICA", "Precio Ante.", "Precio Actual", "Var.", "UnMed.", "Nov.", "Estado", "Observación", "Art_Unmed_Casacomer_ICA", "Grupo", "__PUBLICA__"],
    "PECUARIOS": ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Codigo CPC", "Articulo", "CasaCom.", "RegICA", "Precio Ante.", "Precio Actual", "Var.", "UnMed.", "Nov.", "Estado", "Observación", "Art_Unmed_Casacomer_ICA", "Grupo", "__PUBLICA__"],
    "ELEMENTOS": ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Codigo CPC", "Articulo", "CasaCom.", "RegICA", "Precio Ante.", "Precio Actual", "Var.", "UnMed.", "Nov.", "Estado", "Observación", "Art_Casacomer_ICA_Unmed", "__PUBLICA__"],
    "EMPAQUES":  ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Informante", "Codigo CPC", "Articulo", "Precio Ante.", "Precio Actual", "Var.", "Caracte.", "Nov.", "Estado", "Observación", "Articulo_Caracte", "__PUBLICA__"],
    "ARRIENDOS": ["Articulo_Caracte", "NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Informante", "Codigo CPC", "Articulo", "Precio Ante#", "Precio Actual", "Var#", "Caracte#", "Nov#", "Estado", "Observación", "__PUBLICA__"],
    "SERVICIOS": ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Informante", "Codigo CPC", "Articulo", "Precio Ante#", "Precio Actual", "Var#", "Caracte#", "Nov#", "Estado", "Observación", "Articulo_Caracte", "__PUBLICA__"],
    "JORNALES": ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Informante", "Codigo CPC", "Articulo", "Precio Ante#", "Precio Actual", "Var#", "Caracte#", "Nov#", "Estado", "Observación", "Articulo_Caracte", "__PUBLICA__"],
    "ESPECIES": ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Informante", "Codigo CPC", "Articulo", "Precio Ante#", "Precio Actual", "Var#", "Caracte#", "Nov#", "Estado", "Observación", "Articulo_Caracte", "__PUBLICA__"],
    "PROPAGACION": ["NombreDepartamento", "NombreMunicipio", "CodigoMpio", "CodigoDepto", "Fuente", "Codigo CPC", "Articulo", "CasaCom#", "RegICA", "Precio Ante#", "Precio Actual", "Var#", "UnMed#", "Nov#", "Estado", "Observación", "Articulo_Unmed", "Art_Casacomer_ICA_Unmed", "__PUBLICA__"],
}

_DIAG_ORDEN_DUPLI: dict[str, list[str]] = {
    "AGRICOLAS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "Grupo", "UnMed.", "CasaCom.", "RegICA", "__MES_ANT__", "__MES_ACT__", "VAR", "Estado", "Art_Unmed_Casacomer_ICA", "Observación", "Nov."],
    "PECUARIOS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "Grupo", "UnMed.", "CasaCom.", "RegICA", "__MES_ANT__", "__MES_ACT__", "VAR", "Nov.", "Estado", "Art_Unmed_Casacomer_ICA", "Observación"],
    "ELEMENTOS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "UnMed.", "CasaCom.", "RegICA", "Estado", "Art_Casacomer_ICA_Unmed", "Nov."],
    "EMPAQUES":  ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Estado", "Articulo_Caracte", "Informante", "Caracte.", "Nov."],
    "ARRIENDOS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Codigo CPC", "Caracte#"],
    "SERVICIOS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Codigo CPC", "Caracte#"],
    "JORNALES": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Caracte#"],
    "ESPECIES": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Caracte#"],
    "PROPAGACION": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "UnMed#", "CasaCom#", "RegICA", "Nov#", "Estado", "Articulo_Unmed", "Art_Casacomer_ICA_Unmed"],
}

# Nombres de archivo SAS para cada diagnóstico
_FALTAN_GRUPO_NOMBRES: dict[str, str] = {
    "AGRICOLAS": "FALTAN_GRUPO_AGRICOLA",
    "PECUARIOS": "FALTAN_GRUPO_PECUARIO",
}
_FALTAN_PUBLICA_NOMBRES: dict[str, str] = {
    "AGRICOLAS": "FALTAN_PUBLICA_AGRICOLA",
    "PECUARIOS": "FALTAN_PUBLICA_PECUARIO",
    "ELEMENTOS": "FALTAN_PUBLICA_ELEMENTOS",
    "EMPAQUES":  "FALTAN_PUBLICA_EMPAQUES",
    "ARRIENDOS": "FALTAN_PUBLICA_ARRIENDOS",
    "SERVICIOS": "FALTAN_PUBLICA_SERVICIOS",
    "PROPAGACION": "FALTAN_PUBLICA_MATERIAL",
    "JORNALES":    "FALTAN_PUBLICA_JORNALES",
    "ESPECIES":    "FALTAN_PUBLICA_ESPECIES",
}
_DUPLI_NOMBRES: dict[str, str] = {
    "AGRICOLAS": "INSUMOS_AGRICOLAS_DUPLI",
    "PECUARIOS": "INSUMOS_PECUARIOS_DUPLI",
    "ELEMENTOS": "ELEM_AGROPE_DUPLI",
    "EMPAQUES":  "EMPA_AGROPE_DUPLI",
    "ARRIENDOS": "ARRIENDOS_DUPLI",
    "SERVICIOS": "SERVICIOS_DUPLI",
    "PROPAGACION": "MATERIAL_DUPLI",
    "JORNALES":    "JORNALES_DUPLI",
    "ESPECIES":    "ESPE_PRODUC_DUPLI",
}
_VAR_ATIPICO_NOMBRES: dict[str, str] = {
    "AGRICOLAS": "VAR_ATIPICO_AGRICOLA",
    "PECUARIOS": "VAR_ATIPICO_PECUARIO",
    "ELEMENTOS": "VAR_ATIPICO_ELEMENTOS",
    "EMPAQUES":  "VAR_ATIPICO_EMPAQUES",
    "ARRIENDOS": "VAR_ATIPICO_ARRIENDOS",
    "SERVICIOS": "VAR_ATIPICO_SERVICIOS",
    "PROPAGACION": "VAR_ATIPICO_MATERIAL",
    "JORNALES":    "VAR_ATIPICO_JORNALES",
    "ESPECIES":    "VAR_ATIPICO_ESPECIES",
}
# Hoja de VAR_ATIPICO cuando SAS no usa el nombre del archivo: el programa de
# Material reutiliza la hoja "VAR_ATIPICO_AGRICOLA" (AGO2026).
_VAR_ATIPICO_SHEET: dict[str, str] = {"PROPAGACION": "VAR_ATIPICO_AGRICOLA"}
_DUPLI_SHEET: dict[str, str] = {
    "AGRICOLAS": "INSUMOS_AGRICOLAS_DUPLI",
    "PECUARIOS": "INSUMOS_PECUARIOS_DUPLI",
    "ELEMENTOS": "ELEMENTOS_AGROPECUARIOS_DUPLI",
    "EMPAQUES":  "EMPAQUES_AGROPECUARIOS_DUPLI",
    "ARRIENDOS": "ARRIENDOS_DUPLI",
    "SERVICIOS": "SERVICIOS_DUPLI",
    "PROPAGACION": "MATERIAL_PROPAGACION_DUPLI",
    "JORNALES":    "JORNALES_DUPLI",
    "ESPECIES":    "ESPECIE_PRODUCTIVA_DUPLI",
}


def _diag_col_map(modulo: str, tipo_modulo: str, mes_anterior: str, mes_actual: str, kind: str) -> dict[str, str]:
    """Mapa {etiqueta_salida_SAS: columna_interna} para un módulo/kind dado.

    kind: "faltan" (FALTAN_GRUPO/FALTAN_PUBLICA — Precio Ante./Actual/Var.
          con etiqueta fija) o "dupli" (DUPLI/VAR_ATIPICO — columnas con el
          nombre del mes, VAR siempre en mayúsculas).
    """
    lbl = _DIAG_LABELS[modulo]
    m: dict[str, str] = {
        "CodigoDepto":        "CodigoDepto",
        "CodigoMpio":         "CÓDIGO DIVIPOLA",
        "NombreDepartamento": "NombreDepartamento",
        "NombreMunicipio":    "NombreMunicipio",
        "Fuente":             "FUENTE",
        "Informante":         "INFORMANTE",
        "Codigo CPC":         "CÓDIGO CPC",
        "Articulo":           "ARTÍCULO",
        "Estado":             "ESTADO",
        "Observación":        "OBSERVACION",
        "Grupo":              "Grupo",
        "__PUBLICA__":        "Nombre_Publica",
        # Igual que BASE_INSUMOS: la llave case-preservada con todas sus
        # partes, no LLAVE_ARTICULO (llave interna de cruce en mayúsculas).
        lbl["llave"]:         "_LLAVE_DISPLAY",
    }
    if lbl.get("llave2"):
        m[lbl["llave2"]] = "_LLAVE_DISPLAY_2"
    if tipo_modulo == "caracte":
        m[lbl["caracte"]] = "CARACTERÍSTICA"
    else:
        m[lbl.get("unmed", "UnMed.")] = "UNIDAD DE MEDIDA"
        m[lbl.get("casacom", "CasaCom.")] = "CASA COMERCIAL"
        m["RegICA"] = "REGISTRO ICA"

    m[lbl["nov"]] = "NOV"
    if kind == "faltan":
        m[lbl["precio_ante_faltan"]] = "PRECIO_ANTERIOR_RAW"
        m["Precio Actual"] = "PRECIO"
        m[lbl["var_faltan"]] = "VAR"
    else:
        m["__MES_ANT__"] = "PRECIO_ANTERIOR_RAW"
        m["__MES_ACT__"] = "PRECIO"
        m["VAR"] = "VAR"
        m["REVISA"] = "REVISA"
    return m


# Programas SAS que nombran las columnas de mes en mayúscula (JORNALES JUN2026: MARZO/JUNIO).
_MESES_MAYUSCULA = {"JORNALES"}


def _nombres_mes(modulo: str, mes_anterior: str, mes_actual: str) -> tuple[str, str]:
    if modulo.upper() in _MESES_MAYUSCULA:
        return mes_anterior.upper(), mes_actual.upper()
    return mes_anterior, mes_actual


def _sustituir_orden(orden: list[str], modulo: str, mes_anterior: str, mes_actual: str) -> list[str]:
    col_pub = _NOMBRE_PUBLICA_SAS.get(modulo, "Nombre_Publica")
    return [
        {"__PUBLICA__": col_pub, "__MES_ANT__": mes_anterior, "__MES_ACT__": mes_actual}.get(c, c)
        for c in orden
    ]


# Columna llave visible por la que SAS ordena BASE_INSUMOS y diagnósticos.
_LLAVES_VISIBLES = ("Art_Unmed_Casacomer_ICA", "Articulo_Unmed", "Art_Casacomer_ICA_Unmed", "Articulo_Caracte")

# Módulos cuyo programa SAS guarda Observación con longitud 255 (se trunca y
# se quitan espacios finales). Verificado AGO2026: Arriendos, Servicios,
# Material. JORNALES/ESPECIES por analogía (sin referencia SAS todavía).
_OBSERVACION_255 = {"ARRIENDOS", "SERVICIOS", "PROPAGACION", "JORNALES", "ESPECIES"}


def _truncar_observacion(df: pd.DataFrame, m: str) -> None:
    if m in _OBSERVACION_255 and "OBSERVACION" in df.columns:
        obs = df["OBSERVACION"]
        df["OBSERVACION"] = obs.where(obs.isna(), obs.astype(str).str.slice(0, 255).str.rstrip())
    if m in _VACIOS_COMO_NA:
        for col in ("CASA COMERCIAL", "REGISTRO ICA"):
            if col in df.columns:
                vacio = df[col].isna() | df[col].astype(str).str.strip().isin(["", "nan", "None"])
                df.loc[vacio, col] = "NA"


# Módulos donde SAS muestra "NA" en CasaCom/RegICA vacíos (Material AGO2026,
# Elementos JUL2026).
_VACIOS_COMO_NA = {"PROPAGACION", "ELEMENTOS"}


def _exportar_diag_hoja(
    df: pd.DataFrame,
    orden: list[str],
    col_map: dict[str, str],
    modulo: str,
    mes_anterior: str,
    mes_actual: str,
    ruta: Path,
    sheet_name: str,
) -> int:
    df_out = _armar_diag_hoja(df, orden, col_map, modulo, mes_anterior, mes_actual)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    escribir_excel_multisheet(ruta, {sheet_name[:31]: df_out})
    return len(df_out)


def _armar_diag_hoja(
    df: pd.DataFrame,
    orden: list[str],
    col_map: dict[str, str],
    modulo: str,
    mes_anterior: str,
    mes_actual: str,
) -> pd.DataFrame:
    """Selecciona, renombra (etiquetas SAS) y ordena las filas de BASE_INSUMOS
    y de los archivos de diagnóstico."""
    mes_anterior, mes_actual = _nombres_mes(modulo, mes_anterior, mes_actual)
    orden_sust = _sustituir_orden(orden, modulo, mes_anterior, mes_actual)
    col_map_sust = {
        {"__PUBLICA__": _NOMBRE_PUBLICA_SAS.get(modulo, "Nombre_Publica"),
         "__MES_ANT__": mes_anterior, "__MES_ACT__": mes_actual}.get(k, k): v
        for k, v in col_map.items()
    }
    rename_inv = {v: k for k, v in col_map_sust.items()}
    cols_internos = [col_map_sust[c] for c in orden_sust if c in col_map_sust and col_map_sust[c] in df.columns]
    df_out = df[cols_internos].rename(columns=rename_inv)
    # Reordenar según orden_sust exacto (por si el rename produjo duplicados de nombre)
    cols_finales = [c for c in orden_sust if c in df_out.columns]
    df_out = df_out[cols_finales]
    # SAS ordena por la llave visible del artículo y luego por municipio
    # (estable: los empates conservan el orden de la base liviana).
    llave = next((c for c in _LLAVES_VISIBLES if c in df_out.columns), None)
    if llave:
        df_out = _ordenar_sas(df_out, [llave, "CodigoMpio"])
    return df_out


def exportar_diagnosticos(
    base_completa: pd.DataFrame,
    divipola_raw: pd.DataFrame,
    mappings_grupos: dict,
    mappings_articulos: dict,
    modulo: str,
    periodo: str,
    mes_actual: str,
    mes_anterior: str,
    tipo_modulo: str,
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta FALTAN_GRUPO / FALTAN_PUBLICA / DUPLI / VAR_ATIPICO con la
    misma lógica fila-por-fila y nombres de columna SAS que BASE_INSUMOS,
    en vez de la comparación agregada por municipio que usa el flujo
    principal (quality.detectar_duplicados/detectar_var_atipica).

    SAS calcula estos diagnósticos sobre la base SIN filtrar por Estado/
    Precio, con el precio anterior embebido por fila (columna "Precio
    Ante."), igual que BASE_INSUMOS — no sobre la base filtrada del flujo
    principal ni contra el agregado del período anterior.
    """
    m = modulo.upper()
    df = _construir_base_diagnostico(base_completa, divipola_raw, mappings_grupos, mappings_articulos)
    _agregar_llave_display(df, m)
    _truncar_observacion(df, m)

    carpeta = Path(ruta_reporting) / modulo.lower()
    metas: list[dict] = []

    # DUPLI/VAR_ATIPICO usan como encabezado el nombre del mes del período
    # anterior SEGÚN LA PERIODICIDAD DEL MÓDULO (p.ej. "Marzo" para Elementos/
    # Empaques bimestrales, "Febrero" para Arriendos/Servicios trimestrales),
    # no el mes_anterior mensual genérico — igual que exportar_bases/exportar_anexos.
    periodo_ant_mod = _periodo_anterior_modulo(periodo, m)
    mes_anterior_dupli = _ABBR_A_MES_LARGO.get(periodo_ant_mod[:3], mes_anterior)

    # --- FALTAN_GRUPO (solo módulos multi-grupo) ---
    if m in _FALTAN_GRUPO_NOMBRES:
        df_fg = df[df["Grupo"].isna()]
        col_map = _diag_col_map(m, tipo_modulo, mes_anterior, mes_actual, kind="faltan")
        ruta = carpeta / f"{_FALTAN_GRUPO_NOMBRES[m]}_{periodo}.xlsx"
        n = _exportar_diag_hoja(
            df_fg, _DIAG_ORDEN_FALTAN_GRUPO[m], col_map, m, mes_anterior, mes_actual,
            ruta, _FALTAN_GRUPO_NOMBRES[m],
        )
        metas.append({"tipo": "FALTAN_GRUPO", "archivo": str(ruta), "filas": n})

    # --- FALTAN_PUBLICA ---
    df_fp = df[df["Nombre_Publica"].isna()]
    col_map_fp = _diag_col_map(m, tipo_modulo, mes_anterior, mes_actual, kind="faltan")
    ruta = carpeta / f"{_FALTAN_PUBLICA_NOMBRES[m]}_{periodo}.xlsx"
    n = _exportar_diag_hoja(
        df_fp, _DIAG_ORDEN_FALTAN_PUBLICA[m], col_map_fp, m, mes_anterior, mes_actual,
        ruta, _FALTAN_PUBLICA_NOMBRES[m],
    )
    metas.append({"tipo": "FALTAN_PUBLICA", "archivo": str(ruta), "filas": n})

    # --- DUPLI ---
    llave_dup = [c for c in _LLAVE_DUPLICADOS if c in df.columns]
    es_dupli = df.duplicated(subset=llave_dup, keep="first")
    df_dupli = df[es_dupli]
    col_map_dupli = _diag_col_map(m, tipo_modulo, mes_anterior_dupli, mes_actual, kind="dupli")
    ruta = carpeta / f"{_DUPLI_NOMBRES[m]}_{periodo}.xlsx"
    n = _exportar_diag_hoja(
        df_dupli, _DIAG_ORDEN_DUPLI[m], col_map_dupli, m, mes_anterior_dupli, mes_actual,
        ruta, _DUPLI_SHEET[m],
    )
    metas.append({"tipo": "DUPLI", "archivo": str(ruta), "filas": n})

    # --- VAR_ATIPICO (REVISA: 1 = VAR>=25 o <=-25, 2 = sin VAR, 3 = VAR>=100) ---
    df_var = df.copy()
    df_var["REVISA"] = 0
    df_var.loc[df_var["VAR"].isna(), "REVISA"] = 2
    df_var.loc[
        df_var["VAR"].notna() & ((df_var["VAR"] >= 25.0) | (df_var["VAR"] <= -25.0)), "REVISA"
    ] = 1
    df_var.loc[df_var["VAR"].notna() & (df_var["VAR"] >= 100.0), "REVISA"] = 3
    df_var = df_var[df_var["REVISA"] > 0]
    orden_var = _DIAG_ORDEN_DUPLI[m] + ["REVISA"]
    ruta = carpeta / f"{_VAR_ATIPICO_NOMBRES[m]}_{periodo}.xlsx"
    n = _exportar_diag_hoja(
        df_var, orden_var, col_map_dupli, m, mes_anterior_dupli, mes_actual,
        ruta, _VAR_ATIPICO_SHEET.get(m, _VAR_ATIPICO_NOMBRES[m]),
    )
    metas.append({"tipo": "VAR_ATIPICO", "archivo": str(ruta), "filas": n})

    for meta in metas:
        log.info("exportar_diagnosticos [%s] OK | tipo=%s | archivo=%s | filas=%d",
                  m, meta["tipo"], meta["archivo"], meta["filas"])

    return pd.DataFrame([{"modulo": modulo, "periodo": periodo, **meta} for meta in metas])


def exportar_cuadros(
    base_comparada: pd.DataFrame,
    grupos: list[str],
    modulo: str,
    periodo: str,
    mes_actual: str,
    mes_anterior: str,
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta el cuadro final de publicación, una hoja por grupo.

    Columnas del cuadro final (equivale al output SAS CUADROS_*):
      CÓDIGO CPC | Producto | SUBIO | BAJO | ESTABLE | N.D. | PRECIO_MIN | PRECIO_MAX

    SAS:
      TABULATE Producto × (Positiva, Negativa, Estable, n.d.) → conteo mercados
      MIN(Precio) → PRECIO_MIN
      MAX(Precio) → PRECIO_MAX

    Args:
        base_comparada: DataFrame con Nombre_Publica, TENDENCIA, PRECIO_{mes_actual}.
        grupos: Lista de nombres de grupo.
        modulo: Nombre del módulo.
        periodo: ID del período.
        mes_actual: Nombre del mes actual (para la columna de precio).
        mes_anterior: Nombre del mes anterior (para encabezado).
        ruta_reporting: Directorio raíz de reportes.

    Returns:
        DataFrame de metadatos.
    """
    nombre = _nombre_sas("CUADROS", modulo, periodo)
    ruta = Path(ruta_reporting) / modulo.lower() / nombre
    ruta.parent.mkdir(parents=True, exist_ok=True)

    col_precio_actual = f"PRECIO_{mes_actual}"
    if col_precio_actual not in base_comparada.columns:
        col_precio_actual = "PRECIO_PROMEDIO"

    hojas: dict[str, pd.DataFrame] = {}
    for grupo in grupos:
        df_grupo = base_comparada[base_comparada["Grupo"] == grupo].copy()
        if df_grupo.empty:
            hojas[grupo] = pd.DataFrame()
            continue

        # Conteo de mercados por tendencia
        conteo = (
            df_grupo.groupby(["CÓDIGO CPC", "Nombre_Publica", "TENDENCIA"], observed=True)
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        for tend in ["Positiva", "Negativa", "Estable", "n.d."]:
            if tend not in conteo.columns:
                conteo[tend] = 0

        # Precios mínimo y máximo del período actual
        precios = (
            df_grupo.groupby(["CÓDIGO CPC", "Nombre_Publica"])[col_precio_actual]
            .agg(PRECIO_MIN="min", PRECIO_MAX="max")
            .reset_index()
        )

        cuadro = conteo.merge(precios, on=["CÓDIGO CPC", "Nombre_Publica"], how="left")
        cuadro = cuadro.rename(columns={"Nombre_Publica": "Producto"})
        # SAS combina Estable + n.d. en una sola columna "Estable-n.d." y
        # antepone "N" = total de mercados (Positiva+Negativa+Estable+n.d.)
        cuadro["N"] = cuadro["Positiva"] + cuadro["Negativa"] + cuadro["Estable"] + cuadro["n.d."]
        cuadro["ESTABLE_ND"] = cuadro["Estable"] + cuadro["n.d."]
        cuadro = cuadro.rename(columns={"Positiva": "SUBIO", "Negativa": "BAJO"})
        if modulo.upper() in _SIN_CODIGO_CPC:
            # Sin columna CPC, SAS ordena solo por el nombre del producto.
            cuadro = _ordenar_sas(cuadro, ["Producto"])
        else:
            cuadro = _ordenar_sas(cuadro, ["CÓDIGO CPC", "Producto"], modulo)
        hojas[grupo] = cuadro

    wb = Workbook()
    wb.remove(wb.active)
    m = modulo.upper()
    for nombre_hoja, df_hoja in hojas.items():
        ws = wb.create_sheet(title=_hoja("CUADROS", m, nombre_hoja)[:31])
        if df_hoja.empty:
            continue
        escribir_hoja_pivot_cpc(
            ws,
            df_hoja,
            cpc_col=None if m in _SIN_CODIGO_CPC else "CÓDIGO CPC",
            prod_col="Producto",
            metric_cols=["N", "SUBIO", "BAJO", "ESTABLE_ND", "PRECIO_MIN", "PRECIO_MAX"],
            sub_headers=["N", "Subió", "Bajó", "Estable-n.d.", "Min", "Max"],
            grupos_header=[("TOTAL", 1), ("MERCADOS", 3), (f"PRECIO DE {periodo}", 2)],
            cols_punto_si_cero={"SUBIO", "BAJO", "ESTABLE_ND"},
            cols_numero={"N"},
            total_label="TOTAL",
        )
    wb.save(str(ruta))

    filas_totales = sum(len(df) for df in hojas.values())
    log.info("exportar_cuadros [%s] OK | archivo=%s | hojas=%d | filas=%d",
             modulo, nombre, len(hojas), filas_totales)
    return pd.DataFrame([{
        "archivo": str(ruta),
        "modulo": modulo,
        "periodo": periodo,
        "tipo": "CUADROS",
        "hojas": len(hojas),
        "filas_totales": filas_totales,
    }])


# =============================================================================
# Nomenclatura SAS para archivos de revisión (paso 17)
# =============================================================================

_MAYOR2_NOMBRES: dict[str, str] = {
    "AGRICOLAS":   "INSU_AGRIC_MAYORESQUE2",
    "PECUARIOS":   "INSU_PECUA_MAYORESQUE2",
    "ELEMENTOS":   "ELEM_AGROPE_MAYORESQUE2",
    "EMPAQUES":    "EMPA_AGROPE_MAYORESQUE2",
    "ARRIENDOS":   "ARRIENDOS_MAYORESQUE2",
    "SERVICIOS":   "SERVICIOS_MAYORESQUE2",
    "PROPAGACION": "MATE_PROPAGA_MAYORESQUE2",
    "JORNALES":    "JORNALES_MAYORESQUE2",
    "ESPECIES":    "ESPE_PRODUC_MAYORESQUE2",
}

_MENOR2_NOMBRES: dict[str, str] = {
    "AGRICOLAS":   "INSU_AGRIC_MENORESQUE2",
    "PECUARIOS":   "INSU_PECUA_MENORESQUE2",
    "ELEMENTOS":   "ELEM_AGROPE_MENORESQUE2",
    "EMPAQUES":    "EMPA_AGROPE_MENORESQUE2",
    "ARRIENDOS":   "ARRIENDOS_MENORESQUE2",
    "SERVICIOS":   "SERVICIOS_MENORESQUE2",
    "PROPAGACION": "MATE_PROPAGA_MENORESQUE2",
    "JORNALES":    "JORNALES_MENORESQUE2",
    "ESPECIES":    "ESPE_PRODUC_MENORESQUE2",
}

_MAYMEN_NOMBRES: dict[str, str] = {
    "AGRICOLAS":   "MAY_MEN3_AGRIC",
    "PECUARIOS":   "MAY_MEN2_PECUA",
    "ELEMENTOS":   "ELEM_MAY_MEN2",
    "EMPAQUES":    "EMPA_MAY_MEN2",
    "ARRIENDOS":   "ARRIENDOS_MAYMEN2",
    "SERVICIOS":   "SERVICIOS_MAYMEN2",
    "PROPAGACION": "MAY_MEN2_MATERIAL",
    "JORNALES":    "JORNALES_MAYMEN2",
    "ESPECIES":    "ESPECIES_MAYMEN2",
}

_TABREV_NOMBRES: dict[str, str] = {
    "AGRICOLAS":   "TABREV_AGRI",
    "PECUARIOS":   "INSPEC_TABREV",
    "ELEMENTOS":   "ELEM_TABLASREVISION",
    "EMPAQUES":    "EMPA_TABLASREVISION",
    "ARRIENDOS":   "ARRIENDOS_TABLASREVISION",
    "SERVICIOS":   "SERVICIOS_TABLASREVISION",
    "PROPAGACION": "MATE_TABLASREVISION",
    "JORNALES":    "JORNALES_TABLASREVISION",
    "ESPECIES":    "ESPE_TABLASREVISION",
}

_REVIS4M_NOMBRES: dict[str, str] = {
    "AGRICOLAS":   "REVIS_INSU_AGRI_4MESES",
    "PECUARIOS":   "REVIS_INSU_PECUA_4MESES",
    "ELEMENTOS":   "REVIS_ELEM_AGROPE_4MESES",
    "EMPAQUES":    "REVIS_EMPA_AGROPE_4MESES",
    "ARRIENDOS":   "REVIS_ARRIENDOS_4MESES",
    "SERVICIOS":   "REVIS_SERVICIOS_4MESES",
    "PROPAGACION": "REVIS_MATE_PROPAGA_4MESES",
    "JORNALES":    "REVIS_JORNALES_4MESES",
    "ESPECIES":    "REVIS_ESPE_4MESES",
}

_REVISION_ROOT: dict[str, str] = {
    "AGRICOLAS":   "Revisión insumos agrícolas",
    "PECUARIOS":   "Revisión insumos pecuarios",
    "ELEMENTOS":   "Revisión elementos",
    "EMPAQUES":    "Revisión empaques",
    "ARRIENDOS":   "Revisión arriendos",
    "SERVICIOS":   "Revisión servicios",
    "PROPAGACION": "Revisión material de propagación",
    "JORNALES":    "Revisión jornales",
    "ESPECIES":    "Revisión especies",
}

_MAYOR2_SHEET: dict[str, str] = {
    "AGRICOLAS":   "INSU_AGRICOLAS_MAYORESOIGUALES2",
    "PECUARIOS":   "INSU_PECUARIOS_MAYORESOIGUALES2",
    "ELEMENTOS":   "ELEM_AGROPE_MAYORESOIGUALES2",
    "EMPAQUES":    "EMPA_AGROPE_MAYORESOIGUALES2",
    "ARRIENDOS":   "ARRIENDOS_MAYORESOIGUALES2",
    "SERVICIOS":   "SERVICIOS_MAYORESOIGUALES2",
    "PROPAGACION": "MATE_PROPAGA_MAYORESOIGUALES2",
    "JORNALES":    "JORNALES_MAYORESOIGUALES2",
    "ESPECIES":    "ESPE_PRODUC_MAYORESOIGUALES2",
}

_MENOR2_SHEET: dict[str, str] = {
    "AGRICOLAS":   "INSU_AGRICOLAS_MENORESQUE2",
    "PECUARIOS":   "INSU_PECUARIOS_MENORESQUE2",
    "ELEMENTOS":   "ELEM_AGROPE_MENORESQUE2",
    "EMPAQUES":    "EMPA_AGROPE_MENORESQUE2",
    "ARRIENDOS":   "ARRIENDOS_MENORESQUE2",
    "SERVICIOS":   "SERVICIOS_MENORESQUE2",
    "PROPAGACION": "MATE_PROPAGA_MENORESQUE2",
    "JORNALES":    "JORNALES_MENORESQUE2",
    "ESPECIES":    "ESPE_PRODUC_MENORESQUE2",
}

# Módulos con REVISIÓN TEMÁTICA (etiqueta para el nombre del archivo)
# Meses en español para construir el período anterior
_MES_A_NUM = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12,
}
_NUM_A_MES = {v: k for k, v in _MES_A_NUM.items()}

# Abreviatura (ENE/FEB/...) → nombre completo en español, para encabezados
# de DUPLI/VAR_ATIPICO ("Marzo", "Febrero") según la periodicidad del módulo.
_ABBR_A_MES_LARGO: dict[str, str] = {
    "ENE": "Enero", "FEB": "Febrero", "MAR": "Marzo", "ABR": "Abril",
    "MAY": "Mayo", "JUN": "Junio", "JUL": "Julio", "AGO": "Agosto",
    "SEP": "Septiembre", "OCT": "Octubre", "NOV": "Noviembre", "DIC": "Diciembre",
}

# Salto de períodos por módulo (cuántos meses entre período actual y anterior)
_SALTO_MESES: dict[str, int] = {
    "AGRICOLAS": 1, "PECUARIOS": 1,
    "ELEMENTOS": 2, "EMPAQUES": 2, "PROPAGACION": 2,
    "ARRIENDOS": 3, "SERVICIOS": 3,
    "JORNALES": 3, "ESPECIES": 3,
}

# Nombre SAS de la columna Nombre_Publica según módulo (máx 32 bytes para Excel/SAS)
_NOMBRE_PUBLICA_SAS: dict[str, str] = {
    "AGRICOLAS":   "Nombre_productos_agrícolas_publ",
    "PECUARIOS":   "Nombre_productos_pecuarios_publi",
    "ELEMENTOS":   "Nombre_productos_elementos_publi",
    "EMPAQUES":    "Nombre_productos_empaques_publi",
    "ARRIENDOS":   "Nombre_productos_arriendos_publi",
    "SERVICIOS":   "Nombre_productos_servicios_publi",
    "PROPAGACION": "Nombre_productos_material_publi",
    "JORNALES":    "Nombre_productos_jornales_publi",   # JORNALES JUN2026
    "ESPECIES":    "Nombre_productos_especie_publi",   # ESPECIES JUN2026
}

# Hoja de MAY_MEN cuando SAS no usa "UNION_MAYORES_MENORESQUE2" (MATERIAL AGO2026).
_MAYMEN_SHEET: dict[str, str] = {m: "MAYORES_MENORESQUE2" for m in
                                 ("PROPAGACION", "SERVICIOS", "JORNALES", "ESPECIES")}

# Módulos con más de un Grupo — incluyen la columna "Grupo" en MAYORESQUE2/
# MENORESQUE2/MAYMEN (para módulos de un solo grupo, SAS omite la columna
# por ser un valor constante sin información adicional).
_MULTI_GRUPO_MODULOS = {"AGRICOLAS", "PECUARIOS"}

# Módulos tipo "caracte" que en SAS usan N_INFORMANTE en vez de N_FUENTE en
# los reportes de revisión (MAYORESQUE2/MENORESQUE2/MAYMEN).
_CARACTE_N_INFORMANTE = {"EMPAQUES", "ARRIENDOS", "SERVICIOS", "JORNALES", "ESPECIES"}

# Módulos que SAS no incluye con columna "Codigo CPC" en BASES/ANEXO/
# MAYORESQUE2/MENORESQUE2/MAYMEN.
_SIN_CODIGO_CPC = {"ARRIENDOS", "SERVICIOS"}


def _periodo_anterior_modulo(periodo: str, modulo: str) -> str:
    """Calcula el período anterior según la periodicidad del módulo."""
    mes_str = periodo[:3].upper()
    anio = int(periodo[3:])
    mes_num = _MES_A_NUM.get(mes_str, 1)
    salto = _SALTO_MESES.get(modulo.upper(), 1)
    mes_ant = mes_num - salto
    anio_ant = anio
    if mes_ant < 1:
        mes_ant += 12
        anio_ant -= 1
    return f"{_NUM_A_MES[mes_ant]}{anio_ant}"


def _col_n_fuente_interna(df: pd.DataFrame, modulo: str) -> str | None:
    """Columna interna para N_FUENTE (estándar) o N_INFORMANTE (caracte).

    Salidas anteriores a la corrección no traen N_INFORMANTE — se usa
    N_FUENTE como respaldo.
    """
    if modulo in _CARACTE_N_INFORMANTE and "N_INFORMANTE" in df.columns:
        return "N_INFORMANTE"
    return "N_FUENTE" if "N_FUENTE" in df.columns else None


def _col_mayor2(mayor2: pd.DataFrame, modulo: str = "", periodo: str = "") -> dict:
    """Mapea columnas de mayor2/menor2 a columnas SAS de MAYORESQUE2 con sufijo de período."""
    col_divipola = next((c for c in ["CÓDIGO DIVIPOLA", "CodigoMpio"] if c in mayor2.columns), None)
    m = modulo.upper()
    col_pub_sas = _NOMBRE_PUBLICA_SAS.get(m, "Nombre_Publica")
    sfx = f"_{periodo}" if periodo else ""
    col_n_fuente = f"N_INFORMANTE{sfx}" if m in _CARACTE_N_INFORMANTE else f"N_FUENTE{sfx}"
    return {
        "CodigoDepto":          "CodigoDepto" if "CodigoDepto" in mayor2.columns else None,
        "NombreDepartamento":   "NombreDepartamento" if "NombreDepartamento" in mayor2.columns else None,
        "CodigoMpio":           col_divipola,
        "NombreMunicipio":      "NombreMunicipio" if "NombreMunicipio" in mayor2.columns else None,
        "Codigo CPC":           "CÓDIGO CPC" if "CÓDIGO CPC" in mayor2.columns and m not in _SIN_CODIGO_CPC else None,
        col_pub_sas:            "Nombre_Publica" if "Nombre_Publica" in mayor2.columns else None,
        "Grupo":                "Grupo" if "Grupo" in mayor2.columns and m in _MULTI_GRUPO_MODULOS else None,
        col_n_fuente:           _col_n_fuente_interna(mayor2, m),
        f"N_ARTICULOS{sfx}":    "N_ARTICULOS" if "N_ARTICULOS" in mayor2.columns else None,
        f"PRECIO_PROMEDIO{sfx}":"PRECIO_PROMEDIO" if "PRECIO_PROMEDIO" in mayor2.columns else None,
    }


def _ordenar_mayor_menor(df_out: pd.DataFrame, origen: pd.DataFrame, m: str) -> pd.DataFrame:
    """Orden SAS de MAYORESQUE2/MENORESQUE2: municipio, grupo, CPC, nombre.
    Las filas publicadas por excepción (PUBLICA_SIEMPRE, p.ej. arroz Fedearroz
    en Material) van primero, en bloque, igual que el SET de SAS."""
    df_out = df_out.copy()
    cols = ["CodigoMpio", "Grupo", "Codigo CPC", _NOMBRE_PUBLICA_SAS.get(m, "Nombre_Publica")]
    if "PUBLICA_SIEMPRE" in origen.columns:
        df_out["__bloque"] = (~origen["PUBLICA_SIEMPRE"].fillna(False).astype(bool)).astype(int).values
        return _ordenar_sas(df_out, ["__bloque", *cols], m).drop(columns="__bloque")
    return _ordenar_sas(df_out, cols, m)


def _preparar_mayor_menor(df: pd.DataFrame, modulo: str = "", periodo: str = "") -> pd.DataFrame:
    """Prepara un DataFrame mayor2/menor2 para exportación SAS."""
    col_map = _col_mayor2(df, modulo=modulo, periodo=periodo)
    rename = {v: k for k, v in col_map.items() if v and v in df.columns}
    cols_sel = [v for v in col_map.values() if v and v in df.columns]
    result = df[cols_sel].rename(columns=rename).copy()
    if "CodigoDepto" not in result.columns and "CodigoMpio" in result.columns:
        result.insert(0, "CodigoDepto", result["CodigoMpio"].astype(str).str[:2])
    return result


def exportar_mayo_menores(
    mayor2: pd.DataFrame,
    menor2: pd.DataFrame,
    mayor2_anterior: pd.DataFrame | None,
    modulo: str,
    periodo: str,
    mes_actual: str,
    mes_anterior: str,
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta MAYORESQUE2, MENORESQUE2 y MAY_MEN3 (unión con período anterior).

    SAS equivalente:
      INSUAGRIMAYORQUE2 = POR_MPIO WHERE N_ARTICULOS >= 2
      INSUAGRIMENORESQUE2 = POR_MPIO WHERE N_ARTICULOS < 2
      MAYORESYMENORESQUE2PRECI = (MAYOR U MENOR) LEFT JOIN MAYOR_ANTE ON (mpio, publica)
        + Mercado, Variacion(%), Tendencia

    Args:
        mayor2: Precio promedio por municipio-producto, N_ARTICULOS >= umbral.
        menor2: Precio promedio por municipio-producto, N_ARTICULOS < umbral.
        mayor2_anterior: mayor2 del período anterior (puede ser None).
        modulo: Nombre del módulo en mayúsculas (ej: "AGRICOLAS").
        periodo: ID del período actual (ej: "MAY2026").
        mes_actual: Nombre mes actual (ej: "Mayo").
        mes_anterior: Nombre mes anterior (ej: "Abril").
        ruta_reporting: Directorio raíz de reportes.

    Returns:
        DataFrame de metadatos.
    """
    m = modulo.upper()
    carpeta = Path(ruta_reporting) / modulo.lower()
    carpeta.mkdir(parents=True, exist_ok=True)

    periodo_ant = _periodo_anterior_modulo(periodo, modulo)
    col_pub_sas = _NOMBRE_PUBLICA_SAS.get(m, "Nombre_Publica")

    # --- MAYORESQUE2 ---
    df_mayor = _ordenar_mayor_menor(_preparar_mayor_menor(mayor2, modulo=modulo, periodo=periodo), mayor2, m)
    ruta_mayor = carpeta / f"{_MAYOR2_NOMBRES.get(m, f'{m}_MAYORESQUE2')}_{periodo}.xlsx"
    sheet_mayor = _MAYOR2_SHEET.get(m, "MAYORESOIGUALES2")
    with pd.ExcelWriter(str(ruta_mayor), engine="openpyxl") as w:
        df_mayor.to_excel(w, sheet_name=sheet_mayor[:31], index=False)

    # --- MENORESQUE2 ---
    df_menor = _ordenar_mayor_menor(_preparar_mayor_menor(menor2, modulo=modulo, periodo=periodo), menor2, m)
    if m not in _MULTI_GRUPO_MODULOS:
        # Módulos de un solo grupo: SAS escribe CodigoMpio antes de NombreDepartamento.
        primeras = ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio"]
        df_menor = df_menor[[c for c in primeras if c in df_menor.columns]
                            + [c for c in df_menor.columns if c not in primeras]]
    ruta_menor = carpeta / f"{_MENOR2_NOMBRES.get(m, f'{m}_MENORESQUE2')}_{periodo}.xlsx"
    sheet_menor = _MENOR2_SHEET.get(m, "MENORESQUE2")
    with pd.ExcelWriter(str(ruta_menor), engine="openpyxl") as w:
        df_menor.to_excel(w, sheet_name=sheet_menor[:31], index=False)

    # --- MAY_MEN3: unión mayor+menor con precio anterior + N por período ---
    union = pd.concat([mayor2, menor2], ignore_index=True)

    col_div = next((c for c in ["CÓDIGO DIVIPOLA", "CodigoMpio"] if c in union.columns), None)
    llave_join = [c for c in [col_div, "Nombre_Publica"] if c]

    col_n_fuente = "N_INFORMANTE" if m in _CARACTE_N_INFORMANTE else "N_FUENTE"

    # Renombrar N del período actual antes del join para separar ambos períodos
    n_fuente_act = _col_n_fuente_interna(union, m)
    union = union.drop(columns=[c for c in ("N_FUENTE", "N_INFORMANTE")
                                if c in union.columns and c != n_fuente_act])
    union = union.rename(columns={
        n_fuente_act:  f"{col_n_fuente}_{periodo}",
        "N_ARTICULOS": f"N_ARTICULOS_{periodo}",
    })

    if mayor2_anterior is not None and len(mayor2_anterior) > 0 and llave_join:
        col_div_ant = next(
            (c for c in ["CÓDIGO DIVIPOLA", "CodigoMpio"] if c in mayor2_anterior.columns), None
        )
        llave_ant = [c for c in [col_div_ant, "Nombre_Publica"] if c]
        if set(llave_ant) == set(llave_join):
            cols_ant = llave_ant + ["PRECIO_PROMEDIO"]
            n_fuente_ant = _col_n_fuente_interna(mayor2_anterior, m)
            if n_fuente_ant:
                cols_ant.append(n_fuente_ant)
            if "N_ARTICULOS" in mayor2_anterior.columns:
                cols_ant.append("N_ARTICULOS")
            if "CÓDIGO CPC" in mayor2_anterior.columns and "CÓDIGO CPC" in union.columns:
                cols_ant.append("CÓDIGO CPC")
            ant = mayor2_anterior[cols_ant].rename(columns={
                "PRECIO_PROMEDIO": f"PRECIO_PROMEDIO_{periodo_ant}",
                n_fuente_ant:      f"{col_n_fuente}_{periodo_ant}",
                "N_ARTICULOS":     f"N_ARTICULOS_{periodo_ant}",
                "CÓDIGO CPC":      "_CPC_ANT",
            })
            union = union.merge(ant, left_on=llave_join, right_on=llave_ant, how="left")
            # El MERGE de SAS deja el CPC del período anterior cuando el producto existía.
            if "_CPC_ANT" in union.columns:
                union["CÓDIGO CPC"] = union["_CPC_ANT"].where(union["_CPC_ANT"].notna(), union["CÓDIGO CPC"])
                union = union.drop(columns="_CPC_ANT")

    col_precio_actual   = "PRECIO_PROMEDIO"
    col_precio_anterior = f"PRECIO_PROMEDIO_{periodo_ant}"

    if col_precio_anterior in union.columns:
        mask_valido = union[col_precio_anterior].notna() & (union[col_precio_anterior] != 0)
        union["Variacion(%)"] = np.where(
            mask_valido,
            (union[col_precio_actual] / union[col_precio_anterior] - 1) * 100,
            np.nan,
        )
    else:
        union["Variacion(%)"] = float("nan")

    def _tendencia(v):
        if pd.isna(v):
            return "n.d."
        return "Positiva" if v > 0 else "Negativa" if v < 0 else "Estable"
    union["Tendencia"] = union["Variacion(%)"].apply(_tendencia)

    if "NombreMunicipio" in union.columns and "NombreDepartamento" in union.columns:
        union["Mercado"] = union["NombreMunicipio"] + " (" + union["NombreDepartamento"] + ")"

    if "CodigoDepto" not in union.columns and col_div and col_div in union.columns:
        union.insert(0, "CodigoDepto", union[col_div].astype(str).str[:2])

    rename_union = {col_div: "CodigoMpio", "CÓDIGO CPC": "Codigo CPC"}
    union = union.rename(columns={k: v for k, v in rename_union.items() if k in union.columns})
    union = union.rename(columns={"Nombre_Publica": col_pub_sas})

    # Orden SAS (igual en todos los módulos de referencia JUL/AGO2026):
    # período actual antes que el anterior, y Mercado justo antes de Variacion.
    cols_maymen = [
        "CodigoDepto", "NombreDepartamento", "CodigoMpio", "NombreMunicipio",
        *(["Codigo CPC"] if m not in _SIN_CODIGO_CPC else []),
        col_pub_sas,
        *(["Grupo"] if m in _MULTI_GRUPO_MODULOS else []),
        f"{col_n_fuente}_{periodo}", f"N_ARTICULOS_{periodo}",
        col_precio_actual,
        f"{col_n_fuente}_{periodo_ant}", f"N_ARTICULOS_{periodo_ant}",
        f"PRECIO_PROMEDIO_{periodo_ant}",
        "Mercado", "Variacion(%)", "Tendencia",
    ]
    cols_sel = list(dict.fromkeys(c for c in cols_maymen if c and c in union.columns))
    df_maymen = union[cols_sel].rename(columns={col_precio_actual: f"PRECIO_PROMEDIO_{periodo}"})
    df_maymen = _ordenar_sas(df_maymen, ["CodigoMpio", col_pub_sas])

    ruta_maymen = carpeta / f"{_MAYMEN_NOMBRES.get(m, f'{m}_MAYMEN')}_{periodo}.xlsx"
    with pd.ExcelWriter(str(ruta_maymen), engine="openpyxl") as w:
        df_maymen.to_excel(w, sheet_name=_MAYMEN_SHEET.get(m, "UNION_MAYORES_MENORESQUE2"), index=False)

    filas = len(df_mayor) + len(df_menor) + len(df_maymen)
    log.info(
        "exportar_mayo_menores [%s] OK | MAYOR=%d | MENOR=%d | MAYMEN=%d",
        modulo, len(df_mayor), len(df_menor), len(df_maymen),
    )
    return pd.DataFrame([{
        "modulo": modulo, "periodo": periodo,
        "mayor2_filas": len(df_mayor), "menor2_filas": len(df_menor),
        "maymen_filas": len(df_maymen), "filas_totales": filas,
    }])


# =============================================================================
# Archivos de revisión: REVIS_*_4MESES, TABREV/TABLASREVISION, TEMÁTICA
# =============================================================================
# Traducción directa de los programas SAS de CUADROS (sección "IMPORTAR
# ARCHIVO DE REVISION PRECIO TRES MESES ANTERIORES" y "FRECUENCIAS"):
#   - "Revisión <módulo> <mmmaaaa>.xlsx" es un INSUMO: el 4MESES del período
#     anterior del módulo sin su fecha más antigua (verificado AGO2026: igual
#     en contenido y orden). Se busca en data/01_raw/<P>/REVISIÓN <P>/ y, si no
#     está, en el que dejó la corrida anterior de Python.
#   - 4MESES = ese insumo + precios del mes (base sin duplicados, NODUPKEY).
#   - TABREV = N / promedio / CV por producto y fecha, traspuesto a MES_<m>.
#   - TEMÁTICA = base X6A + conteos y precios nacional/depto/mpio por
#     artículo y por nombre de publicación, con indicador "Revisa".

_TEMATICA_NOMBRES: dict[str, str] = {
    "AGRICOLAS":   "REVISIÓN AGRICOLAS TEMÁTICA",
    "PECUARIOS":   "REVISIÓN PECUARIOS TEMÁTICA",
    "ELEMENTOS":   "REVISIÓN ELEMENTOS TEMÁTICA",
    "PROPAGACION": "REVISIÓN MATERIAL DE PROP TEM",
}
_TEMATICA_SIN_ELSE = {"PROPAGACION"}
_TEMATICA_LLAVE: dict[str, str] = {
    "AGRICOLAS":   "Art_Unmed_Casacomer_ICA",
    "PECUARIOS":   "Art_Unmed_Casacomer_ICA",
    "ELEMENTOS":   "Art_Casacomer_ICA_Unmed",
    "PROPAGACION": "Art_Casacomer_ICA_Unmed",
}

# Llave de NODUPKEY (%LET B=) con la que SAS arma la base del mes (X7).
_REVISION_DEDUP: dict[str, list[str]] = {
    "AGRICOLAS":   ["CodigoMpio", "Fuente", "Codigo CPC", "Articulo", "CasaCom.", "RegICA", "UnMed."],
    "PECUARIOS":   ["CodigoMpio", "Fuente", "Codigo CPC", "Articulo", "CasaCom.", "RegICA", "UnMed."],
    "ELEMENTOS":   ["CodigoMpio", "Fuente", "Codigo CPC", "Articulo", "CasaCom.", "RegICA", "UnMed."],
    "PROPAGACION": ["CodigoMpio", "Fuente", "Codigo CPC", "Articulo", "CasaCom#", "RegICA", "UnMed#"],
    "EMPAQUES":    ["CodigoMpio", "Informante", "Codigo CPC", "Articulo", "Articulo_Caracte"],
    "JORNALES":    ["CodigoMpio", "Informante", "Codigo CPC", "Articulo", "Articulo_Caracte"],
    "ESPECIES":    ["CodigoMpio", "Informante", "Codigo CPC", "Articulo", "Articulo_Caracte"],
    "ARRIENDOS":   ["CodigoMpio", "Informante", "Articulo", "Articulo_Caracte"],
    "SERVICIOS":   ["CodigoMpio", "Informante", "Articulo", "Articulo_Caracte"],
}

# Tipo del Código CPC en 4MESES: "num" ('Codigo CPC'n*1), "txt" o None (sin CPC).
_REVISION_CPC: dict[str, str | None] = {
    "AGRICOLAS": "num", "PECUARIOS": "num", "EMPAQUES": "num", "JORNALES": "num",
    "ELEMENTOS": "txt", "ESPECIES": "txt", "PROPAGACION": "txt",
    "ARRIENDOS": None, "SERVICIOS": None,
}

# Columnas numéricas en la base X6A de SAS por módulo (afecta el orden NODUPKEY):
# CPC en Agrícolas/Pecuarios; RegICA en Pecuarios (todos sus registros ICA son
# números; en Agrícolas hay "PL..." y SAS lo importa como texto).
_COLS_NUMERICAS_BASE: dict[str, tuple[str, ...]] = {
    "AGRICOLAS": ("Codigo CPC",),
    "PECUARIOS": ("Codigo CPC", "RegICA"),
    "ELEMENTOS": ("Codigo CPC",),
}

# TABREV: módulos que agrupan por (CPC, nombre, fecha) en vez de (nombre, fecha)
# y ordenan el resultado por CPC (texto en Material, número en los demás).
_TABREV_AGRUPA_CPC = {"PROPAGACION", "EMPAQUES"}
_TABREV_ORDENA_CPC = {"AGRICOLAS", "PECUARIOS", "PROPAGACION", "EMPAQUES", "ELEMENTOS", "ESPECIES", "JORNALES"}
# Meses retenidos (RETAIN) en orden cronológico: 4 períodos, salvo Elementos
# (MES_&MES5 comentado en SAS); Arriendos/Servicios no retienen (orden de aparición).
_TABREV_N_MESES: dict[str, int | None] = {"ELEMENTOS": 3, "ARRIENDOS": None, "SERVICIOS": None}


def _mes_offset(periodo: str, meses: int) -> tuple[int, int]:
    mes = _MES_A_NUM[periodo[:3].upper()] + meses
    anio = int(periodo[3:])
    while mes < 1:
        mes += 12
        anio -= 1
    while mes > 12:
        mes -= 12
        anio += 1
    return mes, anio


def _nombre_revision(m: str, periodo: str) -> str:
    return f"{_REVISION_ROOT[m]} {periodo[:3].lower()}{periodo[3:]}.xlsx"


def _leer_excel_sas(ruta: Path, cpc_texto: bool = False) -> pd.DataFrame:
    """Lee un .xlsx (también los de SAS, que openpyxl rechaza) como lo importa
    SAS: encabezados vacíos pasan a F<n>. Con cpc_texto el Código CPC se
    conserva como texto (Material: "01121", no 1121)."""
    df = pd.read_excel(ruta, engine="calamine", dtype={"Codigo CPC": str} if cpc_texto else None)
    df.columns = [f"F{i + 1}" if str(c).startswith("Unnamed:") else c for i, c in enumerate(df.columns)]
    return df


def _normalizar_nombre(nombre: str) -> str:
    # Los archivos copiados de OneDrive pueden traer tildes descompuestas (NFD).
    return unicodedata.normalize("NFC", nombre).lower()


def _buscar_revision_previa(m: str, periodo: str, ruta_reporting: str) -> Path | None:
    nombre = _normalizar_nombre(_nombre_revision(m, periodo))
    data_dir = Path(ruta_reporting).parent.parent
    periodo_ant = _periodo_anterior_modulo(periodo, m)
    raw = data_dir / "01_raw" / periodo
    candidatos = [
        *(d for d in (raw.iterdir() if raw.exists() else [])
          if d.is_dir() and _normalizar_nombre(d.name) == _normalizar_nombre(f"REVISIÓN {periodo}")),
        data_dir / "08_reporting" / periodo_ant / m.lower(),
    ]
    for carpeta in candidatos:
        if carpeta.exists():
            for f in carpeta.iterdir():
                if _normalizar_nombre(f.name) == nombre:
                    return f
    return None


def _base_mes_revision(base6a: pd.DataFrame, m: str, mes_actual: str, periodo: str) -> pd.DataFrame:
    """Filas del mes para 4MESES: base X6A sin duplicados (PROC SORT NODUPKEY
    BY &B) y con precio actual, con las columnas fecha/CPC/nombre/precio."""
    llave = [c for c in _REVISION_DEDUP.get(m, []) if c in base6a.columns]
    tmp = base6a.copy()
    if llave:
        orden = tmp[llave].astype(object).where(tmp[llave].notna(), "").astype(str)
        # Columnas que SAS importa como número en la base X6A (ordenan como número:
        # 34641 < 3461101, 7395 < 21268).
        for col in _COLS_NUMERICAS_BASE.get(m, ()):
            if col in llave:
                orden[col] = pd.to_numeric(tmp[col], errors="coerce")
        tmp = tmp.loc[orden.sort_values(llave, kind="mergesort", na_position="first").index]
        tmp = tmp[~orden.loc[tmp.index].duplicated(keep="first")]
    precio = pd.to_numeric(tmp[mes_actual], errors="coerce")
    tmp = tmp[precio.notna()]
    mes, anio = _mes_offset(periodo, 0)
    pub = _NOMBRE_PUBLICA_SAS[m]
    out = pd.DataFrame({"fecha": [date(anio, mes, 1)] * len(tmp)}, index=tmp.index)
    tipo_cpc = _REVISION_CPC.get(m)
    if tipo_cpc == "num":
        out["Codigo CPC"] = pd.to_numeric(tmp["Codigo CPC"], errors="coerce")
    elif tipo_cpc == "txt":
        out["Codigo CPC"] = tmp["Codigo CPC"]
    out[pub] = tmp[pub]
    out["precio actual"] = pd.to_numeric(tmp[mes_actual], errors="coerce")
    return out.reset_index(drop=True)


def _tabrev(cuatro: pd.DataFrame, m: str, periodo: str) -> dict[str, pd.DataFrame]:
    """TABREV/TABLASREVISION: N, PRECIO_PROMEDIO y COEFIC_VARIACIÓN por
    producto y fecha, traspuestos a columnas MES_<mes>, más VAR_m1_m2 y VAR_m2_m3."""
    pub = _NOMBRE_PUBLICA_SAS[m]
    con_cpc = _REVISION_CPC.get(m) is not None and "Codigo CPC" in cuatro.columns
    df = cuatro.copy()
    df["fecha"] = pd.to_datetime(df["fecha"])
    df["precio actual"] = pd.to_numeric(df["precio actual"], errors="coerce")
    claves = (["Codigo CPC"] if (con_cpc and m in _TABREV_AGRUPA_CPC) else []) + [pub, "fecha"]
    g = df.groupby(claves, dropna=False, sort=False)
    res = g.agg(
        N=(pub, "count"),
        PRECIO_PROMEDIO=("precio actual", "mean"),
        _STD=("precio actual", "std"),
    ).reset_index()
    if con_cpc and m not in _TABREV_AGRUPA_CPC:
        res = res.merge(g["Codigo CPC"].max().reset_index(), on=claves, how="left")
    res["COEFIC_VARIACIÓN"] = res["_STD"] / res["PRECIO_PROMEDIO"] * 100
    res = res.drop(columns="_STD")
    res["MES"] = res["fecha"].dt.month

    by = (["Codigo CPC"] if con_cpc else []) + [pub]
    orden_filas = (["Codigo CPC", pub] if (con_cpc and m in _TABREV_ORDENA_CPC) else [pub])
    res = _ordenar_sas(res, [*orden_filas, "fecha"], m)

    # Columnas de mes: retenidas en orden cronológico, o por orden de aparición.
    salto = _SALTO_MESES.get(m, 1)
    n_meses = _TABREV_N_MESES.get(m, 4)
    meses_ciclo = [_mes_offset(periodo, -salto * k)[0] for k in range(4)]  # m1, m2, m3, m5
    if n_meses is None:
        meses_cols = list(dict.fromkeys(res["MES"]))
    else:
        retenidos = list(reversed(meses_ciclo[:n_meses]))
        meses_cols = retenidos + [x for x in dict.fromkeys(res["MES"]) if x not in retenidos]
    m1, m2, m3 = meses_ciclo[:3]

    hojas: dict[str, list[dict]] = {"N": [], "PRECIO_PROMEDIO": [], "COEFIC_VARIACIÓN": []}
    for clave, grupo in res.groupby(by, dropna=False, sort=False):
        clave = clave if isinstance(clave, tuple) else (clave,)
        for fuente in hojas:
            fila = dict(zip(by, clave))
            fila["Fuente"] = fuente
            vals = dict(zip(grupo["MES"], grupo[fuente]))
            for mes in meses_cols:
                fila[f"MES_{mes}"] = vals.get(mes, np.nan)
            a1, a2, a3 = (fila.get(f"MES_{x}", np.nan) for x in (m1, m2, m3))
            fila[f"VAR_{m1}_{m2}"] = (a1 - a2) / a2 * 100 if pd.notna(a1) and pd.notna(a2) and a2 != 0 else np.nan
            fila[f"VAR_{m2}_{m3}"] = (a2 - a3) / a3 * 100 if pd.notna(a2) and pd.notna(a3) and a3 != 0 else np.nan
            hojas[fuente].append(fila)
    cols = [*by, "Fuente", *[f"MES_{x}" for x in meses_cols], f"VAR_{m1}_{m2}", f"VAR_{m2}_{m3}"]
    return {k: pd.DataFrame(v, columns=cols) for k, v in hojas.items()}


def _sas_lt(a: pd.Series, b) -> pd.Series:
    """a < b con la semántica SAS: el faltante es menor que cualquier número."""
    return a.fillna(-np.inf) < (b.fillna(-np.inf) if isinstance(b, pd.Series) else b)


def _tematica(base6a: pd.DataFrame, m: str, mes_anterior: str, mes_actual: str) -> pd.DataFrame:
    """REVISIÓN <MÓDULO> TEMÁTICA: X6A + estadísticos por artículo y por nombre
    de publicación a nivel nacional, departamental y municipal (tabla X8e)."""
    df = base6a.copy()
    llave = _TEMATICA_LLAVE[m]
    pub = _NOMBRE_PUBLICA_SAS[m]
    act = pd.to_numeric(df[mes_actual], errors="coerce")
    ant = pd.to_numeric(df[mes_anterior], errors="coerce")
    con_act = act.where(act > 0)
    con_ant = ant.where(ant > 0)
    niveles = [([], "nacional", "nal"), (["CodigoDepto"], "depto", "depto"),
               (["CodigoDepto", "CodigoMpio"], "mpio", "mpio")]
    cpc_txt = m in _CPC_TEXTO
    cpc_val = (df["Codigo CPC"].astype(object).where(df["Codigo CPC"].notna(), "").astype(str) if cpc_txt
               else pd.to_numeric(df["Codigo CPC"], errors="coerce"))

    def _agg(keys, serie, fn):
        tmp = pd.DataFrame({"__v": serie})
        for k in keys:
            tmp[k] = df[k].astype(object).where(df[k].notna(), "\x00NA")
        return tmp.groupby(keys, dropna=False)["__v"].transform(fn)

    cpc_final = None
    for dim, sufijo, sufijo_var in [(llave, "artículo", "art"), (pub, "publica", "publica")]:
        for nivel, etq, etq_var in niveles:
            keys = [*nivel, dim]
            df[f"Conteo {etq} {sufijo}"] = _agg(keys, df[dim].notna().astype(int), "sum")
            n_precio = _agg(keys, con_act.notna().astype(int), "sum")
            df[f"Conteo precio {etq} {sufijo}"] = n_precio.where(n_precio > 0)
            df[f"Prom {etq} {sufijo}"] = _agg(keys, con_act, "mean")
            df[f"Mínimo {etq} {sufijo}"] = _agg(keys, con_act, "min")
            df[f"Máximo {etq} {sufijo}"] = _agg(keys, con_act, "max")
            df[f"Prom ant {etq} {sufijo}"] = _agg(keys, con_ant, "mean")
            df[f"Mínimo ant {etq} {sufijo}"] = _agg(keys, con_ant, "min")
            df[f"Máximo ant {etq} {sufijo}"] = _agg(keys, con_ant, "max")
            for tipo, col in (("prom", "Prom"), ("mín", "Mínimo"), ("máx", "Máximo")):
                df[f"Varporc prec {tipo} {etq_var} {'art' if dim == llave else 'publica'}"] = (
                    (act / df[f"{col} {etq} {sufijo}"] - 1) * 100)
            if dim == pub and etq == "mpio":
                # El último MERGE (mpio × publicación) deja el MAX(CPC) de la
                # tabla más a la derecha con fila: precio anterior > actual > todas.
                # CPC texto (Material): "" hace de faltante, como el blanco de SAS.
                vacio = "" if cpc_txt else np.nan

                def _max_cpc(serie):
                    res = _agg(keys, serie, "max")
                    return res.replace("", np.nan) if cpc_txt else res

                cpc_b = _max_cpc(cpc_val)
                cpc_c = _max_cpc(cpc_val.where(con_act.notna(), vacio))
                cpc_d = _max_cpc(cpc_val.where(con_ant.notna(), vacio))
                cpc_final = cpc_d.where(cpc_d.notna(), cpc_c.where(cpc_c.notna(), cpc_b))
    if cpc_final is not None:
        df["Codigo CPC"] = cpc_final

    min_ant = df["Mínimo ant nacional artículo"]
    max_nal = df["Máximo nacional artículo"]
    df["Varporc prec Mín ant nal art"] = (act / min_ant - 1) * 100
    df["Varporc prec Máx ant nal art"] = (act / df["Máximo ant nacional artículo"] - 1) * 100
    a = act.fillna(-np.inf)
    ok1 = (a > min_ant.fillna(-np.inf)) & (a < max_nal.fillna(-np.inf))
    ok2 = (a == min_ant.fillna(-np.inf)) & (a == max_nal.fillna(-np.inf))
    ok3 = act.isna()
    rev1 = _sas_lt(df["Varporc prec Mín ant nal art"], -10)
    rev2 = df["Varporc prec Máx ant nal art"].fillna(-np.inf) > 10
    # El programa de Material no tiene el "else Revisa='OK'" final: queda vacío.
    defecto = "" if m in _TEMATICA_SIN_ELSE else "OK"
    df["Revisa"] = np.select([ok1, ok2, ok3, rev1, rev2], ["OK", "OK", "OK", "Revisar", "Revisar"], defecto)

    # Orden final de los PROC SORT encadenados: depto, mpio, publicación, artículo (estable).
    return _ordenar_sas(df, ["CodigoDepto", "CodigoMpio", pub, llave], m)


def exportar_revisiones(
    base_completa: pd.DataFrame,
    divipola_raw: pd.DataFrame,
    mappings_grupos: dict,
    mappings_articulos: dict,
    modulo: str,
    periodo: str,
    mes_actual: str,
    mes_anterior: str,
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta REVIS_*_4MESES, TABREV/TABLASREVISION y (si aplica) la TEMÁTICA
    del módulo, y deja el insumo "Revisión <módulo> <período siguiente>.xlsx"
    (4MESES sin su fecha más antigua) para la próxima corrida."""
    m = modulo.upper()
    carpeta = Path(ruta_reporting) / modulo.lower()
    carpeta.mkdir(parents=True, exist_ok=True)
    base6a, mes_ant_mod = _base_insumos_sas(base_completa, divipola_raw, mappings_grupos,
                                            mappings_articulos, modulo, periodo, mes_actual, mes_anterior)
    mes_ant_mod, mes_actual = _nombres_mes(m, mes_ant_mod, mes_actual)
    metas = []

    # --- 4MESES ---
    previa_ruta = _buscar_revision_previa(m, periodo, ruta_reporting)
    if previa_ruta is None:
        log.warning("exportar_revisiones [%s] | no se encontró '%s' (ni en 01_raw/%s/REVISIÓN %s "
                    "ni en la corrida anterior) — 4MESES solo tendrá el mes actual.",
                    m, _nombre_revision(m, periodo), periodo, periodo)
        previa = pd.DataFrame()
    else:
        previa = _leer_excel_sas(previa_ruta, cpc_texto=_REVISION_CPC.get(m) == "txt")
        log.info("exportar_revisiones [%s] | insumo de revisión: %s (%d filas)", m, previa_ruta, len(previa))
    actual = _base_mes_revision(base6a, m, mes_actual, periodo)
    renombrar = {c: a for c in previa.columns for a in actual.columns if str(c).lower() == a.lower()}
    previa = previa.rename(columns=renombrar)
    if "fecha" in previa.columns:
        previa["fecha"] = pd.to_datetime(previa["fecha"]).dt.date
    cols = [*actual.columns, *[c for c in previa.columns if c not in actual.columns]]
    cuatro = pd.concat([previa, actual], ignore_index=True).reindex(columns=cols)
    ruta_4m = carpeta / f"{_REVIS4M_NOMBRES[m]}_{periodo}.xlsx"
    escribir_excel_multisheet(ruta_4m, {"1.base": cuatro})
    metas.append({"tipo": "4MESES", "archivo": str(ruta_4m), "filas": len(cuatro)})

    # --- Insumo para el siguiente período del módulo ---
    if len(cuatro):
        siguiente = cuatro[pd.to_datetime(cuatro["fecha"]) != pd.to_datetime(cuatro["fecha"]).min()]
        mes_sig, anio_sig = _mes_offset(periodo, _SALTO_MESES.get(m, 1))
        ruta_sig = carpeta / _nombre_revision(m, f"{_NUM_A_MES[mes_sig]}{anio_sig}")
        escribir_excel_multisheet(ruta_sig, {"1.base": siguiente})
        metas.append({"tipo": "REVISION_SIGUIENTE", "archivo": str(ruta_sig), "filas": len(siguiente)})

    # --- TABREV / TABLASREVISION ---
    hojas = _tabrev(cuatro, m, periodo)
    ruta_tab = carpeta / f"{_TABREV_NOMBRES[m]}_{periodo}.xlsx"
    escribir_excel_multisheet(ruta_tab, hojas)
    metas.append({"tipo": "TABREV", "archivo": str(ruta_tab), "filas": len(hojas["N"])})

    # --- TEMÁTICA ---
    if m in _TEMATICA_NOMBRES:
        tem = _tematica(base6a, m, mes_ant_mod, mes_actual)
        ruta_tem = carpeta / f"{_TEMATICA_NOMBRES[m]} {periodo}.xlsx"
        escribir_excel_multisheet(ruta_tem, {"REVISIÓN": tem})
        metas.append({"tipo": "TEMATICA", "archivo": str(ruta_tem), "filas": len(tem)})

    for meta in metas:
        log.info("exportar_revisiones [%s] OK | tipo=%s | archivo=%s | filas=%d",
                 m, meta["tipo"], meta["archivo"], meta["filas"])
    return pd.DataFrame([{"modulo": modulo, "periodo": periodo, **meta} for meta in metas])
