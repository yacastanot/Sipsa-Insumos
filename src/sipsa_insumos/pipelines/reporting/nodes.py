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
from pathlib import Path

import numpy as np
import pandas as pd

from sipsa_insumos.pipelines.quality.nodes import _LLAVE_DUPLICADOS
from sipsa_insumos.utils.excel_writer import aplicar_formato_numerico_precio, escribir_excel_multisheet

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
        df_out = df_out.sort_values(
            [c for c in ["NombreDepartamento", "NombreMunicipio", col_pub_sas] if c in df_out.columns]
        )
        hojas[grupo] = df_out

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
            df_grupo.groupby(["Nombre_Publica", "TENDENCIA"], observed=True)
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        # Asegurar que todas las columnas de tendencia existen
        for tend in ["Positiva", "Negativa", "Estable", "n.d."]:
            if tend not in pivot.columns:
                pivot[tend] = 0
        cols_orden = ["Nombre_Publica"] + [c for c in ["Positiva", "Negativa", "Estable", "n.d."] if c in pivot.columns]
        pivot = pivot[cols_orden]
        pivot["TOTAL"] = pivot[["Positiva", "Negativa", "Estable", "n.d."]].sum(axis=1)
        hojas[grupo] = pivot

    escribir_excel_multisheet(ruta, hojas)

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
            partes = df_grupo["Nombre_Publica"].str.rsplit(",", n=1, expand=True)
            df_grupo["Nombre_insumo"] = partes[0].str.strip()
            df_grupo["Presentación_insumo"] = partes[1].str.strip() if 1 in partes.columns else ""

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
        df_out = df_out.sort_values(
            [c for c in ["NombreDepartamento", "NombreMunicipio", col_pub_sas] if c in df_out.columns]
        )
        hojas[grupo] = df_out

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
    """Exporta BASE_INSUMOS_{MODULO}_{PERIODO}.xlsx — réplica exacta del
    equivalente SAS: la base completa (sin filtrar por Estado ni Precio,
    a diferencia del flujo principal) enriquecida con DIVIPOLA/Grupo/
    Nombre_Publica, con la variación calculada fila por fila usando el
    precio anterior embebido en el propio archivo del mes (columna
    "Precio Ante."), no el promedio agregado del período anterior.

    SAS (paso final del programa de CUADROS):
      MERGE DIVIPOLA base; BY CodigoMpio; IF B;             (conserva sin cruce)
      MERGE base grupos/articulos; BY LLAVE; IF A;          (conserva sin cruce)
      IF Precio Ante./Actual = 0 THEN falta;
      IF Nov. IN ('IA','IN') THEN Precio Ante. = falta;
      VAR = (Actual - Ante) / Ante * 100;

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
    df = _construir_base_diagnostico(base_completa, divipola_raw, mappings_grupos, mappings_articulos)
    col_pub = _NOMBRE_PUBLICA_SAS.get(modulo.upper(), "Nombre_Publica")

    col_map = {
        "CodigoDepto":                 "CodigoDepto",
        "CodigoMpio":                  "CÓDIGO DIVIPOLA",
        "NombreDepartamento":          "NombreDepartamento" if "NombreDepartamento" in df.columns else None,
        "NombreMunicipio":             "NombreMunicipio" if "NombreMunicipio" in df.columns else None,
        "Fuente":                      "FUENTE",
        "Codigo CPC":                  "CÓDIGO CPC",
        "Articulo":                    "ARTÍCULO",
        col_pub:                       "Nombre_Publica",
        "Grupo":                       "Grupo",
        "UnMed.":                      "UNIDAD DE MEDIDA",
        "CasaCom.":                    "CASA COMERCIAL",
        "RegICA":                      "REGISTRO ICA",
        mes_anterior:                  "PRECIO_ANTERIOR_RAW",
        mes_actual:                   "PRECIO",
        "VAR":                         "VAR",
        "Estado":                      "ESTADO",
        "Art_Unmed_Casacomer_ICA":     "LLAVE_ARTICULO",
        "Observación":                 "OBSERVACION",
        "Nov.":                        "NOV",
    }
    cols_sel = [v for v in col_map.values() if v and v in df.columns]
    rename_inv = {v: k for k, v in col_map.items() if v and v in df.columns}
    df_out = df[cols_sel].rename(columns=rename_inv)

    nombre = f"BASE_INSUMOS_{modulo}_{periodo}.xlsx"
    ruta = Path(ruta_reporting) / modulo.lower() / nombre
    ruta.parent.mkdir(parents=True, exist_ok=True)
    escribir_excel_multisheet(ruta, {"INFORMACIÓN INSUMOS": df_out})

    log.info("exportar_base_insumos [%s] OK | archivo=%s | filas=%d", modulo, nombre, len(df_out))
    return pd.DataFrame([{
        "archivo": str(ruta),
        "modulo": modulo,
        "periodo": periodo,
        "tipo": "BASE_INSUMOS",
        "filas_totales": len(df_out),
    }])


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
}

_DIAG_ORDEN_DUPLI: dict[str, list[str]] = {
    "AGRICOLAS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "Grupo", "UnMed.", "CasaCom.", "RegICA", "__MES_ANT__", "__MES_ACT__", "VAR", "Estado", "Art_Unmed_Casacomer_ICA", "Observación", "Nov."],
    "PECUARIOS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "Grupo", "UnMed.", "CasaCom.", "RegICA", "__MES_ANT__", "__MES_ACT__", "VAR", "Nov.", "Estado", "Art_Unmed_Casacomer_ICA", "Observación"],
    "ELEMENTOS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "UnMed.", "CasaCom.", "RegICA", "Estado", "Art_Casacomer_ICA_Unmed", "Nov."],
    "EMPAQUES":  ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Estado", "Articulo_Caracte", "Informante", "Caracte.", "Nov."],
    "ARRIENDOS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Codigo CPC", "Caracte#"],
    "SERVICIOS": ["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Articulo", "__PUBLICA__", "__MES_ANT__", "__MES_ACT__", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte", "Informante", "Codigo CPC", "Caracte#"],
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
}
_DUPLI_NOMBRES: dict[str, str] = {
    "AGRICOLAS": "INSUMOS_AGRICOLAS_DUPLI",
    "PECUARIOS": "INSUMOS_PECUARIOS_DUPLI",
    "ELEMENTOS": "ELEM_AGROPE_DUPLI",
    "EMPAQUES":  "EMPA_AGROPE_DUPLI",
    "ARRIENDOS": "ARRIENDOS_DUPLI",
    "SERVICIOS": "SERVICIOS_DUPLI",
}
_VAR_ATIPICO_NOMBRES: dict[str, str] = {
    "AGRICOLAS": "VAR_ATIPICO_AGRICOLA",
    "PECUARIOS": "VAR_ATIPICO_PECUARIO",
    "ELEMENTOS": "VAR_ATIPICO_ELEMENTOS",
    "EMPAQUES":  "VAR_ATIPICO_EMPAQUES",
    "ARRIENDOS": "VAR_ATIPICO_ARRIENDOS",
    "SERVICIOS": "VAR_ATIPICO_SERVICIOS",
}
_DUPLI_SHEET: dict[str, str] = {
    "AGRICOLAS": "AGRICOLA_DUPLI",
    "PECUARIOS": "PECUARIO_DUPLI",
    "ELEMENTOS": "ELEM_AGROPE_DUPLI",
    "EMPAQUES":  "EMPA_AGROPE_DUPLI",
    "ARRIENDOS": "ARRIENDOS_DUPLI",
    "SERVICIOS": "SERVICIOS_DUPLI",
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
        lbl["llave"]:         "LLAVE_ARTICULO",
    }
    if tipo_modulo == "caracte":
        m[lbl["caracte"]] = "CARACTERÍSTICA"
    else:
        m["UnMed."] = "UNIDAD DE MEDIDA"
        m["CasaCom."] = "CASA COMERCIAL"
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


def _sustituir_orden(orden: list[str], modulo: str, mes_anterior: str, mes_actual: str) -> list[str]:
    col_pub = _NOMBRE_PUBLICA_SAS.get(modulo, "Nombre_Publica")
    return [
        {"__PUBLICA__": col_pub, "__MES_ANT__": mes_anterior, "__MES_ACT__": mes_actual}.get(c, c)
        for c in orden
    ]


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
    ruta.parent.mkdir(parents=True, exist_ok=True)
    escribir_excel_multisheet(ruta, {sheet_name[:31]: df_out})
    return len(df_out)


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
            ruta, f"FALTAN_GRUPO_{m}",
        )
        metas.append({"tipo": "FALTAN_GRUPO", "archivo": str(ruta), "filas": n})

    # --- FALTAN_PUBLICA ---
    df_fp = df[df["Nombre_Publica"].isna()]
    col_map_fp = _diag_col_map(m, tipo_modulo, mes_anterior, mes_actual, kind="faltan")
    ruta = carpeta / f"{_FALTAN_PUBLICA_NOMBRES[m]}_{periodo}.xlsx"
    n = _exportar_diag_hoja(
        df_fp, _DIAG_ORDEN_FALTAN_PUBLICA[m], col_map_fp, m, mes_anterior, mes_actual,
        ruta, f"FALTAN_PUBLICA_{m}",
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
        ruta, _VAR_ATIPICO_NOMBRES[m],
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
        cuadro = cuadro.rename(columns={
            "Nombre_Publica": "Producto",
            "Positiva": "SUBIO",
            "Negativa": "BAJO",
            "Estable": "ESTABLE",
            "n.d.": "N.D.",
        })

        cols_orden = ["CÓDIGO CPC", "Producto", "SUBIO", "BAJO", "ESTABLE", "N.D.", "PRECIO_MIN", "PRECIO_MAX"]
        cols_existentes = [c for c in cols_orden if c in cuadro.columns]
        cuadro = cuadro[cols_existentes].sort_values("Producto")
        hojas[grupo] = cuadro

    # Escribir con formato numérico en columnas de precio
    Path(ruta).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(str(ruta), engine="openpyxl") as writer:
        for nombre_hoja, df_hoja in hojas.items():
            df_hoja.to_excel(writer, sheet_name=nombre_hoja, index=False)
            ws = writer.sheets[nombre_hoja]
            # Aplicar formato de precio a PRECIO_MIN y PRECIO_MAX
            for col_name in ["PRECIO_MIN", "PRECIO_MAX"]:
                if col_name in df_hoja.columns:
                    col_idx = list(df_hoja.columns).index(col_name) + 1
                    aplicar_formato_numerico_precio(ws, col_idx, len(df_hoja))

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
    "PROPAGACION": "MATE_MAY_MEN2",
    "JORNALES":    "JORNALES_MAYMEN2",
    "ESPECIES":    "ESPE_MAYMEN2",
}

_TABREV_NOMBRES: dict[str, str] = {
    "AGRICOLAS":   "TABREV_AGRI",
    "PECUARIOS":   "INSPEC_TABREV",
    "ELEMENTOS":   "ELEM_TABLASREVISION",
    "EMPAQUES":    "EMPA_TABLASREVISION",
    "ARRIENDOS":   "ARRIENDOS_TABLASREVISION",
    "SERVICIOS":   "SERVICIOS_TABLASREVISION",
    "PROPAGACION": "MATERIAL_TABLASREVISION",
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
    "PROPAGACION": "REVIS_MATERIAL_4MESES",
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
    "PROPAGACION": "Revisión material propagación",
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
    "PROPAGACION": "MATERIAL_MAYORESOIGUALES2",
    "JORNALES":    "JORNALES_MAYORESOIGUALES2",
    "ESPECIES":    "ESPE_MAYORESOIGUALES2",
}

_MENOR2_SHEET: dict[str, str] = {
    "AGRICOLAS":   "INSU_AGRICOLAS_MENORESQUE2",
    "PECUARIOS":   "INSU_PECUARIOS_MENORESQUE2",
    "ELEMENTOS":   "ELEM_AGROPE_MENORESQUE2",
    "EMPAQUES":    "EMPA_AGROPE_MENORESQUE2",
    "ARRIENDOS":   "ARRIENDOS_MENORESQUE2",
    "SERVICIOS":   "SERVICIOS_MENORESQUE2",
    "PROPAGACION": "MATERIAL_MENORESQUE2",
    "JORNALES":    "JORNALES_MENORESQUE2",
    "ESPECIES":    "ESPE_MENORESQUE2",
}

# Módulos con REVISIÓN TEMÁTICA (etiqueta para el nombre del archivo)
_REVISION_TEMATICA_LABEL: dict[str, str] = {
    "AGRICOLAS":   "AGRICOLAS",
    "PECUARIOS":   "PECUARIOS",
    "ELEMENTOS":   "ELEMENTOS",
}

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
    "PROPAGACION": "Nombre_material_propagacion_publ",
    "JORNALES":    "Nombre_jornales_publ",
    "ESPECIES":    "Nombre_especies_productivas_publ",
}

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
        col_n_fuente:           "N_FUENTE" if "N_FUENTE" in mayor2.columns else None,
        f"N_ARTICULOS{sfx}":    "N_ARTICULOS" if "N_ARTICULOS" in mayor2.columns else None,
        f"PRECIO_PROMEDIO{sfx}":"PRECIO_PROMEDIO" if "PRECIO_PROMEDIO" in mayor2.columns else None,
    }


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
    df_mayor = _preparar_mayor_menor(mayor2, modulo=modulo, periodo=periodo)
    ruta_mayor = carpeta / f"{_MAYOR2_NOMBRES.get(m, f'{m}_MAYORESQUE2')}_{periodo}.xlsx"
    sheet_mayor = _MAYOR2_SHEET.get(m, "MAYORESOIGUALES2")
    with pd.ExcelWriter(str(ruta_mayor), engine="openpyxl") as w:
        df_mayor.to_excel(w, sheet_name=sheet_mayor[:31], index=False)

    # --- MENORESQUE2 ---
    df_menor = _preparar_mayor_menor(menor2, modulo=modulo, periodo=periodo)
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
    union = union.rename(columns={
        "N_FUENTE":    f"{col_n_fuente}_{periodo}",
        "N_ARTICULOS": f"N_ARTICULOS_{periodo}",
    })

    if mayor2_anterior is not None and len(mayor2_anterior) > 0 and llave_join:
        col_div_ant = next(
            (c for c in ["CÓDIGO DIVIPOLA", "CodigoMpio"] if c in mayor2_anterior.columns), None
        )
        llave_ant = [c for c in [col_div_ant, "Nombre_Publica"] if c]
        if set(llave_ant) == set(llave_join):
            cols_ant = llave_ant + ["PRECIO_PROMEDIO"]
            if "N_FUENTE" in mayor2_anterior.columns:
                cols_ant.append("N_FUENTE")
            if "N_ARTICULOS" in mayor2_anterior.columns:
                cols_ant.append("N_ARTICULOS")
            ant = mayor2_anterior[cols_ant].rename(columns={
                "PRECIO_PROMEDIO": f"PRECIO_PROMEDIO_{periodo_ant}",
                "N_FUENTE":        f"{col_n_fuente}_{periodo_ant}",
                "N_ARTICULOS":     f"N_ARTICULOS_{periodo_ant}",
            })
            union = union.merge(ant, left_on=llave_join, right_on=llave_ant, how="left")

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

    cols_maymen = [
        "CodigoDepto", "NombreDepartamento", "CodigoMpio", "NombreMunicipio", "Mercado",
        *(["Codigo CPC"] if m not in _SIN_CODIGO_CPC else []),
        col_pub_sas,
        *(["Grupo"] if m in _MULTI_GRUPO_MODULOS else []),
        f"{col_n_fuente}_{periodo_ant}", f"N_ARTICULOS_{periodo_ant}",
        f"PRECIO_PROMEDIO_{periodo_ant}",
        f"{col_n_fuente}_{periodo}", f"N_ARTICULOS_{periodo}",
        col_precio_actual,
        "Variacion(%)", "Tendencia",
    ]
    cols_sel = list(dict.fromkeys(c for c in cols_maymen if c and c in union.columns))
    df_maymen = union[cols_sel].rename(columns={col_precio_actual: f"PRECIO_PROMEDIO_{periodo}"})

    ruta_maymen = carpeta / f"{_MAYMEN_NOMBRES.get(m, f'{m}_MAYMEN')}_{periodo}.xlsx"
    with pd.ExcelWriter(str(ruta_maymen), engine="openpyxl") as w:
        df_maymen.to_excel(w, sheet_name="UNION_MAYORES_MENORESQUE2", index=False)

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


def exportar_revision_historica(
    base_enriquecida: pd.DataFrame,
    modulo: str,
    periodo: str,
    mes_actual: str,
    mes_num_actual: int,
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta TABREV (3 hojas), REVIS_4MESES y la hoja rodante de historial.

    La hoja rodante acumula hasta 4 períodos de precios individuales por producto.
    Si existe un archivo previo en la misma carpeta de reportes, se carga y amplía.

    SAS equivalente:
      ARTICULO_PUBLICA_PRECIO2 = ARTICULO_PUBLICA_PRECIO (hist. 3 per.) + PRECIO_ACTUAL
      COMPARA_PRECIO = agg por (Nombre_Publica, FECHA): N, PRECIO_PROMEDIO, CV
      COMPARA_PRECIO3 = pivot wide + variaciones
      TABREV = split por Fuente → 3 hojas
      REVIS_4MESES = ARTICULO_PUBLICA_PRECIO2 (raw)

    Args:
        base_enriquecida: Registros individuales del período actual con PRECIO.
        modulo: Nombre del módulo en mayúsculas.
        periodo: ID del período actual (ej: "MAY2026").
        mes_actual: Nombre mes actual (ej: "Mayo").
        mes_num_actual: Número del mes actual (1-12).
        ruta_reporting: Directorio raíz de reportes.

    Returns:
        DataFrame de metadatos.
    """
    m = modulo.upper()
    carpeta = Path(ruta_reporting) / modulo.lower()
    carpeta.mkdir(parents=True, exist_ok=True)

    col_pub = next(
        (c for c in ["Nombre_Publica", "Nombre_productos_agr_publ"] if c in base_enriquecida.columns),
        None,
    )
    col_cpc = next((c for c in ["CÓDIGO CPC"] if c in base_enriquecida.columns), None)
    col_precio = "PRECIO" if "PRECIO" in base_enriquecida.columns else None
    col_fuente = "FUENTE" if "FUENTE" in base_enriquecida.columns else None
    col_pub_sas = _NOMBRE_PUBLICA_SAS.get(m, "Nombre_Publica")

    if not (col_pub and col_precio):
        log.warning("exportar_revision_historica [%s] | sin columnas requeridas — omitiendo", modulo)
        return pd.DataFrame([{"modulo": modulo, "periodo": periodo, "omitido": True}])

    # Construir registros del período actual
    cols_sel = [c for c in [col_cpc, col_fuente, col_pub, col_precio] if c]
    precio_actual = base_enriquecida[cols_sel].dropna(subset=[col_precio]).copy()
    rename_base: dict[str, str] = {col_pub: "Nombre_Publica", col_precio: "Precio actual"}
    if col_cpc:
        rename_base[col_cpc] = "Codigo CPC"
    precio_actual = precio_actual.rename(columns=rename_base)
    precio_actual["MES_NUM"] = mes_num_actual
    precio_actual["PERIODO"] = periodo

    # Cargar historial del período anterior (si existe)
    periodo_ant = _periodo_anterior_modulo(periodo, modulo)
    hist_path = carpeta / f"revision_historica_{periodo_ant}.parquet"
    if hist_path.exists():
        hist_ant = pd.read_parquet(hist_path)
    else:
        hist_ant = pd.DataFrame(columns=precio_actual.columns)

    # Acumular: agregar período actual y mantener últimos 4 períodos
    hist_curr = pd.concat([hist_ant, precio_actual], ignore_index=True)
    if "PERIODO" in hist_curr.columns:
        periodos_unicos = hist_curr["PERIODO"].dropna().unique()
        if len(periodos_unicos) > 4:
            orden = sorted(periodos_unicos, key=lambda p: (_MES_A_NUM.get(p[:3].upper(), 0), int(p[3:])))
            periodos_keep = orden[-4:]
            hist_curr = hist_curr[hist_curr["PERIODO"].isin(periodos_keep)]

    # Guardar historial actualizado como parquet
    hist_path_curr = carpeta / f"revision_historica_{periodo}.parquet"
    hist_curr.to_parquet(str(hist_path_curr), index=False)

    # ---- REVIS_4MESES ----
    ruta_revis4m = carpeta / f"{_REVIS4M_NOMBRES.get(m, f'REVIS_{m}_4MESES')}_{periodo}.xlsx"
    with pd.ExcelWriter(str(ruta_revis4m), engine="openpyxl") as w:
        hist_curr.to_excel(w, sheet_name="1.base", index=False)

    # ---- Hoja rodante root "Revisión" ----
    root_nombre = _REVISION_ROOT.get(m, f"Revisión {modulo.lower()}")
    ruta_root = carpeta / f"{root_nombre} {periodo.lower()}.xlsx"
    with pd.ExcelWriter(str(ruta_root), engine="openpyxl") as w:
        hist_curr.to_excel(w, sheet_name="1.base", index=False)

    # ---- TABREV: pivot por (FUENTE, Nombre_Publica) × período → MES_2..MES_5 ----
    if "Nombre_Publica" in hist_curr.columns and "PERIODO" in hist_curr.columns:
        # Mapear períodos a índices secuenciales (más antiguo → MES_2, más reciente → MES_5)
        periodos_ord = sorted(
            hist_curr["PERIODO"].dropna().unique(),
            key=lambda p: (_MES_A_NUM.get(p[:3].upper(), 0), int(p[3:])),
        )
        n_per = len(periodos_ord)
        # MES_2 = el más antiguo, MES_{2+n-1} = el más reciente (hasta MES_5)
        mes_idx_start = max(2, 6 - n_per)
        periodo_a_mes = {p: f"MES_{mes_idx_start + i}" for i, p in enumerate(periodos_ord)}
        hist_curr["MES_IDX"] = hist_curr["PERIODO"].map(periodo_a_mes)

        grp_cols = ["Nombre_Publica"]
        if col_fuente and "FUENTE" in hist_curr.columns:
            grp_cols = ["FUENTE"] + grp_cols

        compara = (
            hist_curr.groupby(grp_cols + ["MES_IDX"], dropna=False)["Precio actual"]
            .agg(N="count", PRECIO_PROMEDIO="mean", STD="std")
            .reset_index()
        )
        if "Codigo CPC" in hist_curr.columns:
            cpc_mode = (
                hist_curr.groupby("Nombre_Publica")["Codigo CPC"]
                .agg(lambda x: x.mode().iloc[0] if not x.mode().empty else x.iloc[0])
                .reset_index()
            )
            compara = compara.merge(cpc_mode, on="Nombre_Publica", how="left")

        compara["COEFIC_VARIACION"] = (compara["STD"] / compara["PRECIO_PROMEDIO"] * 100).where(
            compara["N"] >= 2
        )

        cpc_col = "Codigo CPC" if "Codigo CPC" in compara.columns else None
        idx_cols_tabrev = (
            (["Codigo CPC"] if cpc_col else [])
            + (["FUENTE"] if "FUENTE" in compara.columns else [])
            + ["Nombre_Publica"]
        )
        idx_present = [c for c in idx_cols_tabrev if c in compara.columns]

        def _pivot_tabrev_stat(stat_col: str) -> pd.DataFrame:
            piv = compara.pivot_table(
                index=idx_present, columns="MES_IDX", values=stat_col, aggfunc="first"
            ).reset_index()
            piv.columns = [str(c) for c in piv.columns]
            mes_cols = [f"MES_{i}" for i in range(2, 7) if f"MES_{i}" in piv.columns]
            piv = piv[idx_present + mes_cols]
            if stat_col == "PRECIO_PROMEDIO":
                for j in range(1, len(mes_cols)):
                    newer, older = mes_cols[j], mes_cols[j - 1]
                    ni = int(newer.split("_")[1])
                    oi = int(older.split("_")[1])
                    piv[f"VAR_{ni}_{oi}"] = ((piv[newer] / piv[older]) - 1) * 100
            piv = piv.rename(columns={"Nombre_Publica": col_pub_sas})
            return piv

        ruta_tabrev = carpeta / f"{_TABREV_NOMBRES.get(m, f'TABREV_{m}')}_{periodo}.xlsx"
        with pd.ExcelWriter(str(ruta_tabrev), engine="openpyxl") as w:
            _pivot_tabrev_stat("N").to_excel(w, sheet_name="N", index=False)
            _pivot_tabrev_stat("PRECIO_PROMEDIO").to_excel(w, sheet_name="PRECIO_PROMEDIO", index=False)
            _pivot_tabrev_stat("COEFIC_VARIACION").to_excel(w, sheet_name="COEFIC_VARIACIÓN", index=False)
    else:
        log.warning("exportar_revision_historica [%s] | historial vacío — TABREV omitido", modulo)

    log.info(
        "exportar_revision_historica [%s] OK | periodo=%s | filas_hist=%d",
        modulo, periodo, len(hist_curr),
    )
    return pd.DataFrame([{
        "modulo": modulo, "periodo": periodo,
        "filas_historico": len(hist_curr),
        "periodos_acumulados": hist_curr.get("PERIODO", pd.Series()).nunique(),
    }])


def exportar_revision_tematica(
    base_enriquecida: pd.DataFrame,
    mayor2_anterior: pd.DataFrame | None,
    modulo: str,
    periodo: str,
    mes_actual: str,
    grupos: list[str],
    ruta_reporting: str,
) -> pd.DataFrame:
    """Exporta la revisión temática con variaciones por nivel geográfico.

    Solo se genera para módulos con label en _REVISION_TEMATICA_LABEL
    (agricolas, pecuarios, elementos). Para los demás retorna metadatos vacíos.

    SAS equivalente (Insumos_agricolas8e):
      Agrega por 6 niveles: nal/depto/mpio × artículo/publica
      Une al nivel registro individual
      Flag Revisa = "Revisar" si precio > 10% fuera del rango nacional del artículo

    Args:
        base_enriquecida: Registros individuales del período actual con PRECIO.
        mayor2_anterior: mayor2 del período anterior (para rango de referencia).
        modulo: Nombre del módulo.
        periodo: ID del período.
        mes_actual: Nombre mes actual.
        grupos: Lista de grupos del módulo.
        ruta_reporting: Directorio raíz de reportes.

    Returns:
        DataFrame de metadatos.
    """
    m = modulo.upper()
    if m not in _REVISION_TEMATICA_LABEL:
        return pd.DataFrame([{"modulo": modulo, "periodo": periodo, "omitido": True}])

    carpeta = Path(ruta_reporting) / modulo.lower()
    carpeta.mkdir(parents=True, exist_ok=True)

    label = _REVISION_TEMATICA_LABEL[m]
    ruta_tema = carpeta / f"REVISIÓN {label} TEMÁTICA {periodo}.xlsx"

    df = base_enriquecida.copy()

    col_pub = next((c for c in ["Nombre_Publica"] if c in df.columns), None)
    col_art = "ARTÍCULO" if "ARTÍCULO" in df.columns else None
    col_div = next((c for c in ["CÓDIGO DIVIPOLA", "CodigoMpio"] if c in df.columns), None)
    col_precio = "PRECIO" if "PRECIO" in df.columns else None

    if not (col_pub and col_precio):
        log.warning("exportar_revision_tematica [%s] | sin columnas requeridas — omitiendo", modulo)
        return pd.DataFrame([{"modulo": modulo, "periodo": periodo, "omitido": True}])

    if "CodigoDepto" not in df.columns and col_div:
        df["CodigoDepto"] = df[col_div].astype(str).str[:2]

    def _agg_nivel(grp_keys: list[str], prefix: str) -> pd.DataFrame:
        keys = [k for k in grp_keys if k in df.columns]
        if not keys:
            return pd.DataFrame()
        return (
            df.groupby(keys, dropna=False)[col_precio]
            .agg(**{
                f"conteo_{prefix}": "count",
                f"prom_{prefix}": "mean",
                f"min_{prefix}": "min",
                f"max_{prefix}": "max",
            })
            .reset_index()
        )

    def _agg_ant_nivel(grp_keys: list[str], prefix: str) -> pd.DataFrame:
        if mayor2_anterior is None or len(mayor2_anterior) == 0:
            return pd.DataFrame()
        keys = [k for k in grp_keys if k in mayor2_anterior.columns]
        if not keys or "PRECIO_PROMEDIO" not in mayor2_anterior.columns:
            return pd.DataFrame()
        return (
            mayor2_anterior.groupby(keys, dropna=False)["PRECIO_PROMEDIO"]
            .agg(**{
                f"prom_ant_{prefix}": "mean",
                f"min_ant_{prefix}": "min",
                f"max_ant_{prefix}": "max",
            })
            .reset_index()
        )

    # ---- Agregaciones Nombre_Publica ----
    agg_nal_pub   = _agg_nivel([col_pub],                            "nal_publica")
    agg_depto_pub = _agg_nivel(["CodigoDepto", col_pub],             "depto_publica")
    agg_mpio_pub  = _agg_nivel(["CodigoDepto", col_div, col_pub],    "mpio_publica") if col_div else pd.DataFrame()
    ant_nal_pub   = _agg_ant_nivel([col_pub],                        "nal_publica")

    if not ant_nal_pub.empty:
        agg_nal_pub = agg_nal_pub.merge(ant_nal_pub, on=col_pub, how="left")
    else:
        for c in [f"prom_ant_nal_publica", f"min_ant_nal_publica", f"max_ant_nal_publica"]:
            agg_nal_pub[c] = float("nan")

    # ---- Agregaciones ARTÍCULO ----
    agg_nal_art   = _agg_nivel([col_art],                            "nal_articulo")   if col_art else pd.DataFrame()
    agg_depto_art = _agg_nivel(["CodigoDepto", col_art],             "depto_articulo") if col_art else pd.DataFrame()
    agg_mpio_art  = _agg_nivel(["CodigoDepto", col_div, col_art],    "mpio_articulo")  if (col_art and col_div) else pd.DataFrame()
    ant_nal_art   = _agg_ant_nivel([col_art],                        "nal_articulo")   if col_art else pd.DataFrame()

    if col_art and not ant_nal_art.empty:
        agg_nal_art = agg_nal_art.merge(ant_nal_art, on=col_art, how="left")
    elif col_art and not agg_nal_art.empty:
        for c in ["prom_ant_nal_articulo", "min_ant_nal_articulo", "max_ant_nal_articulo"]:
            agg_nal_art[c] = float("nan")

    # ---- Join al nivel de registro individual ----
    result = df.merge(agg_nal_pub, on=col_pub, how="left")
    result = result.merge(agg_depto_pub, on=["CodigoDepto", col_pub], how="left")
    if not agg_mpio_pub.empty:
        result = result.merge(agg_mpio_pub, on=["CodigoDepto", col_div, col_pub], how="left")
    if col_art and not agg_nal_art.empty:
        result = result.merge(agg_nal_art, on=col_art, how="left")
    if col_art and not agg_depto_art.empty:
        result = result.merge(agg_depto_art, on=["CodigoDepto", col_art], how="left")
    if col_art and not agg_mpio_art.empty:
        result = result.merge(agg_mpio_art, on=["CodigoDepto", col_div, col_art], how="left")

    # ---- Variaciones porcentuales ----
    varporc_pairs = [
        ("varporc_prom_nal_publica",    "prom_nal_publica"),
        ("varporc_min_nal_publica",     "min_nal_publica"),
        ("varporc_max_nal_publica",     "max_nal_publica"),
        ("varporc_prom_depto_publica",  "prom_depto_publica"),
        ("varporc_prom_mpio_publica",   "prom_mpio_publica"),
        ("varporc_prom_nal_articulo",   "prom_nal_articulo"),
        ("varporc_min_nal_articulo",    "min_nal_articulo"),
        ("varporc_max_nal_articulo",    "max_nal_articulo"),
        ("varporc_prom_depto_articulo", "prom_depto_articulo"),
        ("varporc_prom_mpio_articulo",  "prom_mpio_articulo"),
    ]
    for new_col, ref in varporc_pairs:
        if ref in result.columns:
            result[new_col] = (result[col_precio] / result[ref] - 1) * 100

    # ---- Flag Revisa (basado en rango nacional de Nombre_Publica) ----
    if "min_ant_nal_publica" in result.columns and "max_ant_nal_publica" in result.columns:
        varporc_min = (result[col_precio] / result["min_ant_nal_publica"] - 1) * 100
        varporc_max = (result[col_precio] / result["max_ant_nal_publica"] - 1) * 100
        result["Revisa"] = np.where(
            result[col_precio].isna(), "OK",
            np.where(
                (result["min_ant_nal_publica"].isna()) | (result["max_ant_nal_publica"].isna()), "OK",
                np.where(varporc_min < -10, "Revisar", np.where(varporc_max > 10, "Revisar", "OK")),
            ),
        )
    else:
        result["Revisa"] = "OK"

    # ---- Renombrar columnas para que coincidan con SAS ----
    rename_map = {
        col_div:      "CodigoMpio",
        col_pub:      "Nombre_Publica",
        col_precio:   f"Precio {mes_actual}",
        "CÓDIGO CPC": "Codigo CPC",
        # Agregaciones Nombre_Publica — nombres de display SAS
        "conteo_nal_publica":            "Conteo nacional publica",
        "prom_nal_publica":              "Prom nacional publica",
        "min_nal_publica":               "Mínimo nacional publica",
        "max_nal_publica":               "Máximo nacional publica",
        "prom_ant_nal_publica":          "Prom ant nacional publica",
        "min_ant_nal_publica":           "Mínimo ant nacional publica",
        "max_ant_nal_publica":           "Máximo ant nacional publica",
        "conteo_depto_publica":          "Conteo depto publica",
        "prom_depto_publica":            "Prom depto publica",
        "min_depto_publica":             "Mínimo depto publica",
        "max_depto_publica":             "Máximo depto publica",
        "conteo_mpio_publica":           "Conteo mpio publica",
        "prom_mpio_publica":             "Prom mpio publica",
        "min_mpio_publica":              "Mínimo mpio publica",
        "max_mpio_publica":              "Máximo mpio publica",
        "varporc_prom_nal_publica":      "Varporc prec prom nal publica",
        "varporc_min_nal_publica":       "Varporc prec mín nal publica",
        "varporc_max_nal_publica":       "Varporc prec máx nal publica",
        "varporc_prom_depto_publica":    "Varporc prec prom depto publica",
        "varporc_prom_mpio_publica":     "Varporc prec prom mpio publica",
        "varporc_min_depto_publica":     "Varporc prec mín depto publica",
        "varporc_max_depto_publica":     "Varporc prec máx depto publica",
        "varporc_min_mpio_publica":      "Varporc prec mín mpio publica",
        "varporc_max_mpio_publica":      "Varporc prec máx mpio publica",
        # Agregaciones ARTÍCULO — nombres de display SAS
        "conteo_nal_articulo":           "Conteo nacional artículo",
        "prom_nal_articulo":             "Prom nacional artículo",
        "min_nal_articulo":              "Mínimo nacional artículo",
        "max_nal_articulo":              "Máximo nacional artículo",
        "prom_ant_nal_articulo":         "Prom ant nacional artículo",
        "min_ant_nal_articulo":          "Mínimo ant nacional artículo",
        "max_ant_nal_articulo":          "Máximo ant nacional artículo",
        "conteo_depto_articulo":         "Conteo depto artículo",
        "prom_depto_articulo":           "Prom depto artículo",
        "min_depto_articulo":            "Mínimo depto artículo",
        "max_depto_articulo":            "Máximo depto artículo",
        "conteo_mpio_articulo":          "Conteo mpio artículo",
        "prom_mpio_articulo":            "Prom mpio artículo",
        "min_mpio_articulo":             "Mínimo mpio artículo",
        "max_mpio_articulo":             "Máximo mpio artículo",
        "varporc_prom_nal_articulo":     "Varporc prec prom nal art",
        "varporc_min_nal_articulo":      "Varporc prec Mín ant nal art",
        "varporc_max_nal_articulo":      "Varporc prec Máx ant nal art",
        "varporc_prom_depto_articulo":   "Varporc prec prom depto art",
        "varporc_prom_mpio_articulo":    "Varporc prec prom mpio art",
        "varporc_min_depto_articulo":    "Varporc prec mín depto art",
        "varporc_max_depto_articulo":    "Varporc prec máx depto art",
        "varporc_min_mpio_articulo":     "Varporc prec mín mpio art",
        "varporc_max_mpio_articulo":     "Varporc prec máx mpio art",
    }
    result = result.rename(columns={k: v for k, v in rename_map.items() if k and k in result.columns})

    with pd.ExcelWriter(str(ruta_tema), engine="openpyxl") as w:
        result.to_excel(w, sheet_name="REVISIÓN", index=False)

    log.info(
        "exportar_revision_tematica [%s] OK | filas=%d | revisiones=%s",
        modulo, len(result),
        int((result.get("Revisa", pd.Series()) == "Revisar").sum()),
    )
    return pd.DataFrame([{
        "modulo": modulo, "periodo": periodo,
        "filas": len(result),
        "para_revisar": int((result.get("Revisa", pd.Series()) == "Revisar").sum()),
    }])
