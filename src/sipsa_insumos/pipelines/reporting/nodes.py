"""Nodos del pipeline de reportes — SIPSA Insumos.

SAS equivalente (pasos 13-16):
  - BASES: filtrar por grupo → BASES_[MODULO]_[PERIODO].XLSX
  - TABLAS: crosstab Producto × Tendencia → TABLAS_[MODULO]_[PERIODO].xlsx
  - ANEXOS: artículo + presentación → ANEXO_[MODULO]_[PERIODO].XLSX
  - CUADROS: tabla de publicación final → CUADROS_[MODULO]_[PERIODO].xlsx

Todos los archivos son multi-hoja (una hoja por grupo).
Se escriben con pd.ExcelWriter(engine='openpyxl') directamente.
El nodo retorna un DataFrame de metadatos.
"""
from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

from sipsa_insumos.utils.excel_writer import aplicar_formato_numerico_precio, escribir_excel_multisheet

log = logging.getLogger(__name__)


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
    nombre = f"BASES_{modulo}_{periodo}.xlsx"
    ruta = Path(ruta_reporting) / modulo.lower() / nombre
    ruta.parent.mkdir(parents=True, exist_ok=True)

    # Detectar columnas de precio actual y anterior (generadas en calcular_variacion_tendencia)
    col_precio_actual = next(
        (c for c in base_comparada.columns if c.startswith("PRECIO_") and "Abril" not in c
         and c not in ("PRECIO_MIN", "PRECIO_MAX")), None
    )
    col_precio_anterior = next(
        (c for c in base_comparada.columns if "Abril" in c or
         (c.startswith("PRECIO_") and c != col_precio_actual
          and c not in ("PRECIO_MIN", "PRECIO_MAX"))), None
    )

    hojas: dict[str, pd.DataFrame] = {}
    for grupo in grupos:
        df_grupo = base_comparada[base_comparada["Grupo"] == grupo].copy()
        if df_grupo.empty:
            hojas[grupo] = pd.DataFrame()
            continue

        # Construir Mercado = "NombreMunicipio (NombreDepartamento)"
        if "NombreMunicipio" in df_grupo.columns and "NombreDepartamento" in df_grupo.columns:
            df_grupo["Mercado"] = (
                df_grupo["NombreMunicipio"] + " (" + df_grupo["NombreDepartamento"] + ")"
            )

        # Seleccionar y ordenar columnas equivalentes al SAS BASES
        col_map = {
            "CodigoDepto":        next((c for c in ["CodigoDepto"] if c in df_grupo.columns), None),
            "NombreDepartamento": "NombreDepartamento" if "NombreDepartamento" in df_grupo.columns else None,
            "CÓDIGO DIVIPOLA":    "CÓDIGO DIVIPOLA" if "CÓDIGO DIVIPOLA" in df_grupo.columns else None,
            "NombreMunicipio":    "NombreMunicipio" if "NombreMunicipio" in df_grupo.columns else None,
            "Mercado":            "Mercado" if "Mercado" in df_grupo.columns else None,
            "CÓDIGO CPC":         "CÓDIGO CPC" if "CÓDIGO CPC" in df_grupo.columns else None,
            "Nombre_Publica":     "Nombre_Publica" if "Nombre_Publica" in df_grupo.columns else None,
            "PRECIO_ANTERIOR":    col_precio_anterior,
            "PRECIO_ACTUAL":      col_precio_actual,
            "VARIACION(%)":       "VARIACION" if "VARIACION" in df_grupo.columns else None,
            "Cuenta":             "N_ARTICULOS" if "N_ARTICULOS" in df_grupo.columns else None,
            "Tendencia":          "TENDENCIA" if "TENDENCIA" in df_grupo.columns else None,
        }
        cols_sel = [v for v in col_map.values() if v and v in df_grupo.columns]
        rename_inv = {v: k for k, v in col_map.items() if v and v in df_grupo.columns}
        df_out = df_grupo[cols_sel].rename(columns=rename_inv)
        df_out = df_out.sort_values(
            [c for c in ["NombreDepartamento", "NombreMunicipio", "Nombre_Publica"] if c in df_out.columns]
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
    nombre = f"TABLAS_{modulo}_{periodo}.xlsx"
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
    nombre = f"ANEXO_{modulo}_{periodo}.xlsx"
    ruta = Path(ruta_reporting) / modulo.lower() / nombre
    ruta.parent.mkdir(parents=True, exist_ok=True)

    col_precio_actual = next(
        (c for c in base_comparada.columns if c.startswith("PRECIO_") and "Abril" not in c
         and c not in ("PRECIO_MIN", "PRECIO_MAX")), None
    )
    col_precio_anterior = next(
        (c for c in base_comparada.columns if "Abril" in c or
         (c.startswith("PRECIO_") and c != col_precio_actual
          and c not in ("PRECIO_MIN", "PRECIO_MAX"))), None
    )

    hojas: dict[str, pd.DataFrame] = {}
    for grupo in grupos:
        df_grupo = base_comparada[base_comparada["Grupo"] == grupo].copy()
        if df_grupo.empty:
            hojas[grupo] = pd.DataFrame()
            continue

        # Mercado = "NombreMunicipio (NombreDepartamento)"
        if "NombreMunicipio" in df_grupo.columns and "NombreDepartamento" in df_grupo.columns:
            df_grupo["Mercado"] = (
                df_grupo["NombreMunicipio"] + " (" + df_grupo["NombreDepartamento"] + ")"
            )

        # Desglosar Nombre_Publica en Nombre_insumo + Presentación_insumo
        # Formato SAS: "Artículo, Presentación" → separar por la última coma
        if "Nombre_Publica" in df_grupo.columns:
            partes = df_grupo["Nombre_Publica"].str.rsplit(",", n=1, expand=True)
            df_grupo["Nombre_insumo"] = partes[0].str.strip()
            df_grupo["Presentación_insumo"] = partes[1].str.strip() if 1 in partes.columns else ""

        col_map = {
            "CodigoDepto":         next((c for c in ["CodigoDepto"] if c in df_grupo.columns), None),
            "NombreDepartamento":  "NombreDepartamento" if "NombreDepartamento" in df_grupo.columns else None,
            "CÓDIGO DIVIPOLA":     "CÓDIGO DIVIPOLA" if "CÓDIGO DIVIPOLA" in df_grupo.columns else None,
            "NombreMunicipio":     "NombreMunicipio" if "NombreMunicipio" in df_grupo.columns else None,
            "CÓDIGO CPC":          "CÓDIGO CPC" if "CÓDIGO CPC" in df_grupo.columns else None,
            "Nombre_insumo":       "Nombre_insumo" if "Nombre_insumo" in df_grupo.columns else None,
            "Presentación_insumo": "Presentación_insumo" if "Presentación_insumo" in df_grupo.columns else None,
            "PRECIO_ANTERIOR":     col_precio_anterior,
            "PRECIO_ACTUAL":       col_precio_actual,
            "VARIACION(%)":        "VARIACION" if "VARIACION" in df_grupo.columns else None,
            "Cuenta":              "N_ARTICULOS" if "N_ARTICULOS" in df_grupo.columns else None,
            "Tendencia":           "TENDENCIA" if "TENDENCIA" in df_grupo.columns else None,
            "Mercado":             "Mercado" if "Mercado" in df_grupo.columns else None,
            "Nombre_Publica":      "Nombre_Publica" if "Nombre_Publica" in df_grupo.columns else None,
        }
        cols_sel = [v for v in col_map.values() if v and v in df_grupo.columns]
        rename_inv = {v: k for k, v in col_map.items() if v and v in df_grupo.columns}
        df_out = df_grupo[cols_sel].rename(columns=rename_inv)
        df_out = df_out.sort_values(
            [c for c in ["NombreDepartamento", "NombreMunicipio", "Nombre_Publica"] if c in df_out.columns]
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
    nombre = f"CUADROS_{modulo}_{periodo}.xlsx"
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
