"""Nodos del pipeline Serie Departamental — SIPSA Insumos.

Réplica de los programas SAS "NNSerie_Departamental_<Módulo>_<MES>.sas"
(OneDrive SIPSA Insumos/03SERIE_DEPTAL/01Programas). Cada módulo sigue el
mismo flujo, con variantes que se declaran en `_MODULOS`:

  1. PROC IMPORT de la base liviana y de la DIVIPOLA (Dep_Mun); CodigoMpio (Z5.),
     CodigoDepto y la llave de publicación (CATX('_', ...)).
  2. MERGE con la DIVIPOLA (IF B) y con las hojas Grupo/Articulo del archivo
     "divipola <módulo>" → FALTAN_*.
  3. Precios 0 → faltante, RENAME a los meses, RETAIN del orden de columnas y
     PROC SORT NODUPKEY → *_DUPLICADOS.
  4. PROC SQL por departamento: N_FUENTE/N_INFORMANTE, N_ARTICULOS y
     PRECIO_PROMEDIO → *_MAYORESQUE2 (N>=2) y *_MENORESQUE2 (N<2).
  5. MERGE con el MAYORESQUE2 departamental del período anterior, Variacion(%),
     Cuenta y Tendencia → ANEXO.

Se emulan las reglas de SAS que cambian el resultado: tipado de PROC IMPORT,
orden de variables del PDV, PROC SORT estable, MERGE con BY (incluidos los
casos muchos-a-muchos) y el "remerge" de PROC SQL.
"""
from __future__ import annotations

import datetime
import logging
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from python_calamine import CalamineWorkbook

from sipsa_insumos.pipelines.reporting.nodes import _periodo_anterior_modulo

log = logging.getLogger(__name__)

_MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
          "Septiembre", "Octubre", "Noviembre", "Diciembre"]
_ABR = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
_DROP_BASE = ["Hist", "Modi", "Apro", "Sup", "Camb", "Deta"]

# Llaves y columnas repetidas entre módulos
_LLAVE_STD = ["CodigoMpio", "Fuente", "Codigo CPC", "Articulo", "CasaCom.", "RegICA", "UnMed."]
_LLAVE_CARACTE = ["CodigoMpio", "Informante", "Codigo CPC", "Articulo", "Articulo_Caracte"]
_ANEXO_ICA = ["CodigoDepto", "NombreDepartamento", "Codigo CPC", "Nombre_insumo", "NOM_ARTICULO",
              "NOM_CASA COMERCIAL", "REGISTRO_ICA", "Presentación_insumo", "@ANT", "@ACT",
              "Variacion(%)", "Cuenta", "Tendencia", "@PUB"]
_ANEXO_SIN_ICA = ["CodigoDepto", "NombreDepartamento", "Codigo CPC", "Nombre_insumo", "NOM_ARTICULO",
                  "Presentación_insumo", "@ANT", "@ACT", "Variacion(%)", "Cuenta", "Tendencia", "@PUB"]

# Especificación de cada programa SAS. Nombres de variable como quedan en SAS
# (con "#" en lugar de "." cuando el programa importa con dbms=excel).
# "@PUB", "@ANT", "@ACT", "@N" se reemplazan por el nombre de publicación, el
# mes anterior/actual y la columna N_FUENTE/N_INFORMANTE del período.
_MODULOS: dict[str, dict] = {
    "AGRICOLAS": dict(
        carpeta="Ins_Agrícolas", clave="agrícolas", motor="xlsx", mes_minuscula=False,
        llave="Art_Unmed_Casacomer_ICA", llave_cols=["Articulo", "UnMed.", "CasaCom.", "RegICA"], largo_llave=118,
        nuevas=["Art_Unmed_Casacomer_ICA", "CodigoMpio", "CodigoDepto"],
        hoja_grupo="Grupo", hoja_articulo="Articulo",
        col_publica=("Nombre_productos_agrícolas_publicación", "Nombre_productos_agrícolas_publ"),
        faltan=("FALTAN_GRUPO_AGRICOLA_{M}.XLSX", "FALTAN_GRUPO_AGRICOLA"),
        precios=("Precio Ante.", "Precio Actual", "Var."),
        retain=["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Articulo",
                "Codigo CPC", "@PUB", "Grupo", "UnMed.", "CasaCom.", "RegICA", "@MANT", "@MACT", "VAR", "Nov#",
                "Estado", "Art_Unmed_Casacomer_ICA", "Observación"],
        llave_dupli=_LLAVE_STD, dupli="AGRÍCOLAS_DUPLICADOS_{M}.XLSX",
        fuente="Fuente", conteo="fuente", n_col="N_FUENTE", grupo=True,
        grp=["CasaCom.", "RegICA", "Articulo"],
        salida_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "CasaCom.", "RegICA", "Articulo",
                     "Grupo", "@N", "N_ARTICULOS", "PRECIO_PROMEDIO"],
        mayor=("INSU_AGRIC_MAYORESQUE2_DEPTO_{M}.XLSX", "INSU_AGRICOLAS_MAYORESOIGUALES2"),
        menor=("INSU_AGRIC_MENORESQUE2_DEPTO_{M}.XLSX", "INSU_AGRICOLAS_MENORESQUE2"),
        ant_num=["Codigo CPC"], llave_ant=["CodigoDepto", "@PUB", "Grupo", "CasaCom.", "RegICA", "Articulo"],
        anexo="ANEXO INSUMOS AGRICOLAS DEPTO {M}.XLSX", anexo_cols=_ANEXO_ICA, estilo_nombre="coma",
        hojas_anexo=["COADYUDANTES", "FERTILIZANTES", "FUNGICIDAS", "HERBICIDAS", "INSECTICIDAS", "BIOINSUMOS"],
    ),
    "PECUARIOS": dict(
        carpeta="Ins_Pecuarios", clave="pecuarios", motor="xlsx", mes_minuscula=True,
        llave="Art_Unmed_Casacomer_ICA", llave_cols=["Articulo", "UnMed.", "CasaCom.", "RegICA"], largo_llave=118,
        nuevas=["Art_Unmed_Casacomer_ICA", "CodigoMpio", "CodigoDepto"],
        hoja_grupo="Grupo", hoja_articulo="Articulo",
        col_publica=("Nombre_productos_pecuarios_publicación", "Nombre_productos_pecuarios_publi"),
        faltan=("FALTAN_GRUPO_PECUARIO_{M}.XLSX", "FALTAN_GRUPO_PECUARIO"),
        precios=("Precio Ante.", "Precio Actual", "Var."),
        retain=["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC",
                "Articulo", "@PUB", "Grupo", "UnMed.", "CasaCom.", "RegICA", "@MANT", "@MACT", "VAR", "Nov.",
                "Estado", "Art_Unmed_Casacomer_ICA", "Observación"],
        llave_dupli=_LLAVE_STD, dupli="PECUARIOS_DUPLICADOS_{M}.XLSX",
        fuente="Fuente", conteo="fuente", n_col="N_FUENTE", grupo=True,
        grp=["CasaCom.", "RegICA", "Articulo"],
        salida_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "CasaCom.", "RegICA", "Articulo",
                     "Grupo", "@N", "N_ARTICULOS", "PRECIO_PROMEDIO"],
        mayor=("INSU_PECUA_MAYORESQUE2_DEPTO_{M}.XLSX", "INSU_PECUARIOS_MAYORESOIGUALES2"),
        menor=("INSU_PECUA_MENORESQUE2_DEPTO_{M}.XLSX", "INSU_PECUARIOS_MENORESQUE2"),
        ant_num=["Codigo CPC", "RegICA"], llave_ant=["CodigoDepto", "@PUB", "Grupo", "CasaCom.", "RegICA", "Articulo"],
        anexo="ANEXO INSUMOS PECUARIOS DEPTO {M}.xlsx", anexo_cols=_ANEXO_ICA, estilo_nombre="coma",
        hojas_anexo=["ALIMENTOS", "ANTIBIOTICOS", "ANTISEPTICOS", "HORMONALES", "INSECTICIDAS", "MEDICAMENTOS",
                     "VITAMINAS"],
    ),
    "ELEMENTOS": dict(
        carpeta="Elementos", clave="elementos", motor="xlsx", mes_minuscula=False,
        llave="Art_Casacomer_ICA_Unmed", llave_cols=["Articulo", "CasaCom.", "RegICA", "UnMed."], largo_llave=453,
        nuevas=["Art_Casacomer_ICA_Unmed", "CodigoMpio", "CodigoDepto"],
        hoja_grupo=None, hoja_articulo="Artículo",
        col_publica=("Nombre productos publicación", "Nombre_productos_elementos_publi"),
        faltan=("FALTAN_PUBLICA_{M}.XLSX", "FALTAN_PUBLICA"),
        precios=("Precio Ante.", "Precio Actual", "Var."),
        retain=["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Articulo",
                "Codigo CPC", "@PUB", "@MANT", "@MACT", "VAR", "Observación", "UnMed.", "CasaCom.", "RegICA",
                "Nov#", "Estado", "Art_Casacomer_ICA_Unmed"],
        llave_dupli=_LLAVE_STD, dupli="ELEMENTOS_DUPLICADOS_{M}.XLSX",
        fuente="Fuente", conteo="fuente", n_col="N_FUENTE", grupo=False,
        grp=["CasaCom.", "RegICA", "Articulo"],
        salida_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "CasaCom.", "RegICA", "Articulo",
                     "@N", "N_ARTICULOS", "PRECIO_PROMEDIO"],
        mayor=("ELEM_AGROPE_MAYORESQUE2_DEPTO_{M}.XLSX", "ELEM_AGROPE_MAYORESOIGUALES2"),
        menor=("ELEM_AGROPE_MENORESQUE2_DEPTO_{M}.XLSX", "ELEM_AGROPE_MENORESQUE2"),
        ant_num=["Codigo CPC"], llave_ant=["CodigoDepto", "@PUB", "CasaCom.", "RegICA", "Articulo"],
        anexo="ANEXO ELEMENTOS AGROPECUARIOS DEPTO {M}.XLSX", anexo_cols=_ANEXO_ICA, estilo_nombre="coma",
        hojas_anexo=["Hoja1"],
    ),
    "EMPAQUES": dict(
        carpeta="Empaques", clave="empaques", motor="xlsx", mes_minuscula=False,
        llave="Articulo_Caracte", llave_cols=["Articulo", "Caracte."], largo_llave=200,
        nuevas=["CodigoMpio", "CodigoDepto", "Articulo_Caracte"],
        hoja_grupo=None, hoja_articulo="Articulo",
        col_publica=("Nombre productos empaques publi", "Nombre_productos_empaques_publi"),
        faltan=("FALTAN_PUBLICA_{M}.XLSX", "FALTAN_PUBLICA"),
        precios=("Precio Ante.", "Precio Actual", "Var."),
        retain=["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC",
                "Articulo", "@PUB", "@MANT", "@MACT", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte"],
        llave_dupli=["CodigoMpio", "Informante", "Articulo", "Articulo_Caracte"],
        dupli="EMPAQUES_DUPLICADOS_{M}.XLSX",
        fuente="Informante", conteo="remerge", n_col="N_INFORMANTE", grupo=False, grp=["Articulo"],
        salida_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "Articulo", "@N", "N_ARTICULOS",
                     "PRECIO_PROMEDIO"],
        mayor=("EMPA_AGROPE_MAYORESQUE2_DEPTO_{M}.XLSX", "EMPA_AGROPE_MAYORESOIGUALES2"),
        menor=("EMPA_AGROPE_MENORESQUE2_DEPTO_{M}.XLSX", "EMPA_AGROPE_MENORESQUE2"),
        ant_num=["Codigo CPC"], llave_ant=["CodigoDepto", "@PUB", "Articulo"],
        anexo="ANEXO EMPAQUES AGROPECUARIOS DEPTO {M}.XLSX", anexo_cols=_ANEXO_SIN_ICA, estilo_nombre="coma",
        hojas_anexo=["Hoja1"],
    ),
    "PROPAGACION": dict(
        carpeta="Material_Propaga", clave="propagación", motor="excel", mes_minuscula=False,
        llave="Articulo_Unmed", llave_cols=["Articulo", "UnMed#"], largo_llave=240,
        nuevas=["Articulo_Unmed", "CodigoMpio", "CodigoDepto", "Art_Casacomer_ICA_Unmed"],
        extra_llaves={"Art_Casacomer_ICA_Unmed": ["Articulo", "CasaCom#", "RegICA", "UnMed#"]},
        hoja_grupo=None, hoja_articulo="Articulo",
        col_publica=("Articulo publicación", "Nombre_productos_material_publi"),
        faltan=("FALTAN_GRUPO_MATPROPAGA_{M}.XLSX", "FALTAN_GRUPO_MATPROPAGA"),
        precios=("Precio Ante#", "Precio Actual", "Var#"), ante_errado=True,
        retain=["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC",
                "Articulo", "@PUB", "@MANT", "@MACT", "VAR", "Observación", "UnMed#", "CasaCom#", "RegICA", "Nov#",
                "Estado", "Articulo_Unmed"],
        llave_dupli=["CodigoMpio", "Fuente", "Codigo CPC", "Articulo", "CasaCom#", "RegICA", "UnMed#"],
        dupli="MATPROPAGA_DUPLICADOS_{M}.XLSX",
        fuente="Fuente", conteo="fuente", n_col="N_FUENTE", grupo=False, grp=["CasaCom#", "Articulo"],
        salida_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "Articulo", "CasaCom#", "@N",
                     "N_ARTICULOS", "PRECIO_PROMEDIO"],
        arroz=True,
        mayor=("MATE_PROPAGA_MAYORESQUE2_{M}.XLSX", "MATE_PROPAGA_MAYORESOIGUALES2"),
        menor=("MATE_PROPAGA_MENORESQUE2_{M}.XLSX", "MATE_PROPAGA_MENORESQUE2"),
        ant_num=[], llave_ant=["CodigoDepto", "@PUB", "CasaCom#", "Articulo"],
        anexo="ANEXO MATERIAL PROPAGACION {M}.xlsx",
        anexo_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "Nombre_insumo", "NOM_ARTICULO",
                    "NOM_CASA COMERCIAL", "Presentación_insumo", "@ANT", "@ACT", "Variacion(%)", "Cuenta",
                    "Tendencia", "@PUB"],
        estilo_nombre="coma_espacio", hojas_anexo=["Material"],
    ),
    "ARRIENDOS": dict(
        carpeta="Arriendos", clave="arriendos", motor="excel", mes_minuscula=False,
        llave="Articulo_Caracte", llave_cols=["Articulo", "Caracte#"], largo_llave=200,
        nuevas=["CodigoMpio", "CodigoDepto", "Articulo_Caracte"], llave_al_inicio=True,
        hoja_grupo=None, hoja_articulo="Articulo",
        col_publica=("articulo publicacion", "Nombre_productos_arriendos_publi"),
        faltan=("FALTAN_PUBLICA_{M}.XLSX", "FALTAN_PUBLICA"),
        precios=("Precio Ante#", "Precio Actual", "Var#"), ante_ia_in=True,
        retain=["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Articulo",
                "Codigo CPC", "@PUB", "@MANT", "@MACT", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte"],
        llave_dupli=_LLAVE_CARACTE, dupli="ARRIENDOS_DUPLICADOS_{M}.XLSX",
        fuente="Informante", conteo="fuente", n_col="N_INFORMANTE", grupo=False, grp=["Articulo"],
        salida_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "Articulo", "@N", "N_ARTICULOS",
                     "PRECIO_PROMEDIO"],
        mayor=("ARRIENDOS_MAYORESQUE2_{M}.XLSX", "ARRIENDOS_MAYORESOIGUALES2"),
        menor=("ARRIENDOS_MENORESQUE2_{M}.XLSX", "ARRIENDOS_MENORESQUE2"),
        ant_num=[], llave_ant=["CodigoDepto", "@PUB", "Articulo"],
        anexo="ANEXO ARRIENDOS {M}.xlsx",
        anexo_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "NOM_ARTICULO", "@ANT", "@ACT",
                    "Variacion(%)", "Cuenta", "Tendencia", "Nombre_insumo", "Presentación_insumo"],
        estilo_nombre="coma", hojas_anexo=["Arriendos"],
    ),
    "SERVICIOS": dict(
        carpeta="Servicios", clave="servicios", motor="excel", mes_minuscula=False,
        llave="Articulo_Caracte", llave_cols=["Articulo", "Caracte#"], largo_llave=200,
        nuevas=["CodigoMpio", "CodigoDepto", "Articulo_Caracte"],
        hoja_grupo=None, hoja_articulo="Articulo",
        col_publica=("articulo publicacion", "Nombre_productos_servicios_publi"),
        faltan=("FALTAN_GRUPO_SERVICIOS_{M}.XLSX", "FALTAN_GRUPO_SERVICIOS"),
        precios=("Precio Ante#", "Precio Actual", "Var#"), ante_ia_in=True,
        retain=["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC",
                "Articulo", "@PUB", "@MANT", "@MACT", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte"],
        llave_dupli=_LLAVE_CARACTE, dupli="SERVICIOS_DUPLICADOS_{M}.XLSX",
        fuente="Informante", conteo="fuente", n_col="N_INFORMANTE", grupo=False, grp=["Articulo"],
        salida_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "Articulo", "@N", "N_ARTICULOS",
                     "PRECIO_PROMEDIO"],
        mayor=("SERVICIOS_MAYORESQUE2_{M}.XLSX", "SERVICIOS_MAYORESOIGUALES2"),
        menor=("SERVICIOS_MENORESQUE2_{M}.XLSX", "SERVICIOS_MENORESQUE2"),
        # El programa cruza con el mes anterior solo por departamento y nombre (sin Articulo).
        ant_num=[], llave_ant=["CodigoDepto", "@PUB"],
        anexo="ANEXO SERVICIOS {M}.xlsx", anexo_cols=_ANEXO_SIN_ICA, estilo_nombre="coma",
        hojas_anexo=["Servicios"],
    ),
    "JORNALES": dict(
        carpeta="Jornales", clave="jornales", motor="excel", mes_minuscula=False,
        llave="Articulo_Caracte", llave_cols=["Articulo", "Caracte#"], largo_llave=264,
        nuevas=["Articulo_Caracte", "CodigoMpio", "CodigoDepto"],
        hoja_grupo=None, hoja_articulo="Articulo",
        col_publica=("articulo publicacion", "Nombre_productos_jornales_publi"),
        faltan=("FALTAN_GRUPO_JORNALES_{M}.XLSX", "FALTAN_GRUPO_JORNALES"),
        precios=("Precio Ante#", "Precio Actual", "Var#"), ante_ia_in=True,
        retain=["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC",
                "Articulo", "@PUB", "@MANT", "@MACT", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte"],
        llave_dupli=_LLAVE_CARACTE, dupli="JORNALES_DUPLICADOS_{M}.XLSX",
        # GROUP BY ...,Informante toma la columna de la base (sin municipio), no New_Fuente.
        fuente="Informante", conteo="informante_crudo", n_col="N_INFORMANTE", grupo=False, grp=["Articulo"],
        salida_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "Articulo", "@N", "N_ARTICULOS",
                     "PRECIO_PROMEDIO"],
        mayor=("JORNALES_MAYORESQUE2_{M}.XLSX", "JORNALES_MAYORESOIGUALES2"),
        menor=("JORNALES_MENORESQUE2_{M}.XLSX", "JORNALES_MENORESQUE2"),
        ant_num=["Codigo CPC"], llave_ant=["CodigoDepto", "@PUB", "Articulo"],
        anexo="ANEXO JORNALES {M}.xlsx",
        anexo_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "@ANT", "@ACT", "Variacion(%)",
                    "Cuenta", "Tendencia", "Nombre_insumo", "NOM_ARTICULO", "Presentación_insumo"],
        estilo_nombre="jornales", hojas_anexo=["Jornales"],
    ),
    "ESPECIES": dict(
        carpeta="Especies", clave="especie", motor="excel", mes_minuscula=False,
        llave="Articulo_Caracte", llave_cols=["Articulo", "Caracte#"], largo_llave=561,
        nuevas=["Articulo_Caracte", "CodigoMpio", "CodigoDepto"],
        hoja_grupo=None, hoja_articulo="Articulo",
        col_publica=("articulo publicacion", "Nombre_productos_especie_publi"),
        faltan=("FALTAN_GRUPO_ESPECIES_{M}.XLSX", "FALTAN_GRUPO_ESPECIES"),
        precios=("Precio Ante#", "Precio Actual", "Var#"), ante_ia_in=True,
        retain=["CodigoDepto", "CodigoMpio", "NombreDepartamento", "NombreMunicipio", "Fuente", "Codigo CPC",
                "Articulo", "@PUB", "@MANT", "@MACT", "VAR", "Observación", "Nov#", "Estado", "Articulo_Caracte"],
        llave_dupli=_LLAVE_CARACTE, dupli="ESPECIES_DUPLICADOS_{M}.XLSX",
        fuente="Informante", conteo="fuente", n_col="N_INFORMANTE", grupo=False, grp=["Articulo"],
        salida_cols=["CodigoDepto", "NombreDepartamento", "Codigo CPC", "@PUB", "Articulo", "@N", "N_ARTICULOS",
                     "PRECIO_PROMEDIO"],
        mayor=("ESPE_PRODUC_MAYORESQUE2_{M}.XLSX", "ESPE_PRODUC_MAYORESOIGUALES2"),
        menor=("ESPE_PRODUC_MENORESQUE2_{M}.XLSX", "ESPE_PRODUC_MENORESQUE2"),
        ant_num=[], llave_ant=["CodigoDepto", "@PUB", "Articulo"],
        anexo="ANEXO ESPECIES PRODUCTIVAS {M}.xlsx", anexo_cols=_ANEXO_SIN_ICA, estilo_nombre="especies",
        hojas_anexo=["Especies"],
    ),
}


# ── Lectura con el tipado de PROC IMPORT ─────────────────────────────────────

def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", str(s))


def _num_texto(v: float) -> str:
    """Número como texto, como lo deja SAS (BEST.) al convertirlo a carácter."""
    f = float(v)
    return str(int(f)) if f.is_integer() else repr(f)


def _es_num(v) -> bool:
    return isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool)


def _serial_excel(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        d = v if isinstance(v, datetime.datetime) else datetime.datetime(v.year, v.month, v.day)
        return (d - datetime.datetime(1899, 12, 30)).total_seconds() / 86400
    return v


def _leer_hoja(ruta: str | Path, hoja: str | None, motor: str) -> pd.DataFrame:
    """Lee una hoja como PROC IMPORT: una columna es numérica si todas sus celdas
    con valor son números; si no, es de caracteres (los números pasan a texto).
    Con dbms=excel los puntos de los nombres de columna quedan como "#".
    Faltantes: NaN en numéricas y "" en caracteres (sin espacios finales)."""
    wb = CalamineWorkbook.from_path(str(ruta))
    nombres = wb.sheet_names
    if hoja is None:
        nombre = nombres[0]
    else:
        nombre = next((n for n in nombres if _nfc(n).lower() == _nfc(hoja).lower()), None)
        if nombre is None:
            raise ValueError(f"No existe la hoja '{hoja}' en {ruta} (hojas: {nombres})")
    filas = wb.get_sheet_by_name(nombre).to_python()
    if not filas:
        return pd.DataFrame()
    cab = [_nfc(c).strip() for c in filas[0]]
    if motor == "excel":
        cab = [c.replace(".", "#") for c in cab]
    ncol = len(cab)
    # Una celda con formato de fecha llega a SAS como su número de serie de Excel.
    cuerpo = [[_serial_excel(v) for v in f] + [""] * (ncol - len(f)) for f in filas[1:]]
    datos = {}
    for j, c in enumerate(cab):
        if not c:
            continue
        vals = [None if (f[j] == "" or f[j] is None) else f[j] for f in cuerpo]
        if any(v is not None for v in vals) and all(v is None or _es_num(v) for v in vals):
            datos[c] = pd.Series([np.nan if v is None else float(v) for v in vals], dtype="float64")
        else:
            datos[c] = pd.Series(
                ["" if v is None else (_num_texto(v) if _es_num(v) else _nfc(v).rstrip()) for v in vals],
                dtype=object)
    return pd.DataFrame(datos)


def _es_numerica(s: pd.Series) -> bool:
    return pd.api.types.is_numeric_dtype(s)


def _catx(df: pd.DataFrame, cols: list[str], largo: int) -> pd.Series:
    """CATX('_', ...) truncado a la longitud de la variable SAS."""
    partes = []
    for c in cols:
        s = df[c]
        if _es_numerica(s):
            partes.append(s.map(lambda v: "" if pd.isna(v) else _num_texto(v)))
        else:
            partes.append(s.astype(str).str.strip())
    return pd.Series(["_".join(p for p in vals if p)[:largo] for vals in zip(*partes)], index=df.index, dtype=object)


# ── Emulación de SAS: PROC SORT, MERGE y PROC SQL ───────────────────────────

def _clave_orden(s: pd.Series) -> pd.Series:
    """Valor de orden SAS: faltante numérico primero; caracteres por byte (UTF-8)."""
    if _es_numerica(s):
        return s.fillna(-np.inf)
    return s.astype(str)


def _ordenar(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    """PROC SORT BY cols (estable)."""
    if df.empty:
        return df.reset_index(drop=True)
    tmp = pd.DataFrame({f"_k{i}": _clave_orden(df[c]).values for i, c in enumerate(cols)})
    orden = tmp.sort_values(list(tmp.columns), kind="mergesort").index
    return df.iloc[orden].reset_index(drop=True)


def _nodupkey(df: pd.DataFrame, cols: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """PROC SORT NODUPKEY OUT= DUPOUT=: primer registro de cada llave y el resto."""
    ord_ = _ordenar(df, cols)
    dup = pd.DataFrame({c: _clave_orden(ord_[c]) for c in cols}).duplicated()
    return ord_[~dup.values].reset_index(drop=True), ord_[dup.values].reset_index(drop=True)


def _faltante(s: pd.Series):
    return np.nan if _es_numerica(s) else ""


def _merge_sas(a: pd.DataFrame, b: pd.DataFrame, by: list[str], salida: str) -> tuple[pd.DataFrame, pd.Series, pd.Series]:
    """DATA step MERGE a b; BY by; — ambos ya ordenados por `by`.

    Replica el PDV: dentro de cada grupo BY se lee la i-ésima observación de
    cada tabla; la que se agota conserva sus últimos valores y, en variables
    comunes, gana la tabla que se lee después (b). Devuelve (tabla, in_a, in_b).
    `salida`: "a" (IF A), "b" (IF B) o "todo".
    """
    cols = list(a.columns) + [c for c in b.columns if c not in a.columns]
    tipos = {c: (a[c] if c in a.columns else b[c]) for c in cols}
    faltantes = {c: _faltante(tipos[c]) for c in cols}

    def llaves(df):
        k = [tuple(v) for v in zip(*[_clave_orden(df[c]).tolist() for c in by])] if len(df) else []
        grupos: dict[tuple, list[int]] = {}
        for i, t in enumerate(k):
            grupos.setdefault(t, []).append(i)
        return grupos

    ga, gb = llaves(a), llaves(b)
    orden_llaves = sorted(set(ga) | set(gb))
    ra, rb = a.to_dict("records"), b.to_dict("records")
    filas, fa, fb = [], [], []
    for k in orden_llaves:
        ia, ib = ga.get(k, []), gb.get(k, [])
        pdv = dict(faltantes)
        for i in range(max(len(ia), len(ib))):
            if i < len(ia):
                pdv.update(ra[ia[i]])
            if i < len(ib):
                pdv.update(rb[ib[i]])
            if (salida == "a" and not ia) or (salida == "b" and not ib):
                continue
            filas.append(dict(pdv))
            fa.append(bool(ia))
            fb.append(bool(ib))
    res = pd.DataFrame(filas, columns=cols)
    for c in cols:
        if _es_numerica(tipos[c]):
            res[c] = pd.to_numeric(res[c], errors="coerce").astype("float64")
        else:
            res[c] = res[c].astype(object)
    return res, pd.Series(fa, dtype=bool), pd.Series(fb, dtype=bool)


def _sql_max(s: pd.Series):
    if _es_numerica(s):
        return s.max()
    v = s[s != ""]
    return v.max() if len(v) else ""


def _contar(s: pd.Series) -> int:
    """COUNT(col): valores no faltantes."""
    return int(s.notna().sum()) if _es_numerica(s) else int((s != "").sum())


# ── Nombre_insumo / Presentación_insumo (SCAN de SAS) ────────────────────────

def _palabras(texto: str) -> list[str]:
    """Palabras de SCAN(texto, n, ','): delimitadores seguidos cuentan como uno."""
    return [p for p in str(texto).split(",") if p != ""]


def _scan(texto: str, n: int) -> str:
    p = _palabras(texto)
    if n < 0:
        n = len(p) + n + 1
    return p[n - 1].strip() if 1 <= n <= len(p) else ""


def _nombre_presentacion(nombre: str, estilo: str) -> tuple[str, str]:
    s1, s2, s3 = _scan(nombre, 1), _scan(nombre, 2), _scan(nombre, 3)
    if estilo == "jornales":
        nom = s1
        pres = (_scan(nombre, -2) + "," + _scan(nombre, -1)) if str(nombre).count(",") == 2 else _scan(nombre, -1)
        return nom, pres
    if estilo == "especies":
        nom = s1 if s3 == "" else s1 + "," + s2 + "," + s3
    else:
        sep = ", " if estilo == "coma_espacio" else ","
        nom = s1 if s3 == "" else s1 + sep + s2
    return nom, _scan(nombre, -1)


# ── Escritura ────────────────────────────────────────────────────────────────

def _celda(v):
    if v is None:
        return None
    if isinstance(v, str):
        return v if v != "" else None
    if _es_num(v):
        return None if pd.isna(v) else float(v)
    return v


def _escribir(ruta: Path, hojas: list[tuple[str, pd.DataFrame]]) -> None:
    """PROC EXPORT DBMS=XLSX: encabezado + filas; faltantes como celda vacía."""
    wb = Workbook(write_only=True)
    for nombre, df in hojas:
        ws = wb.create_sheet(title=nombre[:31])
        ws.append(list(df.columns))
        for fila in df.itertuples(index=False, name=None):
            ws.append([_celda(v) for v in fila])
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(ruta))


# ── Utilidades de período y archivos ─────────────────────────────────────────

def _buscar(carpeta: Path, clave: str, prefijo: str = "") -> Path:
    clave = _nfc(clave).lower()
    cands = [f for f in carpeta.glob("*.xlsx")
             if clave in _nfc(f.name).lower() and _nfc(f.name).lower().startswith(prefijo)
             and not f.name.startswith("~$")]
    if len(cands) != 1:
        raise FileNotFoundError(f"Se esperaba un archivo con '{clave}' en {carpeta}; encontrados: {cands}")
    return cands[0]


def _buscar_nombre(carpetas: list[Path], nombre: str) -> Path | None:
    objetivo = _nfc(nombre).lower()
    for c in carpetas:
        if c.exists():
            for f in c.iterdir():
                if _nfc(f.name).lower() == objetivo:
                    return f
    return None


def _mes_num(periodo: str) -> int:
    return _ABR.index(periodo[:3].upper()) + 1


# ── Nodo principal ───────────────────────────────────────────────────────────

def generar_serie_deptal(
    periodo: str,
    modulo: str,
    ruta_reporting: str,
    ruta_raw: str | None = None,
    carpetas_mes_anterior: list[str] | None = None,
) -> pd.DataFrame:
    """Genera los 5 archivos de la serie departamental de un módulo, idénticos a SAS.

    Args:
        periodo:        Período MMMAAAA (ej: 'AGO2026').
        modulo:         Módulo (AGRICOLAS, PECUARIOS, ELEMENTOS, EMPAQUES, PROPAGACION,
                        ARRIENDOS, SERVICIOS, JORNALES, ESPECIES).
        ruta_reporting: data/08_reporting/<periodo>; la salida queda en
                        <ruta_reporting>/serie_deptal/<carpeta SAS>/.
        ruta_raw:       Carpeta de insumos del período (por defecto data/01_raw/<periodo>).
        carpetas_mes_anterior: Carpetas donde buscar el MAYORESQUE2 departamental del
                        período anterior. Por defecto, la salida Python del período
                        anterior y data/01_raw/<periodo>/SERIE_DEPTAL <periodo>/.

    Returns:
        DataFrame de metadatos con los archivos y el número de filas.
    """
    m = modulo.upper()
    cfg = _MODULOS[m]
    periodo = periodo.upper()
    periodo_ant = _periodo_anterior_modulo(periodo, m)
    mes = periodo.lower() if cfg["mes_minuscula"] else periodo
    mes_ant_arch = periodo_ant.lower() if cfg["mes_minuscula"] else periodo_ant
    m_act, m_ant = _MESES[_mes_num(periodo) - 1], _MESES[_mes_num(periodo_ant) - 1]
    pub = cfg["col_publica"][1]
    n_act, n_art, p_act = f"{cfg['n_col']}_{mes}", f"N_ARTICULOS_{mes}", f"PRECIO_PROMEDIO_{mes}"

    rep = Path(ruta_reporting)
    raw = Path(ruta_raw) if ruta_raw else rep.parent.parent / "01_raw" / periodo
    salida = rep / "serie_deptal" / cfg["carpeta"]

    def sust(cols: list[str]) -> list[str]:
        mapa = {"@PUB": pub, "@MANT": m_ant, "@MACT": m_act, "@N": n_act,
                "N_ARTICULOS": n_art, "PRECIO_PROMEDIO": p_act}
        return [mapa.get(c, c) for c in cols]

    # ── 1. Importación ────────────────────────────────────────────────────────
    ruta_base = _buscar(raw / f"BASES LIVIANAS {periodo}", cfg["clave"])
    ruta_div_mod = _buscar(raw / f"DIVIPOLA {periodo}", cfg["clave"], prefijo="divipola")
    ruta_div = _buscar_nombre([raw / f"DIVIPOLA {periodo}"], "DIVIPOLA.xlsx")
    log.info("[%s] Serie departamental %s | base=%s", m, periodo, ruta_base.name)

    base = _leer_hoja(ruta_base, "Información Insumos", cfg["motor"])
    mpio = base["Municipio"]
    if _es_numerica(mpio):
        cod = mpio.map(lambda v: "" if pd.isna(v) else f"{int(round(v)):05d}")
    else:
        cod = pd.to_numeric(mpio.replace("", np.nan), errors="coerce").map(
            lambda v: "" if pd.isna(v) else f"{int(round(v)):05d}")
    nuevas = {
        "CodigoMpio": cod,
        "CodigoDepto": cod.str.slice(0, 2),
        cfg["llave"]: _catx(base, cfg["llave_cols"], cfg["largo_llave"]),
    }
    for nombre, cols in cfg.get("extra_llaves", {}).items():
        nuevas[nombre] = _catx(base, cols, 200)
    quitar = {c for c in base.columns if c.rstrip(".#") in _DROP_BASE} | {"Municipio"}
    base = base[[c for c in base.columns if c not in quitar]].copy()
    for c in cfg["nuevas"]:
        base[c] = nuevas[c].values

    div = _leer_hoja(ruta_div, "Dep_Mun", cfg["motor"])
    cm = div["CódigoMunicipio"]
    div["CodigoMpio"] = cm.map(lambda v: "" if pd.isna(v) else f"{int(round(float(v))):05d}")
    div["CodigoDepto"] = div["CodigoMpio"].str.slice(0, 2)
    div = div.drop(columns=["CódigoMunicipio"])

    # ── 2. DIVIPOLA, grupos y nombres de publicación ─────────────────────────
    datos, _, _ = _merge_sas(_ordenar(div, ["CodigoMpio"]), _ordenar(base, ["CodigoMpio"]),
                             ["CodigoMpio"], salida="b")
    llave = cfg["llave"]
    faltan = None
    if cfg["hoja_grupo"]:
        grupos = _leer_hoja(ruta_div_mod, cfg["hoja_grupo"], cfg["motor"])
        grupos = grupos[[llave, "Grupo"]]
        grupos = grupos[grupos[llave] != ""]
        datos = _ordenar(datos, [llave])
        datos, _, en_b = _merge_sas(datos, _ordenar(grupos, [llave]), [llave], salida="a")
        faltan = datos[~en_b.values].reset_index(drop=True)

    art = _leer_hoja(ruta_div_mod, cfg["hoja_articulo"], cfg["motor"])
    col_llave_art = next(c for c in art.columns if c.lower() == llave.lower())
    col_nom_art = next(c for c in art.columns if c.lower() == cfg["col_publica"][0].lower())
    art = pd.DataFrame({llave: art[col_llave_art], pub: art[col_nom_art]})
    art = art[art[llave] != ""]
    datos = _ordenar(datos, [llave])
    if cfg.get("llave_al_inicio"):
        # data X; LENGTH Articulo_Caracte $333; set X; — la llave pasa a ser la primera variable.
        datos = datos[[llave] + [c for c in datos.columns if c != llave]]
    datos, _, en_b = _merge_sas(datos, _ordenar(art, [llave]), [llave], salida="a")
    if faltan is None:
        faltan = datos[~en_b.values].reset_index(drop=True)
    ruta_falt = salida / cfg["faltan"][0].format(M=mes)
    _escribir(ruta_falt, [(cfg["faltan"][1], faltan)])

    # ── 3. Precios, orden de variables y duplicados ──────────────────────────
    ante, actual, var = cfg["precios"]
    datos[actual] = datos[actual].where(datos[actual] != 0)
    if cfg.get("ante_errado"):
        # IF 'Precio Ante#'n = 0 THEN 'Ante#'n = .; — el programa crea "Ante#" y no limpia el precio.
        datos["Ante#"] = np.nan
    else:
        datos[ante] = datos[ante].where(datos[ante] != 0)
    if cfg.get("ante_ia_in"):
        datos.loc[datos["Nov#"].isin(["IA", "IN"]), ante] = np.nan
    datos = datos.rename(columns={ante: m_ant, actual: m_act})
    retener = [c for c in sust(cfg["retain"]) if c in datos.columns]
    datos = datos[retener + [c for c in datos.columns if c not in retener and c != var]]
    datos, dupli = _nodupkey(datos, cfg["llave_dupli"])
    ruta_dup = salida / cfg["dupli"].format(M=mes)
    _escribir(ruta_dup, [("DUPLICADOS", dupli)])

    # ── 4. Agregación por departamento ──────────────────────────────────────
    fte = cfg["fuente"]
    datos["_new"] = datos["CodigoMpio"] + datos[fte].astype(str)
    con_precio = datos[datos[m_act].notna()].copy()
    grupo_cols = ["CodigoDepto", pub] + cfg["grp"]
    gk = [_clave_orden(con_precio[c]) for c in grupo_cols]

    if cfg["conteo"] == "remerge":
        # GROUP BY depto, nombre, New_Informante con Articulo sin agregar: SAS
        # vuelve a pegar el resumen a cada registro (remerge).
        k_inf = [_clave_orden(con_precio[c]) for c in ["CodigoDepto", pub]] + [con_precio["_new"]]
        con_precio["_N"] = con_precio.groupby(k_inf, sort=False)["Articulo"].transform(_contar)
        cuenta = con_precio.groupby(gk, sort=False).agg(_NF=("_new", "size"), _NA=("_N", "sum"))
    else:
        k_f = con_precio["_new"] if cfg["conteo"] == "fuente" else _clave_orden(con_precio[fte])
        por_fuente = con_precio.groupby(gk + [k_f], sort=False)["Articulo"].agg(_contar).rename("_N").reset_index()
        cuenta = por_fuente.groupby(list(por_fuente.columns[:len(grupo_cols)]), sort=False).agg(
            _NF=("_N", "size"), _NA=("_N", "sum"))

    agr = {"CodigoDepto": ("CodigoDepto", _sql_max), "NombreDepartamento": ("NombreDepartamento", _sql_max),
           "Codigo CPC": ("Codigo CPC", _sql_max), pub: (pub, _sql_max)}
    for c in cfg["grp"]:
        agr[c] = (c, "first")
    if cfg["grupo"]:
        agr["Grupo"] = ("Grupo", _sql_max)
    agr["_prom"] = (m_act, "mean")
    por = con_precio.groupby(gk, sort=False).agg(**agr)
    por = por.join(cuenta)
    por = por.reset_index(drop=True)
    por[n_act] = por["_NF"].astype(float)
    por[n_art] = por["_NA"].astype(float)
    por[p_act] = por["_prom"].round(7)
    por = por[por[p_act].notna()]
    orden = ["CodigoDepto", "NombreDepartamento"] + (["Grupo"] if cfg["grupo"] else []) + [pub] + cfg["grp"]
    por = _ordenar(por, orden)

    cols_sal = sust(cfg["salida_cols"])
    mayor = por[por[n_art] >= 2][cols_sal].reset_index(drop=True)
    menor = por[por[n_art] < 2][cols_sal].reset_index(drop=True)
    if cfg.get("arroz"):
        arroz = por[pub].str.contains("Arroz", regex=False)
        mayor = pd.concat([por[arroz][cols_sal], por[~arroz & (por[n_art] >= 2)][cols_sal]], ignore_index=True)
        menor = por[~arroz & (por[n_art] < 2)][cols_sal].reset_index(drop=True)
    ruta_may = salida / cfg["mayor"][0].format(M=mes)
    ruta_men = salida / cfg["menor"][0].format(M=mes)
    _escribir(ruta_may, [(cfg["mayor"][1], mayor)])
    _escribir(ruta_men, [(cfg["menor"][1], menor)])

    # ── 5. Mes anterior, variación y ANEXO ───────────────────────────────────
    nombre_ant = cfg["mayor"][0].format(M=mes_ant_arch)
    carpetas = [Path(c) for c in (carpetas_mes_anterior or [])] or [
        rep.parent / periodo_ant / "serie_deptal" / cfg["carpeta"],
        raw / f"SERIE_DEPTAL {periodo}",
    ]
    ruta_ant = _buscar_nombre(carpetas, nombre_ant)
    if ruta_ant is None:
        raise FileNotFoundError(f"[{m}] No se encontró {nombre_ant} (período anterior) en {carpetas}")
    ant = _leer_hoja(ruta_ant, cfg["mayor"][1], cfg["motor"])
    p_ant = next(c for c in ant.columns if c.upper() == f"PRECIO_PROMEDIO_{periodo_ant}")
    ant[p_ant] = ant[p_ant].round(7)
    for c in cfg["ant_num"]:
        # 'Codigo CPC1'n='Codigo CPC'n*1; drop; rename — la variable queda al final.
        v = pd.to_numeric(ant[c].replace("", np.nan), errors="coerce") if not _es_numerica(ant[c]) else ant[c]
        ant = ant.drop(columns=[c])
        ant[c] = v.astype("float64")
    llave_ant = sust(cfg["llave_ant"])
    unido, _, _ = _merge_sas(_ordenar(mayor, llave_ant),
                             _ordenar(ant.drop(columns=["NombreDepartamento"]), llave_ant),
                             llave_ant, salida="a")
    unido["Variacion(%)"] = (unido[p_act] / unido[p_ant].where(unido[p_ant] != 0) - 1) * 100

    conteo = unido.groupby(_clave_orden(unido[pub]), sort=True).agg(
        _cpc=("Codigo CPC", _sql_max), Cuenta=(pub, _contar)).reset_index(drop=True)
    conteo[pub] = sorted(unido[pub].unique(), key=lambda s: s)
    conteo = conteo.rename(columns={"_cpc": "Codigo CPC"})[["Codigo CPC", pub, "Cuenta"]]
    conteo["Cuenta"] = conteo["Cuenta"].astype(float)
    unido, _, _ = _merge_sas(_ordenar(unido, [pub]), conteo, [pub], salida="a")

    v = unido["Variacion(%)"]
    unido["Tendencia"] = np.select([v.isna(), v > 0, v < 0, v == 0], ["n.d.", "Positiva", "Negativa", "Estable"], "")
    np_ = [_nombre_presentacion(n, cfg["estilo_nombre"]) for n in unido[pub]]
    unido["Nombre_insumo"] = [a for a, _ in np_]
    unido["Presentación_insumo"] = [b for _, b in np_]
    unido["NOM_ARTICULO"] = unido["Articulo"].astype(str).str.strip()
    if "CasaCom." in unido.columns or "CasaCom#" in unido.columns:
        casa = "CasaCom." if "CasaCom." in unido.columns else "CasaCom#"
        unido["NOM_CASA COMERCIAL"] = unido[casa].astype(str).str.strip()
    if "RegICA" in unido.columns:
        unido["REGISTRO_ICA"] = unido["RegICA"]

    cols_anexo = [{"@ANT": p_ant, "@ACT": p_act, "@PUB": pub}.get(c, c) for c in cfg["anexo_cols"]]
    hojas = []
    for hoja in cfg["hojas_anexo"]:
        sub = unido[unido["Grupo"] == hoja] if cfg["grupo"] else unido
        hojas.append((hoja, _ordenar(sub, ["CodigoDepto", pub])[cols_anexo]))
    ruta_anexo = salida / cfg["anexo"].format(M=mes)
    _escribir(ruta_anexo, hojas)

    log.info("[%s] Serie departamental %s | MAYORESQUE2=%d MENORESQUE2=%d ANEXO=%d | %s",
             m, periodo, len(mayor), len(menor), sum(len(h) for _, h in hojas), salida)
    return pd.DataFrame([
        {"modulo": m, "periodo": periodo, "archivo": str(r), "filas": n}
        for r, n in [(ruta_falt, len(faltan)), (ruta_dup, len(dupli)), (ruta_may, len(mayor)),
                     (ruta_men, len(menor)), (ruta_anexo, sum(len(h) for _, h in hojas))]
    ])
