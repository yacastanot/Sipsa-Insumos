"""Nodos del pipeline de enriquecimiento — SIPSA Insumos.

SAS equivalente (pasos 3-5):
  MERGE base_liviana DIVIPOLA; by CodigoMpio;
  MERGE base divipola_grupos; by Art_Unmed_Casacomer_ICA;
  MERGE base divipola_articulos; by Art_Unmed_Casacomer_ICA;

Flujo:
  archivo_divipola_grupos ──► [actualizar_mappings_divipola] ──► mappings actualizados (+ persistidos en disco)
  base_bronze + divipola_raw ──► [merge_divipola] ──► base_con_mpio + faltan_divipola
  base_con_mpio + mappings_grupos_actualizado ──► [asignar_grupo] ──► base_con_grupo + faltan_grupo
  base_con_grupo + mappings_articulos_actualizado ──► [asignar_articulo_publica] ──► base_enriquecida + faltan_publica
"""
from __future__ import annotations

import logging
from pathlib import Path

import openpyxl
import pandas as pd
import yaml

from sipsa_insumos.utils.parsers import parsear_llave_divipola

log = logging.getLogger(__name__)

MAPPINGS_GRUPOS_PATH = Path("conf/base/mappings_grupos.yml")
MAPPINGS_ARTICULOS_PATH = Path("conf/base/mappings_articulos.yml")


def _hoja_case_insensitive(nombres: list[str], candidatos: list[str]) -> str | None:
    """Devuelve el nombre real de la primera hoja que coincida (case-insensitive)."""
    for cand in candidatos:
        for h in nombres:
            if h.strip().lower() == cand.lower():
                return h
    return None


def actualizar_mappings_divipola(
    archivo_divipola_grupos: str,
    tipo_llave: str,
    mappings_grupos: dict,
    mappings_articulos: dict,
    modulo: str,
) -> tuple[dict, dict]:
    """Actualiza el mapping con el DIVIPOLA propio del módulo del período.

    Lee la hoja 'Articulo' (Grupo, Nombre_Publica) y, si existe, la hoja
    'Grupo' del archivo DIVIPOLA {modulo} {periodo}.xlsx (columna compuesta
    Art_Unmed_Casacomer_ICA o Art_Casacomer_ICA_Unmed), reconstruye la
    LLAVE_ARTICULO con la misma lógica que usa el pipeline al leer la base
    liviana, y la registra en mappings_grupos.yml / mappings_articulos.yml.

    El DIVIPOLA del período manda, igual que en SAS (que cruza siempre contra
    el DIVIPOLA del mes): las llaves nuevas se agregan y las existentes cuyo
    nombre o grupo cambió se actualizan (p.ej. "Destierro Sl" → "Destierro
    150 Sl"). Las llaves que el DIVIPOLA no trae se conservan. Dentro del
    mismo archivo, si una llave aparece repetida, gana la primera aparición.
    Los cambios quedan tanto en memoria (para que el resto de este mismo
    pipeline los use de inmediato) como persistidos en disco.

    Args:
        archivo_divipola_grupos: Ruta al DIVIPOLA propio del módulo.
        tipo_llave: "unmed" o "casacom_ica_unmed" (ver parsear_llave_divipola).
        mappings_grupos: Dict {"grupos": {...}} cargado desde YAML.
        mappings_articulos: Dict {"articulos_publicacion": {...}} cargado desde YAML.
        modulo: Nombre del módulo para logging.

    Returns:
        (mappings_grupos, mappings_articulos) actualizados.
    """
    grupos_dict: dict[str, str] = dict(mappings_grupos.get("grupos", mappings_grupos))
    articulos_dict: dict[str, str] = dict(
        mappings_articulos.get("articulos_publicacion", mappings_articulos)
    )

    if not archivo_divipola_grupos or not Path(archivo_divipola_grupos).exists():
        log.info(
            "[%s] actualizar_mappings_divipola | sin archivo_divipola_grupos configurado — se omite.",
            modulo,
        )
        return {"grupos": grupos_dict}, {"articulos_publicacion": articulos_dict}

    wb = openpyxl.load_workbook(archivo_divipola_grupos, read_only=True)
    hojas = wb.sheetnames
    wb.close()

    nuevos_grupos = 0
    nuevos_articulos = 0
    cambiados_grupos = 0
    cambiados_articulos = 0
    vistos_art: set[str] = set()
    vistos_grp: set[str] = set()

    def _registrar(dic: dict, vistos: set, llave: str, valor: str) -> str | None:
        """Registra valor; devuelve 'nuevo', 'cambio' o None."""
        if llave in vistos:
            return None
        vistos.add(llave)
        if llave not in dic:
            dic[llave] = valor
            return "nuevo"
        if dic[llave] != valor:
            dic[llave] = valor
            return "cambio"
        return None

    # Hojas con nombre de publicación: "Articulo" y, además, cualquier hoja con
    # llave completa (p.ej. "DivArt_jun2026" de Material, Art_Casacomer_ICA_Unmed).
    # Con tipo_llave "casacom_ica_unmed" solo sirven las hojas de llave completa:
    # la llave corta Articulo_UnMed puede repetirse con nombres distintos
    # (Material JUN2026: "Pasto Brachiaria Decumbens_BOLSA|KILOGRAMO|1").
    hoja_articulo = _hoja_case_insensitive(hojas, ["Articulo", "Artículo"])
    hoja_grupo = _hoja_case_insensitive(hojas, ["Grupo"])
    hojas_nombre = [h for h in [hoja_articulo] if h] + [
        h for h in hojas if h not in (hoja_articulo, hoja_grupo)
    ]
    for hoja_articulo in hojas_nombre:
        df = pd.read_excel(archivo_divipola_grupos, sheet_name=hoja_articulo)
        col_llave = next((c for c in df.columns if "unmed" in str(c).lower() or "casacomer" in str(c).lower()), None)
        col_nombre = next(
            (c for c in df.columns if any(k in str(c).lower() for k in ("nombre", "publica", "product"))),
            None,
        )
        col_grupo = next((c for c in df.columns if str(c).strip().lower() == "grupo"), None)
        if col_llave is None or col_nombre is None:
            if hoja_articulo == hojas_nombre[0] and len(df):
                log.warning(
                    "[%s] actualizar_mappings_divipola | no se encontró columna clave/nombre en hoja '%s'",
                    modulo, hoja_articulo,
                )
            continue
        if tipo_llave == "casacom_ica_unmed" and "casacomer" not in str(col_llave).lower():
            continue
        if col_llave and col_nombre:
            for _, row in df.iterrows():
                llave = parsear_llave_divipola(row[col_llave], tipo_llave)
                if not llave:
                    continue
                r = _registrar(articulos_dict, vistos_art, llave, str(row[col_nombre]).strip())
                nuevos_articulos += r == "nuevo"
                cambiados_articulos += r == "cambio"
                if col_grupo is not None and pd.notna(row[col_grupo]):
                    r = _registrar(grupos_dict, vistos_grp, llave, str(row[col_grupo]).strip())
                    nuevos_grupos += r == "nuevo"
                    cambiados_grupos += r == "cambio"

    if hoja_grupo:
        df = pd.read_excel(archivo_divipola_grupos, sheet_name=hoja_grupo)
        col_llave = next((c for c in df.columns if "unmed" in c.lower() or "casacomer" in c.lower()), None)
        col_grupo = next((c for c in df.columns if "grupo" in c.lower()), None)
        if col_llave and col_grupo:
            for _, row in df.iterrows():
                llave = parsear_llave_divipola(row[col_llave], tipo_llave)
                if llave and pd.notna(row[col_grupo]):
                    r = _registrar(grupos_dict, vistos_grp, llave, str(row[col_grupo]).strip())
                    nuevos_grupos += r == "nuevo"
                    cambiados_grupos += r == "cambio"

    if nuevos_grupos or nuevos_articulos or cambiados_grupos or cambiados_articulos:
        MAPPINGS_GRUPOS_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(MAPPINGS_GRUPOS_PATH, "w", encoding="utf-8") as f:
            yaml.dump({"grupos": dict(sorted(grupos_dict.items()))}, f, allow_unicode=True, default_flow_style=False)
        with open(MAPPINGS_ARTICULOS_PATH, "w", encoding="utf-8") as f:
            yaml.dump(
                {"articulos_publicacion": dict(sorted(articulos_dict.items()))},
                f, allow_unicode=True, default_flow_style=False,
            )

    log.info(
        "[%s] actualizar_mappings_divipola OK | grupos_nuevos=%d | grupos_cambiados=%d | "
        "articulos_nuevos=%d | articulos_cambiados=%d | total_grupos=%d | total_articulos=%d",
        modulo, nuevos_grupos, cambiados_grupos, nuevos_articulos, cambiados_articulos,
        len(grupos_dict), len(articulos_dict),
    )
    return {"grupos": grupos_dict}, {"articulos_publicacion": articulos_dict}


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
    grupos: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Asigna el grupo de insumo a cada artículo usando el mapping YAML.

    SAS: MERGE base divipola_grupos; by Art_Unmed_Casacomer_ICA;
    Equivalente: lookup LLAVE_ARTICULO → Grupo en mappings_grupos.yml.

    Si el módulo tiene un único grupo posible (ej. Elementos, Empaques,
    Propagación), no hay ambigüedad que resolver: cualquier artículo sin
    entrada en el mapping se asigna directamente a ese grupo, sin necesidad
    de mantener el DIVIPOLA de grupos al día para estos módulos.

    Args:
        base_con_mpio: DataFrame con columna 'LLAVE_ARTICULO'.
        mappings_grupos: Dict {llave_articulo: grupo} cargado desde YAML.
        modulo: Nombre del módulo para logging.
        grupos: Lista de grupos configurados para el módulo (params:grupos).

    Returns:
        (base_con_grupo, faltan_grupo): base con columna 'Grupo' + filas sin match.
    """
    grupos_dict: dict[str, str] = mappings_grupos.get("grupos", mappings_grupos)
    df = base_con_mpio.copy()
    df["Grupo"] = df["LLAVE_ARTICULO"].map(grupos_dict)

    if grupos and len(grupos) == 1:
        n_respaldo = df["Grupo"].isna().sum()
        if n_respaldo > 0:
            df.loc[df["Grupo"].isna(), "Grupo"] = grupos[0]
            log.info(
                "asignar_grupo [%s] | %d filas sin mapping asignadas por respaldo de grupo único ('%s')",
                modulo, n_respaldo, grupos[0],
            )

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

    # SAS agrega por (municipio, Nombre_Publica): si el DIVIPOLA asigna grupos
    # distintos a marcas del mismo nombre de publicación (p.ej. "Sales del 6%,
    # 40 kilogramos" VITAMINAS/ALIMENTOS en JUL2026), SAS igual publica una
    # sola fila con un grupo. Se unifica al grupo más frecuente del período
    # para no partir el producto en BASES/CUADROS/MAYORESQUE2.
    if "Grupo" in base_enriquecida.columns and len(base_enriquecida):
        grupo_pub = (
            base_enriquecida.dropna(subset=["Grupo"])
            .groupby("Nombre_Publica")["Grupo"]
            .agg(lambda s: s.value_counts().index[0])
        )
        base_enriquecida["Grupo"] = base_enriquecida["Nombre_Publica"].map(grupo_pub).fillna(
            base_enriquecida["Grupo"])

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
