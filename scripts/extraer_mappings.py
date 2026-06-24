"""Script auxiliar para extraer mappings desde los archivos DIVIPOLA reales.

Uso:
    python scripts/extraer_mappings.py \\
        --divipola "data/01_raw/MAY2026/divipola insumos agrícolas may 2026.xlsx"

Genera entradas para:
    conf/base/mappings_grupos.yml
    conf/base/mappings_articulos.yml

El formato de la columna Art_Unmed_Casacomer_ICA en el DIVIPOLA es:
    ARTÍCULO_UNIDAD|DE|MEDIDA|0|0_CASA_COMERCIAL_REGISTRO_ICA

La clave de mapping usa solo las dos primeras partes (equivale a la
LLAVE_ARTICULO construida por construir_llave_articulo() en utils/parsers.py):
    ARTÍCULO.upper()_UNIDAD|DE|MEDIDA.upper()
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import openpyxl
import pandas as pd


def _encontrar_hoja(ruta: str, candidatos: list[str]) -> str | None:
    """Devuelve el nombre real de la primera hoja que coincida (case-insensitive)."""
    wb = openpyxl.load_workbook(ruta, read_only=True)
    hojas = wb.sheetnames
    wb.close()
    for c in candidatos:
        for h in hojas:
            # normalizar: quitar acentos y pasar a minúsculas para comparar
            h_norm = h.lower().replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
            if h_norm == c.lower():
                return h
    return None


def _extraer_llave_corta(llave_completa: str) -> str:
    """Extrae ARTÍCULO_UNIDAD_DE_MEDIDA de la clave completa del DIVIPOLA.

    Formato completo: ARTÍCULO_UNIDAD|MEDIDA|0|0_CASA_COMERCIAL_REGISTRO_ICA
    La UNIDAD DE MEDIDA siempre contiene '|', lo que la diferencia del ARTÍCULO
    y de los campos siguientes.

    Retorna ARTÍCULO + '_' + UNIDAD_DE_MEDIDA en mayúsculas.
    """
    llave = llave_completa.strip().upper()
    partes = llave.split("_")
    # Buscar el índice de la parte que contiene '|' (UNIDAD DE MEDIDA)
    idx_unidad = next((i for i, p in enumerate(partes) if "|" in p), None)
    if idx_unidad is None:
        return llave  # sin pipe → devolver tal cual
    articulo   = "_".join(partes[:idx_unidad])
    unidad_med = partes[idx_unidad]
    return f"{articulo}_{unidad_med}"


def extraer_grupos(ruta_divipola: str) -> None:
    """Lee la hoja de grupos y genera entradas para mappings_grupos.yml."""
    hoja = _encontrar_hoja(ruta_divipola, ["Grupo", "grupo"])
    if hoja is None:
        print("# (sin hoja 'Grupo': módulo de un solo grupo, no requiere mapping)", file=sys.stderr)
        return

    try:
        df = pd.read_excel(ruta_divipola, sheet_name=hoja, header=0)
    except Exception as exc:
        print(f"ERROR leyendo hoja '{hoja}': {exc}", file=sys.stderr)
        return

    # Detectar columnas
    col_llave = next((c for c in df.columns if "art_unmed" in c.lower()), df.columns[0])
    col_grupo = next((c for c in df.columns if "grupo" in c.lower()), None)

    print(f"\n# Extrayendo grupos desde hoja '{hoja}' de: {Path(ruta_divipola).name}")
    print("\n# Añadir al archivo conf/base/mappings_grupos.yml:")
    print("grupos:")
    vistas = set()
    for _, row in df.iterrows():
        llave = _extraer_llave_corta(str(row[col_llave]))
        if llave in vistas:
            continue
        vistas.add(llave)
        grupo = str(row[col_grupo]).strip() if col_grupo else "DESCONOCIDO"
        print(f'  "{llave}": "{grupo}"')


def extraer_articulos(ruta_divipola: str) -> None:
    """Lee la hoja de artículos y genera entradas para mappings_articulos.yml."""
    hoja = _encontrar_hoja(ruta_divipola, ["Articulo", "articulo", "Artículo", "artículo"])
    if hoja is None:
        print(f"ERROR: no se encontró hoja de artículos en {ruta_divipola}", file=sys.stderr)
        return

    try:
        df = pd.read_excel(ruta_divipola, sheet_name=hoja, header=0)
    except Exception as exc:
        print(f"ERROR leyendo hoja '{hoja}': {exc}", file=sys.stderr)
        return

    col_llave = next((c for c in df.columns if "art_unmed" in c.lower()), df.columns[0])
    # Columna de nombre de publicación: buscar por palabras clave
    col_nombre = next(
        (c for c in df.columns
         if any(k in c.lower() for k in ["nombre", "publica", "product"])),
        df.columns[-1],
    )

    print(f"\n# Extrayendo artículos desde hoja '{hoja}' de: {Path(ruta_divipola).name}")
    print("\n# Añadir al archivo conf/base/mappings_articulos.yml:")
    print("articulos_publicacion:")
    vistas = set()
    for _, row in df.iterrows():
        llave = _extraer_llave_corta(str(row[col_llave]))
        if llave in vistas:
            continue
        vistas.add(llave)
        nombre = str(row[col_nombre]).strip()
        print(f'  "{llave}": "{nombre}"')


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extrae mappings desde archivos DIVIPOLA de SIPSA Insumos."
    )
    parser.add_argument(
        "--divipola",
        required=True,
        help="Ruta al archivo DIVIPOLA de un módulo",
    )
    args = parser.parse_args()

    ruta = args.divipola
    extraer_grupos(ruta)
    extraer_articulos(ruta)


if __name__ == "__main__":
    main()
