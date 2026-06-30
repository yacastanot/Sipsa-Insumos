"""Script para generar mappings de Jornales y Especies Productivas (JUN2026).

Ejecutar DESPUÉS de copiar los archivos DIVIPOLA en la carpeta correspondiente:
  data/01_raw/JUN2026/DIVIPOLA JUN2026/divipola jornales jun 2026.xlsx
  data/01_raw/JUN2026/DIVIPOLA JUN2026/divipola especie productiva jun 2026.xlsx

Uso:
  python scripts/generar_mappings_jun2026.py
"""
import yaml
import pandas as pd
from pathlib import Path

BASE_DIVIPOLA = Path("data/01_raw/JUN2026/DIVIPOLA JUN2026")

MODULOS_CARACTE = [
    {
        "nombre": "jornales",
        "grupo": "JORNALES",
        "archivo": "divipola jornales jun 2026.xlsx",
        "hoja": "Articulo",
        "col_llave": "articulo_caracte",
        "col_nombre": "articulo publicacion",
    },
    {
        "nombre": "especies",
        "grupo": "ESPECIES PRODUCTIVAS",
        "archivo": "divipola especie productiva jun 2026.xlsx",
        "hoja": "Articulo",
        "col_llave": "articulo_caracte",
        "col_nombre": "articulo publicacion",
    },
]

YAML_ARTICULOS = Path("conf/base/mappings_articulos.yml")
YAML_GRUPOS    = Path("conf/base/mappings_grupos.yml")


def cargar_yaml(path: Path, clave: str) -> dict:
    if path.exists():
        with open(path, encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        return data.get(clave, data)
    return {}


def guardar_yaml(path: Path, clave: str, mapping: dict) -> None:
    with open(path, "w", encoding="utf-8") as f:
        yaml.dump({clave: dict(sorted(mapping.items()))}, f, allow_unicode=True, default_flow_style=False)
    print(f"  Guardado: {path} ({len(mapping)} entradas)")


def main():
    art_map = cargar_yaml(YAML_ARTICULOS, "articulos_publicacion")
    grp_map = cargar_yaml(YAML_GRUPOS, "grupos")

    print(f"Mappings actuales: articulos={len(art_map)}, grupos={len(grp_map)}")

    for cfg in MODULOS_CARACTE:
        ruta = BASE_DIVIPOLA / cfg["archivo"]
        if not ruta.exists():
            print(f"\n[AVISO] No encontrado: {ruta}")
            print(f"  Coloca el archivo DIVIPOLA de {cfg['nombre']} en esa ruta y vuelve a ejecutar.")
            continue

        df = pd.read_excel(ruta, sheet_name=cfg["hoja"])
        df.columns = df.columns.str.strip().str.lower()

        col_llave  = cfg["col_llave"]
        col_nombre = cfg["col_nombre"]

        if col_llave not in df.columns or col_nombre not in df.columns:
            print(f"\n[ERROR] Columnas no encontradas en {ruta.name}.")
            print(f"  Columnas disponibles: {list(df.columns)}")
            continue

        df = df.dropna(subset=[col_llave])
        nuevas_art = {}
        nuevas_grp = {}
        for _, row in df.iterrows():
            llave  = str(row[col_llave]).strip().upper()
            nombre = str(row[col_nombre]).strip()
            if llave:
                nuevas_art[llave] = nombre
                nuevas_grp[llave] = cfg["grupo"]

        solapados = set(nuevas_art) & set(art_map)
        print(f"\n{cfg['nombre'].upper()}: {len(nuevas_art)} entradas ({len(solapados)} solapados)")
        art_map.update(nuevas_art)
        grp_map.update(nuevas_grp)

    guardar_yaml(YAML_ARTICULOS, "articulos_publicacion", art_map)
    guardar_yaml(YAML_GRUPOS,    "grupos",                grp_map)
    print("\nMappings actualizados OK. Ahora actualiza globals.yml para JUN2026 y ejecuta:")
    print("  kedro run --pipeline jornales")
    print("  kedro run --pipeline especies")


if __name__ == "__main__":
    main()
