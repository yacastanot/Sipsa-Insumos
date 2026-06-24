"""Pipeline de enriquecimiento — SIPSA Insumos."""
from kedro.pipeline import Pipeline, node

from .nodes import asignar_articulo_publica, asignar_grupo, merge_divipola


def create_pipeline(**kwargs) -> Pipeline:
    return Pipeline(
        [
            node(
                func=merge_divipola,
                inputs=["base_bronze", "divipola_raw"],
                outputs=["base_con_mpio", "faltan_divipola"],
                name="merge_divipola",
            ),
            node(
                func=asignar_grupo,
                inputs=["base_con_mpio", "mappings_grupos", "params:modulo"],
                outputs=["base_con_grupo", "faltan_grupo"],
                name="asignar_grupo",
            ),
            node(
                func=asignar_articulo_publica,
                inputs=["base_con_grupo", "mappings_articulos", "params:modulo"],
                outputs=["base_enriquecida", "faltan_publica"],
                name="asignar_articulo_publica",
            ),
        ]
    )
