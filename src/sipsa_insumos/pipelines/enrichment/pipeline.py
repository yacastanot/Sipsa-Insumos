"""Pipeline de enriquecimiento — SIPSA Insumos."""
from kedro.pipeline import Pipeline, node

from .nodes import (
    actualizar_mappings_divipola,
    asignar_articulo_publica,
    asignar_grupo,
    merge_divipola,
)


def create_pipeline(**kwargs) -> Pipeline:
    return Pipeline(
        [
            node(
                func=actualizar_mappings_divipola,
                inputs=[
                    "params:archivo_divipola_grupos",
                    "params:tipo_llave",
                    "mappings_grupos",
                    "mappings_articulos",
                    "params:modulo",
                ],
                outputs=["mappings_grupos_actualizado", "mappings_articulos_actualizado"],
                name="actualizar_mappings_divipola",
            ),
            node(
                func=merge_divipola,
                inputs=["base_bronze", "divipola_raw"],
                outputs=["base_con_mpio", "faltan_divipola"],
                name="merge_divipola",
            ),
            node(
                func=asignar_grupo,
                inputs=["base_con_mpio", "mappings_grupos_actualizado", "params:modulo", "params:grupos"],
                outputs=["base_con_grupo", "faltan_grupo"],
                name="asignar_grupo",
            ),
            node(
                func=asignar_articulo_publica,
                inputs=["base_con_grupo", "mappings_articulos_actualizado", "params:modulo"],
                outputs=["base_enriquecida", "faltan_publica"],
                name="asignar_articulo_publica",
            ),
        ]
    )
