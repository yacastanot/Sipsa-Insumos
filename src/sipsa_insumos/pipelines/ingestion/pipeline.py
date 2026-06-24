"""Pipeline de ingesta — SIPSA Insumos."""
from kedro.pipeline import Pipeline, node

from .nodes import leer_base_liviana


def create_pipeline(**kwargs) -> Pipeline:
    return Pipeline(
        [
            node(
                func=leer_base_liviana,
                inputs=["params:archivo_liviana", "params:hoja_liviana", "params:periodo"],
                outputs="base_bronze",
                name="leer_base_liviana",
            ),
        ]
    )
