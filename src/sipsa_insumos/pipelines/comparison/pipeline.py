"""Pipeline de comparación interperiódica — SIPSA Insumos."""
from kedro.pipeline import Pipeline, node

from .nodes import calcular_variacion_tendencia


def create_pipeline(**kwargs) -> Pipeline:
    return Pipeline(
        [
            node(
                func=calcular_variacion_tendencia,
                inputs=[
                    "mayor2",
                    "mayor2_anterior",
                    "params:mes_actual",
                    "params:mes_anterior",
                ],
                outputs="base_comparada",
                name="calcular_variacion_tendencia",
            ),
        ]
    )
