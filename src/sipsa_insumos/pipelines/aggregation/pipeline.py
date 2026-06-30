"""Pipeline de agregación — SIPSA Insumos."""
from kedro.pipeline import Pipeline, node

from .nodes import aplicar_secreto_estadistico, calcular_precio_promedio


def create_pipeline(**kwargs) -> Pipeline:
    return Pipeline(
        [
            node(
                func=calcular_precio_promedio,
                inputs=["base_calidad"],
                outputs="precio_promedio",
                name="calcular_precio_promedio",
            ),
            node(
                func=aplicar_secreto_estadistico,
                inputs=["precio_promedio", "params:min_n"],
                outputs=["mayor2", "menor2"],
                name="aplicar_secreto_estadistico",
            ),
        ]
    )
