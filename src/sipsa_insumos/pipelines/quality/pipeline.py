"""Pipeline de calidad — SIPSA Insumos."""
from kedro.pipeline import Pipeline, node

from .nodes import calcular_cv, detectar_duplicados, detectar_var_atipica


def create_pipeline(**kwargs) -> Pipeline:
    return Pipeline(
        [
            node(
                func=detectar_duplicados,
                inputs=["base_enriquecida"],
                outputs=["base_sin_dupli", "duplicados"],
                name="detectar_duplicados",
            ),
            node(
                func=calcular_cv,
                inputs=["base_sin_dupli", "params:modulo", "params:periodo"],
                outputs=["base_con_cv", "cvs_reporte"],
                name="calcular_cv",
            ),
            node(
                func=detectar_var_atipica,
                inputs=[
                    "base_con_cv",
                    "mayor2_anterior",
                    "params:umbral_var_alta",
                    "params:umbral_var_baja",
                    "params:umbral_var_extrema",
                ],
                outputs=["base_calidad", "var_atipico"],
                name="detectar_var_atipica",
            ),
        ]
    )
