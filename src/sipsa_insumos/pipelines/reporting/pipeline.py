"""Pipeline de reportes — SIPSA Insumos."""
from kedro.pipeline import Pipeline, node

from .nodes import exportar_anexos, exportar_bases, exportar_cuadros, exportar_tablas


def create_pipeline(**kwargs) -> Pipeline:
    return Pipeline(
        [
            node(
                func=exportar_bases,
                inputs=["base_comparada", "params:grupos", "params:modulo",
                        "params:periodo", "params:ruta_reporting"],
                outputs="bases_meta",
                name="exportar_bases",
            ),
            node(
                func=exportar_tablas,
                inputs=["base_comparada", "params:grupos", "params:modulo",
                        "params:periodo", "params:ruta_reporting"],
                outputs="tablas_meta",
                name="exportar_tablas",
            ),
            node(
                func=exportar_anexos,
                inputs=["base_comparada", "params:grupos", "params:modulo",
                        "params:periodo", "params:ruta_reporting"],
                outputs="anexos_meta",
                name="exportar_anexos",
            ),
            node(
                func=exportar_cuadros,
                inputs=["base_comparada", "params:grupos", "params:modulo",
                        "params:periodo", "params:mes_actual", "params:mes_anterior",
                        "params:ruta_reporting"],
                outputs="cuadros_meta",
                name="exportar_cuadros",
            ),
        ]
    )
