"""Pipeline de reportes — SIPSA Insumos."""
from kedro.pipeline import Pipeline, node

from .nodes import (
    exportar_anexos,
    exportar_bases,
    exportar_cuadros,
    exportar_mayo_menores,
    exportar_revision_historica,
    exportar_revision_tematica,
    exportar_tablas,
)


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
            node(
                func=exportar_mayo_menores,
                inputs=["mayor2", "menor2", "mayor2_anterior",
                        "params:modulo", "params:periodo",
                        "params:mes_actual", "params:mes_anterior",
                        "params:ruta_reporting"],
                outputs="mayo_menores_meta",
                name="exportar_mayo_menores",
            ),
            node(
                func=exportar_revision_historica,
                inputs=["base_enriquecida",
                        "params:modulo", "params:periodo",
                        "params:mes_actual", "params:mes_num_actual",
                        "params:ruta_reporting"],
                outputs="revision_hist_meta",
                name="exportar_revision_historica",
            ),
            node(
                func=exportar_revision_tematica,
                inputs=["base_enriquecida", "mayor2_anterior",
                        "params:modulo", "params:periodo",
                        "params:mes_actual", "params:grupos",
                        "params:ruta_reporting"],
                outputs="revision_tematica_meta",
                name="exportar_revision_tematica",
            ),
        ]
    )
