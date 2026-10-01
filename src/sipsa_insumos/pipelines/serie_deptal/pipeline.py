"""Pipeline Serie Departamental — SIPSA Insumos.

Genera por módulo los 5 archivos de 03SERIE_DEPTAL (ANEXO, MAYORESQUE2,
MENORESQUE2, DUPLICADOS y FALTAN) en data/08_reporting/<periodo>/serie_deptal/.
Se ejecuta dentro del pipeline de cada módulo y también por separado:

  kedro run --pipeline serie_deptal_agricolas
"""
from __future__ import annotations

from kedro.pipeline import Pipeline, node

from .nodes import generar_serie_deptal


def create_pipeline(modulo: str) -> Pipeline:
    return Pipeline([
        node(
            func=lambda periodo, ruta_reporting, _m=modulo.upper(): generar_serie_deptal(periodo, _m, ruta_reporting),
            inputs=["params:periodo", "params:ruta_reporting"],
            outputs=f"{modulo}.serie_deptal_meta",
            name=f"{modulo}.generar_serie_deptal",
            tags=[modulo, f"serie_deptal_{modulo}"],
        ),
    ])
