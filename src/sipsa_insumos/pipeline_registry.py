"""Registro de pipelines del proyecto SIPSA Insumos.

Pipeline genérico: los 9 módulos de insumos ejecutan el mismo flujo de 6 etapas.
Se instancian vía namespace para que los datasets sean únicos en el catalog.

Módulos:
  agricolas    Insumos Agrícolas — Mensual (6 grupos: COADYUDANTES, FERTILIZANTES,
               FUNGICIDAS, HERBICIDAS, INSECTICIDAS, BIOINSUMOS)
  pecuarios    Insumos Pecuarios — Mensual (7 grupos: ALIMENTOS, ANTIBIOTICOS,
               ANTISEPTICOS, HORMONALES, INSECTICIDAS, MEDICAMENTOS, VITAMINAS)
  arriendos    Arriendos de Tierras — Trimestral (FEB, MAY, AGO, NOV)
  servicios    Servicios Agrícolas — Trimestral (FEB, MAY, AGO, NOV)
  jornales     Jornales Agrícolas — Trimestral (MAR, JUN, SEP, DIC)
  especies     Especies Productivas — Trimestral (MAR, JUN, SEP, DIC)
  elementos    Elementos Agropecuarios — Bimestral (ENE, MAR, MAY, JUL, SEP, NOV)
  empaques     Empaques Agropecuarios — Bimestral (ENE, MAR, MAY, JUL, SEP, NOV)
  propagacion  Material de Propagación — Bimestral (FEB, ABR, JUN, AGO, OCT, DIC)

Pipelines compuestos:
  __default__  Todos los módulos en secuencia.

Uso frecuente:
  kedro run                          # todos los módulos
  kedro run --pipeline agricolas     # solo insumos agrícolas
  kedro run --pipeline propagacion   # solo material de propagación
"""
from __future__ import annotations

from kedro.pipeline import Pipeline, pipeline

from .pipelines.ingestion.pipeline import create_pipeline as ingestion
from .pipelines.enrichment.pipeline import create_pipeline as enrichment
from .pipelines.quality.pipeline import create_pipeline as quality
from .pipelines.aggregation.pipeline import create_pipeline as aggregation
from .pipelines.comparison.pipeline import create_pipeline as comparison
from .pipelines.reporting.pipeline import create_pipeline as reporting

MODULOS = [
    "agricolas",
    "pecuarios",
    "arriendos",
    "servicios",
    "jornales",
    "especies",
    "elementos",
    "empaques",
    "propagacion",
]


def _pipeline_modulo(nombre: str) -> Pipeline:
    """Construye el pipeline completo de un módulo vía namespace.

    Datasets compartidos (sin prefijo de namespace):
      inputs:     divipola_raw, mappings_grupos, mappings_articulos
    Parámetros globales (sin prefijo de namespace):
      parameters: umbral_var_alta, umbral_var_baja, umbral_var_extrema,
                  mes_actual, mes_anterior, periodo, ruta_reporting
    Parámetros de módulo (CON prefijo de namespace, vienen de parameters_<modulo>.yml):
      params:modulo, params:grupos, params:archivo_liviana, params:hoja_liviana
    """
    return (
        pipeline(
            ingestion(),
            namespace=nombre,
            parameters={"periodo"},
            tags=[nombre],
          )
        + pipeline(
            enrichment(),
            namespace=nombre,
            inputs={"divipola_raw", "mappings_grupos", "mappings_articulos"},
            tags=[nombre],
          )
        + pipeline(
            quality(),
            namespace=nombre,
            parameters={"umbral_var_alta", "umbral_var_baja", "umbral_var_extrema"},
            tags=[nombre],
          )
        + pipeline(aggregation(), namespace=nombre, tags=[nombre])
        + pipeline(
            comparison(),
            namespace=nombre,
            parameters={"mes_actual", "mes_anterior"},
            tags=[nombre],
          )
        + pipeline(
            reporting(),
            namespace=nombre,
            parameters={"periodo", "ruta_reporting", "mes_actual", "mes_anterior"},
            tags=[nombre],
          )
    )


def register_pipelines() -> dict[str, Pipeline]:
    """Registra todos los pipelines del proyecto SIPSA Insumos."""
    pipelines: dict[str, Pipeline] = {m: _pipeline_modulo(m) for m in MODULOS}
    pipelines["__default__"] = sum(pipelines[m] for m in MODULOS)
    return pipelines
