"""Pipeline Sin Precio Anterior — SIPSA Insumos.

Genera los archivos REV_SIN_PRECIO_ANTE_[MODULO]_[PERIODO].xlsx
para los registros sin precio anterior (REVISA=2) en cada módulo.

Módulos incluidos:
  sin_precio_ant_agricolas   — mensual (todos los meses)
  sin_precio_ant_pecuarios   — mensual (todos los meses)
  sin_precio_ant_elementos   — bimestral impar (ENE, MAR, MAY, JUL, SEP, NOV)
  sin_precio_ant_propagacion — bimestral par   (FEB, ABR, JUN, AGO, OCT, DIC)

Uso:
  kedro run --pipeline sin_precio_ant              # todos los módulos activos
  kedro run --pipeline sin_precio_ant_agricolas    # solo agrícolas
  kedro run --pipeline sin_precio_ant_elementos    # solo elementos (bimestral impar)
  kedro run --pipeline sin_precio_ant_propagacion  # solo propagación (bimestral par)

El parámetro 'activo' en parameters_sin_precio_ant.yml controla qué módulos
se ejecutan cada período. Si activo=false el nodo escribe metadatos vacíos.
"""
from kedro.pipeline import Pipeline, node

from .nodes import revisar_sin_precio

def _inputs_modulo(mod: str) -> list[str]:
    """Genera la lista de inputs params en el mismo orden que la firma de revisar_sin_precio.

    Orden de argumentos de revisar_sin_precio():
      ruta_historico, hoja_historico, ruta_var_atipico, hoja_var_atipico,
      casacom_col_hist, unmed_col_hist, casacom_col_var, unmed_col_var, nov_col,
      mes_actual (global), periodo (global), modulo, ruta_reporting (global), activo
    """
    p = f"params:sin_precio_ant.{mod}"
    return [
        f"{p}.ruta_historico",
        f"{p}.hoja_historico",
        f"{p}.ruta_var_atipico",
        f"{p}.hoja_var_atipico",
        f"{p}.casacom_col_hist",
        f"{p}.unmed_col_hist",
        f"{p}.casacom_col_var",
        f"{p}.unmed_col_var",
        f"{p}.nov_col",
        "params:mes_actual",        # argumento 10: mes_actual
        "params:periodo",           # argumento 11: periodo
        f"{p}.modulo",              # argumento 12: modulo
        "params:ruta_reporting",    # argumento 13: ruta_reporting
        f"{p}.activo",              # argumento 14: activo (default=True)
    ]


def _nodo_modulo(mod: str, output_key: str) -> node:
    """Crea un nodo Kedro para un módulo Sin Precio Anterior."""
    return node(
        func=revisar_sin_precio,
        inputs=_inputs_modulo(mod),
        outputs=output_key,
        name=f"revisar_sin_precio_{mod}",
        tags=[f"sin_precio_ant_{mod}"],
    )


def create_pipeline_agricolas(**kwargs) -> Pipeline:
    return Pipeline([
        _nodo_modulo("agricolas", "sin_precio_ant.meta_agricolas"),
    ])


def create_pipeline_pecuarios(**kwargs) -> Pipeline:
    return Pipeline([
        _nodo_modulo("pecuarios", "sin_precio_ant.meta_pecuarios"),
    ])


def create_pipeline_elementos(**kwargs) -> Pipeline:
    return Pipeline([
        _nodo_modulo("elementos", "sin_precio_ant.meta_elementos"),
    ])


def create_pipeline_propagacion(**kwargs) -> Pipeline:
    return Pipeline([
        _nodo_modulo("propagacion", "sin_precio_ant.meta_propagacion"),
    ])


def create_pipeline(**kwargs) -> Pipeline:
    """Pipeline completo: los 4 módulos Sin Precio Anterior."""
    return (
        create_pipeline_agricolas()
        + create_pipeline_pecuarios()
        + create_pipeline_elementos()
        + create_pipeline_propagacion()
    )
