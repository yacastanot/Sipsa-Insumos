"""Configuración del proyecto Kedro — SIPSA Insumos."""
from kedro.config import OmegaConfigLoader

from sipsa_insumos.utils.insumos import ruta_vigente_str

CONFIG_LOADER_CLASS = OmegaConfigLoader

CONFIG_LOADER_ARGS = {
    "base_env": "base",
    "default_run_env": "local",
    # ${vigente:'<ruta inicial>'} → versión ajustada del insumo si existe (utils/insumos.py)
    "custom_resolvers": {"vigente": ruta_vigente_str},
}
