"""Versión vigente de los insumos de un período: inicial o ajustada.

Flujo mensual:
  1. Llegan las bases livianas, las DIVIPOLA y los archivos de Sin Precio
     Anterior → procesamiento preliminar (reportes municipales) → temática.
  2. Temática revisa y, algunos meses, reenvía bases livianas y/o DIVIPOLA
     ajustadas → se corren de nuevo los reportes municipales y la serie
     departamental (ejecución definitiva).

Organización en raw:

  data/01_raw/<P>/BASES LIVIANAS <P>/...             ← versión inicial
  data/01_raw/<P>/DIVIPOLA <P>/...                   ← versión inicial
  data/01_raw/<P>/AJUSTADOS <P>/BASES LIVIANAS <P>/  ← versión ajustada (mismo nombre de archivo)
  data/01_raw/<P>/AJUSTADOS <P>/DIVIPOLA <P>/        ← versión ajustada (mismo nombre de archivo)

`ruta_vigente` recibe la ruta de la versión inicial y devuelve la ajustada si
existe un archivo con el mismo nombre; si no, la inicial. Así un módulo sin
ajustes sigue usando su base inicial y no hay que tocar parámetros.
"""
from __future__ import annotations

import logging
import unicodedata
from pathlib import Path

log = logging.getLogger(__name__)


def _norm(nombre: str) -> str:
    return unicodedata.normalize("NFC", nombre).lower()


def carpeta_ajustados(periodo: str, raw: str | Path = "data/01_raw") -> Path:
    return Path(raw) / periodo / f"AJUSTADOS {periodo}"


def ruta_vigente(ruta: str | Path) -> Path:
    """Ruta ajustada del insumo si existe; si no, la ruta inicial recibida.

    Solo aplica a archivos ubicados en data/01_raw/<P>/<subcarpeta>/; cualquier
    otra ruta (o una ruta vacía) se devuelve sin cambios.
    """
    p = Path(ruta)
    if not str(ruta) or len(p.parts) < 3 or p.parent.parent.parent.name != "01_raw":
        return p
    periodo, subcarpeta = p.parent.parent.name, p.parent.name
    carpeta = carpeta_ajustados(periodo, p.parent.parent.parent) / subcarpeta
    if carpeta.is_dir():
        for f in carpeta.iterdir():
            if _norm(f.name) == _norm(p.name) and not f.name.startswith("~$"):
                log.info("Insumo AJUSTADO | %s", f)
                return f
    return p


def ruta_vigente_str(ruta: str) -> str:
    """Variante para el resolver de OmegaConf (`${vigente:'...'}` en el catálogo)."""
    return ruta_vigente(ruta).as_posix()


def version_insumo(ruta: str | Path) -> str:
    return "ajustada" if any(parte.startswith("AJUSTADOS ") for parte in Path(ruta).parts) else "inicial"
