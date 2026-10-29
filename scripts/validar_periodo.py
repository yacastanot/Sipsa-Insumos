"""Validación de cierre del período (ver src/sipsa_insumos/utils/validacion.py).

Uso:
  python scripts/validar_periodo.py            # período de conf/base/globals.yml
  python scripts/validar_periodo.py SEP2026

Imprime el resumen (solo alertas y errores) y guarda
data/08_reporting/<P>/VALIDACION_<P>.json y .xlsx. Sale con código 1 si hay errores.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from sipsa_insumos.utils.validacion import texto_resumen, validar_periodo  # noqa: E402

if __name__ == "__main__":
    out = validar_periodo(sys.argv[1] if len(sys.argv) > 1 else None, root=ROOT)
    print("\n".join(texto_resumen(out)))
    sys.exit(1 if out["estado"] == "error" else 0)
