"""Entry point para ejecutar el proyecto como módulo: python -m sipsa_insumos."""
from kedro.framework.cli.utils import find_run_command

run = find_run_command("sipsa_insumos")

if __name__ == "__main__":
    run()
