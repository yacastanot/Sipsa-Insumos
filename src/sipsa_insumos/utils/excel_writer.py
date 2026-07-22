"""Helper centralizado para escribir archivos Excel multi-hoja con openpyxl."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter


_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
_NUM_FORMAT_PRECIO = "#,##0"


def escribir_excel_multisheet(
    ruta: str | Path,
    hojas: dict[str, pd.DataFrame],
) -> None:
    """Escribe un xlsx multi-hoja sin formato especial (paridad con SAS).

    Args:
        ruta: Ruta destino del archivo .xlsx.
        hojas: Diccionario {nombre_hoja: DataFrame}.
    """
    Path(ruta).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(str(ruta), engine="openpyxl") as writer:
        for nombre_hoja, df in hojas.items():
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)


def aplicar_formato_numerico_precio(ws, col_idx: int, n_filas: int) -> None:
    """Aplica formato de precio colombiano (#,##0) a una columna."""
    for row_idx in range(2, n_filas + 2):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.number_format = _NUM_FORMAT_PRECIO
        cell.alignment = _ALIGN_CENTER
