"""Helper centralizado para escribir archivos Excel multi-hoja con openpyxl."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


_FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
_FILL_HEADER = PatternFill("solid", fgColor="FF4472C4")
_ALIGN_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_FONT_DATA = Font(name="Calibri", size=10)
_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
_NUM_FORMAT_PRECIO = "#,##0"


def escribir_excel_multisheet(
    ruta: str | Path,
    hojas: dict[str, pd.DataFrame],
) -> None:
    """Escribe un xlsx multi-hoja con formato estándar DANE.

    Aplica para cada hoja:
    - Encabezados en negrita, fondo azul, texto blanco, centrado.
    - Datos: fuente Calibri 10, alineación izquierda.
    - Autofit de columnas (aproximado por contenido).
    - freeze_panes en A2.

    Args:
        ruta: Ruta destino del archivo .xlsx.
        hojas: Diccionario {nombre_hoja: DataFrame}.
    """
    Path(ruta).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(str(ruta), engine="openpyxl") as writer:
        for nombre_hoja, df in hojas.items():
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            ws = writer.sheets[nombre_hoja]
            _aplicar_formato_hoja(ws, df)


def _aplicar_formato_hoja(ws, df: pd.DataFrame) -> None:
    """Aplica el formato estándar a una hoja de cálculo."""
    ws.freeze_panes = "A2"

    # Encabezados — fila 1
    for col_idx in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = _FONT_HEADER
        cell.fill = _FILL_HEADER
        cell.alignment = _ALIGN_HEADER

    # Datos — resto de filas
    for row_idx in range(2, len(df) + 2):
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = _FONT_DATA
            cell.alignment = _ALIGN_LEFT

    # Autofit: ancho aprox por contenido (max 60 chars)
    for col_idx, col_name in enumerate(df.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = max(
            len(str(col_name)),
            df.iloc[:, col_idx - 1].astype(str).str.len().max() if len(df) > 0 else 0,
        )
        ws.column_dimensions[col_letter].width = min(max_len + 2, 60)


def aplicar_formato_numerico_precio(ws, col_idx: int, n_filas: int) -> None:
    """Aplica formato de precio colombiano (#,##0) a una columna."""
    for row_idx in range(2, n_filas + 2):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.number_format = _NUM_FORMAT_PRECIO
        cell.alignment = _ALIGN_CENTER
