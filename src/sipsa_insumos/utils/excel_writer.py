"""Helper centralizado para escribir archivos Excel multi-hoja con openpyxl."""
from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
_NUM_FORMAT_PRECIO = "#,##0"
_FONT_HEADER = Font(bold=True, color="112277")
_FILL_HEADER = PatternFill(start_color="EDF2F9", end_color="EDF2F9", fill_type="solid")
_BORDER_THIN = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def _round_sas(valor: float) -> int:
    """Redondea como SAS (mitad hacia arriba), no banker's rounding de Python."""
    return int(Decimal(repr(float(valor))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


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


def escribir_hoja_pivot_cpc(
    ws,
    df: pd.DataFrame,
    cpc_col: str,
    prod_col: str,
    metric_cols: list[str],
    sub_headers: list[str],
    grupos_header: list[tuple[str, int]],
    cols_punto_si_cero: set[str],
    cols_numero: set[str],
    total_label: str,
    cpc_header: str = "Codigo CPC",
    prod_header: str = "PRODUCTOS Y PRESENTACIONES",
    replicar_desfase_sas: bool = True,
) -> None:
    """Escribe una hoja con el layout tipo PROC TABULATE/REPORT de SAS
    (CUADROS/TABLAS): encabezado agrupado a 3 filas (grupo/sub-etiqueta en
    filas 1-2 sobre metric_cols; "Codigo CPC"/"PRODUCTOS Y PRESENTACIONES"
    en fila 3 sobre CPC/Producto), todas con relleno azul claro (#EDF2F9),
    texto azul oscuro (#112277) en negrita y borde fino, datos desde la
    fila 4, Código CPC fusionado verticalmente por grupo de producto, ceros
    mostrados como "." en las columnas de conteo (convención SAS), y fila
    TOTAL/Total final.

    Args:
        ws: worksheet de openpyxl donde escribir.
        df: filas ya ordenadas por [cpc_col, prod_col] — una fila por
            (Código CPC, Producto).
        cpc_col: nombre de columna con el Código CPC.
        prod_col: nombre de columna con el nombre del producto/publicación.
        metric_cols: columnas de métricas en el orden final de salida
            (empiezan en la columna C, después de CPC/Producto).
        sub_headers: etiquetas de la fila 2, alineadas 1:1 con metric_cols.
        grupos_header: [(texto, colspan), ...] para la fila 1, fusionados
            de izquierda a derecha sobre metric_cols; puede cubrir menos
            columnas que metric_cols (las restantes quedan sin etiqueta,
            igual que la columna "Total" de TABLAS).
        cols_punto_si_cero: subconjunto de metric_cols donde 0/NaN se
            muestra como "." (conteos de tendencia — Subió/Bajó/Estable-n.d.).
        cols_numero: subconjunto de metric_cols que siempre se muestran
            como número (nunca "."; ej. N, Total). Columnas de precio
            (MIN/MAX) también van aquí — se detectan por nombre para el
            cálculo del total (min/max en vez de suma) y formato numérico.
        total_label: texto de la fila final ("TOTAL" en CUADROS, "Total"
            en TABLAS — SAS usa mayúsculas distintas en cada uno).
        cpc_col: None para módulos sin Código CPC (Arriendos/Servicios): el
            producto va en la columna A y las métricas desde la B.
        replicar_desfase_sas: SAS escribe los valores del primer producto de
            la hoja en la fila de encabezado "Codigo CPC / PRODUCTOS..." y deja
            vacía la fila del producto. Se replica para paridad exacta.
    """
    if cpc_col is None:
        col_cpc, col_prod = None, 1
    else:
        col_cpc, col_prod = 1, 2
    col_metric_start = col_prod + 1
    n_cols = col_metric_start + len(metric_cols) - 1

    for r in (1, 2):
        for col in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=col)
            cell.font = _FONT_HEADER
            cell.fill = _FILL_HEADER
            cell.border = _BORDER_THIN

    c = col_metric_start
    for texto, span in grupos_header:
        if span > 1:
            ws.merge_cells(start_row=1, start_column=c, end_row=1, end_column=c + span - 1)
        cell = ws.cell(row=1, column=c, value=texto)
        cell.alignment = _ALIGN_CENTER
        c += span

    for i, texto in enumerate(sub_headers):
        cell = ws.cell(row=2, column=col_metric_start + i, value=texto)
        cell.alignment = _ALIGN_CENTER

    for col, texto in ((col_cpc, cpc_header), (col_prod, prod_header)):
        if col is None:
            continue
        cell = ws.cell(row=3, column=col, value=texto)
        cell.font = _FONT_HEADER
        cell.fill = _FILL_HEADER
        cell.border = _BORDER_THIN
        cell.alignment = _ALIGN_CENTER

    fila = 4
    totales_suma: dict[str, float] = {c: 0 for c in cols_punto_si_cero | cols_numero}
    totales_min: dict[str, float] = {}
    totales_max: dict[str, float] = {}

    grupos_cpc = (df.groupby(cpc_col, sort=False, dropna=False) if cpc_col is not None
                  else [(None, df)])
    for cpc, grupo in grupos_cpc:
        fila_inicio = fila
        for _, row in grupo.iterrows():
            ws.cell(row=fila, column=col_prod, value=row[prod_col])
            for i, mcol in enumerate(metric_cols):
                valor = row[mcol]
                col_idx = col_metric_start + i
                es_precio = "MIN" in mcol.upper() or "MAX" in mcol.upper()
                if es_precio:
                    # SAS redondea precios a pesos enteros en este reporte impreso
                    v = None if pd.isna(valor) else _round_sas(valor)
                    ws.cell(row=fila, column=col_idx, value=v)
                    if v is not None:
                        if "MIN" in mcol.upper():
                            totales_min[mcol] = v if mcol not in totales_min else min(totales_min[mcol], v)
                        else:
                            totales_max[mcol] = v if mcol not in totales_max else max(totales_max[mcol], v)
                    ws.cell(row=fila, column=col_idx).number_format = _NUM_FORMAT_PRECIO
                elif mcol in cols_punto_si_cero:
                    es_cero = pd.isna(valor) or valor == 0
                    ws.cell(row=fila, column=col_idx, value="." if es_cero else int(valor))
                    totales_suma[mcol] += 0 if pd.isna(valor) else valor
                else:
                    ws.cell(row=fila, column=col_idx, value=int(valor) if not pd.isna(valor) else 0)
                    totales_suma[mcol] += 0 if pd.isna(valor) else valor
            fila += 1

        if col_cpc is None:
            continue
        cpc_val = cpc
        try:
            cpc_val = int(float(cpc))
        except (TypeError, ValueError):
            pass
        ws.cell(row=fila_inicio, column=col_cpc, value=cpc_val if pd.notna(cpc) else None)
        if fila - fila_inicio > 1:
            ws.merge_cells(start_row=fila_inicio, start_column=col_cpc, end_row=fila - 1, end_column=col_cpc)
            ws.cell(row=fila_inicio, column=col_cpc).alignment = _ALIGN_CENTER

    if col_cpc is not None:
        ws.merge_cells(start_row=fila, start_column=col_cpc, end_row=fila, end_column=col_prod)
    ws.cell(row=fila, column=col_cpc or col_prod, value=total_label)
    for i, mcol in enumerate(metric_cols):
        col_idx = col_metric_start + i
        if "MIN" in mcol.upper():
            ws.cell(row=fila, column=col_idx, value=totales_min.get(mcol))
            ws.cell(row=fila, column=col_idx).number_format = _NUM_FORMAT_PRECIO
        elif "MAX" in mcol.upper():
            ws.cell(row=fila, column=col_idx, value=totales_max.get(mcol))
            ws.cell(row=fila, column=col_idx).number_format = _NUM_FORMAT_PRECIO
        else:
            ws.cell(row=fila, column=col_idx, value=int(totales_suma.get(mcol, 0)))

    if replicar_desfase_sas and fila > 4:
        for i in range(len(metric_cols)):
            origen = ws.cell(row=4, column=col_metric_start + i)
            destino = ws.cell(row=3, column=col_metric_start + i, value=origen.value)
            destino.number_format = origen.number_format
            origen.value = None


def escribir_hoja_freq(
    ws,
    df: pd.DataFrame,
    prod_col: str,
    var_label: str,
    metric_cols: list[str],
    sub_headers: list[str],
) -> None:
    """Escribe una tabla de contingencia con el layout de PROC FREQ de SAS
    (TABLAS de Arriendos/Servicios): título "Procedimiento FREQ", fila
    "Tabla de <var> por Tendencia_mod", encabezado "Frecuencia" y fila Total.
    Los ceros se muestran como 0 (no "."), igual que PROC FREQ.

    Args:
        ws: worksheet de openpyxl.
        df: una fila por producto, ya ordenada (PROC FREQ ordena por valor).
        prod_col: columna con el nombre del producto.
        var_label: nombre SAS de la variable fila (p.ej. Nombre_productos_servicios_publi).
        metric_cols: columnas de conteo en orden de salida (incluye Total al final).
        sub_headers: etiquetas de la fila "Frecuencia" (1.Positiva ... Total).
    """
    n_cols = 1 + len(metric_cols)
    nbsp = " "  # SAS separa el título con espacios no separables
    ws.cell(row=1, column=1, value="Procedimiento FREQ")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws.cell(row=3, column=1, value=nbsp.join(["Tabla", "de", var_label, "por", "Tendencia_mod"]))
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=n_cols)
    ws.cell(row=4, column=1, value=f"{var_label}({var_label})")
    ws.cell(row=4, column=2, value="Tendencia_mod")
    ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=n_cols)
    ws.cell(row=5, column=1, value="Frecuencia")
    for i, texto in enumerate(sub_headers):
        ws.cell(row=5, column=2 + i, value=texto)
    fila = 6
    totales = [0] * len(metric_cols)
    for _, row in df.iterrows():
        ws.cell(row=fila, column=1, value=row[prod_col])
        for i, mcol in enumerate(metric_cols):
            v = 0 if pd.isna(row[mcol]) else int(row[mcol])
            ws.cell(row=fila, column=2 + i, value=v)
            totales[i] += v
        fila += 1
    ws.cell(row=fila, column=1, value="Total")
    for i, v in enumerate(totales):
        ws.cell(row=fila, column=2 + i, value=v)
