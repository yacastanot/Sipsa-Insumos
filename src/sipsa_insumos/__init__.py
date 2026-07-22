"""Paquete sipsa_insumos — Migración SIPSA Insumos SAS → Python + Kedro."""
import pandas.io.formats.excel

# Pandas aplica negrilla al encabezado por defecto al escribir con openpyxl;
# se desactiva aquí para que todos los .xlsx de salida queden sin formato (paridad SAS).
pandas.io.formats.excel.ExcelFormatter.header_style = None
