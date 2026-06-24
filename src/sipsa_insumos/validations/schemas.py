"""Schemas Pandera para validación de datos de entrada — SIPSA Insumos."""
from __future__ import annotations

import pandera.pandas as pa

# Columnas requeridas en la Base Liviana de cualquier módulo de insumos.
COLUMNAS_REQUERIDAS: set[str] = {
    "MES_AÑO",
    "CÓDIGO DIVIPOLA",
    "CÓDIGO CPC",
    "ARTÍCULO",
    "CASA COMERCIAL",
    "REGISTRO ICA",
    "UNIDAD DE MEDIDA",
    "PRECIO",
}

SCHEMA_BASE_LIVIANA = pa.DataFrameSchema(
    columns={
        "CÓDIGO DIVIPOLA": pa.Column(
            str,
            nullable=False,
            checks=pa.Check.str_length(5, 5),
            description="Código DIVIPOLA de 5 dígitos — debe leerse con dtype=str",
        ),
        "CÓDIGO CPC": pa.Column(
            str,
            nullable=False,
            description="Código CPC — debe leerse con dtype=str (evita notación científica)",
        ),
        "PRECIO": pa.Column(
            float,
            nullable=False,
            checks=pa.Check.gt(0),
            coerce=True,
            description="Precio unitario en pesos colombianos — debe ser positivo",
        ),
        "UNIDAD DE MEDIDA": pa.Column(
            str,
            nullable=False,
            checks=pa.Check.str_contains("|"),
            description="Formato pipe: NOMBRE|UNIDAD|CANTIDAD|0|0",
        ),
        "ARTÍCULO": pa.Column(str, nullable=False),
        "CASA COMERCIAL": pa.Column(str, nullable=True),
        "REGISTRO ICA": pa.Column(str, nullable=True, coerce=True),
        "MES_AÑO": pa.Column(str, nullable=False),
    },
    strict=False,   # columnas adicionales son bienvenidas
    coerce=False,
    name="BaseLivianaRaw",
)

SCHEMA_DIVIPOLA = pa.DataFrameSchema(
    columns={
        "CodigoMpio": pa.Column(
            str,
            nullable=False,
            checks=pa.Check.str_length(5, 5),
            description="Código municipio DIVIPOLA de 5 dígitos",
        ),
        "Departamento": pa.Column(str, nullable=False),
    },
    strict=False,
    name="DivopolaRaw",
)
