"""Tests del pipeline de agregación."""
import pandas as pd
import pytest

from sipsa_insumos.pipelines.aggregation.nodes import (
    aplicar_secreto_estadistico,
    calcular_precio_promedio,
)


class TestCalcularPrecioPromedio:
    def test_formula_promedio(self, base_enriquecida_minimal):
        result = calcular_precio_promedio(base_enriquecida_minimal)
        assert "PRECIO_PROMEDIO" in result.columns
        assert "N_ARTICULOS" in result.columns
        assert len(result) >= 1

    def test_promedio_correcto(self):
        df = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001"] * 3,
            "LLAVE_ARTICULO": ["CANECA|LITRO|20|0|0"] * 3,
            "Grupo": ["HERBICIDAS"] * 3,
            "Nombre_Publica": ["Glifosato"] * 3,
            "CÓDIGO CPC": ["3464100000"] * 3,
            "ARTÍCULO": ["GLIFOSATO"] * 3,
            "CodigoDepto": ["11"] * 3,
            "Departamento": ["Cundinamarca"] * 3,
            "PRECIO": [100.0, 110.0, 90.0],
        })
        result = calcular_precio_promedio(df)
        assert abs(result["PRECIO_PROMEDIO"].iloc[0] - 100.0) < 0.01
        assert result["N_ARTICULOS"].iloc[0] == 3

    def test_cuenta_fuentes(self):
        df = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001"] * 3,
            "LLAVE_ARTICULO": ["X|Y|1|0|0"] * 3,
            "Grupo": ["HERBICIDAS"] * 3,
            "Nombre_Publica": ["Producto"] * 3,
            "CÓDIGO CPC": ["123"] * 3,
            "ARTÍCULO": ["ART"] * 3,
            "PRECIO": [100.0, 200.0, 300.0],
        })
        result = calcular_precio_promedio(df)
        assert result["N_ARTICULOS"].iloc[0] == 3


class TestAplicarSecretoEstadistico:
    def test_n_mayor_igual_2_va_a_mayor2(self):
        df = pd.DataFrame({"N_ARTICULOS": [2, 3, 5], "PRECIO_PROMEDIO": [100.0] * 3})
        mayor2, menor2 = aplicar_secreto_estadistico(df)
        assert len(mayor2) == 3
        assert len(menor2) == 0

    def test_n_menor_2_va_a_menor2(self):
        df = pd.DataFrame({"N_ARTICULOS": [1], "PRECIO_PROMEDIO": [100.0]})
        mayor2, menor2 = aplicar_secreto_estadistico(df)
        assert len(mayor2) == 0
        assert len(menor2) == 1

    def test_corte_exacto_en_n2(self):
        df = pd.DataFrame({
            "N_ARTICULOS": [1, 2, 3],
            "PRECIO_PROMEDIO": [100.0] * 3,
        })
        mayor2, menor2 = aplicar_secreto_estadistico(df)
        assert len(mayor2) == 2
        assert len(menor2) == 1

    def test_total_conservado(self):
        df = pd.DataFrame({
            "N_ARTICULOS": [1, 2, 3, 4],
            "PRECIO_PROMEDIO": [100.0] * 4,
        })
        mayor2, menor2 = aplicar_secreto_estadistico(df)
        assert len(mayor2) + len(menor2) == len(df)
