"""Tests del pipeline de comparación interperiódica."""
import pandas as pd
import pytest

from sipsa_insumos.pipelines.comparison.nodes import calcular_variacion_tendencia


def _mayor2(precios: list) -> pd.DataFrame:
    return pd.DataFrame({
        "CÓDIGO DIVIPOLA": [f"1100{i+1}" for i in range(len(precios))],
        "LLAVE_ARTICULO": ["GLIFOSATO 480 SL_CANECA|LITRO|20|0|0"] * len(precios),
        "PRECIO_PROMEDIO": precios,
        "N_ARTICULOS": [3] * len(precios),
    })


class TestCalcularVariacionTendencia:
    def test_sin_anterior_tendencia_nd(self):
        mayor2 = _mayor2([45000.0, 46000.0])
        result = calcular_variacion_tendencia(mayor2, None, "Mayo", "Abril")
        assert (result["TENDENCIA"] == "n.d.").all()

    def test_anterior_vacio_tendencia_nd(self):
        mayor2 = _mayor2([45000.0])
        anterior = pd.DataFrame(columns=["CÓDIGO DIVIPOLA", "LLAVE_ARTICULO", "PRECIO_PROMEDIO"])
        result = calcular_variacion_tendencia(mayor2, anterior, "Mayo", "Abril")
        assert (result["TENDENCIA"] == "n.d.").all()

    def test_variacion_positiva(self):
        mayor2 = _mayor2([50000.0])
        anterior = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001"],
            "LLAVE_ARTICULO": ["GLIFOSATO 480 SL_CANECA|LITRO|20|0|0"],
            "PRECIO_PROMEDIO": [40000.0],
        })
        result = calcular_variacion_tendencia(mayor2, anterior, "Mayo", "Abril")
        assert result["TENDENCIA"].iloc[0] == "Positiva"
        assert result["VARIACION"].iloc[0] == pytest.approx(25.0)

    def test_variacion_negativa(self):
        mayor2 = _mayor2([36000.0])
        anterior = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001"],
            "LLAVE_ARTICULO": ["GLIFOSATO 480 SL_CANECA|LITRO|20|0|0"],
            "PRECIO_PROMEDIO": [40000.0],
        })
        result = calcular_variacion_tendencia(mayor2, anterior, "Mayo", "Abril")
        assert result["TENDENCIA"].iloc[0] == "Negativa"
        assert result["VARIACION"].iloc[0] == pytest.approx(-10.0)

    def test_sin_dato_anterior_tendencia_nd(self):
        mayor2 = _mayor2([45000.0])
        anterior = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["99999"],  # Municipio diferente → sin match
            "LLAVE_ARTICULO": ["GLIFOSATO 480 SL_CANECA|LITRO|20|0|0"],
            "PRECIO_PROMEDIO": [40000.0],
        })
        result = calcular_variacion_tendencia(mayor2, anterior, "Mayo", "Abril")
        assert result["TENDENCIA"].iloc[0] == "n.d."

    def test_agrega_columna_precio_mes_actual(self):
        mayor2 = _mayor2([45000.0])
        result = calcular_variacion_tendencia(mayor2, None, "Mayo", "Abril")
        assert "PRECIO_Mayo" in result.columns

    def test_agrega_columna_precio_mes_anterior(self):
        mayor2 = _mayor2([45000.0])
        anterior = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001"],
            "LLAVE_ARTICULO": ["GLIFOSATO 480 SL_CANECA|LITRO|20|0|0"],
            "PRECIO_PROMEDIO": [42000.0],
        })
        result = calcular_variacion_tendencia(mayor2, anterior, "Mayo", "Abril")
        assert "PRECIO_Abril" in result.columns
