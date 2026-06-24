"""Tests del pipeline de calidad."""
import pandas as pd
import pytest

from sipsa_insumos.pipelines.quality.nodes import (
    calcular_cv,
    detectar_duplicados,
    detectar_var_atipica,
)


class TestDetectarDuplicados:
    def test_sin_duplicados(self, base_enriquecida_minimal):
        sin_dupli, dupli = detectar_duplicados(base_enriquecida_minimal)
        assert len(sin_dupli) == 4
        assert len(dupli) == 0

    def test_detecta_duplicado(self, base_enriquecida_minimal):
        df_dupli = pd.concat([base_enriquecida_minimal, base_enriquecida_minimal.iloc[[0]]])
        sin_dupli, dupli = detectar_duplicados(df_dupli)
        assert len(dupli) == 1
        assert len(sin_dupli) == 4

    def test_mantiene_primer_registro(self, base_enriquecida_minimal):
        df_dupli = pd.concat([base_enriquecida_minimal, base_enriquecida_minimal.iloc[[0]]])
        sin_dupli, _ = detectar_duplicados(df_dupli)
        assert len(sin_dupli) == 4


class TestCalcularCV:
    def test_cv_formula(self):
        df = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001"] * 4,
            "LLAVE_ARTICULO": ["CANECA|LITRO|20|0|0"] * 4,
            "PRECIO": [100.0, 110.0, 90.0, 100.0],
        })
        base_con_cv, cvs = calcular_cv(df)
        cv_calculado = cvs["CV"].iloc[0]
        std = pd.Series([100.0, 110.0, 90.0, 100.0]).std()
        mean = 100.0
        cv_esperado = std / mean * 100
        assert abs(cv_calculado - cv_esperado) < 0.01

    def test_n_menor_2_cv_nan(self):
        df = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001"],
            "LLAVE_ARTICULO": ["CANECA|LITRO|20|0|0"],
            "PRECIO": [100.0],
        })
        _, cvs = calcular_cv(df)
        assert pd.isna(cvs["CV"].iloc[0])

    def test_agrega_columna_cv(self, base_enriquecida_minimal):
        base_con_cv, _ = calcular_cv(base_enriquecida_minimal)
        assert "CV" in base_con_cv.columns

    def test_no_modifica_filas(self, base_enriquecida_minimal):
        base_con_cv, _ = calcular_cv(base_enriquecida_minimal)
        assert len(base_con_cv) == len(base_enriquecida_minimal)


class TestDetectarVarAtipica:
    def _df_con_cv(self):
        return pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001", "05001"],
            "LLAVE_ARTICULO": ["CANECA|LITRO|20|0|0"] * 2,
            "PRECIO": [45000.0, 46000.0],
            "CV": [5.0, 5.0],
        })

    def _anterior(self):
        return pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001", "05001"],
            "LLAVE_ARTICULO": ["CANECA|LITRO|20|0|0"] * 2,
            "PRECIO_PROMEDIO": [42000.0, 44000.0],
        })

    def test_sin_anterior_revisa_2(self):
        df = self._df_con_cv()
        base, atipico = detectar_var_atipica(df, None, 25.0, -25.0, 100.0)
        assert (base["REVISA"] == 2).all()

    def test_variacion_alta_revisa_1(self):
        df = self._df_con_cv()
        anterior = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001", "05001"],
            "LLAVE_ARTICULO": ["CANECA|LITRO|20|0|0"] * 2,
            "PRECIO_PROMEDIO": [30000.0, 30000.0],  # +50% → REVISA=1
        })
        base, atipico = detectar_var_atipica(df, anterior, 25.0, -25.0, 100.0)
        assert (base["REVISA"] == 1).any()

    def test_variacion_extrema_revisa_3(self):
        df = self._df_con_cv()
        anterior = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001", "05001"],
            "LLAVE_ARTICULO": ["CANECA|LITRO|20|0|0"] * 2,
            "PRECIO_PROMEDIO": [20000.0, 20000.0],  # +125% → REVISA=3
        })
        base, atipico = detectar_var_atipica(df, anterior, 25.0, -25.0, 100.0)
        assert (base["REVISA"] == 3).any()

    def test_sin_variacion_revisa_0(self):
        df = self._df_con_cv()
        anterior = pd.DataFrame({
            "CÓDIGO DIVIPOLA": ["11001", "05001"],
            "LLAVE_ARTICULO": ["CANECA|LITRO|20|0|0"] * 2,
            "PRECIO_PROMEDIO": [45000.0, 46000.0],  # Sin variación
        })
        base, atipico = detectar_var_atipica(df, anterior, 25.0, -25.0, 100.0)
        assert (base["REVISA"] == 0).all()
