"""Tests unitarios para sipsa_insumos.utils.parsers."""
import pandas as pd
import pytest

from sipsa_insumos.utils.parsers import (
    agregar_columnas_unidad_medida,
    construir_llave_articulo,
    parsear_unidad_medida,
)


class TestParsearUnidadMedida:
    def test_pipe_completo(self):
        result = parsear_unidad_medida("BULTO|KILOGRAMO|50|0|0")
        assert result == {"nombre_um": "BULTO", "unidad": "KILOGRAMO", "cantidad": "50"}

    def test_frasco_litro(self):
        result = parsear_unidad_medida("FRASCO|LITRO|1|0|0")
        assert result == {"nombre_um": "FRASCO", "unidad": "LITRO", "cantidad": "1"}

    def test_sin_pipe(self):
        result = parsear_unidad_medida("SOLO_NOMBRE")
        assert result["nombre_um"] == "SOLO_NOMBRE"
        assert result["unidad"] == ""
        assert result["cantidad"] == ""

    def test_un_pipe(self):
        result = parsear_unidad_medida("JORNAL|DIA")
        assert result["nombre_um"] == "JORNAL"
        assert result["unidad"] == "DIA"
        assert result["cantidad"] == ""

    def test_valor_none_devuelve_vacios(self):
        result = parsear_unidad_medida(None)
        assert result == {"nombre_um": "", "unidad": "", "cantidad": ""}

    def test_valor_numerico_devuelve_vacios(self):
        result = parsear_unidad_medida(42)
        assert result == {"nombre_um": "", "unidad": "", "cantidad": ""}

    def test_espacios_se_eliminan(self):
        result = parsear_unidad_medida("  BULTO  |  KILOGRAMO  |  50  |0|0")
        assert result == {"nombre_um": "BULTO", "unidad": "KILOGRAMO", "cantidad": "50"}


class TestConstruirLlaveArticulo:
    def test_llave_estandar(self):
        llave = construir_llave_articulo("GLIFOSATO 480 SL", "FRASCO|LITRO|1|0|0")
        assert llave == "GLIFOSATO 480 SL_FRASCO|LITRO|1|0|0"

    def test_normaliza_a_mayusculas(self):
        llave = construir_llave_articulo("glifosato 480 sl", "frasco|litro|1|0|0")
        assert llave == "GLIFOSATO 480 SL_FRASCO|LITRO|1|0|0"

    def test_capitalizado_produce_misma_llave(self):
        llave1 = construir_llave_articulo("Glifosato 480 Sl", "Frasco|Litro|1|0|0")
        llave2 = construir_llave_articulo("GLIFOSATO 480 SL", "FRASCO|LITRO|1|0|0")
        assert llave1 == llave2

    def test_elimina_espacios(self):
        llave = construir_llave_articulo("  UREA  ", "  BULTO|KILOGRAMO|50|0|0  ")
        assert llave == "UREA_BULTO|KILOGRAMO|50|0|0"

    def test_separador_es_guion_bajo(self):
        llave = construir_llave_articulo("ARTICULO", "FRASCO|LITRO|1|0|0")
        assert "_" in llave
        partes = llave.split("_", 1)
        assert partes[0] == "ARTICULO"
        assert partes[1] == "FRASCO|LITRO|1|0|0"

    def test_formato_divipola_real(self):
        # Ejemplo real del DIVIPOLA agrícolas mayo 2026
        llave = construir_llave_articulo("2,4-D AMINA 720 SL MEZFER", "FRASCO|LITRO|1|0|0")
        assert llave == "2,4-D AMINA 720 SL MEZFER_FRASCO|LITRO|1|0|0"


class TestAgregarColumnasUnidadMedida:
    def test_agrega_cuatro_columnas(self):
        df = pd.DataFrame({
            "ARTÍCULO": ["UREA"],
            "UNIDAD DE MEDIDA": ["BULTO|KILOGRAMO|50|0|0"],
        })
        result = agregar_columnas_unidad_medida(df)
        assert "NOMBRE_UM" in result.columns
        assert "UNIDAD" in result.columns
        assert "CANTIDAD" in result.columns
        assert "LLAVE_ARTICULO" in result.columns

    def test_no_modifica_dataframe_original(self):
        df = pd.DataFrame({
            "ARTÍCULO": ["UREA"],
            "UNIDAD DE MEDIDA": ["BULTO|KILOGRAMO|50|0|0"],
        })
        original_cols = list(df.columns)
        agregar_columnas_unidad_medida(df)
        assert list(df.columns) == original_cols

    def test_llave_coincide_con_construir_llave(self):
        df = pd.DataFrame({
            "ARTÍCULO": ["Glifosato 480 SL"],
            "UNIDAD DE MEDIDA": ["frasco|litro|1|0|0"],
        })
        result = agregar_columnas_unidad_medida(df)
        assert result["LLAVE_ARTICULO"].iloc[0] == "GLIFOSATO 480 SL_FRASCO|LITRO|1|0|0"

    def test_multiples_filas(self):
        df = pd.DataFrame({
            "ARTÍCULO": ["UREA", "GLIFOSATO", "JORNAL"],
            "UNIDAD DE MEDIDA": [
                "BULTO|KILOGRAMO|50|0|0",
                "FRASCO|LITRO|1|0|0",
                "JORNAL|DIA|1|0|0",
            ]
        })
        result = agregar_columnas_unidad_medida(df)
        assert len(result) == 3
        assert result["NOMBRE_UM"].tolist() == ["BULTO", "FRASCO", "JORNAL"]
        assert result["CANTIDAD"].tolist() == ["50", "1", "1"]
        assert result["LLAVE_ARTICULO"].tolist() == [
            "UREA_BULTO|KILOGRAMO|50|0|0",
            "GLIFOSATO_FRASCO|LITRO|1|0|0",
            "JORNAL_JORNAL|DIA|1|0|0",
        ]

    def test_pipe_incompleto_no_falla(self):
        df = pd.DataFrame({
            "ARTÍCULO": ["PRODUCTO"],
            "UNIDAD DE MEDIDA": ["SOLO_NOMBRE"],
        })
        result = agregar_columnas_unidad_medida(df)
        assert result["NOMBRE_UM"].iloc[0] == "SOLO_NOMBRE"
        assert result["UNIDAD"].iloc[0] == ""
