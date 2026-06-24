"""Tests del pipeline de enriquecimiento."""
import pandas as pd
import pytest

from sipsa_insumos.pipelines.enrichment.nodes import (
    asignar_articulo_publica,
    asignar_grupo,
    merge_divipola,
)
from sipsa_insumos.utils.parsers import agregar_columnas_unidad_medida


class TestMergeDivipola:
    def test_match_completo(self, base_liviana_minimal, divipola_minimal):
        df = agregar_columnas_unidad_medida(base_liviana_minimal)
        base_con_mpio, faltan = merge_divipola(df, divipola_minimal)
        assert len(base_con_mpio) == 4
        assert len(faltan) == 0

    def test_sin_match_genera_faltan(self):
        base = pd.DataFrame({
            "MES_AÑO": ["MAY2026"],
            "CÓDIGO DIVIPOLA": ["99998"],
            "PRECIO": [100.0],
            "LLAVE_ARTICULO": ["X|Y|1|0|0"],
        })
        divipola = pd.DataFrame({
            "CodigoMpio": ["11001"],
            "CodigoDepto": ["11"],
            "Departamento": ["Cundinamarca"],
        })
        base_con_mpio, faltan = merge_divipola(base, divipola)
        assert len(faltan) == 1
        assert len(base_con_mpio) == 0

    def test_match_parcial(self, divipola_minimal):
        base = pd.DataFrame({
            "MES_AÑO": ["MAY2026", "MAY2026"],
            "CÓDIGO DIVIPOLA": ["11001", "99998"],
            "PRECIO": [100.0, 200.0],
            "LLAVE_ARTICULO": ["X|Y|1|0|0", "X|Y|1|0|0"],
        })
        base_con_mpio, faltan = merge_divipola(base, divipola_minimal)
        assert len(base_con_mpio) == 1
        assert len(faltan) == 1

    def test_agrega_columnas_divipola(self, base_liviana_minimal, divipola_minimal):
        df = agregar_columnas_unidad_medida(base_liviana_minimal)
        base_con_mpio, _ = merge_divipola(df, divipola_minimal)
        assert "CodigoDepto" in base_con_mpio.columns
        assert "Departamento" in base_con_mpio.columns


class TestAsignarGrupo:
    def test_match_directo(self, mappings_grupos_minimal):
        # Clave real: ARTÍCULO_UNIDAD_DE_MEDIDA
        df = pd.DataFrame({"LLAVE_ARTICULO": ["GLIFOSATO 480 SL_CANECA|LITRO|20|0|0"]})
        con_grupo, faltan = asignar_grupo(df, mappings_grupos_minimal, "AGRICOLAS")
        assert len(con_grupo) == 1
        assert con_grupo["Grupo"].iloc[0] == "HERBICIDAS"
        assert len(faltan) == 0

    def test_sin_match_va_a_faltan(self, mappings_grupos_minimal):
        df = pd.DataFrame({"LLAVE_ARTICULO": ["DESCONOCIDO_FRASCO|LITRO|1|0|0"]})
        con_grupo, faltan = asignar_grupo(df, mappings_grupos_minimal, "AGRICOLAS")
        assert len(con_grupo) == 0
        assert len(faltan) == 1

    def test_parcial_match(self, mappings_grupos_minimal):
        df = pd.DataFrame({
            "LLAVE_ARTICULO": [
                "GLIFOSATO 480 SL_CANECA|LITRO|20|0|0",
                "DESCONOCIDO_FRASCO|LITRO|1|0|0",
            ]
        })
        con_grupo, faltan = asignar_grupo(df, mappings_grupos_minimal, "AGRICOLAS")
        assert len(con_grupo) == 1
        assert len(faltan) == 1


class TestAsignarArticuloPublica:
    def test_match_directo(self, mappings_articulos_minimal):
        df = pd.DataFrame({
            "LLAVE_ARTICULO": ["GLIFOSATO 480 SL_CANECA|LITRO|20|0|0"],
            "Grupo": ["HERBICIDAS"],
        })
        con_nombre, faltan = asignar_articulo_publica(df, mappings_articulos_minimal, "AGRICOLAS")
        assert len(con_nombre) == 1
        assert con_nombre["Nombre_Publica"].iloc[0] == "Glifosato 480 SL (20 L)"

    def test_sin_match(self, mappings_articulos_minimal):
        df = pd.DataFrame({
            "LLAVE_ARTICULO": ["SIN_MATCH_FRASCO|LITRO|1|0|0"],
            "Grupo": ["HERBICIDAS"],
        })
        con_nombre, faltan = asignar_articulo_publica(df, mappings_articulos_minimal, "AGRICOLAS")
        assert len(con_nombre) == 0
        assert len(faltan) == 1
