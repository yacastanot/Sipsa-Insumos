"""Fixtures compartidas para todos los tests de SIPSA Insumos."""
import pytest
import pandas as pd


@pytest.fixture
def base_liviana_minimal() -> pd.DataFrame:
    """DataFrame mínimo con la estructura de la Base Liviana."""
    return pd.DataFrame({
        "MES_AÑO":          ["MAY2026"] * 4,
        "CÓDIGO DIVIPOLA":  ["11001", "05001", "76001", "08001"],
        "CÓDIGO CPC":       ["3464100000"] * 4,
        "ARTÍCULO":         ["GLIFOSATO 480 SL"] * 4,
        "CASA COMERCIAL":   ["MONSANTO COLOMBIA"] * 4,
        "REGISTRO ICA":     ["ICA-1234"] * 4,
        "UNIDAD DE MEDIDA": ["CANECA|LITRO|20|0|0"] * 4,
        "PRECIO":           [45000.0, 46000.0, 44000.0, 47000.0],
    })
    # LLAVE_ARTICULO esperada: "GLIFOSATO 480 SL_CANECA|LITRO|20|0|0"


@pytest.fixture
def divipola_minimal() -> pd.DataFrame:
    """DataFrame mínimo con la estructura de DIVIPOLA."""
    return pd.DataFrame({
        "CodigoMpio":  ["11001", "05001", "76001", "08001", "99999"],
        "CodigoDepto": ["11", "05", "76", "08", "99"],
        "Departamento": [
            "Cundinamarca", "Antioquia", "Valle del Cauca", "Atlántico", "Test"
        ],
    })


@pytest.fixture
def mappings_grupos_minimal() -> dict:
    """Mapping mínimo de grupos para tests. Clave = ARTÍCULO_UNIDAD_DE_MEDIDA."""
    return {
        "grupos": {
            "GLIFOSATO 480 SL_CANECA|LITRO|20|0|0": "HERBICIDAS",
            "UREA_BULTO|KILOGRAMO|50|0|0": "FERTILIZANTES",
        }
    }


@pytest.fixture
def mappings_articulos_minimal() -> dict:
    """Mapping mínimo de artículos para tests. Clave = ARTÍCULO_UNIDAD_DE_MEDIDA."""
    return {
        "articulos_publicacion": {
            "GLIFOSATO 480 SL_CANECA|LITRO|20|0|0": "Glifosato 480 SL (20 L)",
            "UREA_BULTO|KILOGRAMO|50|0|0": "Urea (50 kg)",
        }
    }


@pytest.fixture
def base_enriquecida_minimal(base_liviana_minimal, divipola_minimal) -> pd.DataFrame:
    """DataFrame con columnas de enriquecimiento completas."""
    from sipsa_insumos.utils.parsers import agregar_columnas_unidad_medida
    df = agregar_columnas_unidad_medida(base_liviana_minimal)
    divipola = divipola_minimal.rename(columns={"CodigoMpio": "CÓDIGO DIVIPOLA"})
    df = df.merge(divipola, on="CÓDIGO DIVIPOLA", how="left")
    df["Grupo"] = "HERBICIDAS"
    df["Nombre_Publica"] = "Glifosato 480 SL (20 L)"
    return df


@pytest.fixture
def mayor2_anterior_minimal() -> pd.DataFrame:
    """mayor2 del período anterior para tests de comparison."""
    return pd.DataFrame({
        "CÓDIGO DIVIPOLA":  ["11001", "05001", "76001", "08001"],
        "LLAVE_ARTICULO":   ["GLIFOSATO 480 SL_CANECA|LITRO|20|0|0"] * 4,
        "PRECIO_PROMEDIO":  [42000.0, 44000.0, 41000.0, 45000.0],
    })
