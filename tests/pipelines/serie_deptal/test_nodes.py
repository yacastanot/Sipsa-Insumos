"""Tests del pipeline de serie departamental (réplica de 03SERIE_DEPTAL)."""
import datetime

import numpy as np
import pandas as pd
import pytest
from openpyxl import Workbook

from sipsa_insumos.pipelines.serie_deptal.nodes import (
    _catx,
    _leer_hoja,
    _merge_sas,
    _nodupkey,
    _nombre_presentacion,
    generar_serie_deptal,
)


def _xlsx(ruta, hojas: dict[str, list[list]]):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    for nombre, filas in hojas.items():
        ws = wb.create_sheet(nombre)
        for f in filas:
            ws.append(f)
    wb.save(ruta)


class TestLeerHoja:
    def test_tipado_proc_import(self, tmp_path):
        ruta = tmp_path / "x.xlsx"
        _xlsx(ruta, {"H": [["Num", "Mixta", "Precio Ante."], [1, "A", 5], [2, 3119, None]]})
        df = _leer_hoja(ruta, "h", "xlsx")
        assert pd.api.types.is_numeric_dtype(df["Num"])
        # columna con texto y números: carácter, el número queda como texto
        assert df["Mixta"].tolist() == ["A", "3119"]
        assert np.isnan(df["Precio Ante."].iloc[1])

    def test_dbms_excel_cambia_punto_por_numeral(self, tmp_path):
        ruta = tmp_path / "x.xlsx"
        _xlsx(ruta, {"H": [["Nov.", "Precio Ante."], ["P", 1]]})
        assert list(_leer_hoja(ruta, "H", "excel").columns) == ["Nov#", "Precio Ante#"]

    def test_fecha_se_lee_como_serial_excel(self, tmp_path):
        ruta = tmp_path / "x.xlsx"
        _xlsx(ruta, {"H": [["RegICA"], ["ABC"], [datetime.datetime(1908, 7, 15)]]})
        assert _leer_hoja(ruta, "H", "xlsx")["RegICA"].tolist() == ["ABC", "3119"]


class TestEmulacionSAS:
    def test_catx_omite_vacios_y_trunca(self):
        df = pd.DataFrame({"a": ["  X ", "Y"], "b": ["", "Z"], "c": [12.0, np.nan]})
        assert _catx(df, ["a", "b", "c"], 200).tolist() == ["X_12", "Y_Z"]
        assert _catx(df, ["a", "b", "c"], 3).tolist() == ["X_1", "Y_Z"]

    def test_nodupkey_conserva_el_primero_en_orden(self):
        df = pd.DataFrame({"k": ["b", "a", "b"], "v": [1.0, 2.0, 3.0]})
        sin, dup = _nodupkey(df, ["k"])
        assert sin["v"].tolist() == [2.0, 1.0]
        assert dup["v"].tolist() == [3.0]

    def test_merge_muchos_a_muchos_retiene_ultimo_valor(self):
        a = pd.DataFrame({"k": ["x", "x", "x"], "va": [1.0, 2.0, 3.0]})
        b = pd.DataFrame({"k": ["x", "x"], "vb": [10.0, 20.0]})
        res, in_a, in_b = _merge_sas(a, b, ["k"], salida="a")
        # la tabla que se agota conserva su último valor dentro del grupo BY
        assert res["vb"].tolist() == [10.0, 20.0, 20.0]
        assert in_a.all() and in_b.all()

    def test_merge_variable_comun_gana_la_segunda_tabla(self):
        a = pd.DataFrame({"k": ["x", "y"], "v": [1.0, 2.0]})
        b = pd.DataFrame({"k": ["x"], "v": [9.0]})
        res, _, in_b = _merge_sas(a, b, ["k"], salida="a")
        assert res["v"].tolist() == [9.0, 2.0]
        assert in_b.tolist() == [True, False]

    @pytest.mark.parametrize("nombre,estilo,esperado", [
        ("Alisin, 1 litro", "coma", ("Alisin", "1 litro")),
        ("Cilantro, Común, 454 gramos", "coma", ("Cilantro,Común", "454 gramos")),
        ("Cilantro, Común, 454 gramos", "coma_espacio", ("Cilantro, Común", "454 gramos")),
        ("Cosecha de café, sin alimentación, 1 kilogramo", "jornales",
         ("Cosecha de café", "sin alimentación,1 kilogramo")),
        ("Bovino, Cebú, macho, kilogramo", "especies", ("Bovino,Cebú,macho", "kilogramo")),
    ])
    def test_nombre_presentacion(self, nombre, estilo, esperado):
        assert _nombre_presentacion(nombre, estilo) == esperado


class TestGenerarSerieDeptal:
    @pytest.fixture
    def insumos(self, tmp_path):
        raw = tmp_path / "01_raw" / "AGO2026"
        cab = ["Municipio", "Fuente", "Informante", "Codigo CPC", "Articulo", "Precio Ante.", "Precio Actual",
               "Var.", "Caracte.", "Nov.", "Estado", "Observación", "Hist.", "Modi.", "Apro.", "Sup.", "Camb.",
               "Deta."]
        fila = lambda mpio, inf, art, precio: [mpio, "F", inf, 111, art, 0, precio, 0, "X|", "P", "OK", ""] + [""] * 6
        _xlsx(raw / "BASES LIVIANAS AGO2026" / "Arriendos ago 2026.xlsx", {"Información Insumos": [
            cab,
            fila(5001, "A", "PASTOREO", 100),
            fila(5002, "B", "PASTOREO", 200),
            fila(5002, "B", "PASTOREO", 200),   # duplicado
            fila(5001, "A", "CORTE", 50),
            fila(5001, "A", "SIN NOMBRE", 70),  # no está en la hoja de publicación
        ]})
        _xlsx(raw / "DIVIPOLA AGO2026" / "DIVIPOLA.xlsx", {"Dep_Mun": [
            ["CódigoMunicipio", "NombreDepartamento", "NombreMunicipio"],
            [5001, "Antioquia", "Medellín"], [5002, "Antioquia", "Abejorral"]]})
        _xlsx(raw / "DIVIPOLA AGO2026" / "divipola arriendos ago 2026.xlsx", {"articulo": [
            ["articulo_caracte", "articulo publicacion"],
            ["PASTOREO_X|", "Pastoreo, mensual"], ["CORTE_X|", "Corte, mensual"]]})
        ant = tmp_path / "anterior"
        _xlsx(ant / "ARRIENDOS_MAYORESQUE2_MAY2026.XLSX", {"ARRIENDOS_MAYORESOIGUALES2": [
            ["CodigoDepto", "NombreDepartamento", "Codigo CPC", "Nombre_productos_arriendos_publi", "Articulo",
             "N_INFORMANTE_MAY2026", "N_ARTICULOS_MAY2026", "PRECIO_PROMEDIO_MAY2026"],
            ["05", "Antioquia", 111, "Pastoreo, mensual", "PASTOREO", 2, 2, 120]]})
        rep = tmp_path / "08_reporting" / "AGO2026"
        meta = generar_serie_deptal("AGO2026", "ARRIENDOS", str(rep), ruta_raw=str(raw),
                                    carpetas_mes_anterior=[str(ant)])
        return rep / "serie_deptal" / "Arriendos", meta

    def test_genera_los_cinco_archivos(self, insumos):
        carpeta, meta = insumos
        assert len(meta) == 5
        assert all((carpeta / n).exists() for n in [
            "ANEXO ARRIENDOS AGO2026.xlsx", "ARRIENDOS_DUPLICADOS_AGO2026.XLSX", "ARRIENDOS_MAYORESQUE2_AGO2026.XLSX",
            "ARRIENDOS_MENORESQUE2_AGO2026.XLSX", "FALTAN_PUBLICA_AGO2026.XLSX"])

    def test_mayores_menores_y_diagnosticos(self, insumos):
        carpeta, _ = insumos
        mayor = pd.read_excel(carpeta / "ARRIENDOS_MAYORESQUE2_AGO2026.XLSX", dtype={"CodigoDepto": str})
        assert mayor[["CodigoDepto", "Articulo", "N_INFORMANTE_AGO2026", "N_ARTICULOS_AGO2026",
                      "PRECIO_PROMEDIO_AGO2026"]].values.tolist() == [["05", "PASTOREO", 2, 2, 150]]
        menor = pd.read_excel(carpeta / "ARRIENDOS_MENORESQUE2_AGO2026.XLSX")
        assert sorted(menor["Articulo"]) == ["CORTE", "SIN NOMBRE"]
        assert len(pd.read_excel(carpeta / "ARRIENDOS_DUPLICADOS_AGO2026.XLSX")) == 1
        faltan = pd.read_excel(carpeta / "FALTAN_PUBLICA_AGO2026.XLSX")
        assert faltan["Articulo"].tolist() == ["SIN NOMBRE"]
        # data step con LENGTH antes de SET: la llave queda como primera columna
        assert faltan.columns[0] == "Articulo_Caracte"

    def test_anexo_variacion_y_columnas(self, insumos):
        carpeta, _ = insumos
        anexo = pd.read_excel(carpeta / "ANEXO ARRIENDOS AGO2026.xlsx", sheet_name="Arriendos")
        assert list(anexo.columns) == [
            "CodigoDepto", "NombreDepartamento", "Codigo CPC", "Nombre_productos_arriendos_publi", "NOM_ARTICULO",
            "PRECIO_PROMEDIO_MAY2026", "PRECIO_PROMEDIO_AGO2026", "Variacion(%)", "Cuenta", "Tendencia",
            "Nombre_insumo", "Presentación_insumo"]
        fila = anexo.iloc[0]
        assert fila["Variacion(%)"] == pytest.approx(25.0)
        assert (fila["Tendencia"], fila["Nombre_insumo"], fila["Presentación_insumo"]) == ("Positiva", "Pastoreo", "mensual")
