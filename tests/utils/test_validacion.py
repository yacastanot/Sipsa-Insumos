"""Tests de la validación de cierre del período."""
import os
import time

import yaml
from openpyxl import Workbook

from sipsa_insumos.utils.validacion import _mismo_contenido, validar_periodo


def _xlsx(ruta, filas):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    for f in filas:
        wb.active.append(f)
    wb.save(ruta)
    return ruta


def _proyecto(tmp_path):
    """Período ENE2027 con salidas completas de agrícolas y pecuarios (los demás módulos
    de enero no se crean y las pruebas solo revisan estos dos)."""
    conf = tmp_path / "conf" / "base"
    conf.mkdir(parents=True)
    (conf / "globals.yml").write_text('periodo: "ENE2027"\n', encoding="utf-8")
    (conf / "parameters.yml").write_text("validacion: {}\n", encoding="utf-8")
    spa = {"sin_precio_ant": {m: {"activo": False} for m in ("agricolas", "pecuarios")}}
    (conf / "parameters_sin_precio_ant.yml").write_text(yaml.safe_dump(spa), encoding="utf-8")
    raw = tmp_path / "data" / "01_raw" / "ENE2027"
    base = [["Municipio", "Precio Actual"], [5001, 10], [5002, 20]]
    for m, clave in (("agricolas", "Insumos agrícolas"), ("pecuarios", "Insumos pecuarios")):
        p = f"data/01_raw/ENE2027/BASES LIVIANAS ENE2027/{clave} ene 2027.xlsx"
        d = f"data/01_raw/ENE2027/DIVIPOLA ENE2027/divipola {clave.lower()} ene 2027.xlsx"
        (conf / f"parameters_{m}.yml").write_text(
            yaml.safe_dump({m: {"archivo_liviana": p, "archivo_divipola_grupos": d}}, allow_unicode=True),
            encoding="utf-8")
        _xlsx(tmp_path / p, base)
        _xlsx(tmp_path / d, [["Llave", "Nombre"], ["a", "b"]])
    _xlsx(raw / "DIVIPOLA ENE2027" / "DIVIPOLA.xlsx", [["CódigoMunicipio"], [5001]])
    time.sleep(0.05)
    rep = tmp_path / "data" / "08_reporting"
    carpeta_dep = {"agricolas": "Ins_Agrícolas", "pecuarios": "Ins_Pecuarios"}
    for periodo in ("DIC2026", "ENE2027"):
        for m in ("agricolas", "pecuarios"):
            _xlsx(rep / periodo / m / f"BASE_X_{periodo}.xlsx", base)
            _xlsx(rep / periodo / m / f"FALTAN_X_{periodo}.xlsx", [["Llave"]])
            _xlsx(rep / periodo / m / f"MAYORESQUE2_{m}_{periodo}.xlsx", [["CPC", "P"], ["34662", 1.5]])
            for i in range(5):
                _xlsx(rep / periodo / "serie_deptal" / carpeta_dep[m] / f"D{i}_{periodo}.xlsx", [["x"]])
    return tmp_path


def _estados(out, grupo):
    return {(i["modulo"], i["verificacion"]): i["estado"] for i in out["verificaciones"] if i["grupo"] == grupo}


def test_periodo_completo_sin_errores(tmp_path):
    root = _proyecto(tmp_path)
    out = validar_periodo("ENE2027", root=root)
    errores = [i for i in out["verificaciones"]
               if i["estado"] == "error" and i["modulo"] in ("agricolas", "pecuarios", "—")]
    assert not errores, errores
    assert {"agricolas", "pecuarios", "elementos", "empaques"} == set(out["modulos_activos"])
    assert (root / "data" / "08_reporting" / "ENE2027" / "VALIDACION_ENE2027.xlsx").exists()


def test_detecta_archivo_faltante_y_faltan_con_filas(tmp_path):
    root = _proyecto(tmp_path)
    rep = root / "data" / "08_reporting" / "ENE2027"
    os.remove(rep / "agricolas" / "MAYORESQUE2_agricolas_ENE2027.xlsx")
    _xlsx(rep / "pecuarios" / "FALTAN_X_ENE2027.xlsx", [["Llave"], ["sin grupo"]])
    out = validar_periodo("ENE2027", root=root)
    assert _estados(out, "Completitud")[("agricolas", "Reportes municipales")] == "error"
    assert _estados(out, "Diagnósticos")[("pecuarios", "FALTAN_X_ENE2027.xlsx")] == "error"


def test_detecta_base_ajustada_sin_ejecucion_definitiva(tmp_path):
    root = _proyecto(tmp_path)
    time.sleep(0.05)
    _xlsx(root / "data" / "01_raw" / "ENE2027" / "AJUSTADOS ENE2027" / "BASES LIVIANAS ENE2027"
          / "Insumos agrícolas ene 2027.xlsx", [["Municipio", "Precio Actual"], [5001, 11]])
    est = _estados(validar_periodo("ENE2027", root=root), "Ejecución")
    assert est[("agricolas", "Salidas posteriores a los insumos")] == "error"
    assert est[("agricolas", "Salidas hechas con la base vigente")] == "error"


def test_continuidad_con_el_mes_anterior(tmp_path):
    root = _proyecto(tmp_path)
    raw = root / "data" / "01_raw" / "ENE2027"
    # mismo contenido, CPC como número en lugar de texto → idéntico
    _xlsx(raw / "MAYORESQUE2_agricolas_DIC2026.xlsx", [["CPC", "P"], [34662, 1.5]])
    # valor distinto → error
    _xlsx(raw / "MAYORESQUE2_pecuarios_DIC2026.xlsx", [["CPC", "P"], ["34662", 9.9]])
    est = _estados(validar_periodo("ENE2027", root=root), "Continuidad")
    assert est[("agricolas", "MAYORESQUE2_agricolas_DIC2026.xlsx")] == "ok"
    assert est[("pecuarios", "MAYORESQUE2_pecuarios_DIC2026.xlsx")] == "error"


def test_mismo_contenido_distingue_orden(tmp_path):
    a = _xlsx(tmp_path / "a.xlsx", [["k"], [1], [2]])
    b = _xlsx(tmp_path / "b.xlsx", [["k"], [2], [1]])
    c = _xlsx(tmp_path / "c.xlsx", [["k"], [1], [3]])
    assert _mismo_contenido(a, a)[0] == "igual"
    assert _mismo_contenido(a, b)[0] == "orden"
    assert _mismo_contenido(a, c)[0] == "distinto"
