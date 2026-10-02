"""Tests de la selección de insumos iniciales / ajustados."""
from sipsa_insumos.utils.insumos import ruta_vigente, version_insumo


def _crear(ruta):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_bytes(b"x")
    return ruta


def test_usa_inicial_si_no_hay_ajustada(tmp_path):
    ini = _crear(tmp_path / "01_raw" / "SEP2026" / "BASES LIVIANAS SEP2026" / "Elementos sep 2026.xlsx")
    assert ruta_vigente(ini) == ini
    assert version_insumo(ruta_vigente(ini)) == "inicial"


def test_prefiere_ajustada_con_el_mismo_nombre(tmp_path):
    raw = tmp_path / "01_raw" / "SEP2026"
    ini = _crear(raw / "BASES LIVIANAS SEP2026" / "Elementos sep 2026.xlsx")
    aj = _crear(raw / "AJUSTADOS SEP2026" / "BASES LIVIANAS SEP2026" / "ELEMENTOS SEP 2026.xlsx")
    assert ruta_vigente(ini) == aj
    assert version_insumo(ruta_vigente(ini)) == "ajustada"


def test_ajuste_de_un_modulo_no_afecta_otro(tmp_path):
    raw = tmp_path / "01_raw" / "SEP2026"
    ini = _crear(raw / "BASES LIVIANAS SEP2026" / "Insumos agrícolas sep 2026.xlsx")
    _crear(raw / "AJUSTADOS SEP2026" / "BASES LIVIANAS SEP2026" / "Elementos sep 2026.xlsx")
    assert ruta_vigente(ini) == ini


def test_divipola_ajustada(tmp_path):
    raw = tmp_path / "01_raw" / "SEP2026"
    ini = _crear(raw / "DIVIPOLA SEP2026" / "DIVIPOLA.xlsx")
    aj = _crear(raw / "AJUSTADOS SEP2026" / "DIVIPOLA SEP2026" / "DIVIPOLA.xlsx")
    assert ruta_vigente(str(ini)) == aj


def test_rutas_fuera_de_raw_sin_cambios(tmp_path):
    otra = tmp_path / "conf" / "base" / "x.yml"
    assert ruta_vigente(otra) == otra
    assert str(ruta_vigente("")) == "."
