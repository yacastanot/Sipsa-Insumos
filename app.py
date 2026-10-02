"""FastAPI web app — SIPSA Insumos Agropecuarios.

Dos flujos de procesamiento:
  CUADROS          : base liviana por módulo → kedro run --pipeline {modulo}
                     Salida: BASES_*.xlsx en data/08_reporting/{periodo}/{modulo}/
  SIN_PRECIO_ANT   : historico + var_atipico → kedro run --pipeline sin_precio_ant
                     Salida: REV_SIN_PRECIO_ANTE_*.xlsx en data/08_reporting/{periodo}/sin_precio_ant/

Calendario de módulos:
  Mensual (1-12)            : Agrícolas, Pecuarios
  Bimestral impar (1,3,5…) : Elementos, Empaques
  Bimestral par   (2,4,6…) : Propagación
  Trimestral FEB/MAY/AGO/NOV: Arriendos, Servicios
  Trimestral MAR/JUN/SEP/DIC: Jornales, Especies

Sin Precio Anterior (sub-módulos):
  Mensual   : Agrícolas, Pecuarios
  Bim. impar: Elementos
  Bim. par  : Propagación

Credenciales: INSUMOS_USER / INSUMOS_PASS, definidas en .env (ver .env.example).
No hay valor por defecto — si .env no existe o falta alguna clave, arranca
en error en vez de exponer un usuario/clave conocido.
"""
from __future__ import annotations

import asyncio
import io
import zipfile
from urllib.parse import quote
import json as _json
import os
import queue
import re
import secrets
import subprocess
import sys
import threading
from datetime import date
from pathlib import Path

from dotenv import load_dotenv
from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.templating import Jinja2Templates
from markupsafe import Markup
from pydantic import BaseModel

# ── Rutas del proyecto ────────────────────────────────────────────────────────

PROJECT_ROOT  = Path(__file__).parent

load_dotenv(PROJECT_ROOT / ".env")
if not os.environ.get("INSUMOS_USER") or not os.environ.get("INSUMOS_PASS"):
    raise RuntimeError(
        "Faltan INSUMOS_USER / INSUMOS_PASS. Copie .env.example a .env "
        "y complete las credenciales antes de iniciar la aplicación."
    )
CONF_DIR      = PROJECT_ROOT / "conf" / "base"
GLOBALS_YML   = CONF_DIR / "globals.yml"
PARAMS_YML    = CONF_DIR / "parameters.yml"
SPA_PARAMS    = CONF_DIR / "parameters_sin_precio_ant.yml"
REPORTING_DIR = PROJECT_ROOT / "data" / "08_reporting"

app = FastAPI(title="SIPSA Insumos Agropecuarios", docs_url=None, redoc_url=None)
templates = Jinja2Templates(directory=str(PROJECT_ROOT / "templates"))
templates.env.filters["tojson"] = lambda v: Markup(_json.dumps(v, ensure_ascii=False))
security = HTTPBasic()

_pipeline_running = False

# ── Catálogos de dominio ──────────────────────────────────────────────────────

MESES: dict[int, tuple[str, str]] = {
    1:  ("ENE", "Enero"),       2:  ("FEB", "Febrero"),
    3:  ("MAR", "Marzo"),       4:  ("ABR", "Abril"),
    5:  ("MAY", "Mayo"),        6:  ("JUN", "Junio"),
    7:  ("JUL", "Julio"),       8:  ("AGO", "Agosto"),
    9:  ("SEP", "Septiembre"),  10: ("OCT", "Octubre"),
    11: ("NOV", "Noviembre"),   12: ("DIC", "Diciembre"),
}

# Módulos del pipeline de CUADROS con su calendario de actividad
MODULOS: list[dict] = [
    {"id": "agricolas",   "label": "Insumos Agrícolas",       "period_label": "Mensual",           "meses": list(range(1, 13))},
    {"id": "pecuarios",   "label": "Insumos Pecuarios",       "period_label": "Mensual",           "meses": list(range(1, 13))},
    {"id": "elementos",   "label": "Elementos Agropecuarios", "period_label": "Bimestral (impar)", "meses": [1, 3, 5, 7, 9, 11]},
    {"id": "empaques",    "label": "Empaques Agropecuarios",  "period_label": "Bimestral (impar)", "meses": [1, 3, 5, 7, 9, 11]},
    {"id": "propagacion", "label": "Material de Propagación", "period_label": "Bimestral (par)",   "meses": [2, 4, 6, 8, 10, 12]},
    {"id": "arriendos",   "label": "Arriendos de Tierras",    "period_label": "Trimestral",        "meses": [2, 5, 8, 11]},
    {"id": "servicios",   "label": "Servicios Agrícolas",     "period_label": "Trimestral",        "meses": [2, 5, 8, 11]},
    {"id": "jornales",    "label": "Jornales Agrícolas",      "period_label": "Trimestral",        "meses": [3, 6, 9, 12]},
    {"id": "especies",    "label": "Especies Productivas",    "period_label": "Trimestral",        "meses": [3, 6, 9, 12]},
]

# Sub-módulos de Sin Precio Anterior con su calendario
SPA_MODULOS: list[dict] = [
    {"id": "agricolas",   "label": "Insumos Agrícolas",       "meses": list(range(1, 13)),
     "file_hist": "Ins_Agrícolas para revisiones.xlsx",
     "file_var":  "VAR_ATIPICO_AGRICOLA_{PERIODO}.XLSX",
     "var_pipeline": "agricolas/VAR_ATIPICO_AGRICOLA_{PERIODO}.xlsx"},
    {"id": "pecuarios",   "label": "Insumos Pecuarios",       "meses": list(range(1, 13)),
     "file_hist": "Ins_Pecuarios para revisiones.xlsx",
     "file_var":  "VAR_ATIPICO_PECUARIO_{PERIODO}.XLSX",
     "var_pipeline": "pecuarios/VAR_ATIPICO_PECUARIO_{PERIODO}.xlsx"},
    {"id": "elementos",   "label": "Elementos Agropecuarios", "meses": [1, 3, 5, 7, 9, 11],
     "file_hist": "Elementos para revisiones.xlsx",
     "file_var":  "VAR_ATIPICO_ELEMENTOS_{PERIODO}.XLSX",
     "var_pipeline": "elementos/VAR_ATIPICO_ELEMENTOS_{PERIODO}.xlsx"},
    {"id": "propagacion", "label": "Material de Propagación", "meses": [2, 4, 6, 8, 10, 12],
     "file_hist": "Material_propag para revisiones.xlsx",
     "file_var":  "VAR_ATIPICO_MATERIAL_{PERIODO}.XLSX",
     "var_pipeline": "propagacion/VAR_ATIPICO_MATERIAL_{PERIODO}.xlsx"},
]


def modulos_activos(mes_num: int) -> list[str]:
    return [m["id"] for m in MODULOS if mes_num in m["meses"]]


def spa_activos(mes_num: int) -> list[str]:
    return [m["id"] for m in SPA_MODULOS if mes_num in m["meses"]]


# ── Helpers YAML ──────────────────────────────────────────────────────────────

def _set_str(text: str, key: str, value: str) -> str:
    return re.sub(
        rf'^({re.escape(key)}:\s*)"[^"]*"',
        rf'\1"{value}"',
        text, flags=re.MULTILINE,
    )


def _set_int(text: str, key: str, value: int) -> str:
    def _rep(m: re.Match) -> str:
        return f"{m.group(1)}{value}{m.group(2) or ''}"
    return re.sub(
        rf'^({re.escape(key)}:\s*)\d+(\s*#.*)?$',
        _rep, text, flags=re.MULTILINE,
    )


def _read_globals() -> dict:
    result: dict = {}
    for line in GLOBALS_YML.read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if m := re.match(r'^(\w+):\s*"([^"]*)"', s):
            result[m.group(1)] = m.group(2)
        elif m2 := re.match(r'^(\w+):\s*(\d+)', s):
            result[m2.group(1)] = int(m2.group(2))
    return result


def _prev_month(mes: int, anio: int, n: int = 1) -> tuple[int, int]:
    for _ in range(n):
        mes, anio = (12, anio - 1) if mes == 1 else (mes - 1, anio)
    return mes, anio


def _periodo_id(mes: int, anio: int) -> str:
    return f"{MESES[mes][0]}{anio}"


def _write_config(mes_num: int, anio: int) -> dict:
    nombre  = MESES[mes_num][1]
    m_ant, y_ant = _prev_month(mes_num, anio)
    nombre_ant   = MESES[m_ant][1]
    m_bi,  y_bi  = _prev_month(mes_num, anio, 2)
    m_tri, y_tri = _prev_month(mes_num, anio, 3)
    m_m2, _      = _prev_month(mes_num, anio, 2)
    m_m3, _2     = _prev_month(mes_num, anio, 3)

    periodo      = _periodo_id(mes_num, anio)
    p_ant        = _periodo_id(m_ant,  y_ant)
    p_bimestral  = _periodo_id(m_bi,   y_bi)
    p_trimestral = _periodo_id(m_tri,  y_tri)

    g = GLOBALS_YML.read_text(encoding="utf-8")
    for k, v in [
        ("periodo",                    periodo),
        ("periodo_anterior",           p_ant),
        ("periodo_anterior_bimestral", p_bimestral),
        ("periodo_anterior_trimestral",p_trimestral),
        ("mes_actual",                 nombre),
        ("mes_anterior",               nombre_ant),
    ]:
        g = _set_str(g, k, v)
    for k, v in [
        ("mes_num_actual",  mes_num),
        ("mes_num_anterior", m_ant),
        ("anio",            anio),
    ]:
        g = _set_int(g, k, v)
    GLOBALS_YML.write_text(g, encoding="utf-8")

    p = PARAMS_YML.read_text(encoding="utf-8")
    for k, v in [
        ("periodo",         periodo),
        ("periodo_anterior", p_ant),
        ("mes_actual",       nombre),
        ("mes_anterior",     nombre_ant),
        ("fecha_proceso",    date.today().strftime("%Y%m%d")),
    ]:
        p = _set_str(p, k, v)
    for k, v in [
        ("mes_num_actual",  mes_num),
        ("mes_num_anterior", m_ant),
        ("mes_num_menos2",  m_m2),
        ("mes_num_menos3",  m_m3),
        ("anio",            anio),
    ]:
        p = _set_int(p, k, v)
    PARAMS_YML.write_text(p, encoding="utf-8")

    # Actualizar flags 'activo' y rutas del período en parameters_sin_precio_ant.yml
    _update_spa_activos(mes_num)
    _reset_spa_rutas(periodo)

    return {
        "ok":              True,
        "periodo":         periodo,
        "mes_actual":      nombre,
        "mes_anterior":    nombre_ant,
        "modulos_activos": modulos_activos(mes_num),
        "spa_activos":     spa_activos(mes_num),
    }


def _read_param_ruta(modulo_id: str, campo: str) -> str | None:
    """Lee el valor de `campo` en parameters_{modulo_id}.yml (None si no existe)."""
    params_file = CONF_DIR / f"parameters_{modulo_id}.yml"
    if not params_file.exists():
        return None
    m = re.search(rf'^\s+{re.escape(campo)}:\s*"([^"]*)"',
                  params_file.read_text(encoding="utf-8"), flags=re.MULTILINE)
    return m.group(1) if m else None


def _update_param_ruta(modulo_id: str, campo: str, new_path: str) -> None:
    params_file = CONF_DIR / f"parameters_{modulo_id}.yml"
    if not params_file.exists():
        return
    # newline="" conserva los fin de línea originales del YAML (LF).
    with open(params_file, encoding="utf-8", newline="") as f:
        text = f.read()
    text = re.sub(
        rf'^([ \t]+{re.escape(campo)}:[ \t]*)"[^"]*"',
        lambda m: f'{m.group(1)}"{new_path}"',
        text, flags=re.MULTILINE,
    )
    with open(params_file, "w", encoding="utf-8", newline="") as f:
        f.write(text)


def _update_archivo_liviana(modulo_id: str, periodo: str, filename: str) -> None:
    new_path = f"data/01_raw/{periodo}/BASES LIVIANAS {periodo}/{filename}"
    _update_param_ruta(modulo_id, "archivo_liviana", new_path)


def _update_archivo_divipola(modulo_id: str, periodo: str, filename: str) -> None:
    # Módulos caracte tienen archivo_divipola_grupos "" (no soportado por
    # actualizar_mappings_divipola) — no se les asigna ruta.
    if not _read_param_ruta(modulo_id, "archivo_divipola_grupos"):
        return
    new_path = f"data/01_raw/{periodo}/DIVIPOLA {periodo}/{filename}"
    _update_param_ruta(modulo_id, "archivo_divipola_grupos", new_path)


def _nombre_inicial(modulo_id: str, campo: str, periodo: str) -> str | None:
    """Nombre del archivo inicial configurado para el período (la versión ajustada
    debe llevar el mismo nombre para reemplazarlo)."""
    ruta = _read_param_ruta(modulo_id, campo)
    return Path(ruta).name if ruta and f"/{periodo}/" in ruta.replace("\\", "/") else None


def _rutas_desactualizadas(modulo_id: str, periodo: str) -> list[str]:
    """Rutas de insumos del módulo que no apuntan a data/01_raw/{periodo}/.

    Una DIVIPOLA de otro mes no trae los insumos nuevos del período y deja
    filas en FALTAN_GRUPO / FALTAN_PUBLICA.
    """
    problemas = []
    for campo in ("archivo_liviana", "archivo_divipola_grupos"):
        ruta = _read_param_ruta(modulo_id, campo)
        if ruta and f"data/01_raw/{periodo}/" not in ruta.replace("\\", "/"):
            problemas.append(f"{campo} = {ruta}")
    return problemas


def _update_spa_ruta(modulo_id: str, campo: str, nuevo_path: str) -> None:
    """Actualiza ruta_historico o ruta_var_atipico en parameters_sin_precio_ant.yml."""
    if not SPA_PARAMS.exists():
        return
    with open(SPA_PARAMS, encoding="utf-8", newline="") as f:
        text = f.read()
    # Reemplaza dentro del bloque del módulo: busca `    campo: "..."` bajo `  modulo_id:`
    # Usamos un approach de sección: encontramos la sección del módulo y cambiamos el campo
    lines = text.splitlines(keepends=True)
    in_section = False
    depth = 0
    new_lines = []
    for line in lines:
        stripped = line.lstrip()
        indent   = len(line) - len(stripped)
        if re.match(rf'^  {re.escape(modulo_id)}:', line):
            in_section = True
            depth = indent
            new_lines.append(line)
            continue
        if in_section:
            if indent <= depth and stripped and not stripped.startswith("#"):
                in_section = False
            elif re.match(rf'\s+{re.escape(campo)}:\s*"', line):
                line = re.sub(r'"[^"]*"', lambda _m: f'"{nuevo_path}"', line, count=1)
                in_section = False  # reemplazado, salir de la sección
        new_lines.append(line)
    with open(SPA_PARAMS, "w", encoding="utf-8", newline="") as f:
        f.write("".join(new_lines))


def _read_spa_ruta(modulo_id: str, campo: str) -> str | None:
    """Lee ruta_historico / ruta_var_atipico del bloque del módulo en SPA_PARAMS."""
    if not SPA_PARAMS.exists():
        return None
    m = re.search(
        rf'^  {re.escape(modulo_id)}:.*?^\s+{re.escape(campo)}:\s*"([^"]*)"',
        SPA_PARAMS.read_text(encoding="utf-8"), flags=re.MULTILINE | re.DOTALL,
    )
    return m.group(1) if m else None


def _spa_rutas_default(mod: dict, periodo: str) -> dict[str, str]:
    """Rutas por defecto de un sub-módulo SPA para el período.

    ruta_historico: el Excel "para revisiones" cargado en SIN_PRECIO_ANT {periodo}/.
    ruta_var_atipico: el VAR_ATIPICO que genera el propio pipeline del módulo
    (si se carga uno externo desde la app, la carga lo reemplaza).
    """
    root = str(PROJECT_ROOT).replace("\\", "/")
    return {
        "ruta_historico":   f"{root}/data/01_raw/{periodo}/SIN_PRECIO_ANT {periodo}/{mod['file_hist']}",
        "ruta_var_atipico": f"{root}/data/08_reporting/{periodo}/"
                            + mod["var_pipeline"].replace("{PERIODO}", periodo),
    }


def _reset_spa_rutas(periodo: str) -> None:
    """Al cambiar de período, apunta las rutas SPA al período nuevo."""
    for mod in SPA_MODULOS:
        for campo, ruta in _spa_rutas_default(mod, periodo).items():
            _update_spa_ruta(mod["id"], campo, ruta)


def _spa_rutas_problemas(modulo_id: str, periodo: str) -> list[str]:
    """Rutas SPA del módulo que no son del período o cuyo archivo no existe."""
    problemas = []
    for campo in ("ruta_historico", "ruta_var_atipico"):
        ruta = _read_spa_ruta(modulo_id, campo)
        if not ruta:
            continue
        # Rutas con ${globals:periodo} se resuelven al período activo.
        ruta_res = ruta.replace("${globals:periodo}", periodo).replace("\\", "/")
        if f"/{periodo}/" not in ruta_res:
            problemas.append(f"{campo} es de otro período: {ruta}")
        elif not Path(ruta_res).exists():
            problemas.append(f"{campo} no existe: {ruta_res}")
    return problemas


def _spa_activos_config() -> list[str]:
    """Sub-módulos SPA con activo: true en SPA_PARAMS."""
    if not SPA_PARAMS.exists():
        return []
    text = SPA_PARAMS.read_text(encoding="utf-8")
    activos = []
    for mod in SPA_MODULOS:
        m = re.search(rf'^  {re.escape(mod["id"])}:.*?activo:\s*(true|false)',
                      text, flags=re.MULTILINE | re.DOTALL)
        if m and m.group(1) == "true":
            activos.append(mod["id"])
    return activos


def _update_spa_activos(mes_num: int) -> None:
    """Actualiza los flags 'activo' de cada sub-módulo según el mes."""
    if not SPA_PARAMS.exists():
        return
    activos = set(spa_activos(mes_num))
    text = SPA_PARAMS.read_text(encoding="utf-8")
    for mod in SPA_MODULOS:
        nuevo = "true" if mod["id"] in activos else "false"
        # Reemplaza `activo: true/false` dentro del bloque del módulo
        # Busca la aparición después de `  modulo_id:`
        pattern = rf'(  {re.escape(mod["id"])}:.*?activo:\s*)(true|false)'
        text = re.sub(pattern, rf'\g<1>{nuevo}', text, flags=re.DOTALL)
    SPA_PARAMS.write_text(text, encoding="utf-8")


def _set_spa_activo(modulo_id: str, activo: bool) -> None:
    """Ajuste manual del flag 'activo' de un único sub-módulo SPA."""
    if not SPA_PARAMS.exists():
        return
    nuevo = "true" if activo else "false"
    text = SPA_PARAMS.read_text(encoding="utf-8")
    pattern = rf'(  {re.escape(modulo_id)}:.*?activo:\s*)(true|false)'
    text = re.sub(pattern, rf'\g<1>{nuevo}', text, count=1, flags=re.DOTALL)
    SPA_PARAMS.write_text(text, encoding="utf-8")


# ── Autenticación ─────────────────────────────────────────────────────────────

def _check_auth(credentials: HTTPBasicCredentials = Depends(security)) -> str:
    exp_user = os.environ["INSUMOS_USER"]
    exp_pass = os.environ["INSUMOS_PASS"]
    ok = (
        secrets.compare_digest(credentials.username.encode(), exp_user.encode())
        and secrets.compare_digest(credentials.password.encode(), exp_pass.encode())
    )
    if not ok:
        raise HTTPException(401, "Credenciales incorrectas",
                            headers={"WWW-Authenticate": "Basic"})
    return credentials.username


# ── Modelos ───────────────────────────────────────────────────────────────────

class ConfigRequest(BaseModel):
    mes_num: int
    anio: int


class SpaActivoRequest(BaseModel):
    modulo_id: str
    activo: bool


# ── Rutas ─────────────────────────────────────────────────────────────────────

@app.get("/favicon.ico", include_in_schema=False)
async def favicon() -> Response:
    return Response(status_code=204)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request, _: str = Depends(_check_auth)) -> HTMLResponse:
    cfg     = _read_globals()
    mes_num = int(cfg.get("mes_num_actual", 5))
    anio    = int(cfg.get("anio", 2026))
    return templates.TemplateResponse(request, "index.html", {
        "cfg":        cfg,
        "mes_num":    mes_num,
        "anio":       anio,
        "modulos":    MODULOS,
        "spa_modulos": SPA_MODULOS,
        "activos":    modulos_activos(mes_num),
        "spa_activos": spa_activos(mes_num),
    })


@app.post("/configure")
async def configure(body: ConfigRequest, _: str = Depends(_check_auth)) -> dict:
    if not (1 <= body.mes_num <= 12):
        raise HTTPException(400, "mes_num debe estar entre 1 y 12")
    if not (2020 <= body.anio <= 2040):
        raise HTTPException(400, "Año fuera del rango permitido (2020–2040)")
    return _write_config(body.mes_num, body.anio)


@app.post("/configure/spa-activo")
async def configure_spa_activo(body: SpaActivoRequest, _: str = Depends(_check_auth)) -> dict:
    valid_ids = {m["id"] for m in SPA_MODULOS}
    if body.modulo_id not in valid_ids:
        raise HTTPException(400, f"Módulo SPA desconocido: {body.modulo_id}")
    _set_spa_activo(body.modulo_id, body.activo)
    return {"ok": True, "modulo_id": body.modulo_id, "activo": body.activo}


@app.post("/upload/cuadros/{modulo_id}")
async def upload_cuadros(
    modulo_id: str,
    file: UploadFile = File(...),
    ajustada: bool = False,
    _: str = Depends(_check_auth),
) -> dict:
    """Base liviana del período. Con ?ajustada=true es la versión que temática
    reenvía después de revisar: se guarda en AJUSTADOS <P>/ con el nombre de la
    base inicial y el proceso la toma en su lugar (utils/insumos.ruta_vigente)."""
    valid_ids = {m["id"] for m in MODULOS}
    if modulo_id not in valid_ids:
        raise HTTPException(400, f"Módulo desconocido: {modulo_id}")
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Solo se aceptan archivos Excel (.xlsx, .xls)")

    cfg     = _read_globals()
    periodo = str(cfg.get("periodo", "MAY2026"))
    contents = await file.read()
    if ajustada:
        filename = _nombre_inicial(modulo_id, "archivo_liviana", periodo) or file.filename
        dest_dir = PROJECT_ROOT / "data" / "01_raw" / periodo / f"AJUSTADOS {periodo}" / f"BASES LIVIANAS {periodo}"
        dest_dir.mkdir(parents=True, exist_ok=True)
        (dest_dir / filename).write_bytes(contents)
        return {"ok": True, "filename": filename, "modulo": modulo_id, "version": "ajustada",
                "size_kb": round(len(contents) / 1024, 1)}
    dest_dir = PROJECT_ROOT / "data" / "01_raw" / periodo / f"BASES LIVIANAS {periodo}"
    dest_dir.mkdir(parents=True, exist_ok=True)
    (dest_dir / file.filename).write_bytes(contents)
    _update_archivo_liviana(modulo_id, periodo, file.filename)
    return {"ok": True, "filename": file.filename, "modulo": modulo_id, "version": "inicial",
            "size_kb": round(len(contents) / 1024, 1)}


@app.post("/upload/divipola/{tipo}")
async def upload_divipola(
    tipo: str,   # "master" | id de módulo en MODULOS
    file: UploadFile = File(...),
    ajustada: bool = False,
    _: str = Depends(_check_auth),
) -> dict:
    valid_ids = {"master"} | {m["id"] for m in MODULOS}
    if tipo not in valid_ids:
        raise HTTPException(400, f"Tipo de DIVIPOLA desconocido: {tipo}")
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Solo se aceptan archivos Excel (.xlsx, .xls)")

    cfg     = _read_globals()
    periodo = str(cfg.get("periodo", "MAY2026"))
    contents = await file.read()
    # El maestro DIVIPOLA.xlsx tiene ruta fija en el catalog — se normaliza el nombre.
    filename = "DIVIPOLA.xlsx" if tipo == "master" else file.filename
    if ajustada:
        # Versión reenviada por temática: mismo nombre que la inicial, en AJUSTADOS <P>/.
        if tipo != "master":
            filename = _nombre_inicial(tipo, "archivo_divipola_grupos", periodo) or filename
        dest_dir = PROJECT_ROOT / "data" / "01_raw" / periodo / f"AJUSTADOS {periodo}" / f"DIVIPOLA {periodo}"
        dest_dir.mkdir(parents=True, exist_ok=True)
        (dest_dir / filename).write_bytes(contents)
        return {"ok": True, "filename": filename, "tipo": tipo, "version": "ajustada",
                "size_kb": round(len(contents) / 1024, 1)}
    dest_dir = PROJECT_ROOT / "data" / "01_raw" / periodo / f"DIVIPOLA {periodo}"
    dest_dir.mkdir(parents=True, exist_ok=True)
    (dest_dir / filename).write_bytes(contents)
    if tipo != "master":
        _update_archivo_divipola(tipo, periodo, filename)
    return {"ok": True, "filename": filename, "tipo": tipo, "version": "inicial",
            "size_kb": round(len(contents) / 1024, 1)}


@app.post("/upload/mayoresque2/{modulo_id}")
async def upload_mayoresque2(
    modulo_id: str,
    file: UploadFile = File(...),
    _: str = Depends(_check_auth),
) -> dict:
    valid_ids = {m["id"] for m in MODULOS}
    if modulo_id not in valid_ids:
        raise HTTPException(400, f"Módulo desconocido: {modulo_id}")
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Solo se aceptan archivos Excel (.xlsx, .xls)")

    cfg     = _read_globals()
    periodo = str(cfg.get("periodo", "MAY2026"))
    # Los MAYORESQUE2 de referencia van sueltos en la raíz del período (no en subcarpeta).
    dest_dir = PROJECT_ROOT / "data" / "01_raw" / periodo
    dest_dir.mkdir(parents=True, exist_ok=True)
    contents = await file.read()
    (dest_dir / file.filename).write_bytes(contents)
    return {"ok": True, "filename": file.filename, "modulo": modulo_id,
            "size_kb": round(len(contents) / 1024, 1)}


@app.post("/upload/spa/{modulo_id}/{tipo}")
async def upload_spa(
    modulo_id: str,
    tipo: str,           # "historico" | "var_atipico"
    file: UploadFile = File(...),
    _: str = Depends(_check_auth),
) -> dict:
    valid_ids = {m["id"] for m in SPA_MODULOS}
    if modulo_id not in valid_ids:
        raise HTTPException(400, f"Módulo SPA desconocido: {modulo_id}")
    if tipo not in ("historico", "var_atipico"):
        raise HTTPException(400, "tipo debe ser 'historico' o 'var_atipico'")
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Solo se aceptan archivos Excel (.xlsx, .xls)")

    cfg     = _read_globals()
    periodo = str(cfg.get("periodo", "MAY2026"))
    dest_dir = PROJECT_ROOT / "data" / "01_raw" / periodo / f"SIN_PRECIO_ANT {periodo}"
    dest_dir.mkdir(parents=True, exist_ok=True)
    contents = await file.read()
    dest_path = dest_dir / file.filename
    dest_path.write_bytes(contents)

    campo = "ruta_historico" if tipo == "historico" else "ruta_var_atipico"
    _update_spa_ruta(modulo_id, campo, str(dest_path).replace("\\", "/"))
    return {"ok": True, "filename": file.filename, "modulo": modulo_id, "tipo": tipo,
            "size_kb": round(len(contents) / 1024, 1)}


@app.get("/status")
async def status(_: str = Depends(_check_auth)) -> dict:
    return {"running": _pipeline_running}


# ── Árbol de salidas, organizado como la carpeta de resultados de SAS ──────────
#   <P>/
#     REVISIÓN <P>/
#       AGRICOLAS_<P>/ ... (reportes municipales de cada módulo)
#       DEPARTAMENTAL_<P>/Ins_Agrícolas/ ... (serie departamental)
#       Revisión insumos agrícolas <mes siguiente>.xlsx ... (sueltos, como en SAS)
#     SIN PRECIO ANTERIOR <P>/
#     PRELIMINAR (enviado a temática)/  (misma estructura; solo si hubo ejecución definitiva)
_CARPETA_MODULO_SAS = {
    "agricolas": "AGRICOLAS", "pecuarios": "PECUARIOS", "propagacion": "MATERIAL",
    "arriendos": "ARRIENDOS", "servicios": "SERVICIOS", "elementos": "ELEMENTOS",
    "empaques": "EMPAQUES", "jornales": "JORNALES", "especies": "ESPECIES",
}
_ORDEN_MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]


def _es_salida(f: Path) -> bool:
    return f.is_file() and f.suffix.lower() in (".xlsx", ".xls") and not f.name.startswith("~$")


def _nodo(nombre: str, archivos: list[Path] | None = None, carpetas: list[dict] | None = None) -> dict:
    archivos = sorted(archivos or [], key=lambda f: f.name.lower())
    carpetas = [c for c in (carpetas or []) if c["total"]]
    return {
        "nombre": nombre,
        "archivos": [{"nombre": f.name,
                      "ruta": f.relative_to(REPORTING_DIR).as_posix(),
                      "kb": round(f.stat().st_size / 1024, 1)} for f in archivos],
        "carpetas": carpetas,
        "total": len(archivos) + sum(c["total"] for c in carpetas),
    }


def _arbol_revision(base: Path, periodo: str, nombre: str) -> dict:
    """REVISIÓN <P>: módulos, DEPARTAMENTAL y los archivos 'Revisión ...' sueltos."""
    modulos, sueltos = [], []
    for mod, carpeta_sas in _CARPETA_MODULO_SAS.items():
        d = base / mod
        if not d.is_dir():
            continue
        archivos = [f for f in d.iterdir() if _es_salida(f)]
        sueltos += [f for f in archivos if f.name.lower().startswith("revisi")]
        modulos.append(_nodo(f"{carpeta_sas}_{periodo}",
                             [f for f in archivos if not f.name.lower().startswith("revisi")]))
    deptal = base / "serie_deptal"
    departamental = _nodo(f"DEPARTAMENTAL_{periodo}", carpetas=[
        _nodo(d.name, [f for f in d.iterdir() if _es_salida(f)])
        for d in sorted(deptal.iterdir()) if d.is_dir()
    ] if deptal.is_dir() else [])
    return _nodo(nombre, sueltos, modulos + [departamental])


def _arbol_periodo(d: Path) -> dict:
    periodo = d.name
    spa = d / "sin_precio_ant"
    carpetas = [
        _arbol_revision(d, periodo, f"REVISIÓN {periodo}"),
        _nodo(f"SIN PRECIO ANTERIOR {periodo}",
              [f for f in spa.iterdir() if _es_salida(f)] if spa.is_dir() else []),
    ]
    pre = d / "_PRELIMINAR"
    if pre.is_dir():
        carpetas.append(_nodo("PRELIMINAR (enviado a temática)",
                              carpetas=[_arbol_revision(pre, periodo, f"REVISIÓN {periodo}")]))
    return _nodo(periodo, carpetas=carpetas)


def _clave_periodo(nombre: str) -> tuple[int, int]:
    mes, anio = nombre[:3].upper(), nombre[3:]
    return (int(anio) if anio.isdigit() else 0,
            _ORDEN_MESES.index(mes) + 1 if mes in _ORDEN_MESES else 0)


def _arbol_salidas() -> list[dict]:
    if not REPORTING_DIR.exists():
        return []
    periodos = [d for d in REPORTING_DIR.iterdir() if d.is_dir() and _clave_periodo(d.name)[1]]
    arbol = [_arbol_periodo(d) for d in sorted(periodos, key=lambda d: _clave_periodo(d.name), reverse=True)]
    return [n for n in arbol if n["total"]]


def _buscar_nodo(ruta: str) -> tuple[dict, list[dict]] | None:
    """Nodo del árbol por su ruta de nombres ('SEP2026/REVISIÓN SEP2026/...')."""
    nivel, camino = _arbol_salidas(), []
    for parte in [p for p in ruta.split("/") if p]:
        nodo = next((c for c in nivel if c["nombre"] == parte), None)
        if nodo is None:
            return None
        camino.append(nodo)
        nivel = nodo["carpetas"]
    return (camino[-1], camino) if camino else None


@app.get("/outputs")
async def list_outputs(_: str = Depends(_check_auth)) -> dict:
    return {"periodo_actual": str(_read_globals().get("periodo", "")), "arbol": _arbol_salidas()}


@app.get("/download/zip/{ruta:path}")
async def download_zip(ruta: str, _: str = Depends(_check_auth)) -> StreamingResponse:
    """Descarga una carpeta del árbol como .zip, con la misma estructura de carpetas."""
    encontrado = _buscar_nodo(ruta)
    if encontrado is None:
        raise HTTPException(404, "Carpeta no encontrada")
    nodo, _camino = encontrado
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        def agregar(n: dict, prefijo: str) -> None:
            for a in n["archivos"]:
                z.write(REPORTING_DIR / a["ruta"], f"{prefijo}{a['nombre']}")
            for c in n["carpetas"]:
                agregar(c, f"{prefijo}{c['nombre']}/")
        agregar(nodo, f"{nodo['nombre']}/")
    buf.seek(0)
    nombre_zip = (nodo["nombre"] if len(_camino) == 1 else f"{_camino[0]['nombre']} - {nodo['nombre']}") + ".zip"
    return StreamingResponse(buf, media_type="application/zip", headers={
        "Content-Disposition": f"attachment; filename*=UTF-8''{quote(nombre_zip)}"})


@app.get("/download/cuadros/{filename:path}")
async def download_cuadros(filename: str, _: str = Depends(_check_auth)) -> FileResponse:
    path = (REPORTING_DIR / filename).resolve()
    if not str(path).startswith(str(REPORTING_DIR.resolve())):
        raise HTTPException(403, "Acceso denegado")
    if not path.exists():
        raise HTTPException(404, "Archivo no encontrado")
    return FileResponse(str(path), filename=path.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/download/spa/{filename}")
async def download_spa(filename: str, _: str = Depends(_check_auth)) -> FileResponse:
    spa_dir = REPORTING_DIR / str(_read_globals().get("periodo", "")) / "sin_precio_ant"
    path    = (spa_dir / filename).resolve()
    if not str(path).startswith(str(spa_dir.resolve())):
        raise HTTPException(403, "Acceso denegado")
    if not path.exists():
        raise HTTPException(404, "Archivo no encontrado")
    return FileResponse(str(path), filename=filename,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/run")
async def run_pipeline(
    pipeline_name: str = Form("__default__"),
    modulos: str = Form(""),
    _: str = Depends(_check_auth),
) -> StreamingResponse:
    global _pipeline_running
    if _pipeline_running:
        raise HTTPException(409, "El pipeline ya está en ejecución")

    line_queue: queue.Queue = queue.Queue()

    def _run_one(cmd: list[str], label: str) -> int:
        line_queue.put(f"▶ {label}")
        proc = subprocess.Popen(
            cmd, cwd=str(PROJECT_ROOT),
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace",
        )
        for line in proc.stdout:
            line_queue.put(line.rstrip())
        proc.wait()
        return proc.returncode

    def _run_worker() -> None:
        try:
            if pipeline_name == "all_active":
                # Respeta el ajuste manual del usuario si se envía; si no, usa el
                # calendario sugerido según el mes configurado.
                if modulos.strip():
                    valid_ids = {m["id"] for m in MODULOS}
                    activos = [m.strip() for m in modulos.split(",") if m.strip() in valid_ids]
                else:
                    cfg     = _read_globals()
                    mes_num = int(cfg.get("mes_num_actual", 5))
                    activos = modulos_activos(mes_num)
                periodo = str(_read_globals().get("periodo", ""))
                desactualizados = {m: p for m in activos if (p := _rutas_desactualizadas(m, periodo))}
                if desactualizados:
                    line_queue.put(f"✖ Hay insumos que no corresponden a {periodo}. "
                                   "Cargue la base liviana / DIVIPOLA del período:")
                    for mod, problemas in desactualizados.items():
                        for p in problemas:
                            line_queue.put(f"   {mod.upper()}: {p}")
                    line_queue.put(("__DONE__", "insumos de otro período"))
                    return
                for mod in activos:
                    rc = _run_one(
                        [sys.executable, "-m", "kedro", "run", "--pipeline", mod],
                        f"Módulo: {mod.upper()}",
                    )
                    if rc != 0:
                        line_queue.put(("__DONE__", rc))
                        return
                line_queue.put(("__DONE__", 0))
            else:
                if pipeline_name.startswith("sin_precio_ant"):
                    periodo = str(_read_globals().get("periodo", ""))
                    sufijo  = pipeline_name.removeprefix("sin_precio_ant").lstrip("_")
                    spa_ids = [sufijo] if sufijo else _spa_activos_config()
                    problemas = {m: p for m in spa_ids if (p := _spa_rutas_problemas(m, periodo))}
                    if problemas:
                        line_queue.put(f"✖ Insumos de Sin Precio Anterior no listos para {periodo}. "
                                       "Cargue el histórico / VAR_ATIPICO del período "
                                       "o corra primero el módulo correspondiente:")
                        for mod, lista in problemas.items():
                            for p in lista:
                                line_queue.put(f"   {mod.upper()}: {p}")
                        line_queue.put(("__DONE__", "insumos SPA de otro período"))
                        return
                cmd = [sys.executable, "-m", "kedro", "run"]
                if pipeline_name != "__default__":
                    cmd += ["--pipeline", pipeline_name]
                rc = _run_one(cmd, pipeline_name.upper())
                line_queue.put(("__DONE__", rc))
        except Exception as exc:
            line_queue.put(("__DONE__", str(exc)))

    async def generate():
        global _pipeline_running
        _pipeline_running = True
        thread = threading.Thread(target=_run_worker, daemon=True)
        thread.start()
        loop = asyncio.get_event_loop()
        try:
            while True:
                item = await loop.run_in_executor(None, line_queue.get)
                if isinstance(item, tuple) and item[0] == "__DONE__":
                    rc = item[1]
                    yield ("data: __SUCCESS__\n\n" if rc == 0
                           else f"data: __ERROR__{rc}\n\n")
                    break
                yield f"data: {item}\n\n"
        except Exception as exc:
            yield f"data: __ERROR__{exc}\n\n"
        finally:
            _pipeline_running = False

    return StreamingResponse(generate(), media_type="text/event-stream")
