"""FastAPI web app — SIPSA Insumos Agropecuarios.

Dos flujos de procesamiento:
  CUADROS          : base liviana por módulo → kedro run --pipeline {modulo}
                     Salida: BASES_*.xlsx en data/08_reporting/
  SIN_PRECIO_ANT   : historico + var_atipico → kedro run --pipeline sin_precio_ant
                     Salida: REV_SIN_PRECIO_ANTE_*.xlsx en data/08_reporting/sin_precio_ant/

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

Credenciales: INSUMOS_USER / INSUMOS_PASS (por defecto sipsa / cambiar_esta_clave)
"""
from __future__ import annotations

import asyncio
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

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.templating import Jinja2Templates
from markupsafe import Markup
from pydantic import BaseModel

# ── Rutas del proyecto ────────────────────────────────────────────────────────

PROJECT_ROOT  = Path(__file__).parent
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
     "file_var":  "VAR_ATIPICO_AGRICOLA_{PERIODO}.XLSX"},
    {"id": "pecuarios",   "label": "Insumos Pecuarios",       "meses": list(range(1, 13)),
     "file_hist": "Ins_Pecuarios para revisiones.xlsx",
     "file_var":  "VAR_ATIPICO_PECUARIO_{PERIODO}.XLSX"},
    {"id": "elementos",   "label": "Elementos Agropecuarios", "meses": [1, 3, 5, 7, 9, 11],
     "file_hist": "Elementos para revisiones.xlsx",
     "file_var":  "VAR_ATIPICO_ELEMENTOS_{PERIODO}.XLSX"},
    {"id": "propagacion", "label": "Material de Propagación", "meses": [2, 4, 6, 8, 10, 12],
     "file_hist": "Material_propag para revisiones.xlsx",
     "file_var":  "VAR_ATIPICO_MATERIAL_{PERIODO}.XLSX"},
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

    # Actualizar flags 'activo' en parameters_sin_precio_ant.yml
    _update_spa_activos(mes_num)

    return {
        "ok":              True,
        "periodo":         periodo,
        "mes_actual":      nombre,
        "mes_anterior":    nombre_ant,
        "modulos_activos": modulos_activos(mes_num),
        "spa_activos":     spa_activos(mes_num),
    }


def _update_archivo_liviana(modulo_id: str, periodo: str, filename: str) -> None:
    params_file = CONF_DIR / f"parameters_{modulo_id}.yml"
    if not params_file.exists():
        return
    new_path = f"data/01_raw/{periodo}/BASES LIVIANAS {periodo}/{filename}"
    text = params_file.read_text(encoding="utf-8")
    text = re.sub(
        r'^(\s+archivo_liviana:\s*)"[^"]*"',
        rf'\1"{new_path}"',
        text, flags=re.MULTILINE,
    )
    params_file.write_text(text, encoding="utf-8")


def _update_spa_ruta(modulo_id: str, campo: str, nuevo_path: str) -> None:
    """Actualiza ruta_historico o ruta_var_atipico en parameters_sin_precio_ant.yml."""
    if not SPA_PARAMS.exists():
        return
    text = SPA_PARAMS.read_text(encoding="utf-8")
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
                line = re.sub(r'"[^"]*"', f'"{nuevo_path}"', line)
                in_section = False  # reemplazado, salir de la sección
        new_lines.append(line)
    SPA_PARAMS.write_text("".join(new_lines), encoding="utf-8")


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
    exp_user = os.environ.get("INSUMOS_USER", "sipsa")
    exp_pass = os.environ.get("INSUMOS_PASS", "cambiar_esta_clave")
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
    _: str = Depends(_check_auth),
) -> dict:
    valid_ids = {m["id"] for m in MODULOS}
    if modulo_id not in valid_ids:
        raise HTTPException(400, f"Módulo desconocido: {modulo_id}")
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Solo se aceptan archivos Excel (.xlsx, .xls)")

    cfg     = _read_globals()
    periodo = str(cfg.get("periodo", "MAY2026"))
    dest_dir = PROJECT_ROOT / "data" / "01_raw" / periodo / f"BASES LIVIANAS {periodo}"
    dest_dir.mkdir(parents=True, exist_ok=True)
    contents = await file.read()
    (dest_dir / file.filename).write_bytes(contents)
    _update_archivo_liviana(modulo_id, periodo, file.filename)
    return {"ok": True, "filename": file.filename, "modulo": modulo_id,
            "size_kb": round(len(contents) / 1024, 1)}


@app.post("/upload/divipola/{tipo}")
async def upload_divipola(
    tipo: str,   # "master" | id de módulo en MODULOS
    file: UploadFile = File(...),
    _: str = Depends(_check_auth),
) -> dict:
    valid_ids = {"master"} | {m["id"] for m in MODULOS}
    if tipo not in valid_ids:
        raise HTTPException(400, f"Tipo de DIVIPOLA desconocido: {tipo}")
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Solo se aceptan archivos Excel (.xlsx, .xls)")

    cfg     = _read_globals()
    periodo = str(cfg.get("periodo", "MAY2026"))
    dest_dir = PROJECT_ROOT / "data" / "01_raw" / periodo / f"DIVIPOLA {periodo}"
    dest_dir.mkdir(parents=True, exist_ok=True)
    contents = await file.read()
    # El maestro DIVIPOLA.xlsx tiene ruta fija en el catalog — se normaliza el nombre.
    filename = "DIVIPOLA.xlsx" if tipo == "master" else file.filename
    (dest_dir / filename).write_bytes(contents)
    return {"ok": True, "filename": filename, "tipo": tipo,
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
    dest_dir = PROJECT_ROOT / "data" / "01_raw" / periodo / "SIN_PRECIO_ANT"
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


@app.get("/outputs")
async def list_outputs(_: str = Depends(_check_auth)) -> dict:
    if not REPORTING_DIR.exists():
        return {"cuadros": [], "spa": []}

    spa_dir = REPORTING_DIR / "sin_precio_ant"
    cuadros = sorted(
        [f for f in REPORTING_DIR.rglob("*.xlsx") if spa_dir not in f.parents],
        key=lambda f: f.stat().st_mtime, reverse=True,
    )
    spa = sorted(
        [f for f in spa_dir.glob("*.xlsx")] if spa_dir.exists() else [],
        key=lambda f: f.stat().st_mtime, reverse=True,
    )
    return {
        "cuadros": [str(f.relative_to(REPORTING_DIR)).replace("\\", "/") for f in cuadros],
        "spa":     [f.name for f in spa],
    }


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
    spa_dir = REPORTING_DIR / "sin_precio_ant"
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
