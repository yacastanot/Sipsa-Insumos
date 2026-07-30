# Instalación

## Requisitos del sistema

| Requisito | Detalle |
|-----------|---------|
| Python    | 3.9 o superior (probado con 3.13) |
| pip       | incluido con Python |
| S.O.      | Windows (rutas con backslash y `.bat` de arranque asumen Windows; el código Python en sí es portable) |
| Excel     | los archivos de entrada/salida son `.xlsx` — no se requiere Excel instalado, `openpyxl` los lee/escribe |

No hay base de datos ni servicios externos: todo el estado vive en archivos
(`conf/`, `data/`).

## Pasos de instalación

```bash
git clone <url-del-repo>
cd Sipsa-Insumos

# 1. Entorno virtual
python -m venv .venv
.venv\Scripts\activate          # Windows
# source .venv/bin/activate     # Linux/Mac (no probado)

# 2. Dependencias (incluye kedro, pandas, fastapi, pandera, etc. — ver pyproject.toml)
pip install -e ".[dev]"

# 3. Variables de entorno
cp .env.example .env
# Editar .env y completar INSUMOS_USER / INSUMOS_PASS
```

## Variables de `.env.example`

| Variable       | Propósito                                    | Valor por defecto |
|----------------|-----------------------------------------------|--------------------|
| `INSUMOS_USER` | Usuario de acceso a la interfaz web (HTTP Basic Auth) | *(vacío, obligatorio)* |
| `INSUMOS_PASS` | Contraseña de acceso a la interfaz web       | *(vacío, obligatorio)* |
| `PORT`         | Puerto donde escucha uvicorn                 | `8080` |

Detalle completo de configuración (no solo `.env`) en
[03_configuracion.md](03_configuracion.md).

## Verificar que la instalación funciona

```bash
kedro registry list
```

Debe listar los 9 módulos de insumos más `sin_precio_ant` y sus
sub-pipelines. Si falla con `ModuleNotFoundError`, revisar que el venv esté
activo y que `pip install -e .` se haya ejecutado desde la raíz del repo
(donde está `pyproject.toml`).

## Primera ejecución

```bash
kedro run --pipeline agricolas
```

Genera las salidas del módulo Agrícolas para el período configurado en
`conf/base/globals.yml` (`periodo:`), en `data/08_reporting/<periodo>/agricolas/`.

## Levantar la interfaz web

```bash
iniciar_app.bat
```

Abre `http://localhost:8080` (o el puerto configurado en `.env`). Pide login
con `INSUMOS_USER`/`INSUMOS_PASS`. El script falla explícitamente si no
existe `.env` — no hay valores hardcodeados como fallback.
