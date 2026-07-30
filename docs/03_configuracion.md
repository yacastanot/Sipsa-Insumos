# Configuración

Este proyecto no tiene un único `config.py`: Kedro separa la configuración en
archivos YAML bajo `conf/base/`, que cumplen ese mismo rol de "fuente única de
verdad" — el código nunca hardcodea una ruta, un umbral o un período. Todo
sale de aquí.

## Variables de entorno (`.env`)

| Variable       | Tipo   | Por defecto | Propósito |
|----------------|--------|-------------|-----------|
| `INSUMOS_USER` | str    | *(ninguno)* | Usuario HTTP Basic Auth de la interfaz web |
| `INSUMOS_PASS` | str    | *(ninguno)* | Contraseña HTTP Basic Auth |
| `PORT`         | int    | `8080`      | Puerto de uvicorn |

`.env` nunca se comitea (está en `.gitignore`); `.env.example` sí, sin
valores. `app.py` carga `.env` con `python-dotenv` al iniciar
(`load_dotenv(PROJECT_ROOT / ".env")`).

## Período activo — `globals.yml` y `parameters.yml`

Ambos se editan juntos al empezar cada mes. `globals.yml` alimenta rutas de
`catalog_*.yml` vía interpolación `${globals:clave}`; `parameters.yml`
alimenta los nodos directamente vía `params:clave`.

```yaml
# conf/base/globals.yml
periodo: "JUL2026"                      # período que se está procesando
periodo_anterior: "JUN2026"             # mensual
periodo_anterior_bimestral: "MAY2026"   # Elementos/Empaques (impar) o Propagación (par)
periodo_anterior_trimestral: "MAR2026"  # Arriendos/Servicios o Jornales/Especies
mes_actual: "Julio"
mes_anterior: "Junio"
mes_num_actual: 7
```

`parameters.yml` repite `periodo`/`mes_actual`/etc. (Kedro no comparte
variables entre `globals` y `parameters` automáticamente) y agrega umbrales
comunes:

```yaml
umbral_var_alta: 25.0       # VAR >= esto → REVISA=1
umbral_var_baja: -25.0      # VAR <= esto → REVISA=1
umbral_var_extrema: 100.0   # VAR >= esto → REVISA=3
ruta_reporting: "data/08_reporting/${globals:periodo}"
```

## Parámetros por módulo — `parameters_<modulo>.yml`

Un archivo por cada uno de los 9 módulos (`parameters_agricolas.yml`,
`parameters_pecuarios.yml`, …). Define de dónde lee ese módulo y cómo arma su
llave de artículo:

```yaml
# conf/base/parameters_agricolas.yml
agricolas:
  modulo: "AGRICOLAS"
  archivo_liviana: "data/01_raw/JUL2026/BASES LIVIANAS JUL2026/Insumos agrícolas jul 2026.xlsx"
  hoja_liviana: "Información Insumos"
  archivo_divipola_grupos: "data/01_raw/JUL2026/DIVIPOLA JUL2026/divipola insumos agrícolas jul 2026.xlsx"
  tipo_modulo: "estandar"       # "estandar" (UnMed/CasaCom/RegICA) | "caracte" (Caracte., sin UnMed)
  tipo_llave: "unmed"           # "unmed" | "casacom_ica_unmed" (Elementos/Pecuarios)
  min_n: 2                      # mínimo de fuentes para publicar (secreto estadístico)
  grupos: ["COADYUDANTES", "FERTILIZANTES", ...]
```

**Importante**: al escribir un override parcial (ver más abajo) hay que
repetir el diccionario completo — Kedro reemplaza la clave de nivel superior
entera (`agricolas:`), no hace merge profundo campo por campo.

## Catálogo de datasets — `catalog.yml` y `catalog_<modulo>.yml`

`catalog.yml` define los datasets compartidos (DIVIPOLA maestro, mappings).
`catalog_<modulo>.yml` define, namespaced (`agricolas.base_bronze`,
`agricolas.mayor2`, …), cada dataset intermedio y final de ese módulo, con
sus rutas interpoladas por `${globals:periodo}`.

## Mappings — `mappings_grupos.yml` / `mappings_articulos.yml`

Diccionarios `LLAVE_ARTICULO → Grupo` y `LLAVE_ARTICULO → Nombre de
publicación`, compartidos entre todos los módulos. Se actualizan
automáticamente en cada corrida (nodo `actualizar_mappings_divipola`) leyendo
el archivo DIVIPOLA propio del módulo — de forma **no destructiva**: solo
agrega llaves nuevas, nunca sobreescribe una existente. Si una llave quedó
mal (por un bug de parseo ya corregido), hay que borrarla manualmente del
YAML para que se regenere correcta en la siguiente corrida.

## `conf/local/` — overrides temporales sin tocar la configuración activa

Para regenerar la salida de un período histórico (ej. una línea base de
comparación) sin alterar el período activo en `conf/base/`, se crean
versiones temporales de `globals.yml`, `parameters.yml` y
`parameters_<modulo>.yml` dentro de `conf/local/` (Kedro los superpone
automáticamente sobre `conf/base/`). Al terminar, se borran — `conf/local/`
debe quedar vacío y `git status` limpio en `conf/`. Ver el flujo completo en
[05_flujo_datos.md](05_flujo_datos.md).

## `logging.yml`

Configuración estándar de logging de Kedro (formato de consola). No suele
necesitar cambios.
