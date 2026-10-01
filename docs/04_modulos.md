# Módulos

Un bloque por archivo de `src/`, `app.py` y `scripts/`. Los 9 módulos de
insumos (agricolas, pecuarios, elementos, empaques, arriendos, servicios,
propagacion, jornales, especies) comparten el mismo código de pipeline —
solo cambian sus parámetros (ver [03_configuracion.md](03_configuracion.md)),
así que cada etapa se documenta una sola vez.

## `src/sipsa_insumos/pipeline_registry.py`

**Propósito**: registra los pipelines de Kedro. Construye, para cada uno de
los 9 módulos, la secuencia namespaced de las 6 etapas
(`_pipeline_modulo(nombre)`), y agrega condicionalmente los sub-pipelines
`BASE_INSUMOS` y `exportar_diagnosticos` para los módulos ya migrados a esa
lógica (`MODULOS_BASE_INSUMOS`, `MODULOS_DIAGNOSTICOS`).

**Cómo ejecutarlo**: no se ejecuta directo — `kedro run --pipeline <nombre>`
lo consulta. Ver nombres disponibles con `kedro registry list`.

**Ejemplo de llamada entre módulos**:
```python
resultado = pipeline(ingestion(), namespace="agricolas", parameters={"periodo"})
```

---

## `src/sipsa_insumos/pipelines/ingestion/`

**Propósito**: lee el Excel "base liviana" del mes y lo deja listo para
enriquecer.

- `leer_base_liviana(archivo_liviana, hoja_liviana, periodo, tipo_modulo,
  tipo_llave)` — filtra por `Estado in (ANALIZADO CENTRAL, APROBADO)` y
  `Precio > 0`, normaliza nombres de columna (`_RENAME_ESTANDAR` o
  `_RENAME_CARACTE` según `tipo_modulo`), arma `LLAVE_ARTICULO`.
- `leer_base_completa(archivo_liviana, hoja_liviana, periodo, tipo_modulo,
  tipo_llave)` — igual pero **sin** filtrar Estado/Precio y conservando
  columnas de control de calidad (`Precio Ante.`, `Nov.`, `Estado`,
  `Observación`). La usan `BASE_INSUMOS` y los diagnósticos
  (FALTAN_GRUPO/FALTAN_PUBLICA/DUPLI/VAR_ATIPICO), que replican el cálculo
  fila-por-fila de SAS en vez de una comparación agregada por municipio.

**Parámetros clave**: `tipo_modulo` (`"estandar"` con UnMed./CasaCom./RegICA,
o `"caracte"` con Caracte./Informante), `tipo_llave` (`"unmed"` o
`"casacom_ica_unmed"`).

---

## `src/sipsa_insumos/pipelines/enrichment/`

**Propósito**: cruza la base con DIVIPOLA y con los mappings de
Grupo/Nombre de publicación.

- `actualizar_mappings_divipola(...)` — primer nodo; si el módulo trae su
  propio archivo DIVIPOLA (`archivo_divipola_grupos`), agrega llaves nuevas
  a `mappings_grupos.yml`/`mappings_articulos.yml` (no destructivo).
- `merge_divipola(base_bronze, divipola_raw)` — cruce por código de
  municipio; produce también `faltan_divipola` (filas sin municipio, no se
  exporta a disco — dataset transitorio).
- `asignar_grupo(...)` — mapea `LLAVE_ARTICULO → Grupo`; si el módulo tiene
  un solo grupo posible (`params:grupos` con 1 elemento), usa ese como
  respaldo cuando no hay match.
- `asignar_articulo_publica(...)` — mapea `LLAVE_ARTICULO → Nombre_Publica`.

---

## `src/sipsa_insumos/pipelines/quality/`

**Propósito**: control de calidad sobre la base ya enriquecida.

- `detectar_duplicados(base_enriquecida)` — llave compuesta
  (`_LLAVE_DUPLICADOS`: municipio, fuente, informante, CPC, artículo, casa
  comercial, ICA, unidad, precio) — dos fuentes distintas con el mismo
  precio son observaciones válidas, no duplicados.
- `calcular_cv(base_sin_dupli, modulo, periodo)` — coeficiente de variación
  por municipio-producto.
- `detectar_var_atipica(base_con_cv, mayor2_anterior, umbral_var_alta,
  umbral_var_baja, umbral_var_extrema)` — compara contra el precio promedio
  del período anterior y clasifica `REVISA` (1 = variación alta, 2 = sin
  precio anterior, 3 = variación extrema ≥100%). Si `mayor2_anterior` es
  `None`/vacío (primer período), todo queda en `REVISA=2`.

---

## `src/sipsa_insumos/pipelines/aggregation/`

**Propósito**: precio promedio por municipio-producto y secreto estadístico.

- `calcular_precio_promedio(base_calidad)` — agrupa por (municipio,
  Nombre_Publica, Grupo); produce `N_FUENTE`, `N_ARTICULOS`,
  `PRECIO_PROMEDIO`, `PRECIO_MIN`, `PRECIO_MAX`.
- `aplicar_secreto_estadistico(precio_promedio, min_n)` — separa en `mayor2`
  (publicable) y `menor2` (por debajo del mínimo de fuentes) según
  `params:<modulo>.min_n`.

---

## `src/sipsa_insumos/pipelines/comparison/`

**Propósito**: variación interperiódica.

- `calcular_variacion_tendencia(mayor2, mayor2_anterior, mes_actual,
  mes_anterior)` — `VARIACION = (actual/anterior - 1) * 100`; clasifica
  `TENDENCIA` (Positiva/Negativa/Estable/n.d.).

---

## `src/sipsa_insumos/pipelines/reporting/`

**Propósito**: exporta todos los `.xlsx` finales replicando el formato SAS
exacto (columnas, orden, etiquetas por módulo). Es el archivo de nodos más
grande del proyecto — contiene los diccionarios de configuración por módulo
(`_NOMBRE_PUBLICA_SAS`, `_DIAG_LABELS`, `_DIAG_ORDEN_*`, etc.) que capturan
las particularidades de cada programa SAS.

- `exportar_bases` / `exportar_anexos` / `exportar_cuadros` / `exportar_tablas`
  — reportes principales, una hoja por grupo.
- `exportar_mayo_menores` — `MAYORESQUE2`/`MENORESQUE2`/`MAYMEN` (unión con
  período anterior).
- `exportar_base_insumos` — `BASE_INSUMOS_{MODULO}_{PERIODO}.xlsx`, réplica
  fila-por-fila con `Precio Ante.` embebido (solo módulos en
  `MODULOS_BASE_INSUMOS`).
- `exportar_diagnosticos` — `FALTAN_GRUPO`/`FALTAN_PUBLICA`/`DUPLI`/
  `VAR_ATIPICO`, misma lógica fila-por-fila que `exportar_base_insumos`
  (solo módulos en `MODULOS_DIAGNOSTICOS`).
- `exportar_revision_historica` / `exportar_revision_tematica` — reportes de
  revisión histórica (TABREV, TEMÁTICA) — **estructura aún no validada
  exhaustivamente contra SAS**, ver pendientes en
  [05_flujo_datos.md](05_flujo_datos.md).

---

## `src/sipsa_insumos/pipelines/sin_precio_ant/`

**Propósito**: pipeline auxiliar que corre *después* del principal. Para
cada módulo activo ese período, cruza el histórico "para revisiones"
(insumo externo) con el `VAR_ATIPICO_{MODULO}_{PERIODO}.xlsx` que el propio
proceso generó, filtrando `REVISA=2` (sin precio anterior).

- `revisar_sin_precio(...)` — detecta automáticamente si `VAR_ATIPICO` viene
  en formato "nuevo" (`exportar_diagnosticos`, columnas SAS directas) o
  "viejo" (comparación agregada, columnas internas) según si existe la
  columna `CÓDIGO DIVIPOLA` — para no romper los módulos que aún no migraron
  a la lógica fila-por-fila.

**Cómo ejecutarlo**: `kedro run --pipeline sin_precio_ant` (todos los
módulos activos) o `--pipeline sin_precio_ant_<modulo>` para uno solo. El
flag `activo` en `parameters_sin_precio_ant.yml` decide si un módulo
corresponde ese período (ver calendario en
[05_flujo_datos.md](05_flujo_datos.md)).

---

## `src/sipsa_insumos/pipelines/serie_deptal/`

**Propósito**: réplica de los programas SAS de `03SERIE_DEPTAL`
(`NNSerie_Departamental_<Módulo>_<MES>.sas`). Por módulo genera los 5
archivos departamentales: `ANEXO ... DEPTO`, `*_MAYORESQUE2`,
`*_MENORESQUE2`, `*_DUPLICADOS` y `FALTAN_*`, en
`data/08_reporting/<periodo>/serie_deptal/<carpeta SAS>/`.

- `generar_serie_deptal(periodo, modulo, ruta_reporting, ...)` — un solo
  nodo genérico; las variantes de cada programa (llave de publicación,
  conteo por fuente o informante, cruce con el mes anterior, columnas del
  ANEXO, nombres de archivo) están en el diccionario `_MODULOS`.
- Emula las reglas de SAS que cambian el resultado: tipado de `PROC IMPORT`
  (`dbms=xlsx` vs `dbms=excel`, celdas de fecha como número de serie),
  orden de variables del PDV (`RETAIN`, `LENGTH` antes de `SET`),
  `PROC SORT NODUPKEY`, `MERGE ... BY` con grupos muchos-a-muchos
  (`_merge_sas`) y el "remerge" de `PROC SQL` (Empaques).
- No usa rutas en parámetros: busca la base liviana y la DIVIPOLA del
  período por nombre en `data/01_raw/<periodo>/`.
- **Mes anterior**: toma el `*_MAYORESQUE2` departamental de la salida
  Python del período anterior; si no existe (primer mes), lo busca en
  `data/01_raw/<periodo>/SERIE_DEPTAL <periodo>/` (copiarlo ahí desde las
  salidas SAS).

**Cómo ejecutarlo**: corre dentro de `kedro run --pipeline <modulo>` y
también por separado con `kedro run --pipeline serie_deptal_<modulo>`.

---

## `src/sipsa_insumos/utils/parsers.py`

**Propósito**: parseo de la columna pipe-delimitada `UNIDAD DE MEDIDA`
(`NOMBRE|UNIDAD|CANTIDAD|0|0`) y construcción de `LLAVE_ARTICULO` en sus dos
modos (`unmed`, `casacom_ica_unmed`). No conoce ningún módulo específico —
es capacidad genérica reutilizada por los 9.

```python
from sipsa_insumos.utils.parsers import construir_llave_articulo
llave = construir_llave_articulo("Glifosato 480 SL", "FRASCO|LITRO|1|0|0")
# "GLIFOSATO 480 SL_FRASCO|LITRO|1|0|0"
```

## `src/sipsa_insumos/utils/excel_writer.py`

**Propósito**: escritura de `.xlsx` multi-hoja sin formato especial (sin
negrilla, sin relleno de color) para paridad exacta con la salida de SAS.

```python
from sipsa_insumos.utils.excel_writer import escribir_excel_multisheet
escribir_excel_multisheet(ruta, {"HOJA1": df1, "HOJA2": df2})
```

## `src/sipsa_insumos/validations/schemas.py`

**Propósito**: contrato Pandera (`SCHEMA_BASE_LIVIANA`) que valida la base
liviana recién leída — columnas obligatorias, tipos, longitud de código
DIVIPOLA. Falla rápido y explícito si el Excel de entrada cambia de formato.

## `src/sipsa_insumos/settings.py`

**Propósito**: wiring de Kedro — declara `OmegaConfigLoader` como cargador de
configuración y `default_run_env="local"` (para que `conf/local/` se
superponga automáticamente sobre `conf/base/`). No suele necesitar cambios.

---

## `app.py`

**Propósito**: interfaz web FastAPI para operar el proceso sin terminal —
cambiar período activo, subir archivos de entrada (bases livianas, DIVIPOLA,
SIN_PRECIO_ANT, MAYORESQUE2), activar/desactivar módulos, lanzar
`kedro run` como subproceso y transmitir el log en vivo por streaming.

**Cómo ejecutarlo**: `iniciar_app.bat` (Windows) o `uvicorn app:app --reload`.
Requiere `.env` con `INSUMOS_USER`/`INSUMOS_PASS` (ver
[03_configuracion.md](03_configuracion.md)).

---

## `scripts/`

Scripts puntuales, no parte del pipeline de Kedro — se ejecutan a mano
cuando hace falta.

- `extraer_mappings.py` — genera entradas de `mappings_grupos.yml`/
  `mappings_articulos.yml` a partir de un archivo DIVIPOLA real.
- `generar_mappings_jun2026.py` — variante puntual para Jornales/Especies en
  su primer período (nombre atado a esa fecha; revisar antes de reusar).
- `validar_vs_sas.py` — compara cada `.xlsx` de `data/08_reporting/` con su
  contraparte SAS hoja por hoja, columna por columna. Uso:
  `python scripts/validar_vs_sas.py --modulo agricolas`.

## `tests/`

Espejo de `src/sipsa_insumos/pipelines/<etapa>/` y `utils/`. Marcadores
`integration` (requiere archivos reales en disco) y `slow` (`kedro run`
completo) — ver `pyproject.toml`. Correr con `pytest`.
