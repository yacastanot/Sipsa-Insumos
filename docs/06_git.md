# Convenciones de Git

## Estado actual

El repositorio hoy trabaja sobre una sola rama (`main`), sin ramas
intermedias de desarrollo/staging. Los commits se hacen directo a `main` y
se push-ean a `origin` (GitHub) solo cuando se confirma explícitamente. Esto
es adecuado para el tamaño actual del equipo; las reglas de abajo están
pensadas para escalar sin fricción cuando se sume más gente.

## Mensajes de commit

Formato observado y a mantener: línea de título corta (qué cambió, en
imperativo o sustantivo, sin punto final), línea en blanco, cuerpo con
bullets `-` explicando el *por qué* de cada cambio cuando no es obvio.

```
Paridad exacta SAS: diagnósticos fila-por-fila, BASE_INSUMOS, mappings y Sin Precio Anterior

- Nueva lógica fila-por-fila para FALTAN_GRUPO/FALTAN_PUBLICA/DUPLI/VAR_ATIPICO
  en los 6 módulos validados, reemplazando la comparación agregada por municipio.
- Corrige REGISTRO ICA truncando sufijos decimales legítimos (".5")...
```

No usar mensajes genéricos ("fix", "update", "wip"). Si el commit toca
varios archivos por razones distintas, el cuerpo debe explicar cada una.

## Cuándo comitear

- Nunca comitear sin que el usuario lo pida explícitamente para esa tanda de
  cambios — una aprobación anterior no cubre cambios posteriores.
- Preferir un commit nuevo a `--amend`, salvo pedido explícito.
- Revisar `git status`/`git diff` antes de `git add` — nunca `git add -A`
  a ciegas; si aparece algo inesperado (credenciales, archivo de datos), no
  se agrega sin confirmar.
- `data/`, `.venv/`, `__pycache__/`, `.env` nunca se comitean (ver
  `.gitignore`) — si `git status` los muestra como no ignorados, es una señal
  de que algo cambió mal, no de que haya que agregarlos.

## Estrategia de ramas recomendada al crecer el equipo

Cuando haya más de un desarrollador activo, adoptar ramas por feature en vez
de commitear directo a `main`:

```
feature/nombre-corto  →  PR contra main  →  revisión por pares  →  merge
```

- Una rama por cambio lógico (`feature/base-insumos-pecuarios`,
  `fix/registro-ica-decimal`), no una rama por persona.
- El PR incluye código + documentación del mismo cambio (docstrings,
  `docs/04_modulos.md` si se agregó/movió un módulo,
  `docs/03_configuracion.md` si hay parámetros nuevos) — ver la regla de
  oro en la sección siguiente.
- Nada se mergea a `main` sin que el pipeline relevante (`kedro run
  --pipeline <modulo>`) se haya corrido y, cuando aplique, validado contra
  la salida SAS de referencia (`scripts/validar_vs_sas.py`).

## Documentación y Git — regla de oro

**La documentación va en el mismo PR (o el mismo commit, mientras se trabaje
en `main` directo) que el código que describe.** Un cambio de lógica en
`reporting/nodes.py` que agrega un módulo a `MODULOS_DIAGNOSTICOS` debe venir
con la actualización correspondiente en `docs/04_modulos.md` y
`docs/05_flujo_datos.md`, no "después".

Checklist antes de comitear un cambio de código:

- [ ] Docstrings de las funciones nuevas o modificadas
- [ ] `docs/04_modulos.md` si se agregó/cambió el propósito de un archivo
- [ ] `docs/03_configuracion.md` si hay parámetros o variables `.env` nuevas
- [ ] `.env.example` con las claves nuevas (sin valores)
- [ ] `pyproject.toml` con la dependencia exacta si se agregó una librería
