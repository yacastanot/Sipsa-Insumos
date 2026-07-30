# SIPSA Insumos

Migración a Python/Kedro del procesamiento mensual de precios de insumos
agropecuarios (SIPSA Insumos, DANE). Reemplaza el flujo original en SAS
reproduciendo sus salidas byte-a-byte, con una interfaz web para operarlo
sin necesidad de la terminal.

## Inicio rápido

```bash
git clone <url-del-repo>
cd Sipsa-Insumos
cp .env.example .env          # completar INSUMOS_USER / INSUMOS_PASS
python -m venv .venv && .venv\Scripts\activate
pip install -e ".[dev]"
kedro run --pipeline agricolas   # o: iniciar_app.bat (interfaz web)
```

## Requisitos

| Requisito       | Versión / detalle                          |
|-----------------|---------------------------------------------|
| Python          | ≥ 3.9                                       |
| Kedro           | 0.19.x                                      |
| S.O.            | Windows (rutas y `.bat` probados en Windows) |
| Servicios externos | Ninguno — todo corre local sobre archivos Excel/Parquet |

## Documentación completa

- Arquitectura → [docs/01_arquitectura.md](docs/01_arquitectura.md)
- Instalación → [docs/02_instalacion.md](docs/02_instalacion.md)
- Configuración → [docs/03_configuracion.md](docs/03_configuracion.md)
- Módulos → [docs/04_modulos.md](docs/04_modulos.md)
- Flujo de datos → [docs/05_flujo_datos.md](docs/05_flujo_datos.md)
- Convenciones de Git → [docs/06_git.md](docs/06_git.md)

## Estado del proyecto

**En desarrollo** — 6 de 9 módulos (Agrícolas, Pecuarios, Elementos, Empaques,
Arriendos, Servicios) validados byte-a-byte contra SAS. Propagación, Jornales
y Especies operativos, sin validación exhaustiva aún.
