# Verificación OEM — Proyecto

Resumen rápido
- Paquete principal: `verificacion_oem/` (módulos: `config.py`, `excel_handler.py`, `query_builder.py`, `scraper.py`, `logger.py`, `processor.py`, `cli.py`).
- Tests: agrupados en `tests/` (ejecutar con `pytest`).
- Backups / legacy: `archive/` contiene versiones antiguas y el scraper legacy.
- Copia funcional preservada en la raíz: `oem_verify_OK.py`.

Estructura (relevante)

- [verificacion_oem](verificacion_oem): código reorganizado y listo para importar como paquete.
- [tests](tests): pruebas unitarias del proyecto.
- [archive](archive): backups de versiones antiguas (mantener sólo como respaldo).

Requisitos

- El `requirements.txt` activo está en: [verificacion_oem/requirements.txt](verificacion_oem/requirements.txt).
- Se recomienda usar el entorno virtual local `.venv/` provisto (no modificarlo si quieres preservar reproducibilidad).

Uso (Windows — PowerShell)

1. Activar el entorno virtual:

```powershell
& ".\.venv\Scripts\Activate.ps1"
```

2. Ejecutar tests:

```powershell
python -m pytest -q
```

3. Ejecutar la CLI del paquete (por ejemplo):

```powershell
python -m verificacion_oem.cli --help
```

Notas sobre archivos y limpieza

- Ya movimos/copiaste las versiones antiguas a `archive/`; la copia funcional con sufijo `_OK` se conserva en la raíz.
- Se eliminó `__pycache__` y la carpeta `test_codes.py` vacía.
- Si querés, puedo crear un `pyproject.toml` o un `Makefile`/`ps1` con comandos de desarrollo.

Contacto / siguientes pasos

- Puedo actualizar este README con más detalles (ejemplos de uso, formato del Excel de entrada, variables de entorno, etc.). Indicame qué preferís añadir.

Ejemplo: ejecución y resultado

1) Ejecutar desde PowerShell (activar `.venv` primero):

```powershell
& ".\.venv\Scripts\Activate.ps1"
python -m verificacion_oem.cli --folder "C:\Users\<usuario>\OneDrive\Desktop\Verificacion_OEM" --batch 500 --start 0 --autosave 50 --delay 0.5 --max-pages 20 --per-page 30
```

2) Salida esperada (resumen de lo que verás en la consola):

- Mensajes iniciales: `Usando carpeta: <ruta>` y `Leyendo: <archivo.xlsx>`.
- Progreso por fila: `Fila 1/300 -> orig=0 para=1 | pieza='MOTOR' | E='ABC123' | F='XXX111'` (cada fila muestra el conteo obtenido).
- Mensajes de autosave (cada `--autosave` filas):
	`Autosave: guardado parcial hasta fila 150 -> <base>_validated_partial_0_150.xlsx`.
- Mensaje final cuando termina la ejecución completa:
	`Resultados guardados en: <base>_validated.xlsx`.

Archivos generados

- Parciales: `<base>_validated_partial_{start}_{end}.xlsx` (guardados automáticamente según `--autosave`).
- Final: `<base>_validated.xlsx` (archivo con las columnas `Validacion Original` y `Validacion Paralelo`).

Reanudar después de una interrupción

Si el proceso se interrumpe, podés reanudar desde la fila siguiente con la opción `--start`.

Pasos recomendados para reanudar:

1. Revisá el último archivo parcial generado (`<base>_validated_partial_{s}_{e}.xlsx`) en la misma carpeta; abrilo y verificá la última fila procesada (por ejemplo `e`).
2. Ejecutá el programa indicando `--start <e>` para continuar desde la fila siguiente. Ejemplo:

```powershell
& ".\.venv\Scripts\Activate.ps1"
python -m verificacion_oem.cli --folder "C:\Users\<usuario>\OneDrive\Desktop\Verificacion_OEM" --start 150 --batch 500 --autosave 50
```

Notas útiles

- Si preferís no generar parciales, usá `--autosave 0`.
- El script añade/actualiza las columnas `Validacion Original` y `Validacion Paralelo` en los archivos parciales y en el archivo final.
- Conservá la copia `_OK` en la raíz como respaldo si necesitás una versión sin cambios.
