"""
CLI para Verificación OEM.
"""
import argparse
import os
import requests
from pathlib import Path

from .config import CounterConfig
from .excel_handler import (
    locate_unique_excel,
    read_workbook,
    save_results,
    validate_dataframe
)
from .processor import process


def find_default_folder() -> Path:
    """
    Busca la carpeta por defecto 'Verificacion OEM' en ubicaciones comunes.

    Returns:
        Path a la carpeta encontrada

    Raises:
        FileNotFoundError: Si no encuentra la carpeta
    """
    home = Path.home()

    candidates = [
        home / "Desktop" / "Verificacion OEM",
        home / "OneDrive" / "Desktop" / "Verificacion OEM",
    ]

    # Agregar OneDrive si existe en variables de entorno
    onedrive = os.environ.get("OneDrive")
    if onedrive:
        candidates.insert(0, Path(onedrive) / "Desktop" / "Verificacion OEM")

    # Agregar directorio actual
    candidates.append(Path.cwd())

    for candidate in candidates:
        if candidate.exists() and candidate.is_dir():
            return candidate

    raise FileNotFoundError(
        "No se encontró la carpeta 'Verificacion OEM' en Desktop/OneDrive. "
        "Usa --folder <ruta> para especificar ubicación personalizada."
    )


def main():
    """Función principal del CLI."""
    parser = argparse.ArgumentParser(
        description="Verificación de OEM (Original/Paralelo) en ecooparts.com",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python -m verificacion_oem.cli --batch 100
  python -m verificacion_oem.cli --batch 500 --resume
  python -m verificacion_oem.cli --batch 200 --proxy http://user:pass@proxy:8080
  python -m verificacion_oem.cli --folder /ruta/custom --batch 1000 --verbose
        """
    )

    # Argumentos generales
    parser.add_argument(
        "--folder",
        type=str,
        default=None,
        help="Carpeta donde está el Excel (por defecto Desktop/Verificacion OEM)"
    )
    parser.add_argument(
        "--batch",
        type=int,
        default=500,
        help="Número de filas a procesar (por defecto 500)"
    )
    parser.add_argument(
        "--start",
        type=int,
        default=0,
        help="Fila inicial 0-indexed (por defecto 0)"
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.15,
        help="Delay base entre filas en segundos (por defecto 0.15)"
    )
    parser.add_argument(
        "--autosave",
        type=int,
        default=50,
        help="Guardar parcial cada N filas, 0 desactiva (por defecto 50)"
    )
    # Configuración Playwright
    parser.add_argument(
        "--timeout",
        type=int,
        default=30000,
        help="Timeout Playwright en ms (por defecto 30000)"
    )
    parser.add_argument(
        "--proxy",
        type=str,
        default=None,
        help="Proxy HTTP(S) ej: http://user:pass@host:port"
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=20,
        help="Máximo de páginas por búsqueda (por defecto 20)"
    )
    parser.add_argument(
        "--per-page",
        type=int,
        default=30,
        help="Resultados por página (por defecto 30)"
    )

    # Flags
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Mostrar logs detallados de scraping"
    )
    parser.add_argument(
        "--headful",
        action="store_true",
        help="Mostrar navegador (por defecto headless)"
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Continuar desde última ejecución (lee run_log.jsonl)"
    )

    parser.add_argument(
        "--require-country",
        type=str,
        default=None,
        help=(
            "Requerir que la IP pública pertenezca a un país o grupo. "
            "Acepta código ISO (ej: ES), lista CSV (ej: ES,PT) o 'EU' para Unión Europea. Si no coincide, se cancela."
        ),
    )

    # Outputs
    parser.add_argument(
        "--out-dir",
        type=str,
        default=None,
        help="Carpeta para logs (por defecto ./run_artifacts junto al Excel)"
    )

    args = parser.parse_args()

    # Si se requiere país específico, verificar la IP pública
    if getattr(args, "require_country", None):
        # Lista de códigos ISO de la Unión Europea (miembros actuales)
        EU_COUNTRIES = {
            "AT","BE","BG","HR","CY","CZ","DK","EE","FI","FR","DE",
            "GR","HU","IE","IT","LV","LT","LU","MT","NL","PL","PT",
            "RO","SK","SI","ES","SE",
        }

        def parse_allowed_codes(val: str):
            v = (val or "").strip()
            if not v:
                return set()
            if v.upper() == "EU":
                return EU_COUNTRIES
            return {p.strip().upper() for p in v.split(",") if p.strip()}

        def get_public_country_code(timeout: float = 3.0) -> str:
            try:
                r = requests.get("https://ipapi.co/json/", timeout=timeout)
                if r.status_code != 200:
                    return ""
                data = r.json()
                return (data.get("country_code", "") or "").upper()
            except Exception:
                return ""

        allowed = parse_allowed_codes(args.require_country)
        cc = get_public_country_code()
        if not cc or cc not in allowed:
            allowed_display = ",".join(sorted(list(allowed)))
            plural = "país" if len(allowed) == 1 else "países"
            print(f"❌ No estás conectado desde {allowed_display} ({plural}). IP detectada: {cc or 'desconocida'}. Cancelo para ahorrar datos.")
            return 2

    # Determinar carpeta de trabajo
    if args.folder:
        folder = Path(args.folder)
    else:
        try:
            folder = find_default_folder()
        except FileNotFoundError as e:
            print(f"❌ Error: {e}")
            return 1

    print(f"📁 Carpeta: {folder}")

    # Buscar Excel
    try:
        excel_path = locate_unique_excel(folder)
    except (FileNotFoundError, Exception) as e:
        print(f"❌ Error: {e}")
        return 1

    print(f"📄 Excel: {excel_path.name}")

    # Leer Excel
    try:
        df = read_workbook(excel_path)
    except Exception as e:
        print(f"❌ Error leyendo Excel: {e}")
        return 1

    # Validar estructura
    errors = validate_dataframe(df)
    if errors:
        print("❌ Errores en el archivo Excel:")
        for err in errors:
            print(f"  • {err}")
        return 1

    print(f"✓ Excel válido: {len(df)} filas")

    # Configurar scraper
    cfg = CounterConfig(
        headless=not args.headful,
        timeout_ms=args.timeout,
        proxy=args.proxy,
        max_pages=args.max_pages,
        per_page=args.per_page,
    )

    # Preparar out_dir
    out_dir = Path(args.out_dir) if args.out_dir else None

    # Procesar
    print(f"\n🚀 Iniciando procesamiento...")
    print(f"   Batch: {args.batch} filas")
    print(f"   Start: fila {args.start + 1}")
    print(f"   Delay: {args.delay}s")
    print(f"   Resume: {'✓' if args.resume else '✗'}")
    print()

    try:
        df_out = process(
            df,
            excel_path,
            batch=args.batch,
            start=args.start,
            delay=args.delay,
            autosave_interval=args.autosave,
            counter_cfg=cfg,
            verbose=args.verbose,
            out_dir=out_dir,
            resume=args.resume,
        )
    except KeyboardInterrupt:
        print("\n⚠ Proceso interrumpido por usuario (Ctrl+C)")
        print("💡 Tip: Usa --resume para continuar desde donde quedó")
        return 130
    except Exception as e:
        print(f"\n❌ Error durante procesamiento: {e}")
        import traceback
        traceback.print_exc()
        return 1

    # Guardar resultados finales
    try:
        out_path = save_results(excel_path, df_out)
        print(f"\n✅ Resultados guardados: {out_path.name}")
    except Exception as e:
        print(f"\n❌ Error guardando resultados: {e}")
        return 1

    print("\n🎉 Proceso completado exitosamente")
    return 0


if __name__ == "__main__":
    exit(main())
