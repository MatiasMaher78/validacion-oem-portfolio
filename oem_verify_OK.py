import os
import time
import argparse
import random
import string
import re
import base64
from dataclasses import dataclass
from typing import Dict, Optional, Set

import pandas as pd

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except Exception:
    sync_playwright = None  # type: ignore
    PlaywrightTimeoutError = Exception  # type: ignore


DEFAULT_BATCH = 500


# ----------------------------
# Excel I/O
# ----------------------------
def locate_unique_excel(folder: str) -> str:
    """
    Busca el Excel "de entrada" dentro de la carpeta.
    - Ignora temporales (~$)
    - Ignora salidas generadas por el script (*_validated*.xlsx)
    """
    candidates = [f for f in os.listdir(folder) if f.lower().endswith((".xls", ".xlsx"))]
    files = [f for f in candidates if not f.startswith("~$") and not f.startswith(".")]

    # prefer input-like files (exclude validated outputs)
    preferred = [f for f in files if "_validated" not in f.lower()]
    if preferred:
        files = preferred

    if not files and candidates:
        raise FileExistsError(
            f"Se detectaron sólo archivos temporales en {folder} (ej: archivo abierto en Excel): {candidates}. "
            "Cierra el archivo en Excel o indica la ruta con --folder <ruta>."
        )

    if not files:
        raise FileNotFoundError(f"No se encontró ningún .xls/.xlsx en {folder}")

    if len(files) > 1:
        raise FileExistsError(f"Se encontraron varios archivos Excel en {folder}: {files}")

    return os.path.join(folder, files[0])


def read_workbook(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    df = df.fillna("")
    return df


def get_col(df: pd.DataFrame, names, idx):
    # Buscar por nombre (ignorando mayúsculas) o por posición
    cols = {c.lower(): c for c in df.columns}
    for n in names:
        if n.lower() in cols:
            return df[cols[n.lower()]]
    # fallback por índice (0-based)
    if df.shape[1] > idx:
        return df.iloc[:, idx]
    return pd.Series([""] * len(df))


# ----------------------------
# Ecooparts URL builder
# ----------------------------
_FORBIDDEN_CHARS_RE = re.compile(r"[ \t\r\n/\-\,\.]")


def should_search(code: str) -> bool:
    """
    Regla pedida:
    - Si el código contiene espacio, '/', '-', ',' o '.' => NO buscar (resultado = 0).
    - No "normalizamos" ni corregimos; sólo decidimos si buscamos o no.
    """
    if not code:
        return False
    s = str(code)
    if s.strip() == "":
        return False
    return _FORBIDDEN_CHARS_RE.search(s) is None


def _b64(s: str) -> str:
    return base64.b64encode(s.encode("utf-8")).decode("ascii")


def build_ecooparts_search_url(code: str, *, page: int = 1, per_page: int = 30) -> str:
    """
    Construye una URL como las reales (incluye token 'toen').
    """
    c = str(code).strip()
    token = "".join(random.choices(string.ascii_lowercase + string.digits, k=22))

    params = {
        "pag": "pro",
        "busval": _b64(f"|{c}|ninguno|producto|-1|0|0|0|0||0|0|0|0"),
        "filval": "",
        "panu": _b64(str(page)),
        "tebu": _b64(c),
        "ord": _b64("ninguno"),
        "valo": _b64("-1"),
        "ubic": "",
        "toen": _b64(token),
        "veid": _b64("0"),
        "qregx": _b64(str(per_page)),
        "tmin": _b64("1"),
        "ttseu": "",
        "txbu": _b64(c),
        "ivevh": "",
        "ivevhmat": "",
        "ivevhsel": "",
        "ivevhcsver": "",
        "ivevhse": "",
        "oem": "",
        "vin": "",
    }

    ordered_keys = [
        "pag", "busval", "filval", "panu", "tebu", "ord", "valo", "ubic", "toen", "veid",
        "qregx", "tmin", "ttseu", "txbu", "ivevh", "ivevhmat", "ivevhsel", "ivevhcsver",
        "ivevhse", "oem", "vin"
    ]
    query = "&".join(f"{k}={params[k]}" for k in ordered_keys)
    return f"https://ecooparts.com/recambios-automovil-segunda-mano/?{query}"


# ----------------------------
# Playwright counter
# ----------------------------
_PRODUCT_LINK_SELECTORS = (
    'a[href*="recambio-automovil-segunda-mano/"],'
    'a[href*="/en/used-auto-part/"],'
    'a[href*="/used-auto-part/"],'
    'a[href*="/pt/peca-auto-usada/"],'
    'a[href*="/peca-auto-usada/"]'
)


@dataclass
class CounterConfig:
    headless: bool = True
    timeout_ms: int = 30000
    proxy: Optional[str] = None  # e.g. http://user:pass@host:port
    max_pages: int = 10
    max_count: int = 30  # Límite superior de conteo (optimiza paradas tempranas)
    per_page: int = 30
    scroll_rounds: int = 10
    scroll_wait_ms: int = 800


class EcoopartsCounter:
    """
    Reutiliza un único navegador / contexto para muchas consultas.
    Incluye caché por código para acelerar repetidos.
    """
    def __init__(self, cfg: CounterConfig):
        self.cfg = cfg
        self._pw = None
        self._browser = None
        self._context = None
        self._page = None
        self.cache: Dict[str, int] = {}

    def start(self):
        if sync_playwright is None:
            raise RuntimeError(
                "Playwright no está instalado. Ejecuta: "
                "pip install playwright && python -m playwright install chromium"
            )

        if self._pw:
            return

        self._pw = sync_playwright().start()
        launch_args = {"headless": self.cfg.headless}
        if self.cfg.proxy:
            launch_args["proxy"] = {"server": self.cfg.proxy}

        self._browser = self._pw.chromium.launch(**launch_args)
        self._context = self._browser.new_context(
            locale="es-ES",
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0 Safari/537.36"
            ),
            viewport={"width": 1366, "height": 768},
        )
        self._page = self._context.new_page()
        self._page.set_default_timeout(self.cfg.timeout_ms)

    def close(self):
        try:
            if self._context:
                self._context.close()
        finally:
            self._context = None
        try:
            if self._browser:
                self._browser.close()
        finally:
            self._browser = None
        try:
            if self._pw:
                self._pw.stop()
        finally:
            self._pw = None
        self._page = None

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()

    def _ensure_page(self):
        if self._page is None:
            self.start()

    def _try_accept_cookies(self):
        """
        Best-effort: intenta cerrar/aceptar banner de cookies si bloquea el render.
        """
        if not self._page:
            return
        candidates = [
            'button:has-text("Aceptar")',
            'button:has-text("ACEPTAR")',
            'button:has-text("Acepto")',
            'button:has-text("Entendido")',
            'button:has-text("Accept")',
        ]
        for sel in candidates:
            try:
                loc = self._page.locator(sel).first
                if loc.is_visible():
                    loc.click(timeout=2000)
                    break
            except Exception:
                continue

    def _collect_links_current_view(self) -> Set[str]:
        assert self._page is not None
        loc = self._page.locator(_PRODUCT_LINK_SELECTORS)
        hrefs: Set[str] = set()
        try:
            urls = loc.evaluate_all("els => els.map(e => e.href)")
            for u in urls:
                if isinstance(u, str) and u:
                    hrefs.add(u)
        except Exception:
            pass
        return hrefs

    def _scroll_to_load_more(self, *, verbose: bool = False) -> Set[str]:
        """
        Soporta "infinite scroll" / carga progresiva.
        """
        assert self._page is not None
        all_links = self._collect_links_current_view()
        stable_rounds = 0

        for i in range(self.cfg.scroll_rounds):
            before = len(all_links)
            self._page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            self._page.wait_for_timeout(self.cfg.scroll_wait_ms)
            self._try_accept_cookies()
            all_links |= self._collect_links_current_view()
            after = len(all_links)

            if verbose:
                print(f"[verbose] scroll {i+1}/{self.cfg.scroll_rounds}: {before} -> {after}")

            if after == before:
                stable_rounds += 1
                if stable_rounds >= 2:
                    break
            else:
                stable_rounds = 0

        return all_links

    def _click_next_if_exists(self) -> bool:
        """
        Intenta detectar un botón/enlace de "Siguiente" y clickearlo.
        Devuelve True si pudo clickear algo.
        """
        assert self._page is not None
        next_selectors = [
            'a[rel="next"]',
            'a:has-text("Siguiente")',
            'a:has-text("Next")',
            'li.next a',
            'a.pagination-next',
        ]
        for sel in next_selectors:
            try:
                loc = self._page.locator(sel).first
                if loc.count() and loc.is_visible():
                    aria_disabled = loc.get_attribute("aria-disabled")
                    class_attr = (loc.get_attribute("class") or "").lower()
                    if aria_disabled == "true" or "disabled" in class_attr:
                        continue
                    loc.click()
                    return True
            except Exception:
                continue
        return False

    def count(self, code: str, *, verbose: bool = False) -> int:
        c = str(code or "").strip()
        if not should_search(c):
            return 0

        if c in self.cache:
            return self.cache[c]

        self._ensure_page()
        assert self._page is not None

        total_links: Set[str] = set()

        # 1) Entrar a la búsqueda (página 1)
        url = build_ecooparts_search_url(c, page=1, per_page=self.cfg.per_page)
        if verbose:
            print(f"[verbose] goto: {url}")

        try:
            self._page.goto(url, wait_until="domcontentloaded")
            try:
                self._page.wait_for_load_state("networkidle", timeout=min(self.cfg.timeout_ms, 20000))
            except Exception:
                pass

            self._try_accept_cookies()

            # Esperar a que aparezca algún link de producto
            try:
                self._page.wait_for_selector(_PRODUCT_LINK_SELECTORS, timeout=min(self.cfg.timeout_ms, 20000))
            except PlaywrightTimeoutError:
                if verbose:
                    snippet = (self._page.content() or "")[:3000]
                    snippet = snippet.replace("\n", " ")
                    print("[verbose] timeout esperando resultados; snippet head:", snippet[:500])
                self.cache[c] = 0
                return 0

            links = self._scroll_to_load_more(verbose=verbose)
            total_links |= links
            if verbose:
                print(f"[verbose] page=1 links={len(links)} total={len(total_links)}")

            # 2) Paginación (si existe): click en "Siguiente" y acumular
            for page_idx in range(2, self.cfg.max_pages + 1):
                before_total = len(total_links)
                before_view = self._collect_links_current_view()

                if not self._click_next_if_exists():
                    break

                # Esperar que cambie algo (URL o el set visible)
                changed = False
                for _ in range(10):
                    try:
                        self._page.wait_for_load_state("networkidle", timeout=3000)
                    except Exception:
                        pass
                    self._try_accept_cookies()
                    cur_view = self._collect_links_current_view()
                    if cur_view != before_view:
                        changed = True
                        break
                    self._page.wait_for_timeout(400)

                if verbose and not changed:
                    print(f"[verbose] page={page_idx}: no se detectó cambio tras click 'Siguiente'")

                links = self._scroll_to_load_more(verbose=verbose)
                total_links |= links
                if verbose:
                    print(f"[verbose] page={page_idx} links={len(links)} total={len(total_links)}")

                # Si alcanzamos el tope de interés, no hace falta seguir
                if hasattr(self.cfg, 'max_count') and self.cfg.max_count > 0:
                    if len(total_links) >= self.cfg.max_count:
                        if verbose:
                            print(f"[verbose] page={page_idx}: alcanzado max_count={self.cfg.max_count}, terminando")
                        break

                if len(total_links) == before_total:
                    break

        except Exception as ex:
            if verbose:
                print(f"[verbose] error en count(): {ex}")

        cnt = len(total_links)
        # Capear al máximo configurado (0..max_count)
        if hasattr(self.cfg, 'max_count') and self.cfg.max_count > 0:
            if cnt > self.cfg.max_count:
                cnt = self.cfg.max_count
        self.cache[c] = cnt
        return cnt


# ----------------------------
# Main processing
# ----------------------------
def _coerce_int_series(s: pd.Series) -> pd.Series:
    out = []
    for v in s.astype(str).fillna("").tolist():
        v = v.strip()
        if v == "":
            out.append(0)
        else:
            try:
                out.append(int(float(v)))
            except Exception:
                out.append(0)
    return pd.Series(out)


def process(
    df: pd.DataFrame,
    original_path: str,
    *,
    batch: int = DEFAULT_BATCH,
    start: int = 0,
    delay: float = 0.5,
    autosave_interval: int = 50,
    counter_cfg: Optional[CounterConfig] = None,
    verbose: bool = False,
) -> pd.DataFrame:
    pieza = get_col(df, ["Pieza"], 1)
    ref_original = get_col(df, ["Ref. Original (Concesionarios)"], 4)
    ref_paralelo = get_col(df, ["Ref. Paralelo (Recambistas)"], 5)

    n = len(df)
    end = min(n, start + batch)

    # mantener lo ya calculado si existe
    out_orig = _coerce_int_series(df["Validacion Original"]) if "Validacion Original" in df.columns else pd.Series([0] * n)
    out_para = _coerce_int_series(df["Validacion Paralelo"]) if "Validacion Paralelo" in df.columns else pd.Series([0] * n)

    base_name = os.path.splitext(os.path.basename(original_path))[0]
    folder = os.path.dirname(original_path)

    if counter_cfg is None:
        counter_cfg = CounterConfig()

    with EcoopartsCounter(counter_cfg) as counter:
        for i in range(start, end):
            p = str(pieza.iloc[i]).strip()
            e = str(ref_original.iloc[i]).strip()
            f = str(ref_paralelo.iloc[i]).strip()

            valid_orig = counter.count(e, verbose=verbose) if e else 0
            valid_para = counter.count(f, verbose=verbose) if f else 0

            # Normalizar valores a 0..max_count
            max_c = getattr(counter_cfg, 'max_count', 30)
            def _normalize(v):
                try:
                    vv = int(v)
                except Exception:
                    vv = 0
                if vv <= 0:
                    return 0
                return min(vv, max_c)

            out_val_orig = _normalize(valid_orig)
            out_val_para = _normalize(valid_para)

            out_orig.iloc[i] = int(out_val_orig)
            out_para.iloc[i] = int(out_val_para)

            print(f"Fila {i+1}/{n} -> orig={out_val_orig} para={out_val_para} | pieza='{p}' | E='{e}' | F='{f}'")

            processed = i - start + 1
            if autosave_interval and (processed % autosave_interval == 0 or i == end - 1):
                df_out = df.copy()
                df_out["Validacion Original"] = out_orig
                df_out["Validacion Paralelo"] = out_para
                partial_name = os.path.join(folder, f"{base_name}_validated_partial_{start}_{i+1}.xlsx")
                try:
                    df_out.to_excel(partial_name, index=False)
                    print(f"Autosave: guardado parcial hasta fila {i+1} -> {partial_name}")
                except Exception as ex:
                    print(f"Advertencia: no se pudo guardar autosave: {ex}")

            time.sleep(delay)

    df_out = df.copy()
    df_out["Validacion Original"] = out_orig
    df_out["Validacion Paralelo"] = out_para
    return df_out


def save_results(original_path: str, df_out: pd.DataFrame) -> str:
    folder = os.path.dirname(original_path)
    base, _ = os.path.splitext(os.path.basename(original_path))
    out_path = os.path.join(folder, f"{base}_validated.xlsx")
    df_out.to_excel(out_path, index=False)
    return out_path


def main():
    parser = argparse.ArgumentParser(description="Verificación de OEM (Original/Paralelo) en ecooparts.com")
    parser.add_argument("--folder", default=None, help="Carpeta donde está el Excel (por defecto Desktop/Verificacion OEM)")
    parser.add_argument("--batch", type=int, default=DEFAULT_BATCH)
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--delay", type=float, default=0.5, help="Delay entre filas (segundos)")
    parser.add_argument("--autosave", type=int, default=50, help="Guardar resultados parciales cada N filas (0 desactiva)")
    parser.add_argument("--timeout", type=int, default=30000, help="Timeout Playwright en ms (ej: 45000)")
    parser.add_argument("--proxy", type=str, default=None, help="Proxy HTTP(S) (Playwright) ej: http://user:pass@host:port")
    parser.add_argument("--max-pages", type=int, default=10, help="Máximo de páginas a recorrer por código (por defecto 10)")
    parser.add_argument("--per-page", type=int, default=30, help="Resultados por página (qregx). Por defecto 30")
    parser.add_argument("--verbose", action="store_true", help="Logs de depuración por búsqueda")
    parser.add_argument("--headful", action="store_true", help="Muestra el navegador (no recomendado)")

    args = parser.parse_args()

    # localizar carpeta
    if args.folder:
        folder = args.folder
    else:
        home = os.path.expanduser("~")
        candidates = [
            os.path.join(home, "Desktop", "Verificacion OEM"),
            os.path.join(home, "OneDrive", "Desktop", "Verificacion OEM"),
        ]
        one = os.environ.get("OneDrive")
        if one:
            candidates.insert(0, os.path.join(one, "Desktop", "Verificacion OEM"))
        candidates.append(os.getcwd())

        folder = next((c for c in candidates if os.path.isdir(c)), None)
        if folder is None:
            raise FileNotFoundError(
                "No se encontró la carpeta 'Verificacion OEM' en Desktop/OneDrive. "
                "Indica la ruta con --folder <ruta>"
            )

    print(f"Usando carpeta: {folder}")

    excel_path = locate_unique_excel(folder)
    print(f"Leyendo: {excel_path}")
    df = read_workbook(excel_path)

    cfg = CounterConfig(
        headless=not args.headful,   # por defecto headless (no abre ventana)
        timeout_ms=args.timeout,
        proxy=args.proxy,
        max_pages=args.max_pages,
        per_page=args.per_page,
    )

    df_out = process(
        df,
        excel_path,
        batch=args.batch,
        start=args.start,
        delay=args.delay,
        autosave_interval=args.autosave,
        counter_cfg=cfg,
        verbose=args.verbose,
    )
    out_path = save_results(excel_path, df_out)
    print(f"Resultados guardados en: {out_path}")


if __name__ == "__main__":
    main()
