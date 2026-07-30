"""
Módulo de scraping para contar resultados en Ecooparts.
"""
import base64
import random
import string
from typing import Dict, Set

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except Exception:
    sync_playwright = None
    PlaywrightTimeoutError = Exception

from .config import CounterConfig


# Selectores CSS para enlaces de productos
_PRODUCT_LINK_SELECTORS = (
    'a[href*="recambio-automovil-segunda-mano/"],'
    'a[href*="/en/used-auto-part/"],'
    'a[href*="/used-auto-part/"],'
    'a[href*="/pt/peca-auto-usada/"],'
    'a[href*="/peca-auto-usada/"]'
)


def _b64(s: str) -> str:
    """Codifica string en base64."""
    return base64.b64encode(s.encode("utf-8")).decode("ascii")


def build_ecooparts_search_url(
    query_text: str,
    *,
    page: int = 1,
    per_page: int = 30
) -> str:
    """
    Construye URL de búsqueda para Ecooparts con todos los parámetros.

    Args:
        query_text: Texto a buscar (puede incluir espacios)
        page: Número de página (1-indexed)
        per_page: Resultados por página

    Returns:
        URL completa con parámetros codificados
    """
    c = str(query_text).strip()
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
        "pag", "busval", "filval", "panu", "tebu", "ord", "valo", "ubic",
        "toen", "veid", "qregx", "tmin", "ttseu", "txbu", "ivevh",
        "ivevhmat", "ivevhsel", "ivevhcsver", "ivevhse", "oem", "vin"
    ]

    query = "&".join(f"{k}={params[k]}" for k in ordered_keys)
    return f"https://ecooparts.com/recambios-automovil-segunda-mano/?{query}"


class EcoopartsCounter:
    """
    Contador de resultados en Ecooparts usando Playwright.

    Mantiene un navegador abierto para múltiples búsquedas y usa
    caché en memoria para evitar búsquedas duplicadas.

    Attributes:
        cfg: Configuración del contador
        cache: Caché de resultados {query_text: count}
        cache_hits: Número de hits en caché
        cache_misses: Número de misses en caché
    """

    def __init__(self, cfg: CounterConfig):
        self.cfg = cfg
        self._pw = None
        self._browser = None
        self._context = None
        self._page = None
        self.cache: Dict[str, int] = {}

        # Métricas
        self.cache_hits = 0
        self.cache_misses = 0

    def start(self):
        """Inicializa Playwright y el navegador."""
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
        # Block heavy resources (images, fonts, stylesheets) to save memory and bandwidth
        # instalar bloqueo de requests según configuración
        if getattr(self.cfg, "block_images", False):
            self._install_request_blocking()
        self._page = self._context.new_page()
        self._page.set_default_timeout(self.cfg.timeout_ms)

    def close(self):
        """Cierra recursos de Playwright."""
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
        """Asegura que el navegador esté iniciado."""
        if self._page is None:
            self.start()

    def _install_request_blocking(self):
        """
        Bloquea recursos pesados para reducir consumo (imágenes, fuentes, media).
        Mantiene CSS/JS para no romper render ni carga dinámica.
        """
        assert self._context is not None

        def _handler(route):
            try:
                rtype = route.request.resource_type
                if rtype in {"image", "media", "font"}:
                    try:
                        route.abort()
                    except Exception:
                        try:
                            route.continue_()
                        except Exception:
                            pass
                else:
                    try:
                        route.continue_()
                    except Exception:
                        pass
            except Exception:
                # Best-effort: si falla el handler, no bloquear.
                try:
                    route.continue_()
                except Exception:
                    pass

        try:
            self._context.route("**/*", _handler)
        except Exception:
            # route may not be available in some environments; ignore and continue
            pass

    def _try_accept_cookies(self):
        """Intenta aceptar el banner de cookies si aparece."""
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
        """Recolecta todos los enlaces de productos visibles."""
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
        Hace scroll para cargar más resultados (lazy loading).

        Args:
            verbose: Si True, imprime logs de depuración

        Returns:
            Set con todos los enlaces únicos encontrados
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
                print(f"[scroll] ronda {i+1}/{self.cfg.scroll_rounds}: {before} -> {after}")

            if after == before:
                stable_rounds += 1
                if stable_rounds >= 2:
                    break
            else:
                stable_rounds = 0

        return all_links

    def count(self, query_text: str, *, verbose: bool = False) -> int:
        """
        Cuenta el número total de resultados para un query.

        Args:
            query_text: Texto a buscar
            verbose: Si True, imprime logs detallados

        Returns:
            Número de productos únicos encontrados
        """
        q = str(query_text or "").strip()
        if q == "":
            return 0

        # Revisar caché
        if q in self.cache:
            self.cache_hits += 1
            return self.cache[q]

        self.cache_misses += 1
        self._ensure_page()
        assert self._page is not None

        total_links: Set[str] = set()

        # Recorrer múltiples páginas
        for page_num in range(1, self.cfg.max_pages + 1):
            url = build_ecooparts_search_url(
                q,
                page=page_num,
                per_page=self.cfg.per_page
            )

            if verbose:
                print(f"[página {page_num}] {url}")

            try:
                self._page.goto(url, wait_until="domcontentloaded")

                try:
                    self._page.wait_for_load_state(
                        "networkidle",
                        timeout=min(self.cfg.timeout_ms, 20000)
                    )
                except Exception:
                    pass

                self._try_accept_cookies()

                try:
                    self._page.wait_for_selector(
                        _PRODUCT_LINK_SELECTORS,
                        timeout=min(self.cfg.timeout_ms, 20000)
                    )
                except PlaywrightTimeoutError:
                    if verbose:
                        print(f"[página {page_num}] Sin resultados (timeout)")
                    break

                page_links = self._scroll_to_load_more(verbose=verbose)

                if verbose:
                    print(f"[página {page_num}] Enlaces encontrados: {len(page_links)}")

                if not page_links:
                    break

                before_total = len(total_links)
                total_links |= page_links
                added = len(total_links) - before_total

                # Si alcanzamos el tope de interés (p.ej. 30), no hace falta seguir
                if getattr(self.cfg, "max_count", 0) and len(total_links) >= int(self.cfg.max_count):
                    if verbose:
                        print(f"[página {page_num}] Alcanzado max_count={self.cfg.max_count}, terminando temprano")
                    break

                # Si en página 2+ no hay nuevos enlaces, terminar
                if page_num > 1 and added == 0:
                    if verbose:
                        print(f"[página {page_num}] Sin nuevos enlaces, terminando")
                    break

                # Si hay menos resultados que per_page, es la última página
                if len(page_links) < self.cfg.per_page:
                    if verbose:
                        print(f"[página {page_num}] Última página detectada")
                    break

            except Exception as ex:
                if verbose:
                    print(f"[página {page_num}] Error: {ex}")
                break

        cnt = len(total_links)
        # Restringir al máximo configurado para ahorrar trabajo y devolver solo 0..max_count
        maxc = int(getattr(self.cfg, "max_count", 0) or 0)
        if maxc > 0 and cnt > maxc:
            cnt = maxc
        self.cache[q] = cnt
        return cnt
