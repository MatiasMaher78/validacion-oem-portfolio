import os
import time
import argparse
import re
from urllib.parse import quote_plus
import base64
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import unquote

DEFAULT_BATCH = 500


def locate_unique_excel(folder):
    candidates = [f for f in os.listdir(folder) if f.lower().endswith((".xls", ".xlsx"))]
    # Ignorar archivos temporales de Excel que empiezan con '~$' y archivos ocultos
    files = [f for f in candidates if not f.startswith('~$') and not f.startswith('.')]

    # Si sólo hay archivos temporales (por ejemplo el archivo está abierto en Excel),
    # intenta utilizar la lista original (candidates) para mostrar un mensaje útil.
    if not files and candidates:
        # show clearer error to user instead of silently picking temp file
        raise FileExistsError(
            f"Se detectaron sólo archivos temporales en {folder} (ej: archivo abierto en Excel): {candidates}. "
            "Cierra el archivo en Excel o indica la ruta con --folder <ruta>."
        )

    if not files:
        raise FileNotFoundError(f"No se encontró ningún .xls/.xlsx en {folder}")

    if len(files) > 1:
        raise FileExistsError(f"Se encontraron varios archivos Excel en {folder}: {files}")

    return os.path.join(folder, files[0])


def read_workbook(path):
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    df = df.fillna("")
    return df


def get_col(df, names, idx):
    # Buscar por nombre (ignorando mayúsculas) o por posición
    cols = {c.lower(): c for c in df.columns}
    for n in names:
        if n.lower() in cols:
            return df[cols[n.lower()]]
    # fallback por índice (0-based)
    if df.shape[1] > idx:
        return df.iloc[:, idx]
    return pd.Series([""] * len(df))


def parse_search_results(html, query):
    soup = BeautifulSoup(html, 'html.parser')
    selectors = [
        "ul.products li.product",
        "div.products div.product",
        "article",
        "div.search-results",
        "div.post",
        "div.result",
        "li.product",
        "div.product",
        "div[itemtype='http://schema.org/Product']",
    ]
    for sel in selectors:
        elems = soup.select(sel)
        if elems:
            links = set()
            for e in elems:
                for a in e.find_all('a', href=True):
                    links.add(a['href'])
            if links:
                return len(links)

    main = soup.find('main') or soup.find('body')
    if main:
        links = set()
        for a in main.find_all('a', href=True):
            text = (a.get_text(' ', strip=True) or '').lower()
            href = a['href'].lower()
            if query.lower() in text or query.lower() in href:
                links.add(a['href'])
        if links:
            return len(links)

    # fallback: contar ocurrencias textuales
    return len(re.findall(re.escape(str(query)), html, flags=re.IGNORECASE))


def normalize_ref(s: str) -> str:
    if not s:
        return ""
    # Keep only alphanumeric characters and lowercase for comparison
    return re.sub(r'[^0-9a-z]', '', str(s).lower())


def build_ecooparts_url(query):
    """
    Construye la URL de búsqueda de ecooparts usando el formato exacto del sitio.
    Basado en ingeniería inversa de URLs reales:
    - busval = base64 de "|{codigo}|ninguno|producto|-1|0|0|0|0||0|0|0|0"
    - tebu = base64 del código
    - txbu = base64 del código
    """
    code = str(query).strip()
    # tebu y txbu son simplemente el código en base64
    tebu = base64.b64encode(code.encode()).decode()
    txbu = tebu
    # busval tiene formato especial: |{codigo}|ninguno|producto|-1|0|0|0|0||0|0|0|0
    busval_raw = f"|{code}|ninguno|producto|-1|0|0|0|0||0|0|0|0"
    busval = base64.b64encode(busval_raw.encode()).decode()
    # ord = "ninguno" en base64
    ord_b64 = base64.b64encode("ninguno".encode()).decode()
    # valo = "-1" en base64
    valo_b64 = base64.b64encode("-1".encode()).decode()
    # panu = "1" en base64
    panu_b64 = base64.b64encode("1".encode()).decode()

    url = (
        f"https://ecooparts.com/recambios-automovil-segunda-mano/"
        f"?pag=pro&busval={busval}&filval=&panu={panu_b64}&tebu={tebu}"
        f"&ord={ord_b64}&valo={valo_b64}&ubic=&txbu={txbu}"
    )
    return url


def search_count(session, query, proxies=None, headers=None, timeout=15, max_pages=10, verbose=False):
    """
    Busca un código OEM en ecooparts y devuelve el número de publicaciones encontradas.
    Usa la URL con parámetros codificados en base64 (formato real del sitio).
    """
    if not query or str(query).strip() == "":
        return 0

    if headers is None:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                          '(KHTML, like Gecko) Chrome/120.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8',
        }

    norm_query = normalize_ref(query)

    # Usar la URL con formato exacto del sitio
    url = build_ecooparts_url(query)

    try:
        r = session.get(url, proxies=proxies, headers=headers, timeout=timeout)
        html = r.text

        # Primero intentar parsear con parse_search_results
        count = parse_search_results(html, query)
        if verbose:
            print(f"[verbose] base-url: {url} -> parsed: {count}")
        if count > 0:
            return count

        # Forzar paginación: intentar páginas 1..max_pages y acumular enlaces únicos
        links_found = set()
        for page in range(1, max_pages + 1):
            if page == 1:
                page_url = url
            else:
                page_url = url + f"&pag={page}"
            try:
                rp = session.get(page_url, proxies=proxies, headers=headers, timeout=timeout)
                htmlp = rp.text
                s = BeautifulSoup(htmlp, 'html.parser')
                page_links = set()
                # buscar productos en selectores conocidos
                selectors = [
                    "ul.products li.product",
                    "div.products div.product",
                    "article",
                    "li.product",
                    "div.product",
                ]
                for sel in selectors:
                    elems = s.select(sel)
                    for e in elems:
                        for a in e.find_all('a', href=True):
                            href = a['href']
                            page_links.add(href)

                # también buscar enlaces cuyo texto/href contenga la referencia
                for a in s.find_all('a', href=True):
                    href = a['href']
                    text = (a.get_text(' ', strip=True) or '')
                    try:
                        href_dec = unquote(href)
                    except Exception:
                        href_dec = href
                    if norm_query and (norm_query in normalize_ref(text) or norm_query in normalize_ref(href_dec)):
                        page_links.add(href)

                new_links = page_links - links_found
                links_found.update(page_links)
                if verbose:
                    print(f"[verbose] page {page} -> page_links={len(page_links)} new={len(new_links)} total={len(links_found)} url={page_url}")
                # si no hay nuevos enlaces en esta página, terminar
                if not new_links and page > 1:
                    break
                # pequeña espera
                time.sleep(0.2)
            except Exception:
                if verbose:
                    print(f"[verbose] error fetching page {page_url}")
                break

        if links_found:
            return len(links_found)

        # Último recurso: contar ocurrencias del código en el HTML inicial
        occurrences = len(re.findall(re.escape(str(query)), html, flags=re.IGNORECASE))
        if verbose:
            print(f"[verbose] occurrences in base html: {occurrences}")
        if occurrences > 0:
            return max(1, occurrences // 3)

        return 0

    except Exception as e:
        if verbose:
            print(f"[verbose] exception in search_count: {e}")
        return 0
