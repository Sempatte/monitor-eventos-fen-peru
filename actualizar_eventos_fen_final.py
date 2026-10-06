#!/usr/bin/env python3
# -*- coding: utf-8 -*-


from __future__ import annotations

import argparse
import collections
import csv
from concurrent.futures import ThreadPoolExecutor, as_completed
import ctypes
import hashlib
from io import BytesIO
import json
import logging
import os
import shutil
import re
import sys
import threading
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse, unquote

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.table import TableColumn
import urllib3

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# -----------------------------------------------------------------------------
# CONFIGURACION
# -----------------------------------------------------------------------------

DEFAULT_WORKBOOK = Path(__file__).resolve().parent / "data" / "Eventos Fenomeno del Niño (2).xlsx"

# Los logs NO se guardan junto al Excel de SharePoint.
DEFAULT_LOG_DIR = Path(__file__).resolve().parent / "logs"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152 Safari/537.36"
)
TIMEOUT = 30
SLEEP_SECONDS = 0.35
OPEN_RETRIES = 5
OPEN_RETRY_SECONDS = 5
SAVE_RETRIES = 5
SAVE_RETRY_SECONDS = 5
PDF_MAX_BYTES = 25 * 1024 * 1024
PDF_MAX_PAGES = 40
PDF_LEAD_CHARS = 5000
MAX_WORKERS_WEEKLY = 4
MAX_WORKERS_BACKFILL = 8
THREAD_LOCAL = threading.local()

SOURCES = [
    {
        "institucion": "INDECI",
        "index_urls": [
            "https://www.gob.pe/institucion/indeci/noticias",
            "https://www.gob.pe/institucion/indeci/noticias?sheet=2",
            "https://www.gob.pe/institucion/indeci/noticias?sheet=3",
        ],
        "allowed_host": "www.gob.pe",
        "path_contains": "/institucion/indeci/noticias/",
    },
    {
        "institucion": "COEN - INDECI",
        "index_urls": [
            "https://portal.indeci.gob.pe/emergencias/",
            "https://portal.indeci.gob.pe/emergencias/page/2/",
            "https://portal.indeci.gob.pe/emergencias/page/3/",
        ],
        "allowed_host": "portal.indeci.gob.pe",
        "path_contains": "/emergencias/",
    },
    {
        "institucion": "Contraloria General de la Republica",
        "index_urls": [
            "https://www.gob.pe/institucion/contraloria/noticias",
            "https://www.gob.pe/institucion/contraloria/noticias?sheet=2",
        ],
        "allowed_host": "www.gob.pe",
        "path_contains": "/institucion/contraloria/noticias/",
    },
    {
        "institucion": "ENFEN",
        "index_urls": [
            "https://enfen.imarpe.gob.pe/noticias/",
            "https://enfen.imarpe.gob.pe/comunicados/",
        ],
        "allowed_host": "enfen.imarpe.gob.pe",
        # Los posts de ENFEN usan URLs tipo /2026/09/28/titulo-del-post/.
        # Evita capturar paginas de navegacion como /noticias/, /videos/, etc.
        "path_regex": r"^/20\d{2}/\d{2}/\d{2}/[^/]+/?$",
    },
]

DIRECT_FEN_PATTERNS = [
    r"\bfenomeno\s+(?:de\s+)?el\s+nino\b",
    r"\bel\s+nino\s+costero\b",
    r"\bnino\s+costero\b",
    r"\benfen\b",
    # Siglas: COEN/INDECI escriben a veces 'FEN' (sin 'Fenomeno El Nino'). Se compara en minusculas
    # y sin tildes; 'fen' solo coincide como palabra completa.
    r"\bfen\b",
    r"\bregion\s+nino\s*(?:1\+2|3\.4)\b",
]

# Peligros hidrometeorologicos. A proposito NO incluye palabras genericas como
# "afectado" o "danos": esas palabras por si solas producian muchos falsos positivos.
HAZARD_PATTERNS = [
    r"\blluvias?\s+intensas?\b",
    r"\bprecipitacion(?:es)?\b",
    r"\binundacion(?:es)?\b",
    r"\bhuaicos?\b",
    r"\bhuaycos?\b",
    r"\bdesborde(?:s)?\b",
    r"\bdesbordamiento(?:s)?\b",
    r"\bactivacion\s+de\s+quebrada(?:s)?\b",
    r"\bmovimientos?\s+en\s+masa\b",
    r"\bdeslizamiento(?:s)?\b",
    r"\baluvion(?:es)?\b",
    r"\bcrecida(?:s)?\b",
    r"\bavenida(?:s)?\s+(?:de\s+)?(?:rio|quebrada)\b",
]

# Evidencia de que el peligro YA ocurrio / genero impacto. Para ingresar a
# Eventos_reales o Eventos_candidatos se exige una senal de este tipo.
EVENT_OCCURRED_PATTERNS = [
    r"\bse\s+registro\b",
    r"\bse\s+registraron\b",
    r"\bocurrio\b",
    r"\bocurrieron\b",
    r"\bse\s+produjo\b",
    r"\bse\s+produjeron\b",
    r"\bse\s+presento\b",
    r"\bse\s+presentaron\b",
    r"\bprovoco\b",
    r"\bprovocaron\b",
    r"\bcauso\b",
    r"\bcausaron\b",
    r"\bgenero\b",
    r"\bgeneraron\b",
    r"\bafecto\b",
    r"\bafectaron\b",
    r"\bdano\b",
    r"\bdanaron\b",
    r"\bdesbordo\b",
    r"\bdesbordaron\b",
    r"\binundo\b",
    r"\binundaron\b",
    r"\bcolapso\b",
    r"\bcolapsaron\b",
    r"\binterrumpio\b",
    r"\binterrumpieron\b",
    r"\bdejo\s+(?:a\s+)?\d+\s+(?:personas?|familias?|viviendas?)\b",
    r"\breporto\b.{0,120}\b(?:lluv|inund|huaico|huayco|desborde|deslizamiento)\b",
    r"\breportaron\b.{0,120}\b(?:lluv|inund|huaico|huayco|desborde|deslizamiento)\b",
    r"\b(?:personas?|familias?|viviendas?|vias?|cultivos?)\s+(?:afectad|damnificad|destruid|colapsad)",
    # Formas que faltaban (nota INDECI 1356852 sobre Arequipa: 'ocasionaron', 'emergencias
    # ocurridas', 'danos', 'se han reportado', 'fallecio' no se reconocian como ocurrencia).
    r"\bocasion(?:o|aron)\b",
    r"\bocurrid[ao]s?\b",
    r"\bsucedio\b",
    r"\bsucedieron\b",
    r"\bfallecio\b",
    r"\bfallecieron\b",
    r"\bperdi(?:o|eron)\s+la\s+vida\b",
    r"\bse\s+(?:han\s+)?(?:reportado|registrado)\b",
    r"\bse\s+reportaron\b",
    r"\bdanos\b",
    r"\bse\s+activ(?:o|aron)\b",
    r"\bresultaron\s+(?:heridos?|afectad|damnificad|muert)",
]

# Noticias preventivas/institucionales que pueden contener muchas palabras de
# riesgo, pero NO son eventos observados.
NON_EVENT_PATTERNS = [
    r"\brecomienda(?:n)?\s+medidas?\s+de\s+preparacion\b",
    r"\bmedidas?\s+de\s+preparacion\b",
    r"\bpreparacion\s+(?:frente|ante)\b",
    r"\bacciones?\s+preventivas?\b",
    r"\bfortalec(?:e|en)\s+(?:la\s+)?(?:coordinacion|acciones|capacidades|cooperacion)\b",
    r"\bcapacitacion\b",
    r"\bcapacita(?:n|cion)?\b",
    r"\bsimulacro\b",
    r"\bejercicio\s+multinacional\b",
    r"\bconferencia\b",
    r"\breunion\s+multisectorial\b",
    r"\bcooperacion\s+tecnica\b",
    r"\bcertificacion\s+de\s+competencias\b",
    r"\bconvoca\b.*\bproceso\b",
    r"\bposibles?\s+impactos?\b",
    r"\bante\s+la\s+posible\s+ocurrencia\b",
    r"\bpronostico\b",
    r"\bescenario\s+de\s+riesgo\b",
    r"\binspeccion(?:o|aron|a|es)?\b",
    r"\bsupervision(?:es)?\b",
    r"\bobra(?:s)?\b",
    r"\binfraestructura\b",
    r"\bmantenimiento\b",
    r"\bservicios?\s+publicos?\b",
]

# -----------------------------------------------------------------------------
# MODELOS / UTILIDADES
# -----------------------------------------------------------------------------

@dataclass
class Article:
    institucion: str
    url: str
    title: str
    text: str
    published: Optional[datetime]
    direct_fen: bool
    direct_fen_primary: bool
    hazard: bool
    hazard_primary: bool
    event_occurred: bool
    event_occurred_primary: bool
    non_event: bool
    # En COEN-INDECI se conserva por separado el contenido HTML y el PDF oficial.
    html_text: str = ""
    pdf_url: Optional[str] = None
    pdf_text: str = ""
    pdf_pages: int = 0
    # Fecha real de ocurrencia cuando COEN la publica. Si no existe, se usa published.
    event_date: Optional[datetime] = None
    # Nivel de evidencia de vinculo con el FEN (se calcula en el hilo principal).
    fen_level: str = ""
    fen_note: str = ""


def strip_accents(value: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", value)
        if unicodedata.category(c) != "Mn"
    )


def norm_text(value: Optional[str]) -> str:
    value = strip_accents(value or "").lower()
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def candidate_event_context_ok(article_text: str, article_title: str) -> bool:
    """Exige que peligro e impacto esten cercanos en el titulo/lead.

    Reduce falsos positivos de notas de obras/inspecciones que mencionan
    peligros hidrometeorologicos solo como contexto.
    """
    primary = norm_text(f"{article_title} {article_text[:1600]}")
    hazard_hits = []
    for pattern in HAZARD_PATTERNS:
        for m in re.finditer(pattern, primary):
            hazard_hits.append(m.start())

    occurred_hits = []
    for pattern in EVENT_OCCURRED_PATTERNS:
        for m in re.finditer(pattern, primary):
            occurred_hits.append(m.start())

    if not hazard_hits or not occurred_hits:
        return False

    # Deben aparecer razonablemente cerca para asumir que describen el mismo hecho.
    return min(abs(h - o) for h in hazard_hits for o in occurred_hits) <= 350


def normalize_url(url: str) -> str:
    p = urlparse(url.strip())
    clean_query = [
        (k, v) for k, v in parse_qsl(p.query)
        if not k.lower().startswith("utm_")
    ]
    path = re.sub(r"/+$", "", p.path) or "/"
    return urlunparse(
        (p.scheme.lower(), p.netloc.lower(), path, "", urlencode(clean_query), "")
    )


def setup_logging(log_dir: Path) -> Path:
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        # Fallback util si C:\Automatizaciones aun no existe/no tiene permisos.
        log_dir = Path.home() / "FEN_logs"
        log_dir.mkdir(parents=True, exist_ok=True)

    log_path = log_dir / f"actualizacion_fen_{datetime.now():%Y%m%d_%H%M%S}.log"
    handlers = [logging.FileHandler(log_path, encoding="utf-8")]
    # pythonw.exe no expone stdout/stderr. Solo agrega consola si existe.
    if sys.stdout is not None:
        handlers.append(logging.StreamHandler(sys.stdout))

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=handlers,
    )
    return log_path


# gob.pe limita las rafagas: responde HTTP 200 con una pagina casi vacia (~7.9 KB, sin fecha
# ni contenido) en vez de la noticia. Sin control, un backfill perdia en silencio la mayor
# parte de INDECI y Contraloria (56-76 % de los articulos salian "sin fecha").
GOB_PE_SEMAPHORE = threading.BoundedSemaphore(3)
GOB_PE_MIN_ARTICLE_BYTES = 15000
# Contadores de limitacion (lectura/escritura desde varios hilos: se protegen con un lock).
THROTTLE_STATS = {"reintentos": 0, "no_recuperadas": 0, "indices_vacios": 0}
THROTTLE_LOCK = threading.Lock()


def _is_gob_pe_article(url: str) -> bool:
    parsed = urlparse(url)
    return parsed.netloc.lower().endswith("gob.pe") and re.search(r"/noticias/\d", parsed.path) is not None


def request_html(session: requests.Session, url: str) -> str:
    """Obtiene HTML con reintentos solo para errores transitorios.

    404/410 no se reintentan: en paginacion significan normalmente que ya no
    existe esa pagina y reintentarlos solo hace mas lento el backfill.
    Las noticias de gob.pe se piden con concurrencia limitada y, si llega la pagina
    vacia de limitacion, se reintenta con espera creciente.
    """
    last_exc = None
    gob_pe_article = _is_gob_pe_article(url)
    for attempt in range(1, 6 if gob_pe_article else 4):
        try:
            if gob_pe_article:
                with GOB_PE_SEMAPHORE:
                    response = session.get(url, timeout=TIMEOUT, verify=False)
                    time.sleep(0.15)
            else:
                response = session.get(url, timeout=TIMEOUT, verify=False)
            if response.status_code in (404, 410):
                logging.info("HTTP %s sin reintento: %s", response.status_code, url)
                return ""
            response.raise_for_status()
            content_type = response.headers.get("content-type", "").lower()
            if "html" not in content_type and "text" not in content_type:
                return ""
            if gob_pe_article and len(response.text) < GOB_PE_MIN_ARTICLE_BYTES:
                with THROTTLE_LOCK:
                    THROTTLE_STATS["reintentos"] += 1
                if attempt < 5:
                    time.sleep(attempt * 3)
                    continue
                with THROTTLE_LOCK:
                    THROTTLE_STATS["no_recuperadas"] += 1
                logging.warning("[GOBPE-LIMITADO] pagina vacia tras 5 intentos: %s", url)
            return response.text
        except requests.RequestException as exc:
            last_exc = exc
            status = getattr(getattr(exc, "response", None), "status_code", None)
            # Otros 4xx tampoco suelen mejorar con reintentos.
            if status is not None and 400 <= status < 500 and status not in (408, 429):
                logging.warning("HTTP %s sin reintento para %s: %s", status, url, exc)
                return ""
            logging.warning(
                "HTTP intento %s/3 fallo para %s: %s", attempt, url, exc
            )
            time.sleep(attempt * 2)
    raise last_exc  # type: ignore[misc]


def allowed_article_url(url: str, cfg: dict) -> bool:
    p = urlparse(url)
    if p.netloc.lower() != cfg["allowed_host"]:
        return False

    path = p.path.lower()

    # Las paginas de paginacion (/page/N/) son indices, no articulos: su menu y
    # lateral citan avisos de lluvias y generaban "eventos" falsos.
    if re.search(r"/page/\d+/?$", path):
        return False

    # Nunca tratar los propios indices como articulos.
    index_paths = {
        urlparse(index_url).path.rstrip("/") or "/"
        for index_url in cfg.get("index_urls", [])
    }
    if (path.rstrip("/") or "/") in index_paths:
        return False

    if "path_regex" in cfg:
        return re.search(cfg["path_regex"], path, flags=re.I) is not None

    if "path_contains" in cfg:
        return cfg["path_contains"] in path

    return any(x in path for x in cfg.get("path_contains_any", []))


def discover_links(session: requests.Session, cfg: dict) -> set[str]:
    """Descubrimiento simple usado por el modo semanal."""
    links: set[str] = set()
    total = len(cfg["index_urls"])
    for pos, index_url in enumerate(cfg["index_urls"], start=1):
        try:
            logging.info("  Leyendo indice %s/%s: %s", pos, total, index_url)
            html = request_html(session, index_url)
        except Exception as exc:
            logging.warning("No se pudo leer indice %s: %s", index_url, exc)
            continue
        if not html:
            continue

        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            url = normalize_url(urljoin(index_url, a["href"]))
            if allowed_article_url(url, cfg):
                links.add(url)
        time.sleep(SLEEP_SECONDS)
    return links


SPANISH_MONTHS = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5,
    "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9,
    "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}


def parse_date_string(value: str) -> Optional[datetime]:
    value = value.strip().replace("Z", "+00:00")
    if not value:
        return None

    try:
        return datetime.fromisoformat(value)
    except Exception:
        pass

    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, fmt)
        except Exception:
            pass

    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(20\d{2})\b", value)
    if m:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))

    # Fechas habituales en gob.pe: "5 de octubre de 2026".
    # Se busca dentro de cadenas mas largas, por ejemplo
    # "Publicado el 28 de setiembre de 2026 - 11:57 a. m.".
    normalized = norm_text(value)
    m = re.search(
        r"\b(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(20\d{2})\b",
        normalized,
    )
    if m and m.group(2) in SPANISH_MONTHS:
        return datetime(
            int(m.group(3)), SPANISH_MONTHS[m.group(2)], int(m.group(1))
        )

    # Variante usada por ENFEN en algunos listados: "28 Septiembre, 2026".
    m = re.search(
        r"\b(\d{1,2})\s+([a-z]+),?\s+(20\d{2})\b",
        normalized,
    )
    if m and m.group(2) in SPANISH_MONTHS:
        return datetime(
            int(m.group(3)), SPANISH_MONTHS[m.group(2)], int(m.group(1))
        )

    return None


def parse_url_date(url: str) -> Optional[datetime]:
    """Fallback para fechas incluidas en URLs COEN y WordPress/ENFEN."""
    path = urlparse(url).path

    # COEN: ...-26-3-2026-... / ...-26-03-2026-...
    for m in re.finditer(r"(?:^|[-/])(\d{1,2})-(\d{1,2})-(20\d{2})(?:[-/]|$)", path):
        day, month, year = map(int, m.groups())
        try:
            return datetime(year, month, day)
        except ValueError:
            continue

    # WordPress / ENFEN: /2026/09/28/titulo/
    m = re.search(r"/(20\d{2})/(\d{1,2})/(\d{1,2})(?:/|$)", path)
    if m:
        year, month, day = map(int, m.groups())
        try:
            return datetime(year, month, day)
        except ValueError:
            pass

    return None


def parse_visible_date(soup: BeautifulSoup) -> Optional[datetime]:
    """Busca una fecha visible en la pagina.

    gob.pe suele mostrar fechas como:
        28 de setiembre de 2026 - 11:57 a. m.
    aunque no siempre las expone en JSON-LD/meta de forma consistente.
    """
    # Primero revisa elementos cortos donde normalmente vive la fecha.
    candidates: list[str] = []
    for tag in soup.find_all(["time", "p", "span", "div", "small"], limit=250):
        txt = tag.get_text(" ", strip=True)
        if txt and len(txt) <= 180:
            candidates.append(txt)

    # Fallback: texto completo de la pagina.
    candidates.append(soup.get_text(" ", strip=True))

    for txt in candidates:
        parsed = parse_date_string(txt)
        if parsed:
            return parsed

    return None


def parse_meta_date(soup: BeautifulSoup) -> Optional[datetime]:
    selectors = [
        ("meta", {"property": "article:published_time"}, "content"),
        ("meta", {"name": "date"}, "content"),
        ("meta", {"name": "DC.date"}, "content"),
        ("meta", {"itemprop": "datePublished"}, "content"),
    ]
    for tag_name, attrs, attr in selectors:
        tag = soup.find(tag_name, attrs=attrs)
        if tag and tag.get(attr):
            parsed = parse_date_string(str(tag.get(attr)))
            if parsed:
                return parsed
    return None

def _jsonld_items(data):
    if isinstance(data, list):
        for item in data:
            yield from _jsonld_items(item)
    elif isinstance(data, dict):
        yield data
        graph = data.get("@graph")
        if graph:
            yield from _jsonld_items(graph)


def parse_jsonld_date(soup: BeautifulSoup) -> Optional[datetime]:
    for tag in soup.find_all("script", attrs={"type": "application/ld+json"}):
        raw = tag.string or tag.get_text(" ", strip=True)
        if not raw:
            continue
        try:
            data = json.loads(raw)
        except Exception:
            continue

        for item in _jsonld_items(data):
            value = item.get("datePublished") or item.get("dateCreated")
            if value:
                parsed = parse_date_string(str(value))
                if parsed:
                    return parsed
    return None


COEN_REPORT_TITLE_PATTERNS = [
    r"\breporte\s+complementario\b",
    r"\breporte\s+preliminar\b",
    r"\breporte\s+de\s+emergencia\b",
    r"\breporte\s+de\s+situacion\b",
]

COEN_HYDROMET_TYPES = [
    r"\blluvias?\s+intensas?\b",
    r"\bprecipitacion(?:es)?\b",
    r"\binundacion(?:es)?\b",
    r"\bhuaicos?\b",
    r"\bhuaycos?\b",
    r"\bdesborde(?:s)?\b",
    r"\bdesbordamiento(?:s)?\b",
    r"\bactivacion\s+de\s+quebrada(?:s)?\b",
    r"\bmovimientos?\s+en\s+masa\b",
    r"\bdeslizamiento(?:s)?\b",
    r"\baluvion(?:es)?\b",
    r"\bcrecida(?:s)?\b",
    r"\bavenida(?:s)?\s+(?:de\s+)?(?:rio|quebrada)\b",
    r"\bvientos?\s+fuertes?\b",
    r"\berosion(?:es)?\b",
]

COEN_NON_HYDROMET_TYPES = [
    r"\bincendio(?:s)?\b",
    r"\bsismo\b",
    r"\bterremoto\b",
    r"\baccidente\b",
    r"\baccidente\s+de\s+transporte\b",
    r"\berupcion(?:es)?\b",
    r"\bincendio\s+forestal\b",
    r"\bincendio\s+urbano\b",
]


def extract_coen_report_number_from_url(url: str) -> Optional[str]:
    """Obtiene el numero de reporte directamente del slug COEN."""
    path = unquote(urlparse(url).path)
    normalized = norm_text(path)
    patterns = [
        r"reporte\s+(?:complementario|preliminar|de\s+emergencia|de\s+situacion)[^0-9]{0,30}(\d{3,6})",
        r"(?:reporte|informe)[^0-9]{0,20}(\d{3,6})",
    ]
    for pattern in patterns:
        m = re.search(pattern, normalized)
        if m:
            return m.group(1)
    return None


def coen_url_worth_fetching(url: str) -> bool:
    """True si el SLUG de la URL COEN describe un reporte de emergencia hidrometeorologico.

    Evita descargar (HTML + PDF) las ~2/3 partes del indice que son incendios, sismos, heladas,
    boletines y avisos. Exige numero de reporte en la URL y un tipo hidrometeorologico sin
    tipo excluyente (mismas listas que usa coen_is_hydromet_report sobre el titulo).
    """
    if extract_coen_report_number_from_url(url) is None:
        return False
    slug = norm_text(unquote(urlparse(url).path.rstrip("/").rsplit("/", 1)[-1]).replace("-", " "))
    return coen_is_hydromet_report(slug, "")


def extract_coen_title(soup: BeautifulSoup, page_url: str = "") -> str:
    """Extrae el titulo del reporte COEN correspondiente a ESTA URL."""
    report_no = extract_coen_report_number_from_url(page_url)

    # 1) Metadatos del articulo: suelen ser mas confiables que encabezados genericos.
    meta_candidates = []
    for attrs in (
        {"property": "og:title"},
        {"name": "twitter:title"},
        {"name": "title"},
    ):
        tag = soup.find("meta", attrs=attrs)
        if tag and tag.get("content"):
            meta_candidates.append(str(tag.get("content")).strip())

    title_tag = soup.find("title")
    if title_tag:
        meta_candidates.append(title_tag.get_text(" ", strip=True))

    # 2) Encabezados que contengan el mismo numero del reporte de la URL.
    heading_candidates = []
    for tag in soup.find_all(["h1", "h2", "h3", "h4", "strong", "b"], limit=250):
        txt = re.sub(r"\s+", " ", tag.get_text(" ", strip=True)).strip()
        if txt:
            heading_candidates.append(txt)

    # Sin numero de reporte en la URL este bucle devolveria el primer encabezado de la
    # pagina, que es el del carrusel de ultimos reportes; se salta (ver paso 3).
    for txt in (meta_candidates + heading_candidates) if report_no else []:
        n = norm_text(txt)
        if not any(re.search(pattern, n) for pattern in COEN_REPORT_TITLE_PATTERNS):
            continue
        if not re.search(rf"\b{re.escape(report_no)}\b", n):
            continue
        return txt

    # 3) Sin numero en la URL (boletines, avisos, sismos...): SOLO encabezados del cuerpo
    #    propio de la pagina. Los demas encabezados pertenecen al carrusel de ultimos
    #    reportes y asignaban a estas paginas el titulo de OTRO evento.
    if not report_no:
        body = soup.select_one("div.post-content")
        if body is not None:
            for tag in body.find_all(["h1", "h2", "h3", "h4", "strong", "b"], limit=60):
                txt = re.sub(r"\s+", " ", tag.get_text(" ", strip=True)).strip()
                if txt and any(re.search(pattern, norm_text(txt)) for pattern in COEN_REPORT_TITLE_PATTERNS):
                    return txt

    # 4) Construccion desde el slug como ultimo recurso. Esto evita tomar el reporte
    #    mas reciente del listado cuando la pagina usa encabezados compartidos.
    slug = unquote(urlparse(page_url).path.rstrip("/").split("/")[-1])
    slug = re.sub(r"[-_]+", " ", slug)
    slug = re.sub(r"\s+", " ", slug).strip()
    return slug[:220]


def extract_coen_report_number(value: str, page_url: str = "") -> Optional[str]:
    """Numero de reporte, priorizando la URL de la pagina."""
    from_url = extract_coen_report_number_from_url(page_url) if page_url else None
    if from_url:
        return from_url
    n = norm_text(value)
    m = re.search(
        r"reporte\s+(?:complementario|preliminar|de\s+emergencia|de\s+situacion)"
        r"[^0-9]{0,40}(\d{3,6})",
        n,
    )
    return m.group(1) if m else None


def _anchor_label(a) -> str:
    parts = [
        a.get_text(" ", strip=True),
        str(a.get("title", "")),
        str(a.get("aria-label", "")),
        str(a.get("data-title", "")),
    ]
    return norm_text(" ".join(parts))


_SIG_STOPWORDS = {"en", "el", "la", "los", "las", "de", "del", "y", "distrito", "provincia", "departamento", "centro", "poblado"}


def coen_event_signature(page_url: str, title: str = "") -> set[str]:
    """Tokens (tipo de emergencia + lugar) que identifican el EVENTO de un reporte COEN.

    Los complementarios enlazan al PDF de la ultima actualizacion, que lleva OTRO
    numero de reporte; lo que se mantiene igual es el tipo de evento y el lugar.
    """
    candidates = []
    if page_url:
        candidates.append(norm_text(unquote(urlparse(page_url).path.rstrip("/").rsplit("/", 1)[-1])))
    if title:
        candidates.append(norm_text(title))
    for text in candidates:
        m = re.search(r"\bhoras\s+(?:(?:reporte|informe)\s+n\s*o?\s*\d+\s+)?(.+)$", text)
        if m:
            phrase = re.split(r"\s+indeci\s+tarea\b", m.group(1))[0]
            tokens = {t.rstrip("s") for t in phrase.split() if len(t) >= 3 and not t.isdigit() and t not in _SIG_STOPWORDS}
            if len(tokens) >= 3:
                return tokens
    return set()


def _signature_ratio(signature: set[str], text: str) -> float:
    if not signature:
        return 0.0
    words = {w.rstrip("s") for w in re.findall(r"[a-z0-9]+", text)}
    return len(signature & words) / len(signature)


def find_coen_pdf_url(
    soup: BeautifulSoup,
    page_url: str,
    report_title: str,
) -> Optional[str]:
    """Selecciona el PDF exacto del reporte actual sin descargar PDFs ajenos."""
    report_no = extract_coen_report_number(report_title, page_url)
    if not report_no:
        logging.warning(
            "[PDF-NO-COINCIDENCIA] COEN | no se pudo identificar numero de reporte | url=%s",
            page_url,
        )
        return None

    candidates: list[tuple[int, int, str, str]] = []
    signature = coen_event_signature(page_url, report_title)

    for a in soup.find_all("a", href=True):
        href_raw = str(a.get("href", "")).strip()
        if not href_raw:
            continue
        label = _anchor_label(a)
        absolute = normalize_url(urljoin(page_url, href_raw))
        href_decoded = norm_text(unquote(absolute))
        is_pdf = ".pdf" in href_decoded
        is_download = "descargar archivo" in label or label == "descargar" or "descargar archivo" in norm_text(str(a))
        if not (is_pdf or is_download):
            continue

        # Solo PDFs cuyo ENLACE tenga el numero del reporte, o (complementarios) el mismo
        # evento: tipo + lugar iguales aunque el PDF sea una actualizacion con otro numero.
        number_match = re.search(rf"\b{re.escape(report_no)}\b", href_decoded) is not None
        filename = norm_text(unquote(urlparse(absolute).path.rsplit("/", 1)[-1]))
        event_match = _signature_ratio(signature, filename) >= 0.8
        if not (number_match or event_match):
            continue

        score = 1000 if number_match else 800
        file_no = re.search(r"(?<!\d)(\d{4,6})(?!\d)", filename)
        latest = int(file_no.group(1)) if file_no else 0
        if is_download:
            score += 200
        if f"reporte-complementario-n" in href_decoded or "reporte-preliminar-n" in href_decoded:
            score += 50
        candidates.append((score, latest, absolute, label))

    if not candidates:
        logging.warning(
            "[PDF-NO-COINCIDENCIA] COEN | reporte=%s | no se encontro enlace PDF del mismo reporte/evento | titulo=%s",
            report_no,
            report_title[:140],
        )
        return None

    candidates.sort(key=lambda x: (-x[0], -x[1], x[2]))
    best_score, _best_latest, best_url, best_label = candidates[0]
    logging.debug(
        "[PDF-SELECCIONADO] COEN | reporte=%s | score=%s | etiqueta=%s | url=%s",
        report_no,
        best_score,
        best_label,
        best_url,
    )
    return best_url


def coen_pdf_matches_article(title: str, pdf_text: str, page_url: str = "") -> bool:
    """Validacion secundaria: el PDF debe corresponder al reporte actual."""
    if not pdf_text.strip():
        return False
    report_no = extract_coen_report_number(title, page_url)
    pdf_norm = norm_text(pdf_text[:16000])
    if report_no and re.search(rf"\b{re.escape(report_no)}\b", pdf_norm):
        return True
    signature = coen_event_signature(page_url, title)
    if signature:
        # Actualizacion del mismo evento (otro numero de reporte): tipo + lugar deben coincidir.
        return _signature_ratio(signature, pdf_norm) >= 0.8
    if report_no:
        return False
    return any(re.search(pattern, pdf_norm) for pattern in COEN_REPORT_TITLE_PATTERNS)


def request_pdf_bytes(session: requests.Session, url: str) -> bytes:
    """Descarga un PDF con reintentos y limite de tamanio."""
    last_exc = None
    for attempt in range(1, 4):
        try:
            response = session.get(
                url,
                timeout=TIMEOUT,
                verify=False,
                allow_redirects=True,
            )
            response.raise_for_status()
            data = response.content
            if len(data) > PDF_MAX_BYTES:
                raise ValueError(
                    f"PDF demasiado grande ({len(data) / 1024 / 1024:.1f} MB; max {PDF_MAX_BYTES / 1024 / 1024:.0f} MB)"
                )
            content_type = response.headers.get("content-type", "").lower()
            if "pdf" not in content_type and not data.startswith(b"%PDF"):
                raise ValueError(f"El enlace no devolvio un PDF (content-type={content_type})")
            return data
        except Exception as exc:
            last_exc = exc
            logging.warning("PDF intento %s/3 fallo para %s: %s", attempt, url, exc)
            time.sleep(attempt * 2)
    raise last_exc  # type: ignore[misc]


def extract_pdf_text(pdf_bytes: bytes) -> tuple[str, int]:
    """Extrae texto de un PDF digital. No hace OCR.

    Inserta marcadores de pagina. Son importantes para COEN: sus encabezados y
    pies institucionales se repiten en cada pagina y no deben mezclarse con la
    seccion UBICACION del evento.
    """
    if PdfReader is None:
        raise RuntimeError("Falta instalar la dependencia 'pypdf'.")
    reader = PdfReader(BytesIO(pdf_bytes))
    page_count = len(reader.pages)
    texts: list[str] = []
    for idx, page in enumerate(reader.pages[:PDF_MAX_PAGES]):
        try:
            txt = page.extract_text() or ""
        except Exception as exc:
            logging.warning("No se pudo extraer texto de pagina PDF %s: %s", idx + 1, exc)
            txt = ""
        if txt.strip():
            texts.append(f"[PAGINA PDF {idx + 1}]\n{txt}")
    return "\n".join(texts).strip(), page_count


def parse_occurrence_date(value: str) -> Optional[datetime]:
    compact = norm_text(value)
    patterns = [
        r"fecha\s+y\s+hora\s+de\s+la\s+ocurrencia.{0,300}?(\d{1,2}/\d{1,2}/20\d{2})",
        r"fecha\s+de\s+(?:la\s+)?ocurrencia.{0,300}?(\d{1,2}/\d{1,2}/20\d{2})",
        r"fecha\s+de\s+ocurrencia.{0,300}?(\d{1,2}/\d{1,2}/20\d{2})",
    ]
    for pattern in patterns:
        m = re.search(pattern, compact)
        if m:
            return parse_date_string(m.group(1))
    return None


def parse_coen_report_date(title: str) -> Optional[datetime]:
    """Fallback: fecha impresa en el titulo oficial del reporte COEN.

    Se usa solo cuando no fue posible extraer la fecha de ocurrencia de HECHOS.
    Es preferible a la fecha de publicacion de la pagina web para un backfill
    historico, aunque sigue siendo una fecha de reporte y no siempre de ocurrencia.
    """
    n = norm_text(title)
    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(20\d{2})\b", n)
    if not m:
        return None
    try:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def coen_non_event(primary_compact: str) -> bool:
    """Exclusiones de gestion/preparacion para COEN."""
    management_patterns = [
        r"\binspeccion(?:o|aron|a|es)?\b",
        r"\bsupervision(?:es)?\b",
        r"\bobra(?:s)?\b",
        r"\binfraestructura\b",
        r"\bmantenimiento\b",
        r"\bservicios?\s+publicos?\b",
    ]
    generic_patterns = [
        pattern for pattern in NON_EVENT_PATTERNS
        if pattern not in management_patterns
    ]
    return any(re.search(pattern, primary_compact) for pattern in generic_patterns)


def coen_is_hydromet_report(title: str, html_text: str) -> bool:
    """Determina si el reporte COEN es hidrometeorologico por su TEMA principal.

    Importante: NO usa todo el HTML para esta decision. Las paginas COEN suelen
    incluir documentos relacionados (avisos de lluvia, boletines, sismos, etc.)
    que podrian disparar falsamente la descarga del PDF.
    """
    title_norm = norm_text(title)
    return any(re.search(pattern, title_norm) for pattern in COEN_HYDROMET_TYPES) and not any(re.search(pattern, title_norm) for pattern in COEN_NON_HYDROMET_TYPES)


def _new_http_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept-Language": "es-PE,es;q=0.9,en;q=0.7",
    })
    return session


def get_thread_session() -> requests.Session:
    """Una Session por thread; requests.Session no se comparte entre workers."""
    session = getattr(THREAD_LOCAL, "session", None)
    if session is None:
        session = _new_http_session()
        THREAD_LOCAL.session = session
    return session


def extract_article_threaded(institucion: str, url: str) -> Optional[Article]:
    return extract_article(get_thread_session(), institucion, url)


def extract_article(
    session: requests.Session, institucion: str, url: str
) -> Optional[Article]:
    try:
        html = request_html(session, url)
    except Exception as exc:
        logging.warning("No se pudo leer articulo %s: %s", url, exc)
        return None

    if not html:
        return None

    soup = BeautifulSoup(html, "html.parser")
    published = parse_jsonld_date(soup) or parse_meta_date(soup)
    if not published:
        time_tag = soup.find("time")
        if time_tag:
            published = parse_date_string(
                time_tag.get("datetime", "") or time_tag.get_text(" ", strip=True)
            )
    if not published:
        published = parse_visible_date(soup)
    if not published:
        published = parse_url_date(url)
    if institucion == "COEN - INDECI":
        # La fecha visible/meta de las paginas COEN es la del dia de consulta (boletines
        # y avisos de 2024 salian con la fecha de hoy); la URL lleva la fecha del reporte.
        published = parse_url_date(url) or published

    if institucion == "COEN - INDECI":
        title = extract_coen_title(soup, url)
    else:
        title_tag = soup.find("h1") or soup.find("title")
        title = title_tag.get_text(" ", strip=True) if title_tag else ""

    # El enlace al archivo debe encontrarse antes de eliminar nav/header/footer.
    pdf_url = None

    # Eliminamos ruido solo despues de extraer fecha/titulo/enlace PDF.
    for tag in soup(["script", "style", "nav", "footer", "header", "noscript"]):
        tag.decompose()

    main = None
    if institucion == "COEN - INDECI":
        # El primer <article> de portal.indeci.gob.pe es el carrusel con los ULTIMOS reportes
        # (de otros eventos); el reporte propio (HECHOS, tabla de ocurrencia) vive en
        # div.post-content. Usar el carrusel contaminaba el contexto primario y las fechas.
        main = soup.select_one("div.post-content")
    main = main or soup.find("main") or soup.find("article") or soup.body or soup
    html_text = main.get_text(" ", strip=True)

    # Solo descargamos el PDF de COEN si el HTML ya apunta a un peligro relevante
    # o menciona FEN. Evita descargar incendios, sismos, accidentes, etc.
    preliminary = norm_text(f"{title} {html_text}")
    preliminary_relevant = (
        coen_is_hydromet_report(title, html_text)
        if institucion == "COEN - INDECI"
        else (
            any(re.search(pattern, preliminary) for pattern in HAZARD_PATTERNS)
            or any(re.search(pattern, preliminary) for pattern in DIRECT_FEN_PATTERNS)
        )
    )

    # Solo los reportes/informes de emergencia (con numero en la URL) tienen PDF propio;
    # boletines y avisos no son eventos y no deben disparar la busqueda del PDF.
    if (
        institucion == "COEN - INDECI"
        and preliminary_relevant
        and extract_coen_report_number_from_url(url)
    ):
        pdf_url = find_coen_pdf_url(soup, url, title)

    pdf_text = ""
    pdf_pages = 0
    if pdf_url and preliminary_relevant:
        try:
            data = request_pdf_bytes(session, pdf_url)
            pdf_text, pdf_pages = extract_pdf_text(data)
            if pdf_text and not coen_pdf_matches_article(title, pdf_text):
                logging.warning(
                    "[PDF-DESCARTADO] COEN | el archivo no coincide con el reporte | titulo=%s | url=%s",
                    title[:120],
                    pdf_url,
                )
                pdf_text = ""
                pdf_pages = 0
                pdf_url = None
            elif pdf_text:
                logging.info(
                    "[PDF-OK] COEN | paginas=%s | caracteres=%s | %s",
                    pdf_pages,
                    len(pdf_text),
                    title[:120],
                )
            else:
                logging.warning("[PDF-SIN-TEXTO] COEN | %s | %s", title[:120], pdf_url)
        except Exception as exc:
            logging.warning("[PDF-ERROR] COEN | %s | %s", title[:120], exc)

    # El texto completo enriquecido se usa para ubicacion, magnitud y detalle.
    text = html_text
    if pdf_text:
        text = f"{html_text}\n\n[CONTENIDO PDF OFICIAL COEN]\n{pdf_text}"

    compact = norm_text(f"{title} {text}")

    # Para COEN incorporamos el inicio del PDF al contexto primario porque ahi
    # suelen estar HECHOS, UBICACION y EVALUACION DE DANOS.
    title_compact = norm_text(title)
    lead_compact = norm_text(html_text[:1600])
    pdf_lead_compact = norm_text(pdf_text[:PDF_LEAD_CHARS]) if pdf_text else ""
    primary_compact = f"{title_compact} {lead_compact} {pdf_lead_compact}".strip()

    direct_fen = any(re.search(pattern, compact) for pattern in DIRECT_FEN_PATTERNS)
    direct_fen_primary = any(
        re.search(pattern, primary_compact) for pattern in DIRECT_FEN_PATTERNS
    )
    hazard = any(re.search(pattern, compact) for pattern in HAZARD_PATTERNS)
    hazard_primary = any(
        re.search(pattern, primary_compact) for pattern in HAZARD_PATTERNS
    )
    event_occurred = any(
        re.search(pattern, compact) for pattern in EVENT_OCCURRED_PATTERNS
    )
    event_occurred_primary = any(
        re.search(pattern, primary_compact) for pattern in EVENT_OCCURRED_PATTERNS
    )

    if institucion == "COEN - INDECI" and not coen_is_hydromet_report(title, ""):
        # El TIPO de emergencia del titulo manda: un incendio forestal en 'Huayco Pampa'
        # o un texto que cita 'lluvias' como contexto no es un evento hidrometeorologico.
        hazard = False
        hazard_primary = False

    if institucion == "COEN - INDECI":
        non_event = coen_non_event(primary_compact)
    else:
        non_event = any(
            re.search(pattern, primary_compact) for pattern in NON_EVENT_PATTERNS
        )

    event_date = None
    if institucion == "COEN - INDECI":
        event_date = (
            parse_occurrence_date(html_text)
            or parse_occurrence_date(pdf_text)
            or parse_coen_report_date(title)
        )

    return Article(
        institucion=institucion,
        url=url,
        title=title.strip(),
        text=text.strip(),
        published=published,
        direct_fen=direct_fen,
        direct_fen_primary=direct_fen_primary,
        hazard=hazard,
        hazard_primary=hazard_primary,
        event_occurred=event_occurred,
        event_occurred_primary=event_occurred_primary,
        non_event=non_event,
        html_text=html_text.strip(),
        pdf_url=pdf_url,
        pdf_text=pdf_text,
        pdf_pages=pdf_pages,
        event_date=event_date,
    )


# -----------------------------------------------------------------------------
# EXCEL
# -----------------------------------------------------------------------------

REQUIRED_EVENT_HEADERS = [
    "Evento",
    "Fecha",
    "Magnitud / severidad observada",
    "Escala_Magnitud",
    "Departamento",
    "Provincia",
    "Distrito afectado",
    "Zona",
    "Fuente oficial",
    "Detalle",
]


def header_map(ws) -> dict[str, int]:
    return {
        str(ws.cell(1, c).value).strip(): c
        for c in range(1, ws.max_column + 1)
        if ws.cell(1, c).value is not None
    }


def open_workbook_with_retry(path: Path):
    last_exc = None
    for attempt in range(1, OPEN_RETRIES + 1):
        try:
            return load_workbook(path)
        except (PermissionError, OSError) as exc:
            last_exc = exc
            logging.warning(
                "No se pudo abrir Excel. Intento %s/%s: %s",
                attempt,
                OPEN_RETRIES,
                exc,
            )
            time.sleep(OPEN_RETRY_SECONDS)
    raise last_exc  # type: ignore[misc]


def validate_workbook(wb) -> None:
    required_sheets = {"Eventos_reales", "Distritos Zonas"}
    missing_sheets = required_sheets.difference(wb.sheetnames)
    if missing_sheets:
        raise ValueError(
            "Faltan hojas requeridas: " + ", ".join(sorted(missing_sheets))
        )

    headers = header_map(wb["Eventos_reales"])
    missing_headers = [h for h in REQUIRED_EVENT_HEADERS if h not in headers]
    if missing_headers:
        raise ValueError(
            "Faltan columnas en Eventos_reales: " + ", ".join(missing_headers)
        )


def create_empty_workbook(path: Path) -> None:
    """Crea el Excel base (con las 3 hojas y sus columnas) cuando no existe en el directorio de trabajo."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Eventos_reales"
    sheets = {
        "Eventos_reales": REQUIRED_EVENT_HEADERS,
        "Distritos Zonas": ["Distrito (original)", "Etiqueta dashboard", "Departamento inferido"],
    }
    for name, headers in sheets.items():
        sheet = ws if name == "Eventos_reales" else wb.create_sheet(name)
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(1, col, header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        sheet.freeze_panes = "A2"
        for col in range(1, len(headers) + 1):
            sheet.column_dimensions[sheet.cell(1, col).column_letter].width = 28
    ensure_candidates_sheet(wb)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    logging.warning(
        "El Excel no existia; se creo uno nuevo en %s. La hoja 'Distritos Zonas' esta vacia: "
        "sin ese maestro territorial, Departamento/Distrito/Zona quedaran vacios.",
        path,
    )


CANONICAL_DEPARTMENTS = {
    "AMAZONAS": "Amazonas", "ANCASH": "Áncash", "APURIMAC": "Apurímac", "AREQUIPA": "Arequipa",
    "AYACUCHO": "Ayacucho", "CAJAMARCA": "Cajamarca", "CALLAO": "Callao", "CUSCO": "Cusco",
    "CUZCO": "Cusco", "HUANCAVELICA": "Huancavelica", "HUANUCO": "Huánuco", "ICA": "Ica",
    "JUNIN": "Junín", "LA LIBERTAD": "La Libertad", "LAMBAYEQUE": "Lambayeque", "LIMA": "Lima",
    "LORETO": "Loreto", "MADRE DE DIOS": "Madre de Dios", "MOQUEGUA": "Moquegua",
    "PASCO": "Pasco", "CERRO DE PASCO": "Pasco", "PIURA": "Piura", "PUNO": "Puno",
    "SAN MARTIN": "San Martín", "TACNA": "Tacna", "TUMBES": "Tumbes", "UCAYALI": "Ucayali",
}


def canonical_department(name: str) -> str:
    """Nombre oficial del departamento (el maestro trae CUZCO / CERRO DE PASCO).

    Debe coincidir con como lo escriben los reportes (CUSCO, PASCO), porque esos
    nombres se usan para restringir homonimos por departamento.
    """
    name = (name or "").strip()
    return CANONICAL_DEPARTMENTS.get(strip_accents(name).upper(), name.title())


DEFAULT_UBIGEO = Path(__file__).resolve().parent / "data" / "ubigeo_distrito.csv"
UBIGEO_COLUMN = "Ubigeo"
# (departamento, provincia, distrito) normalizados -> codigo INEI de 6 digitos.
UBIGEO_INDEX: dict[tuple[str, str, str], str] = {}


def load_ubigeo(path: Path) -> list[dict]:
    """Catalogo oficial de distritos (INEI): CSV con columnas inei, departamento, provincia, distrito."""
    if not path.exists():
        logging.warning("No existe el catalogo de ubigeos (%s); se usa solo el maestro Distritos Zonas.", path)
        return []
    rows: list[dict] = []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for rec in csv.DictReader(fh):
            code = str(rec.get("inei", "")).strip()
            if not re.fullmatch(r"\d{6}", code):
                continue
            district = strip_accents(str(rec.get("distrito", "")).strip()).upper()
            rows.append(
                {
                    "ubigeo": code,
                    "department": canonical_department(str(rec.get("departamento", "")).strip()),
                    "province": strip_accents(str(rec.get("provincia", "")).strip()).title(),
                    "district": district,
                    "district_norm": norm_text(district),
                }
            )
    logging.info("Catalogo de ubigeos: %s distritos (%s)", len(rows), path.name)
    return rows


def ubigeo_code(department: str, province: str, district: str) -> str:
    return UBIGEO_INDEX.get((norm_text(department), norm_text(province), norm_text(district)), "")


def location_ubigeo(article: "Article", department: str, province: str, district: str) -> str:
    """Ubigeo del distrito afectado; la tabla del PDF decide la provincia si el nombre se repite."""
    if district and article.pdf_text:
        province = coen_province_for(article.pdf_text, district) or province
    return ubigeo_code(department, province, district) if district else ""


def merge_ubigeo_into_catalog(catalog: list[dict], ubigeo_rows: list[dict]) -> None:
    """Completa el maestro con el catalogo INEI.

    - Distritos que faltan en 'Distritos Zonas' (Lima metropolitana, nombres truncados del
      maestro) se agregan con la Zona de su departamento (la hoja Metodologia_y_fuentes define
      la zona a nivel departamento).
    - Los que ya estaban reciben provincia y ubigeo cuando son unicos en el departamento.
    - Un mismo nombre en dos provincias del mismo departamento (p. ej. 'Pampas' en Ancash)
      queda sin provincia/ubigeo: la provincia la decide el PDF del reporte.
    """
    zones_by_dept: dict[str, set[str]] = collections.defaultdict(set)
    for item in catalog:
        if item["zone"]:
            zones_by_dept[norm_text(item["department"])].add(item["zone"])

    UBIGEO_INDEX.clear()
    by_key: dict[tuple[str, str], list[dict]] = collections.defaultdict(list)
    for rec in ubigeo_rows:
        UBIGEO_INDEX[(norm_text(rec["department"]), norm_text(rec["province"]), rec["district_norm"])] = rec["ubigeo"]
        by_key[(norm_text(rec["department"]), rec["district_norm"])].append(rec)

    # El maestro trae filas repetidas (p. ej. 'Pampas' x4 en Ancash): se deduplican y cada
    # una recibe provincia/ubigeo.
    deduped: dict[tuple[str, str], dict] = {}
    for item in catalog:
        deduped.setdefault((norm_text(item["department"]), item["district_norm"]), item)
    catalog[:] = list(deduped.values())
    known = deduped
    added = 0
    for key, recs in by_key.items():
        province = recs[0]["province"] if len({r["province"] for r in recs}) == 1 else ""
        code = recs[0]["ubigeo"] if len(recs) == 1 else ""
        item = known.get(key)
        if item is not None:
            item["province"], item["ubigeo"] = province, code
            continue
        zones = zones_by_dept.get(key[0], set())
        catalog.append(
            {
                "district": recs[0]["district"],
                "district_norm": key[1],
                "department": recs[0]["department"],
                "zone": next(iter(zones)) if len(zones) == 1 else "",
                "province": province,
                "ubigeo": code,
            }
        )
        added += 1
    logging.info("Catalogo territorial: %s distritos agregados desde el ubigeo INEI (total %s)", added, len(catalog))


def load_location_catalog(wb, ubigeo_rows: Optional[list[dict]] = None) -> list[dict]:
    ws = wb["Distritos Zonas"]
    headers = header_map(ws)
    required = [
        "Distrito (original)",
        "Etiqueta dashboard",
        "Departamento inferido",
    ]
    for col in required:
        if col not in headers:
            raise ValueError(f"Falta columna '{col}' en Distritos Zonas")

    catalog: list[dict] = []
    for row in range(2, ws.max_row + 1):
        original = str(ws.cell(row, headers["Distrito (original)"]).value or "").strip()
        department = str(
            ws.cell(row, headers["Departamento inferido"]).value or ""
        ).strip()
        zone = str(ws.cell(row, headers["Etiqueta dashboard"]).value or "").strip()
        if not original:
            continue

        # Mantiene la logica del archivo original, pero evita depender de formulas.
        district = re.sub(r"-[A-ZÁÉÍÓÚÑ ]+$", "", original, flags=re.I).strip()
        catalog.append(
            {
                "district": district,
                "district_norm": norm_text(district),
                "department": canonical_department(department),
                "zone": zone,
                "province": "",
                "ubigeo": "",
            }
        )

    if ubigeo_rows:
        merge_ubigeo_into_catalog(catalog, ubigeo_rows)
    return sorted(catalog, key=lambda x: len(x["district_norm"]), reverse=True)


COEN_LOCATION_NOISE_LINE_PATTERNS = [
    r"^distribucion\s*:",
    r"^centro\s+de\s+operaciones\s+de\s+emergencia\s+nacional\b",
    r"^av\.?\s+el\s+sol\b",
    r"^tel\.?\s*\+?511\b",
    r"^coenperu\b",
    r"^p\s*a\s*g\s*i\s*n\s*a\b",
    r"^www\.indeci\.gob\.pe\b",
]


def clean_coen_pdf_location_text(value: str) -> str:
    """Quita boilerplate institucional del PDF antes de inferir ubicaciones.

    Nunca usamos domicilios, telefonos, cabeceras/pies ni datos de contacto como
    evidencia territorial del evento. Esto evita falsos positivos como CHORRILLOS,
    que aparece en la direccion del COEN/INDECI pero no es el distrito afectado.
    """
    cleaned: list[str] = []
    for raw_line in (value or "").splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if not line:
            cleaned.append("")
            continue
        n = norm_text(line)
        if any(re.search(pattern, n) for pattern in COEN_LOCATION_NOISE_LINE_PATTERNS):
            continue
        # Tambien elimina lineas de contacto/domicilio genericas.
        if "@" in line or re.search(r"\b(?:telefono|telefono|correo|direccion)\s*:", n):
            continue
        cleaned.append(line)
    return "\n".join(cleaned)


def extract_coen_pdf_section(pdf_text: str, section_number: int, heading: str) -> str:
    """Extrae una seccion COEN sin cruzar a otra pagina.

    Los marcadores [PAGINA PDF N] insertados durante la extraccion impiden que
    UBICACION capture el encabezado/pie de la pagina siguiente.
    """
    cleaned = clean_coen_pdf_location_text(pdf_text)
    normalized = "\n".join(norm_text(line) for line in cleaned.splitlines())
    heading_norm = norm_text(heading)
    pattern = (
        rf"(?is)(?:^|\n)\s*{section_number}\.?\s*{heading_norm}\s*:?\s*"
        rf"(.*?)"
        rf"(?=\n\s*(?:\[pagina pdf \d+\]|{section_number + 1}\.?\s+[a-z]))"
    )
    m = re.search(pattern, normalized)
    if m:
        return m.group(1).strip()

    # Fallback para PDFs antiguos extraidos sin marcadores de pagina.
    fallback = re.search(
        rf"(?is)(?:^|\n)\s*{section_number}\.?\s*{heading_norm}\s*:?\s*(.{{0,2200}}?)"
        rf"(?=\n\s*{section_number + 1}\.?\s+[a-z]|$)",
        normalized,
    )
    return fallback.group(1).strip() if fallback else ""


def _coen_explicit_district_match(context: str, district_norm: str) -> bool:
    """Exige una relacion linguistica explicita entre 'distrito' y el nombre."""
    if not context or not district_norm:
        return False
    d = re.escape(district_norm)
    patterns = [
        rf"\bdistrito\s+(?:de\s+)?{d}\b",
        rf"\bdistrito\s+y\s+provincia\s+de\s+{d}\b",
        rf"\bdistrito\s*,?\s*provincia\s+(?:y\s+departamento\s+)?(?:de\s+)?{d}\b",
        rf"\bdist\.?\s+{d}\b",
    ]
    return any(re.search(p, context) for p in patterns)


def _coen_district_in_list(context: str, district_norm: str) -> bool:
    """Reconoce listas tipo 'distritos de A, B y C' en un contexto acotado."""
    if not context or not district_norm:
        return False
    for m in re.finditer(r"\bdistritos\s+(?:de\s+)?", context):
        segment = context[m.end(): m.end() + 260]
        segment = re.split(r"[.;:\n]", segment, maxsplit=1)[0]
        if re.search(rf"\b{re.escape(district_norm)}\b", segment):
            return True
    return False


def _coen_department_hints(
    title_scope: str,
    location_scope: str,
    catalog: list[dict],
) -> set[str]:
    """Obtiene departamentos explicitamente mencionados en titulo/UBICACION.

    Se usa para desambiguar distritos homonimos. Ejemplo: SANTA CRUZ existe como
    distrito en Ica, pero si el reporte indica ANCASH no puede asignarse a Ica.
    """
    departments = sorted(
        {norm_text(item.get("department", "")) for item in catalog if item.get("department")},
        key=len,
        reverse=True,
    )

    # El titulo es la fuente mas fuerte para el departamento del evento.
    title_hits = {
        dep for dep in departments
        if dep and re.search(rf"\b{re.escape(dep)}\b", title_scope)
    }
    if title_hits:
        return title_hits

    # UBICACION es segunda opcion. Ya viene aislada del resto del PDF, por lo que
    # no arrastra el domicilio institucional de COEN/INDECI.
    return {
        dep for dep in departments
        if dep and re.search(rf"\b{re.escape(dep)}\b", location_scope)
    }


def _coen_has_non_district_role(context: str, district_norm: str) -> bool:
    """Detecta cuando un nombre aparece explicitamente con OTRO rol geografico.

    Un nombre que figura como centro poblado, sector o provincia no debe promoverse
    a distrito solo porque tambien exista como distrito en el maestro.
    """
    if not context or not district_norm:
        return False
    d = re.escape(district_norm)
    patterns = [
        rf"\bcentro\s+poblado\s+(?:de\s+)?{d}\b",
        rf"\bsector\s+(?:de\s+)?{d}\b",
        rf"\bprovincia\s+(?:de\s+)?{d}\b",
        rf"\bdepartamento\s+(?:de\s+)?{d}\b",
        rf"\bdpto\.?\s+(?:de\s+)?{d}\b",
    ]
    return any(re.search(pattern, context) for pattern in patterns)


def _coen_explicit_catalog_matches(
    context: str,
    catalog: list[dict],
    department_hints: set[str],
) -> list[dict]:
    """Devuelve solo distritos ligados linguisticamente a la palabra DISTRITO."""
    matches: list[dict] = []
    for item in catalog:
        dep = norm_text(item.get("department", ""))
        if department_hints and dep not in department_hints:
            continue
        d = item["district_norm"]
        if not d:
            continue
        if _coen_explicit_district_match(context, d) or _coen_district_in_list(context, d):
            matches.append(item)
    return matches


def _dedupe_location_items(items: list[dict]) -> list[dict]:
    result: list[dict] = []
    seen: set[tuple[str, str]] = set()
    for item in items:
        key = (norm_text(item.get("department", "")), item["district_norm"])
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


_TITLE_DEPARTMENTS = sorted(
    ((norm_text(key), canonical) for key, canonical in CANONICAL_DEPARTMENTS.items()),
    key=lambda kv: len(kv[0]),
    reverse=True,
)


def department_zone(catalog: list[dict], department: str) -> str:
    """Zona del departamento segun el maestro.

    La hoja Metodologia_y_fuentes define la zona a nivel DEPARTAMENTO (proxy V0 de
    CENEPRED): todos los distritos de un departamento comparten etiqueta. Si el
    maestro no la define de forma unica, se devuelve vacio.
    """
    dep_norm = norm_text(department)
    zones = {item["zone"] for item in catalog if norm_text(item["department"]) == dep_norm and item["zone"]}
    return next(iter(zones)) if len(zones) == 1 else ""


def coen_title_location(title: str) -> Optional[tuple[str, str]]:
    """(departamento, DISTRITO) desde '... EN EL DISTRITO DE X - DEPARTAMENTO' del titulo COEN.

    Sirve cuando el distrito NO esta en el maestro 'Distritos Zonas' (p. ej. Lima
    metropolitana): el titulo oficial sigue siendo evidencia autoritativa del lugar.
    No asigna Zona (solo el maestro puede hacerlo). Titulos con varios distritos
    ('A, B y C') se descartan por ambiguos.
    """
    text = norm_text(title)
    text = re.sub(r"\s+indeci\s+tarea\b.*$", "", text)
    m = re.search(r"\bdistrito\s+(?:de|del)\s+(.+)$", text)
    if not m:
        return None
    rest = m.group(1).strip(" -")
    for dep_norm, dep in _TITLE_DEPARTMENTS:
        if rest.endswith(" " + dep_norm) or rest.endswith("-" + dep_norm):
            district = rest[: -len(dep_norm)].strip(" -")
            if len(district) < 3 or "," in district or re.search(r"\by\b", district) and "distritos" in text:
                return None
            return dep, district.upper()
    return None


def infer_coen_locations(article: Article, catalog: list[dict]) -> list[tuple[str, str, str, str]]:
    """Inferencia territorial jerarquica y conservadora para reportes COEN.

    Regla principal: NO se decide la ubicacion buscando nombres sueltos en el PDF.
    Se respeta la jerarquia semantica del documento:

      1. Titulo: "... EN EL DISTRITO DE X - DEPARTAMENTO" -> autoritativo.
      2. HECHOS: menciones explicitas "distrito de X" -> alta confianza.
      3. UBICACION: solo corrobora candidatos ya identificados o aporta el
         departamento para desambiguar homonimos.
      4. Nombres marcados como CENTRO POBLADO / SECTOR / PROVINCIA / DEPARTAMENTO
         no se convierten en distrito.

    Precision > recall: si no hay una relacion explicita con DISTRITO, se deja la
    ubicacion vacia para revision en vez de asignar una homonimia incorrecta.
    """
    title_scope = norm_text(article.title)
    facts_scope = norm_text(extract_coen_pdf_section(article.pdf_text, 1, r"HECHOS"))
    location_scope = norm_text(extract_coen_pdf_section(article.pdf_text, 2, r"UBICACION"))
    html_lead_scope = norm_text(f"{article.title} {(article.html_text or '')[:1800]}")

    department_hints = _coen_department_hints(title_scope, location_scope, catalog)

    def to_locations(items: list[dict], evidence: str) -> list[tuple[str, str, str, str]]:
        items = _dedupe_location_items(items)
        out: list[tuple[str, str, str, str]] = []
        for item in items:
            out.append((item["department"], item.get("province", ""), item["district"].upper(), item["zone"]))
            logging.info(
                "[UBICACION-COEN] distrito=%s | departamento=%s | evidencia=%s | %s",
                item["district"].upper(),
                item["department"],
                evidence,
                article.title[:100],
            )
        return out

    # 1) TITULO: si contiene un distrito explicito, termina aqui. Esto evita que
    #    un centro poblado del cuerpo del PDF compita con el distrito del evento.
    title_matches = _coen_explicit_catalog_matches(title_scope, catalog, department_hints)
    if title_matches:
        return to_locations(title_matches, "titulo+departamento")

    # 2) HECHOS: segunda fuente autoritativa. Excluye nombres que en ese mismo
    #    contexto estan etiquetados como centro poblado/sector/provincia.
    facts_matches = [
        item for item in _coen_explicit_catalog_matches(facts_scope, catalog, department_hints)
        if not (
            _coen_has_non_district_role(facts_scope, item["district_norm"])
            and not _coen_explicit_district_match(facts_scope, item["district_norm"])
        )
    ]
    if facts_matches:
        return to_locations(facts_matches, "hechos+departamento")

    # 3) HTML/lead: util cuando el PDF no se pudo leer. Tambien exige DISTRITO.
    html_matches = _coen_explicit_catalog_matches(html_lead_scope, catalog, department_hints)
    if html_matches:
        return to_locations(html_matches, "html_explicit+departamento")

    # 4) UBICACION solo puede CORROBORAR un distrito que tenga relacion explicita
    #    con la palabra distrito. No aceptamos nombres aislados ni repetidos: un
    #    centro poblado puede compartir nombre con un distrito de otra region.
    location_matches = _coen_explicit_catalog_matches(location_scope, catalog, department_hints)
    location_matches = [
        item for item in location_matches
        if not (
            _coen_has_non_district_role(location_scope, item["district_norm"])
            and not _coen_explicit_district_match(location_scope, item["district_norm"])
        )
    ]
    if location_matches:
        return to_locations(location_matches, "ubicacion_explicit+departamento")

    # 5) TITULO sin maestro: el distrito y el departamento figuran explicitos en el titulo
    #    pero el distrito no esta en 'Distritos Zonas'. Se registran sin Zona.
    title_loc = coen_title_location(article.title)
    if title_loc:
        dep, district = title_loc
        zone = department_zone(catalog, dep)
        logging.info(
            "[UBICACION-COEN] distrito=%s | departamento=%s | zona=%s | evidencia=titulo_fuera_de_maestro | %s",
            district, dep, zone or "(sin zona)", article.title[:100],
        )
        return [(dep, "", district, zone)]

    logging.warning(
        "[UBICACION-COEN-SIN-MATCH] no se identifico distrito con evidencia jerarquica | departamento_hint=%s | %s",
        ",".join(sorted(department_hints)) or "SIN_HINT",
        article.title[:120],
    )
    return []

def infer_locations(article: Article, catalog: list[dict]) -> list[tuple[str, str, str, str]]:
    """Devuelve TODOS los distritos afectados detectados en la publicacion/PDF.

    COEN usa una metodologia especial y conservadora para separar la ubicacion
    del evento de domicilios institucionales, fuentes y firmas del documento.
    Para el resto de fuentes se mantiene la inferencia por menciones explicitas.
    """
    if article.institucion == "COEN - INDECI":
        return infer_coen_locations(article, catalog)

    primary_txt = norm_text(f"{article.title} {article.html_text[:2500] or article.text[:2500]}")
    html_scope = norm_text(f"{article.title} {article.html_text or article.text}")
    found: dict[tuple[str, str], tuple[str, str, str, str]] = {}

    dept_norms = {norm_text(c["department"]) for c in catalog if c["department"]}
    entries_per_name = collections.Counter(c["district_norm"] for c in catalog)
    dept_mentions = {d for d in dept_norms if re.search(rf"\b{re.escape(d)}\b", html_scope)}

    def add_item(item: dict, explicit: bool) -> None:
        name = item["district_norm"]
        # Un nombre que tambien es departamento ('Huancavelica', 'Callao') solo es distrito
        # si el texto dice 'distrito de X'; en listas ('Cajamarca, Huancavelica y Lambayeque')
        # son departamentos.
        if name in dept_norms and not explicit:
            return
        # Homonimos ('El Porvenir', 'San Isidro'): solo con el departamento mencionado.
        if entries_per_name[name] > 1 and norm_text(item["department"]) not in dept_mentions:
            return
        found[(norm_text(item["department"]), name)] = (
            item["department"],
            item.get("province", ""),
            item["district"].upper(),
            item["zone"],
        )

    # HTML/noticia: menciones explicitas "distrito de X".
    for item in catalog:
        d = re.escape(item["district_norm"])
        if re.search(rf"\bdistrito\s+(?:de\s+)?{d}\b", html_scope):
            add_item(item, explicit=True)

    # Listas: "distritos de Piura, Castilla y Catacaos". Cada elemento de la lista debe ser
    # EXACTAMENTE un distrito (no basta con que la palabra aparezca cerca: 'marco', 'quinua').
    for match in re.finditer(r"\bdistritos\s+(?:de\s+)?", html_scope):
        segment = html_scope[match.end(): match.end() + 320]
        segment = re.split(r"[.;:]", segment, maxsplit=1)[0]
        segment = re.split(r"\s+(?:para|ante|por|con|donde|que|luego|tras|durante|debido|en\s+el\s+marco)\b", segment, maxsplit=1)[0]
        parts = {part.strip() for part in re.split(r",|\sy\s|\se\s", segment)}
        for item in catalog:
            if item["district_norm"] in parts:
                add_item(item, explicit=False)

    # Lista introducida por 'los distritos afectados ... son A, B y C': por construccion cada
    # elemento es un distrito (incluso 'Arequipa', que tambien es departamento).
    for match in re.finditer(r"\bdistritos\s+(?:afectados|damnificados)\b[^.;:]{0,100}?\bson\s+", html_scope):
        segment = html_scope[match.end(): match.end() + 500]
        segment = re.split(r"[.;:]", segment, maxsplit=1)[0]
        padded = f", {segment.strip()},"
        for item in catalog:
            # Coincidencia como elemento completo de la lista (respeta nombres con 'y',
            # p. ej. 'Jose Luis Bustamante y Rivero').
            if re.search(rf"(?:,|\sy\s|\se\s)\s*{re.escape(item['district_norm'])}\s*(?:,|\sy\s|\se\s)", padded):
                add_item(item, explicit=True)

    if found:
        return list(found.values())

    # Fallback: un nombre suelto NO es evidencia ('marco', 'quinua', 'San Isidro' de la
    # firma de INDECI, 'La Punta'). Solo se acepta con su departamento al lado:
    # 'Lampa (Ayacucho)', 'Ayna, Ayacucho'.
    for item in catalog:
        if len(item["district_norm"]) < 5:
            continue
        dep = re.escape(norm_text(item["department"]))
        if re.search(rf"\b{re.escape(item['district_norm'])}\s*(?:\(|,)\s*{dep}\b", primary_txt):
            add_item(item, explicit=True)

    return list(found.values())


def infer_location(article: Article, catalog: list[dict]) -> tuple[str, str, str, str]:
    """Compatibilidad hacia atras: devuelve solo la primera ubicacion detectada."""
    locations = infer_locations(article, catalog)
    return locations[0] if locations else ("", "", "", "")


_NAME_LOWER_WORDS = {"de", "del", "en", "el", "la", "los", "las", "y", "por", "a", "al"}


def coen_event_phrase(title: str) -> str:
    """'Lluvias Intensas en el Distrito de San Marcos - Ancash' a partir del titulo COEN.

    Quita el encabezado (tipo de reporte, numero, hora) y el sufijo del portal; los
    reportes complementarios de un mismo evento quedan con el mismo nombre.
    """
    m = re.search(
        r"\bhoras\b\s*(?:\(\s*(?:reporte|informe)[^)]*\)|(?:reporte|informe)\s+n\s*[o°º.]*\s*\d+)?\s*(.+)$",
        re.sub(r"\s+", " ", title or ""),
        flags=re.I,
    )
    if not m:
        return ""
    phrase = re.sub(r"\s*-?\s*INDECI\s+Tarea de Todos.*$", "", m.group(1), flags=re.I).strip(" -")
    if len(phrase) < 10:
        return ""
    words = phrase.lower().split()
    return " ".join(w if (i and w in _NAME_LOWER_WORDS) else w.capitalize() for i, w in enumerate(words))


def coen_event_type(title: str) -> str:
    """'Lluvias intensas' a partir de 'Lluvias Intensas en el Distrito de X - Dpto' (sin el lugar)."""
    phrase = coen_event_phrase(title)
    kind = re.split(r"\s+(?:en\s+)?(?:el\s+|los\s+|la\s+)?(?:distritos?|departamentos?|provincias?)\b", phrase, flags=re.I)[0].strip()
    kind = " ".join(_ACCENT_RESTORE.get(w, w) for w in kind.lower().split())
    return (kind[:1].upper() + kind[1:]) if kind else ""


# Los titulos COEN llegan a veces sin tildes (slug de la URL) y a veces con ellas; se
# normalizan para que el mismo tipo de evento no aparezca escrito de dos maneras.
_ACCENT_RESTORE = {
    "inundacion": "inundación", "inundaciones": "inundaciones", "rio": "río", "rios": "ríos",
    "erosion": "erosión", "activacion": "activación", "saturacion": "saturación",
    "precipitacion": "precipitación", "precipitaciones": "precipitaciones",
    "quebrada": "quebrada", "aluvion": "aluvión", "deslizamientos": "deslizamientos",
}


def infer_event_name(article: Article) -> str:
    """Nombre corto del evento (tipo), como en la hoja original; el lugar va en sus columnas."""
    if article.institucion == "COEN - INDECI":
        kind = coen_event_type(article.title)
        if kind:
            return kind[:180]
    title = re.sub(r"\s+", " ", article.title).strip()
    return title[:180] if title else "Evento relacionado al Fenomeno El Nino"


def effective_event_date(article: Article):
    """Fecha del hecho si COEN la informa; de lo contrario fecha de publicacion."""
    value = article.event_date or article.published
    return value.date() if value else datetime.now().date()


DAMAGE_CATEGORIES = [
    ("viviendas", r"viviend"),
    ("vías y puentes", r"\bvias?\b|carretera|puente|camino"),
    ("agricultura", r"cultiv|agricol|hectare"),
    ("educación", r"educaci|colegio"),
    ("salud", r"establecimiento de salud|posta medica|centro de salud"),
    ("servicios básicos", r"agua potable|desague|electric|canal|riego|saneamiento|alcantarillado"),
]
LIFE_IMPACT_PATTERN = r"fallecid|desaparecid|\bherid|damnificad|personas? afectad|poblacion afectad"
NO_LIFE_IMPACT_PATTERN = r"no se reportan? da.os (?:a la vida|personales|a las personas)|sin da.os personales"


def coen_damage_scope(pdf_text: str) -> str:
    """Seccion '3. EVALUACION DE DANOS' del reporte COEN (hasta '4. ACCIONES')."""
    if not pdf_text:
        return ""
    normalized = "\n".join(norm_text(line) for line in pdf_text.splitlines())
    m = re.search(r"evaluacion de danos\s*:?(.*?)(?=\n\s*4\.?\s*acciones|\Z)", normalized, flags=re.S)
    return m.group(1) if m else ""


def damage_counts(scope: str) -> int:
    """Suma de cifras de las filas 'DIST. X ...' de la tabla de danos (0 si no hay)."""
    total = 0
    for m in re.finditer(r"\bdist\.?\s+[a-z .'\-]+?((?:\s+\d+)+)\s*$", scope, flags=re.M):
        total += sum(int(n) for n in m.group(1).split())
    return total


def infer_magnitude(article: Article) -> tuple[str, str]:
    """Magnitud descriptiva y Escala_Magnitud segun la hoja Metodologia_y_fuentes.

    Alta: personas damnificadas/afectadas/fallecidas y/o danos en 3+ tipos de infraestructura.
    Media: danos materiales relevantes (2+ tipos o cifras significativas).
    Baja: afectacion puntual/localizada o activacion sin danos amplios.
    """
    scope = coen_damage_scope(article.pdf_text) if article.pdf_text else ""
    if not scope:
        scope = norm_text(f"{article.title} {article.text[:3000]}")
    no_life = re.search(NO_LIFE_IMPACT_PATTERN, scope) is not None
    life = re.search(LIFE_IMPACT_PATTERN, scope) is not None and not no_life
    categories = [name for name, pattern in DAMAGE_CATEGORIES if re.search(pattern, scope)]
    total = damage_counts(scope)

    if life or len(categories) >= 3:
        level = "Alta"
    elif len(categories) >= 2 or total >= 10:
        level = "Media"
    else:
        level = "Baja"

    hazard = coen_event_type(article.title) if article.institucion == "COEN - INDECI" else ""
    hazard = hazard or "Evento hidrometeorológico"
    if life:
        impact = "afectación a personas reportada"
    elif categories:
        impact = "daños en " + ", ".join(categories)
    else:
        impact = "sin daños cuantificados"
    return f"{hazard}; {impact}", level


def article_key(article: Article, department: str, district: str) -> str:
    date_part = effective_event_date(article).isoformat()
    raw = "|".join(
        [
            norm_text(infer_event_name(article)),
            date_part,
            norm_text(department),
            norm_text(district),
        ]
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def row_key(event: str, date_value, department: str, district: str) -> str:
    if isinstance(date_value, datetime):
        date_part = date_value.date().isoformat()
    elif hasattr(date_value, "isoformat"):
        date_part = date_value.isoformat()
    else:
        date_part = str(date_value or "")[:10]

    raw = "|".join(
        [norm_text(event), date_part, norm_text(department), norm_text(district)]
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def ensure_candidates_sheet(wb):
    name = "Eventos_candidatos"
    if name in wb.sheetnames:
        return wb[name]

    ws = wb.create_sheet(name)
    headers = [
        "Evento",
        "Fecha",
        "Institucion",
        "Departamento",
        "Distrito",
        "Fuente oficial",
        "Motivo revision",
    ]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(1, col, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="7030A0")
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 70
    ws.column_dimensions["G"].width = 70
    return ws


def existing_state(events_ws, candidates_ws) -> tuple[set[str], set[str]]:
    urls: set[str] = set()
    keys: set[str] = set()

    eh = header_map(events_ws)
    for row in range(2, events_ws.max_row + 1):
        url = str(events_ws.cell(row, eh["Fuente oficial"]).value or "").strip()
        if url:
            urls.add(normalize_url(url))
        keys.add(
            row_key(
                str(events_ws.cell(row, eh["Evento"]).value or ""),
                events_ws.cell(row, eh["Fecha"]).value,
                str(events_ws.cell(row, eh["Departamento"]).value or ""),
                str(events_ws.cell(row, eh["Distrito afectado"]).value or ""),
            )
        )

    ch = header_map(candidates_ws)
    if "Fuente oficial" in ch:
        for row in range(2, candidates_ws.max_row + 1):
            url = str(candidates_ws.cell(row, ch["Fuente oficial"]).value or "").strip()
            if url:
                urls.add(normalize_url(url))

    return urls, keys


# -----------------------------------------------------------------------------
# NIVEL DE EVIDENCIA FEN
# -----------------------------------------------------------------------------

FEN_LEVEL_EXPLICIT = "1 - Causa explícita El Niño/ENFEN"
FEN_LEVEL_DECREE = "2 - Declaratoria FEN vigente"
FEN_LEVEL_ALERT = "3 - Alerta ENFEN vigente"
FEN_LEVEL_NONE = "4 - Sin vínculo FEN"
FEN_REAL_LEVELS = (FEN_LEVEL_EXPLICIT, FEN_LEVEL_DECREE)
FEN_LEVEL_COLUMN = "Nivel_evidencia_FEN"

# Una mencion a El Nino dentro de una nota de declaratoria ('D.S. N° 124-2026-PCM ... peligro
# inminente ... asociadas al Fenomeno El Nino') no atribuye el EVENTO al FEN: solo describe
# una declaratoria que puede ser posterior al hecho. Por eso se separa de la causa explicita.
DECREE_CONTEXT_PATTERN = r"pcm|decreto supremo|peligro inminente|estado de emergencia|declarad"
DECREE_ID_PATTERN = r"\b(\d{2,3}-\d{4}-pcm)\b"
DECREE_WINDOW_PATTERN = (
    r"(?:vigente\s+)?(?:desde|del)\s+(?:el\s+)?(\d{1,2})\s+(?:de\s+)?([a-z]+)(?:\s+(?:de\s+)?(\d{4}))?"
    r"\s+(?:hasta|al)\s+(?:(?:el|al)\s+)*(\d{1,2})\s+(?:de\s+)?([a-z]+)(?:\s+(?:de\s+)?(\d{4}))?"
)


def parse_decree_windows(window: str, fallback_year: int) -> list[tuple[str, "datetime.date", "datetime.date"]]:
    """(decreto, inicio, fin) de las vigencias que el texto indica junto a un decreto."""
    ids = re.findall(DECREE_ID_PATTERN, window)
    decree = ids[0].upper() if ids else "DS s/n"
    out = []
    for m in re.finditer(DECREE_WINDOW_PATTERN, window):
        d1, m1, y1, d2, m2, y2 = m.groups()
        if m1 not in SPANISH_MONTHS or m2 not in SPANISH_MONTHS:
            continue
        end_year = int(y2 or y1 or fallback_year)
        start_year = int(y1 or end_year)
        try:
            start = datetime(start_year, SPANISH_MONTHS[m1], int(d1)).date()
            end = datetime(end_year, SPANISH_MONTHS[m2], int(d2)).date()
            if start > end:
                start = datetime(start_year - 1, SPANISH_MONTHS[m1], int(d1)).date()
        except ValueError:
            continue
        out.append((decree, start, end))
    return out


def assess_fen_evidence(article: Article, event_day, alert_start) -> tuple[str, str]:
    """Nivel de evidencia de vinculo con el FEN y la nota que lo justifica.

    1 causa explicita (mencion fuera de una nota de declaratoria); 2 declaratoria FEN cuya
    vigencia cubre la fecha del evento; 3 evento dentro de la alerta ENFEN vigente;
    4 sin vinculo. Una declaratoria que no cubre la fecha del evento NO cuenta como nivel 2.
    """
    primary = norm_text(
        f"{article.title} {(article.html_text or article.text)[:1600]} "
        f"{article.pdf_text[:PDF_LEAD_CHARS] if article.pdf_text else ''}"
    )
    decree_windows = []
    for pattern in DIRECT_FEN_PATTERNS:
        for m in re.finditer(pattern, primary):
            window = primary[max(0, m.start() - 300): m.end() + 250]
            if not re.search(DECREE_CONTEXT_PATTERN, window):
                return FEN_LEVEL_EXPLICIT, primary[max(0, m.start() - 110): m.end() + 110].strip()
            decree_windows.append(window)

    note = ""
    for window in decree_windows:
        for decree, start, end in parse_decree_windows(window, event_day.year):
            if start <= event_day <= end:
                return FEN_LEVEL_DECREE, f"{decree}, vigente del {start:%d/%m/%Y} al {end:%d/%m/%Y}"
            note = f"{decree} (vigente del {start:%d/%m/%Y} al {end:%d/%m/%Y}) no cubre la fecha del evento"
        if not note:
            note = "menciona una declaratoria FEN sin vigencia verificable"
    if alert_start and event_day >= alert_start:
        return FEN_LEVEL_ALERT, note
    return FEN_LEVEL_NONE, note


def detect_enfen_alert_start(session: requests.Session):
    """Inicio de la alerta de El Nino Costero vigente, segun los comunicados de ENFEN.

    Se toman las fechas (de la URL) de las publicaciones ENFEN que hablan de 'alerta' y
    'Nino Costero' y se retrocede mientras no haya huecos de mas de 60 dias.
    """
    cfg = next((c for c in SOURCES if c["institucion"] == "ENFEN"), None)
    if cfg is None:
        return None
    today = datetime.now().date()
    try:
        urls = make_source_index_urls(cfg, "backfill", today, today, 60)
        links = discover_links_with_pagination(session, cfg, urls, 60)
    except Exception as exc:
        logging.warning("No se pudo determinar el inicio de la alerta ENFEN: %s", exc)
        return None
    days = set()
    for url in links:
        slug = norm_text(unquote(urlparse(url).path).replace("-", " "))
        when = parse_url_date(url)
        if when and re.search(r"alerta.*nino costero|nino costero.*alerta", slug):
            days.add(when.date())
    ordered = sorted(days, reverse=True)
    if not ordered:
        return None
    start = ordered[0]
    for day in ordered[1:]:
        if (start - day).days > 60:
            break
        start = day
    return start


def ensure_column(ws, name: str, width: int = 34) -> int:
    """Agrega una columna al final (y a la Tabla de Excel, si la hay) si todavia no existe."""
    headers = header_map(ws)
    if name in headers:
        return headers[name]
    col = max(headers.values()) + 1 if headers else 1
    cell = ws.cell(1, col, name)
    previous = ws.cell(1, col - 1) if col > 1 else None
    if previous is not None and previous.has_style:
        cell.font = previous.font.copy()
        cell.fill = previous.fill.copy()
        cell.alignment = previous.alignment.copy()
        cell.border = previous.border.copy()
    ws.column_dimensions[get_column_letter(col)].width = width
    for table in ws.tables.values():
        m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", table.ref)
        if m and column_index_from_string(m.group(3)) == col - 1:
            table.tableColumns.append(TableColumn(id=len(table.tableColumns) + 1, name=name))
            table.ref = f"{m.group(1)}{m.group(2)}:{get_column_letter(col)}{m.group(4)}"
            if table.autoFilter is not None:
                table.autoFilter.ref = table.ref
    return col


# -----------------------------------------------------------------------------
# DECRETOS DE ESTADO DE EMERGENCIA POR PELIGRO INMINENTE (FEN)
# -----------------------------------------------------------------------------

DEFAULT_DECREE_DIR = Path(__file__).resolve().parent / "data" / "decretos"
DEFAULT_DECREE_INDEX = DEFAULT_DECREE_DIR / "decretos_fen.json"

# Vigencia de cada decreto (plazo de 60 dias calendario). Los PDF del cuerpo de algunos decretos
# son imagenes y no se pueden leer; las fechas salen de la edicion de El Peruano / notas COEN y
# se pueden corregir editando decretos_fen.json (o esta tabla y regenerando con --build-decretos).
DECREE_SOURCES = [
    {
        "id": "097-2026-PCM",
        "pdf": "ds-097-2026-pcm-anexo.pdf",
        "inicio": "2026-07-02",
        "plazo_dias": 60,
        "nota": "Inicio segun notas COEN (02-03/07); verificar edicion de El Peruano.",
        "url": "https://www.gob.pe/institucion/pcm/normas-legales/8372996-097-2026-pcm",
    },
    {
        "id": "124-2026-PCM",
        "pdf": "8597538-1-ds-n-124-2026-pcm.pdf",
        "inicio": "2026-09-01",
        "plazo_dias": 60,
        "nota": "Edicion extraordinaria El Peruano del 01/09/2026; 60 dias calendario.",
        "url": "https://www.gob.pe/institucion/pcm/normas-legales/8597538-124-2026-pcm",
    },
]


def parse_decree_annex(pdf_path: Path, ubigeo_rows: list[dict]) -> tuple[list[dict], int]:
    """Distritos del ANEXO de un decreto: (lista de registros ubigeo, total numerado en el anexo).

    El anexo lista 'N DISTRITO' por departamento/provincia. Para cada departamento se buscan los
    distritos del catalogo INEI precedidos por su numero de orden ('16 SAN FRANCISCO DEL YESO'),
    de modo que un nombre de provincia igual a un distrito no cuente.
    """
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    full = "\n".join((page.extract_text() or "") for page in reader.pages)
    start = full.upper().find("DISTRITOS DECLARADOS EN ESTADO DE EMERGENCIA")
    annex = full[start:] if start >= 0 else full

    dept_names = {norm_text(canonical) for canonical in CANONICAL_DEPARTMENTS.values()}
    dept_names |= {"callao", "la libertad", "madre de dios", "san martin"}
    segments: dict[str, list[str]] = collections.defaultdict(list)
    current = None
    for line in annex.splitlines():
        norm_line = norm_text(line).strip()
        if norm_line in dept_names:
            current = norm_line
            continue
        if current:
            segments[current].append(norm_line)

    by_dept: dict[str, list[dict]] = collections.defaultdict(list)
    for rec in ubigeo_rows:
        by_dept[norm_text(rec["department"])].append(rec)

    found: dict[str, dict] = {}
    for dept, lines in segments.items():
        text = " ".join(lines)
        for rec in by_dept.get(dept, []):
            pattern = rf"(?<![a-z0-9])\d{{1,4}}\s+{re.escape(rec['district_norm'])}(?![a-z])"
            if re.search(pattern, text):
                found[rec["ubigeo"]] = rec
    # Numeros de orden del anexo (se excluyen anios como 2026/2027 del titulo).
    numbers = [int(n) for n in re.findall(r"(?<![0-9])(\d{1,4})\s+[A-Z]{3,}", annex) if int(n) < 1900]
    return list(found.values()), (max(numbers) if numbers else 0)


def build_decree_index(ubigeo_path: Path, decree_dir: Path, output: Path) -> None:
    """Lee los anexos descargados en data/decretos y escribe decretos_fen.json."""
    ubigeo_rows = load_ubigeo(ubigeo_path)
    if not ubigeo_rows:
        raise SystemExit("Hace falta el catalogo de ubigeos para construir el indice de decretos.")
    index = []
    for src in DECREE_SOURCES:
        pdf = decree_dir / src["pdf"]
        if not pdf.exists():
            logging.warning("Falta el PDF del decreto %s: %s", src["id"], pdf)
            continue
        districts, declared = parse_decree_annex(pdf, ubigeo_rows)
        start = datetime.strptime(src["inicio"], "%Y-%m-%d").date()
        end = start + timedelta(days=int(src["plazo_dias"]) - 1)
        index.append(
            {
                "id": src["id"],
                "inicio": start.isoformat(),
                "fin": end.isoformat(),
                "plazo_dias": src["plazo_dias"],
                "nota": src["nota"],
                "url": src["url"],
                "distritos_en_anexo": declared,
                "distritos_identificados": len(districts),
                "ubigeos": sorted(rec["ubigeo"] for rec in districts),
            }
        )
        logging.info(
            "%s: %s distritos en el anexo, %s identificados contra el catalogo INEI (vigencia %s a %s)",
            src["id"], declared, len(districts), start, end,
        )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    logging.info("Indice de decretos guardado: %s", output)


DECREE_INDEX: list[dict] = []


def load_decree_index(path: Path) -> list[dict]:
    """Carga decretos_fen.json (si existe) y deja los conjuntos de ubigeos listos para consultar."""
    DECREE_INDEX.clear()
    if not path.exists():
        logging.warning("Sin indice de decretos (%s): solo se usaran las notas de los reportes. "
                        "Generalo con --build-decretos.", path)
        return DECREE_INDEX
    for item in json.loads(path.read_text(encoding="utf-8")):
        DECREE_INDEX.append(
            {
                **item,
                "_inicio": datetime.strptime(item["inicio"], "%Y-%m-%d").date(),
                "_fin": datetime.strptime(item["fin"], "%Y-%m-%d").date(),
                "_ubigeos": set(item["ubigeos"]),
            }
        )
    logging.info("Decretos FEN cargados: %s", ", ".join(d["id"] for d in DECREE_INDEX) or "ninguno")
    return DECREE_INDEX


def decree_covering(ubigeo: str, event_day) -> Optional[dict]:
    """Decreto FEN cuyo anexo incluye el distrito y cuya vigencia cubre el dia del evento."""
    if not ubigeo:
        return None
    for item in DECREE_INDEX:
        if ubigeo in item["_ubigeos"] and item["_inicio"] <= event_day <= item["_fin"]:
            return item
    return None


def coen_province_for(pdf_text: str, district: str) -> str:
    """Provincia del distrito segun las tablas estructuradas del PDF COEN (o '' si no consta).

    1) Tabla de danos: bloques 'PROV. X' seguidos de filas 'DIST. Y <cifras>'.
    2) Tabla '2. UBICACION': fila 'DEPARTAMENTO PROVINCIA DISTRITO ... <ubigeo>'.
    """
    district_norm = norm_text(district)
    if not pdf_text or not district_norm:
        return ""
    lines = [norm_text(line) for line in pdf_text.splitlines()]

    province = ""
    for line in lines:
        m = re.match(r"^prov\.?\s+(.+)$", line)
        if m:
            province = m.group(1).strip()
            continue
        m = re.match(r"^dist\.?\s+(.+?)(?:\s+\d+)*$", line)
        if m and province and m.group(1).strip() == district_norm:
            return province.title()

    departments = [key for key, _canonical in _TITLE_DEPARTMENTS]
    for line in lines:
        m = re.match(r"^(.+?)\s+\d{6}$", line)
        if not m:
            continue
        row = m.group(1)
        dep = next((d for d in departments if row.startswith(d + " ")), None)
        if not dep:
            continue
        rest = row[len(dep):].strip()
        idx = rest.find(district_norm)
        if idx > 0 and rest[idx - 1] == " ":
            return rest[:idx].strip().title()
    return ""


def fen_evidence_snippet(article: Article, width: int = 110) -> str:
    """Fragmento del texto oficial donde aparece la referencia explicita a El Nino/ENFEN."""
    text = norm_text(f"{article.title} {article.pdf_text or ''} {article.html_text or article.text}")
    for pattern in DIRECT_FEN_PATTERNS:
        m = re.search(pattern, text)
        if m:
            return text[max(0, m.start() - width): m.end() + width].strip()
    return ""


def append_event_for_location(
    events_ws,
    article: Article,
    location: tuple[str, str, str, str],
) -> tuple[str, str]:
    """Agrega UNA fila a Eventos_reales para un distrito concreto."""
    headers = header_map(events_ws)
    dep, prov, district, zone = location
    if district and article.pdf_text:
        # La tabla del PDF describe ESTE evento (y desambigua nombres repetidos en el
        # departamento); el catalogo solo aporta la provincia cuando es unica.
        prov = coen_province_for(article.pdf_text, district) or prov
    ubigeo = ubigeo_code(dep, prov, district) if district else ""
    magnitude, severity = infer_magnitude(article)
    event = infer_event_name(article)
    date_value = effective_event_date(article)

    impact = magnitude.split("; ", 1)[1] if "; " in magnitude else magnitude
    place = ", ".join(x for x in (district.title() if district else "", dep) if x)
    detail = f"{article.institucion} reportó {magnitude.split('; ')[0].lower()}"
    detail += f" en {place}" if place else ""
    detail += f"; {impact}."
    if article.fen_level == FEN_LEVEL_DECREE:
        detail += f" Contexto FEN: declaratoria {article.fen_note}."
    elif article.fen_level == FEN_LEVEL_EXPLICIT and article.fen_note:
        detail += f" Atribución a El Niño/ENFEN en la fuente: «…{article.fen_note}…»."
    if article.pdf_url and article.pdf_text:
        detail += " Fuente: reporte oficial COEN (PDF)."

    row = next_data_row(events_ws)
    values = {
        "Evento": event,
        "Fecha": date_value,
        "Magnitud / severidad observada": magnitude,
        "Escala_Magnitud": severity,
        "Departamento": dep,
        "Provincia": prov,
        "Distrito afectado": district,
        "Zona": zone,
        "Fuente oficial": article.url,
        "Detalle": detail[:1000],
    }

    for header, value in values.items():
        events_ws.cell(row, headers[header], value)

    events_ws.cell(row, headers["Fecha"]).number_format = "dd/mm/yyyy"
    if FEN_LEVEL_COLUMN in headers:
        events_ws.cell(row, headers[FEN_LEVEL_COLUMN], article.fen_level)
    if UBIGEO_COLUMN in headers:
        cell = events_ws.cell(row, headers[UBIGEO_COLUMN], ubigeo)
        cell.number_format = "@"
    set_tables_last_row(events_ws, row, only_grow=True)
    return dep, district


def append_event(events_ws, article: Article, catalog: list[dict]) -> list[tuple[str, str]]:
    """Compatibilidad: agrega una fila por cada distrito detectado."""
    locations = infer_locations(article, catalog) or [("", "", "", "")]
    return [append_event_for_location(events_ws, article, loc) for loc in locations]


def append_candidate(
    candidates_ws,
    article: Article,
    department: str,
    district: str,
    reason: str,
    province: str = "",
) -> None:
    row = next_data_row(candidates_ws)
    if district and article.pdf_text:
        province = coen_province_for(article.pdf_text, district) or province
    ubigeo = ubigeo_code(department, province, district) if district else ""
    values = [
        infer_event_name(article),
        effective_event_date(article),
        article.institucion,
        department,
        district,
        article.url,
        reason,
    ]
    for col, value in enumerate(values, 1):
        candidates_ws.cell(row, col, value)
    if article.published:
        candidates_ws.cell(row, 2).number_format = "dd/mm/yyyy"
    ubigeo_col = header_map(candidates_ws).get(UBIGEO_COLUMN)
    if ubigeo_col:
        candidates_ws.cell(row, ubigeo_col, ubigeo).number_format = "@"
    level_col = header_map(candidates_ws).get(FEN_LEVEL_COLUMN)
    if level_col:
        candidates_ws.cell(row, level_col, article.fen_level if not article.fen_note else f"{article.fen_level}")
    set_tables_last_row(candidates_ws, row, only_grow=True)


def save_workbook_safely(wb, path: Path) -> None:
    temp_path = path.with_name(f"{path.stem}.__tmp__{path.suffix}")
    last_exc = None

    try:
        if temp_path.exists():
            temp_path.unlink()
    except OSError:
        pass

    try:
        wb.save(temp_path)
    except Exception as exc:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except OSError:
            pass
        raise OSError(f"No se pudo generar archivo temporal: {exc}") from exc

    for attempt in range(1, SAVE_RETRIES + 1):
        try:
            os.replace(temp_path, path)
            return
        except (PermissionError, OSError) as exc:
            last_exc = exc
            logging.warning(
                "No se pudo reemplazar el Excel. Intento %s/%s: %s",
                attempt,
                SAVE_RETRIES,
                exc,
            )
            time.sleep(SAVE_RETRY_SECONDS)

    try:
        if temp_path.exists():
            temp_path.unlink()
    except OSError:
        pass
    raise last_exc  # type: ignore[misc]


def write_pad_summary(log_dir: Path, payload: dict) -> Path:
    """Escribe un resumen JSON simple para que PAD pueda consumirlo si desea."""
    summary_path = log_dir / f"resumen_fen_{datetime.now():%Y%m%d_%H%M%S}.json"
    summary_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    return summary_path


def emit_pad_summary(payload: dict, summary_path: Optional[Path] = None) -> None:
    """Emite una linea capturable por consola cuando se usa python.exe."""
    if sys.stdout is None:
        return
    parts = [
        "PAD_SUMMARY",
        f"status={payload.get('status', 'UNKNOWN')}",
        f"eventos={payload.get('eventos_reales_agregados', 0)}",
        f"candidatos={payload.get('eventos_candidatos_agregados', 0)}",
        f"fen_no_evento={payload.get('fen_sin_evento', 0)}",
        f"duplicados={payload.get('duplicados', 0)}",
        f"errores_fetch={payload.get('articulos_no_recuperados', 0)}",
        f"dry_run={payload.get('dry_run', False)}",
    ]
    if summary_path is not None:
        parts.append(f"summary_json={summary_path}")
    print("|".join(parts), flush=True)


def mostrar_resultado(titulo: str, mensaje: str, es_error: bool = False) -> None:
    """Muestra un popup nativo de Windows sin requerir consola."""
    if os.name != "nt":
        return
    try:
        # MB_OK (0x0) + MB_ICONERROR (0x10) / MB_ICONINFORMATION (0x40)
        icono = 0x10 if es_error else 0x40
        ctypes.windll.user32.MessageBoxW(0, mensaje, titulo, icono)
    except Exception:
        logging.exception("No se pudo mostrar la ventana de resultado.")


def resumen_exito(payload: dict, log_path: Path, sin_novedades: bool = False) -> str:
    """Construye el resumen visible al usuario al finalizar."""
    if payload.get("mode") == "backfill":
        encabezado = (
            "Backfill historico finalizado correctamente.\n\n"
            if not sin_novedades
            else "Backfill historico finalizado correctamente.\n\nNo se encontraron eventos en el rango indicado.\n\n"
        )
    else:
        encabezado = (
            "La revision finalizo correctamente.\n\n"
            if sin_novedades
            else "Actualizacion del Fenomeno del Nino finalizada correctamente.\n\n"
        )
    if sin_novedades:
        encabezado += "No se encontraron nuevos eventos para incorporar.\n\n"
    return (
        encabezado
        + f"Filas nuevas en Eventos_reales: {payload.get('eventos_reales_agregados', 0)}\n"
        + f"Noticias/eventos con altas: {payload.get('articulos_evento_agregados', 0)}\n"
        + f"Candidatos para revision: {payload.get('eventos_candidatos_agregados', 0)}\n"
        + f"FEN sin evento observado: {payload.get('fen_sin_evento', 0)}\n"
        + f"Duplicados ignorados: {payload.get('duplicados', 0)}\n"
        + f"Articulos con error: {payload.get('articulos_no_recuperados', 0)}\n"
        + f"PDF COEN leidos: {payload.get('pdf_coen_leidos', 0)}\n"
        + f"Enlaces no vistos revisados: {payload.get('enlaces_no_vistos', 0)}\n\n"
        + f"Log:\n{log_path}"
    )


def resumen_dry_run(payload: dict, log_path: Path) -> str:
    rango = ""
    if payload.get("mode") == "backfill":
        rango = (
            f"Rango: {payload.get('from_date')} a {payload.get('to_date')}\n"
            "Modo: construccion historica desde cero (simulado).\n\n"
        )
    return (
        "DRY RUN finalizado correctamente.\n"
        "El Excel NO fue modificado.\n\n"
        + rango
        + f"Filas de Eventos_reales que se agregarian: {payload.get('eventos_reales_agregados', 0)}\n"
        f"Noticias/eventos con altas: {payload.get('articulos_evento_agregados', 0)}\n"
        f"Candidatos que se agregarian: {payload.get('eventos_candidatos_agregados', 0)}\n"
        f"Duplicados: {payload.get('duplicados', 0)}\n"
        f"Articulos con error: {payload.get('articulos_no_recuperados', 0)}\n"
        f"PDF COEN leidos: {payload.get('pdf_coen_leidos', 0)}\n\n"
        f"Log:\n{log_path}"
    )


# -----------------------------------------------------------------------------
# PROCESO PRINCIPAL
# -----------------------------------------------------------------------------



# -----------------------------------------------------------------------------
# BACKFILL / PAGINACION
# -----------------------------------------------------------------------------

def make_source_index_urls(cfg: dict, mode: str, start_date: datetime.date, end_date: datetime.date, max_pages: int) -> list[str]:
    """Genera indices suficientes para un backfill.

    En modo weekly conserva los indices configurados en SOURCES. En backfill
    genera paginas de forma conservadora para gob.pe/COEN y usa los indices
    configurados para ENFEN; la deteccion de paginas adicionales se
    hace tambien mediante enlaces de paginacion encontrados en cada indice.
    """
    if mode != "backfill":
        return list(cfg["index_urls"])

    base_urls = list(cfg["index_urls"])
    institution = cfg["institucion"]
    urls: list[str] = []

    # En backfill partimos de la primera pagina de cada fuente y seguimos la paginacion.
    # Esto evita generar decenas de URLs iniciales y luego explorarlas todas como si fueran semillas.
    if base_urls:
        urls.append(base_urls[0])

    # Para las fuentes paginadas conocidas, tambien agregamos paginas deterministas como respaldo.
    if institution in {"INDECI", "Contraloria General de la Republica"}:
        # sheet=2...N en gob.pe
        base = base_urls[0].split("?", 1)[0]
        for sheet in range(2, max_pages + 1):
            urls.append(f"{base}?sheet={sheet}")
    elif institution == "COEN - INDECI":
        base = base_urls[0].rstrip("/")
        for page in range(2, max_pages + 1):
            urls.append(f"{base}/page/{page}/")
    elif institution == "ENFEN":
        # WordPress suele soportar /page/N/ sobre los indices.
        for base in base_urls:
            base = base.rstrip("/")
            for page in range(2, max_pages + 1):
                urls.append(f"{base}/page/{page}/")
    # Deduplicar preservando orden.
    return list(dict.fromkeys(urls))


def discover_links_with_pagination(session: requests.Session, cfg: dict, index_urls: list[str], max_pages: int) -> set[str]:
    """Descubre articulos y sigue paginacion con un maximo TOTAL por fuente.

    Corta temprano cuando encuentra varias paginas inexistentes/vacias consecutivas.
    Esto evita recorrer page/4, page/5, ... cuando una fuente ya termino.
    """
    links: set[str] = set()
    visited_indexes: set[str] = set()
    queued_indexes: set[str] = set(index_urls)
    queue: list[str] = list(index_urls)
    page_limit = max(1, max_pages)
    consecutive_empty = 0
    max_consecutive_empty = 2

    while queue and len(visited_indexes) < page_limit:
        index_url = queue.pop(0)
        if index_url in visited_indexes:
            continue
        visited_indexes.add(index_url)

        try:
            logging.info("  Leyendo indice %s/%s: %s", len(visited_indexes), page_limit, index_url)
            html = request_html(session, index_url)
        except Exception as exc:
            logging.warning("No se pudo leer indice %s: %s", index_url, exc)
            consecutive_empty += 1
            if consecutive_empty >= max_consecutive_empty:
                logging.info("  Fin de paginacion tras %s paginas vacias/error consecutivas.", consecutive_empty)
                break
            continue

        if not html:
            consecutive_empty += 1
            if consecutive_empty >= max_consecutive_empty:
                logging.info("  Fin de paginacion tras %s paginas inexistentes/vacias consecutivas.", consecutive_empty)
                break
            continue

        consecutive_empty = 0
        soup = BeautifulSoup(html, "html.parser")
        new_links_before = len(links)
        for a in soup.find_all("a", href=True):
            absolute = normalize_url(urljoin(index_url, a["href"]))
            if allowed_article_url(absolute, cfg):
                links.add(absolute)
                continue

            label = norm_text(" ".join([
                a.get_text(" ", strip=True),
                str(a.get("aria-label", "")),
                str(a.get("title", "")),
            ]))
            path = urlparse(absolute).path.lower()
            query = dict(parse_qsl(urlparse(absolute).query))
            is_pagination = (
                re.search(r"/(?:page|pagina)/\d+/?$", path) is not None
                or "sheet" in query
                or "siguiente" in label
                or label in {"next", "older"}
            )
            if is_pagination and urlparse(absolute).netloc.lower() == cfg["allowed_host"]:
                if (
                    absolute not in visited_indexes
                    and absolute not in queued_indexes
                    and len(visited_indexes) + len(queue) < page_limit
                ):
                    queue.append(absolute)
                    queued_indexes.add(absolute)

        if len(links) == new_links_before:
            logging.info("  Indice sin nuevos articulos: %s", index_url)
        time.sleep(SLEEP_SECONDS)

    return links


NUMBERED_INDEX_SOURCES = {"INDECI", "Contraloria General de la Republica", "COEN - INDECI"}


def numbered_index_url(cfg: dict, page: int) -> str:
    """URL del indice numerado N (gob.pe: ?sheet=N, COEN: /page/N/)."""
    base = cfg["index_urls"][0]
    if cfg["institucion"] == "COEN - INDECI":
        base = base.rstrip("/")
        return base + "/" if page <= 1 else f"{base}/page/{page}/"
    base = base.split("?", 1)[0]
    return base if page <= 1 else f"{base}?sheet={page}"


def index_links(session: requests.Session, cfg: dict, index_url: str) -> Optional[set[str]]:
    """Enlaces de articulo de un indice. None si la pagina no existe/no se pudo leer."""
    try:
        html = request_html(session, index_url)
    except Exception as exc:
        logging.warning("No se pudo leer indice %s: %s", index_url, exc)
        return None
    if not html:
        return None
    soup = BeautifulSoup(html, "html.parser")
    links: set[str] = set()
    for a in soup.find_all("a", href=True):
        url = normalize_url(urljoin(index_url, a["href"]))
        if allowed_article_url(url, cfg):
            links.add(url)
    return links


def _index_links_threaded(cfg: dict, index_url: str) -> Optional[set[str]]:
    """Lee un indice numerado reintentando si llega vacio.

    gob.pe responde a veces con una pagina casi vacia (sin error HTTP) cuando se le
    consulta en rafaga; sin reintento esa pagina se perdia en silencio.
    """
    links = None
    for attempt in range(1, 4):
        links = index_links(get_thread_session(), cfg, index_url)
        if links:
            return links
        time.sleep(2 * attempt)
    with THROTTLE_LOCK:
        THROTTLE_STATS["indices_vacios"] += 1
    logging.warning("[INDICE-VACIO] %s | sigue vacio tras 3 intentos: %s", cfg["institucion"], index_url)
    return links


def find_backfill_last_page(
    session: requests.Session, cfg: dict, start_date, hard_max: int
) -> int:
    """Ultima pagina del indice que aun contiene publicaciones >= start_date.

    Los indices estan ordenados de mas nuevo a mas antiguo, asi que se busca por
    biseccion la primera pagina cuyo contenido ya es anterior a start_date. Las
    paginas fijas (barra lateral repetida en todas las paginas) se excluyen de la
    muestra para no sesgar la fecha.
    """
    sidebar = index_links(session, cfg, numbered_index_url(cfg, 1)) or set()
    cache: dict[int, Optional[datetime]] = {}

    def newest(page: int):
        """Fecha MAS RECIENTE de la pagina. None = vacia/inexistente; datetime.max = indeterminada.

        El indice COEN se ordena por ultima actualizacion, no por fecha del reporte: un evento
        viejo con un complementario nuevo sube a las primeras paginas. Por eso una pagina solo
        esta "fuera de rango" si TODOS sus elementos son anteriores (el maximo, no el minimo).
        """
        if page in cache:
            return cache[page]
        links = index_links(session, cfg, numbered_index_url(cfg, page))
        if links is not None and page > 1:
            links = links - sidebar
        if not links:
            cache[page] = None
            return None
        dates = [d for d in (parse_url_date(u) for u in links) if d is not None]
        if not dates:
            ordered = sorted(links)
            step = max(1, len(ordered) // 4)
            for url in ordered[::step][:4]:
                art = extract_article(session, cfg["institucion"], url)
                if art and art.published:
                    dates.append(art.published)
        cache[page] = max(dates) if dates else datetime.max
        return cache[page]

    def past_range(page: int) -> bool:
        d = newest(page)
        return d is None or d.date() < start_date

    lo, hi = 1, 1
    if not past_range(1):
        lo = 1
        hi = 2
        while hi < hard_max and not past_range(hi):
            lo, hi = hi, min(hi * 2, hard_max)
        while hi - lo > 1:
            mid = (lo + hi) // 2
            if past_range(mid):
                hi = mid
            else:
                lo = mid
    # Margen: el orden por actualizacion deja eventos del rango algo mas abajo.
    margin = 5 if cfg["institucion"] == "COEN - INDECI" else 1
    last = min(hi + margin, hard_max)
    logging.info(
        "%s: el rango desde %s llega hasta la pagina ~%s (tope %s)",
        cfg["institucion"], start_date, last, hard_max,
    )
    return last


def discover_numbered_index_links(cfg: dict, last_page: int, workers: int) -> set[str]:
    """Lee las paginas 1..last_page en paralelo (solo I/O; sin tocar el workbook)."""
    urls = [numbered_index_url(cfg, n) for n in range(1, last_page + 1)]
    links: set[str] = set()
    done = 0
    with ThreadPoolExecutor(max_workers=max(1, min(workers, 8))) as pool:
        futures = {pool.submit(_index_links_threaded, cfg, u): u for u in urls}
        for fut in as_completed(futures):
            res = fut.result()
            done += 1
            if res:
                links |= res
            if done % 50 == 0 or done == len(urls):
                logging.info("  %s: indices leidos %s/%s (%s enlaces)", cfg["institucion"], done, len(urls), len(links))
    return links


def writable_output_path(path: Path) -> Path:
    """Si el Excel de salida esta abierto (Excel lo bloquea), usa un nombre con hora en vez de fallar."""
    if not path.exists():
        return path
    try:
        with open(path, "ab"):
            return path
    except PermissionError:
        alternative = path.with_name(f"{path.stem}_{datetime.now():%H%M%S}{path.suffix}")
        logging.warning(
            "El Excel de salida esta abierto en otro programa (%s); se guardara como %s. "
            "Cierra el archivo anterior para reutilizar el nombre.",
            path.name, alternative.name,
        )
        return alternative


def clear_sheet_data_preserve_format(ws) -> None:
    """Elimina las filas de datos conservando encabezado, anchos, congelado y la Tabla de Excel.

    Solo borrar los valores dejaba ws.max_row intacto y los eventos nuevos se escribian
    debajo de las filas vacias, fuera de la Tabla.
    """
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)
    set_tables_last_row(ws, 2)


def next_data_row(ws) -> int:
    """Primera fila libre tras el ultimo dato real (ignora filas vacias con solo formato)."""
    for r in range(ws.max_row, 1, -1):
        if any(c.value not in (None, "") for c in ws[r]):
            return r + 1
    return 2


def set_tables_last_row(ws, last_row: int, only_grow: bool = False) -> None:
    """Ajusta el rango de las Tablas de Excel de la hoja para que llegue a last_row."""
    for table in ws.tables.values():
        m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", table.ref)
        if not m:
            continue
        end = max(last_row, int(m.group(2)) + 1)
        if only_grow and end <= int(m.group(4)):
            continue
        table.ref = f"{m.group(1)}{m.group(2)}:{m.group(3)}{end}"
        if table.autoFilter is not None:
            table.autoFilter.ref = table.ref


def prepare_backfill_workbook(input_path: Path, output_path: Path) -> None:
    """Copia el workbook y deja Eventos_reales/candidatos vacios, conservando estructura y formato."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(input_path, output_path)
    wb = load_workbook(output_path)
    validate_workbook(wb)

    clear_sheet_data_preserve_format(wb["Eventos_reales"])

    if "Eventos_candidatos" in wb.sheetnames:
        clear_sheet_data_preserve_format(wb["Eventos_candidatos"])

    wb.save(output_path)
    logging.info("Workbook historico preparado (tablas vacias): %s", output_path)


def parse_iso_date_arg(value: str, arg_name: str) -> datetime.date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"{arg_name} debe tener formato YYYY-MM-DD: {value}") from exc

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Actualiza Eventos_reales desde fuentes oficiales."
    )
    parser.add_argument(
        "workbook",
        nargs="?",
        default=str(DEFAULT_WORKBOOK),
        help="Ruta del XLSX. Si se omite usa la ruta sincronizada de OneDrive.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Procesa todo pero no guarda cambios.",
    )
    parser.add_argument(
        "--lookback-days",
        type=int,
        default=14,
        help="Solo considera publicaciones de los ultimos N dias en modo weekly. Default: 14.",
    )
    parser.add_argument(
        "--from-date",
        help="Fecha inicial inclusiva para backfill, formato YYYY-MM-DD.",
    )
    parser.add_argument(
        "--to-date",
        help="Fecha final inclusiva para backfill, formato YYYY-MM-DD.",
    )
    parser.add_argument(
        "--backfill-year",
        type=int,
        help="Atajo para backfill de todo un anio calendario, por ejemplo 2026.",
    )
    parser.add_argument(
        "--backfill-max-pages",
        type=int,
        default=1500,
        help=(
            "Tope de seguridad de paginas por fuente en backfill. INDECI/Contraloria/COEN se "
            "recorren solo hasta la pagina donde termina el rango de fechas; ENFEN usa este tope. "
            "Default: 1500."
        ),
    )
    parser.add_argument(
        "--threads",
        type=int,
        default=None,
        help="Numero de workers HTTP. Default: 4 weekly / 8 backfill. Recomendado maximo: 12.",
    )
    parser.add_argument(
        "--output",
        help="Ruta del XLSX de salida. Opcional en backfill; por defecto (y si es ruta relativa) se crea en la carpeta del script.",
    )
    parser.add_argument(
        "--ubigeo",
        default=str(DEFAULT_UBIGEO),
        help="CSV del catalogo INEI de distritos (inei, departamento, provincia, distrito). Default: data/ubigeo_distrito.csv.",
    )
    parser.add_argument(
        "--decretos",
        default=str(DEFAULT_DECREE_INDEX),
        help="JSON con los decretos FEN (distritos del anexo y vigencia). Se genera con --build-decretos.",
    )
    parser.add_argument(
        "--build-decretos",
        action="store_true",
        help="Lee los PDF de data/decretos y genera el indice de decretos FEN; no ejecuta el scraping.",
    )
    parser.add_argument(
        "--max-new",
        type=int,
        default=None,
        help=(
            "Maximo de filas que se guardan por ejecucion. Se revisan todas las fuentes. "
            "Default: 100 en weekly; sin tope en backfill."
        ),
    )
    parser.add_argument(
        "--log-dir",
        default=str(DEFAULT_LOG_DIR),
        help="Carpeta local para logs.",
    )
    parser.add_argument(
        "--no-popup",
        action="store_true",
        help="No muestra ventana final. Util para ejecuciones totalmente silenciosas.",
    )
    args = parser.parse_args()
    if args.max_new is None:
        # Un backfill con tope de 100 filas se trunca en silencio (paso en la corrida de junio-octubre).
        args.max_new = 10**9 if (args.backfill_year or args.from_date or args.to_date) else 100

    # Determina modo operativo. Backfill requiere rango exacto y genera un workbook NUEVO.
    backfill = bool(args.backfill_year or args.from_date or args.to_date)
    mode = "backfill" if backfill else "weekly"

    if args.backfill_year:
        if args.from_date or args.to_date:
            raise SystemExit("No combines --backfill-year con --from-date/--to-date.")
        from_date = datetime(args.backfill_year, 1, 1).date()
        to_date = datetime(args.backfill_year, 12, 31).date()
    elif args.from_date or args.to_date:
        if not (args.from_date and args.to_date):
            raise SystemExit("En backfill debes indicar ambos: --from-date y --to-date.")
        from_date = parse_iso_date_arg(args.from_date, "--from-date")
        to_date = parse_iso_date_arg(args.to_date, "--to-date")
        if from_date > to_date:
            raise SystemExit("--from-date no puede ser posterior a --to-date.")
    else:
        from_date = None
        to_date = None

    log_path = setup_logging(Path(args.log_dir))
    if args.build_decretos:
        build_decree_index(Path(args.ubigeo), DEFAULT_DECREE_DIR, Path(args.decretos))
        return 0
    workbook_path = Path(args.workbook)

    output_path = None
    if backfill:
        if args.dry_run:
            logging.info("BACKFILL en dry-run: se analizara sin crear/guardar workbook.")
        # El backfill se guarda en el workspace (carpeta del script), no junto al Excel de OneDrive/SharePoint.
        script_dir = Path(__file__).resolve().parent
        if args.output:
            output_path = Path(args.output)
            if not output_path.is_absolute():
                output_path = script_dir / output_path
        else:
            output_path = script_dir / (
                f"{workbook_path.stem}_backfill_{from_date:%Y%m%d}_{to_date:%Y%m%d}{workbook_path.suffix}"
            )

    logging.info("=== INICIO ACTUALIZACION FEN ===")
    if backfill:
        logging.info("Backfill: workbook de salida independiente; no se modifica el Excel oficial de entrada.")
    logging.info("Modo: %s", mode)
    logging.info("Excel entrada: %s", workbook_path)
    if output_path:
        logging.info("Excel salida: %s", output_path)
    if backfill:
        logging.info("Rango backfill: %s a %s", from_date, to_date)
    else:
        logging.info("Lookback: %s dias", args.lookback_days)
    logging.info("Dry-run: %s", args.dry_run)
    process_started = time.perf_counter()

    if PdfReader is None:
        logging.error("Falta la dependencia pypdf. Ejecuta: python -m pip install pypdf")
        if not args.no_popup:
            mostrar_resultado(
                "Actualizacion FEN - Error",
                "Falta instalar la dependencia pypdf.\n\nEjecuta:\npython -m pip install pypdf\n\n"
                f"Log:\n{log_path}",
                es_error=True,
            )
        return 1

    if not workbook_path.exists():
        try:
            create_empty_workbook(workbook_path)
        except Exception as exc:
            logging.exception("No existe el Excel y no se pudo crear: %s", exc)
            logging.info("Log: %s", log_path)
            if not args.no_popup:
                mostrar_resultado(
                    "Actualizacion FEN - Error",
                    "No se encontro el archivo Excel y no se pudo crear.\n\n"
                    f"{workbook_path}\n\nDetalle: {exc}\n\nLog:\n{log_path}",
                    es_error=True,
                )
            return 2

    if backfill and output_path is not None and not args.dry_run:
        try:
            output_path = writable_output_path(output_path)
            if output_path.resolve() != workbook_path.resolve():
                prepare_backfill_workbook(workbook_path, output_path)
            else:
                raise ValueError("En backfill --output debe ser distinto del workbook de entrada.")
        except Exception as exc:
            logging.exception("No se pudo preparar el workbook de backfill: %s", exc)
            if not args.no_popup:
                mostrar_resultado(
                    "Actualizacion FEN - Error",
                    f"No se pudo preparar el workbook historico.\n\nDetalle: {exc}\n\nLog:\n{log_path}",
                    es_error=True,
                )
            return 3

    active_workbook_path = output_path if backfill and output_path is not None and not args.dry_run else workbook_path

    try:
        wb = open_workbook_with_retry(active_workbook_path)
    except Exception as exc:
        logging.exception("No se pudo abrir el Excel: %s", exc)
        logging.info("Log: %s", log_path)
        if not args.no_popup:
            mostrar_resultado(
                "Actualizacion FEN - Error",
                "No se pudo abrir el archivo Excel.\n\n"
                f"Detalle: {exc}\n\n"
                "Verifica que Excel no este bloqueando el archivo y que OneDrive este disponible.\n\n"
                f"Log:\n{log_path}",
                es_error=True,
            )
        return 2

    try:
        validate_workbook(wb)
        if backfill and output_path is not None and not args.dry_run:
            workbook_path = active_workbook_path
        events_ws = wb["Eventos_reales"]
        candidates_ws = ensure_candidates_sheet(wb)
        ensure_column(events_ws, FEN_LEVEL_COLUMN)
        ensure_column(candidates_ws, FEN_LEVEL_COLUMN)
        ensure_column(events_ws, UBIGEO_COLUMN, width=10)
        ensure_column(candidates_ws, UBIGEO_COLUMN, width=10)
        catalog = load_location_catalog(wb, load_ubigeo(Path(args.ubigeo)))
        load_decree_index(Path(args.decretos))
    except Exception as exc:
        logging.exception("Estructura Excel invalida: %s", exc)
        logging.info("Log: %s", log_path)
        if not args.no_popup:
            mostrar_resultado(
                "Actualizacion FEN - Error",
                "La estructura del Excel no es valida.\n\n"
                f"Detalle: {exc}\n\n"
                f"Log:\n{log_path}",
                es_error=True,
            )
        return 3

    if backfill:
        # En backfill se construye una tabla nueva desde cero. Incluso en dry-run
        # simulamos que no hay URLs/filas existentes para no omitir el historico.
        existing_urls, existing_keys = set(), set()
    else:
        existing_urls, existing_keys = existing_state(events_ws, candidates_ws)

    session = _new_http_session()
    enfen_alert_start = detect_enfen_alert_start(session)
    logging.info("Alerta ENFEN vigente desde: %s", enfen_alert_start or "no determinada")
    fen_level_counts: dict[str, collections.Counter] = {"reales": collections.Counter(), "candidatos": collections.Counter()}
    default_workers = MAX_WORKERS_BACKFILL if backfill else MAX_WORKERS_WEEKLY
    workers = args.threads if args.threads is not None else default_workers
    workers = max(1, min(int(workers), 12))
    logging.info("Threads HTTP: %s", workers)

    discovered: list[tuple[dict, str]] = []
    source_stats: dict[str, int] = {}
    prefiltered_by_url = 0
    prefiltered_by_type = 0
    skipped_by_source: dict[str, collections.Counter] = {
        "antiguedad": collections.Counter(),
        "sin_fecha": collections.Counter(),
        "irrelevantes": collections.Counter(),
    }

    for cfg in SOURCES:
        if backfill and cfg["institucion"] in NUMBERED_INDEX_SOURCES:
            last_page = find_backfill_last_page(
                session, cfg, from_date, max(1, args.backfill_max_pages)
            )
            links = discover_numbered_index_links(cfg, last_page, workers)
        elif backfill:
            index_urls = make_source_index_urls(
                cfg, mode, from_date, to_date, min(max(1, args.backfill_max_pages), 60)
            )
            links = discover_links_with_pagination(
                session, cfg, index_urls, min(max(1, args.backfill_max_pages), 60)
            )
        else:
            links = discover_links(session, cfg)

        source_stats[cfg["institucion"]] = len(links)
        logging.info(
            "%s: %s enlaces descubiertos en indices (%s)",
            cfg["institucion"],
            len(links),
            mode,
        )
        for url in sorted(links):
            if url in existing_urls:
                continue
            # COEN/ENFEN llevan la fecha en la URL: evita descargar HTML+PDF de reportes fuera del rango.
            # Un evento no puede ocurrir despues de su reporte, por lo que ambos extremos son seguros.
            if backfill and cfg["institucion"] in {"COEN - INDECI", "ENFEN"}:
                url_date = parse_url_date(url)
                if url_date and not (from_date <= url_date.date() <= to_date):
                    prefiltered_by_url += 1
                    skipped_by_source["antiguedad"][cfg["institucion"]] += 1
                    continue
            # El slug COEN ya dice el tipo de emergencia: incendios, sismos, heladas, boletines y
            # avisos (sin numero de reporte) no son eventos hidrometeorologicos; no se descargan.
            if cfg["institucion"] == "COEN - INDECI" and not coen_url_worth_fetching(url):
                prefiltered_by_type += 1
                skipped_by_source["irrelevantes"][cfg["institucion"]] += 1
                continue
            discovered.append((cfg, url))

    logging.info("Enlaces no vistos previamente: %s", len(discovered))

    if backfill:
        start_filter = from_date
        end_filter = to_date
    else:
        start_filter = datetime.now().date() - timedelta(days=max(args.lookback_days, 0))
        end_filter = datetime.now().date()
    inserted = 0  # filas agregadas a Eventos_reales
    event_articles = 0  # publicaciones que generaron al menos una fila real
    candidates = 0  # filas agregadas a Eventos_candidatos
    candidate_articles = 0
    fen_non_event = 0
    skipped_old = prefiltered_by_url
    skipped_no_date = 0
    skipped_non_event = 0
    skipped_irrelevant = prefiltered_by_type
    skipped_duplicate = 0
    skipped_cap = 0
    fetch_failures = 0
    pdf_links_found = 0
    pdfs_read = 0
    pdfs_without_text = 0

    # Descarga/parseo HTML+PDF en paralelo; toda escritura en Excel queda en el hilo principal.
    extraction_started = time.perf_counter()
    extracted_results: list[tuple[dict, str, Optional[Article]]] = []
    total_discovered = len(discovered)
    if total_discovered:
        logging.info("Procesando %s articulos con %s threads...", total_discovered, workers)
        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="fen-http") as executor:
            future_map = {
                executor.submit(extract_article_threaded, cfg["institucion"], url): (cfg, url)
                for cfg, url in discovered
            }
            completed = 0
            for future in as_completed(future_map):
                cfg, url = future_map[future]
                completed += 1
                try:
                    article = future.result()
                except Exception as exc:
                    logging.warning("[ERROR-THREAD] %s | %s | %s", cfg["institucion"], url, exc)
                    article = None
                extracted_results.append((cfg, url, article))
                if completed == 1 or completed % 25 == 0 or completed == total_discovered:
                    logging.info("  Progreso articulos: %s/%s", completed, total_discovered)

    logging.info("Extraccion paralela finalizada en %.1f s", time.perf_counter() - extraction_started)

    # Clasificacion, deduplicacion y escritura: secuencial para proteger openpyxl.
    for cfg, url, article in extracted_results:
        if not article:
            fetch_failures += 1
            continue

        if article.pdf_url:
            pdf_links_found += 1
            if article.pdf_text:
                pdfs_read += 1
            elif article.institucion == "COEN - INDECI" and article.hazard:
                pdfs_without_text += 1

        # En produccion nunca insertamos una publicacion sin fecha verificable.
        # Evita backfills accidentales de articulos antiguos encontrados en indices.
        if not article.published:
            skipped_no_date += 1
            skipped_by_source["sin_fecha"][article.institucion] += 1
            if article.direct_fen or article.hazard:
                logging.info("[SIN-FECHA] %s | %s", article.institucion, article.title[:140])
            continue

        article_date = article.event_date.date() if (article.event_date and cfg["institucion"] == "COEN - INDECI") else article.published.date()

        if backfill:
            if article_date < start_filter or article_date > end_filter:
                skipped_old += 1
                skipped_by_source["antiguedad"][article.institucion] += 1
                continue
        else:
            if article_date < start_filter:
                skipped_old += 1
                skipped_by_source["antiguedad"][article.institucion] += 1
                continue

        if not article.direct_fen and not article.hazard:
            skipped_irrelevant += 1
            skipped_by_source["irrelevantes"][article.institucion] += 1
            continue

        locations = infer_locations(article, catalog)
        if not locations:
            # Se conserva una fila sin distrito cuando la noticia cumple todos los
            # criterios de evento pero la ubicacion no puede mapearse al maestro.
            locations = [("", "", "", "")]

        # La URL sigue siendo un atajo de deduplicacion entre ejecuciones. Dentro de
        # una misma noticia, la deduplicacion real se hace por evento+fecha+distrito.
        if normalize_url(article.url) in existing_urls:
            skipped_duplicate += 1
            continue

        # Preparacion, coordinacion, simulacros, pronosticos, etc. se excluyen
        # aunque mencionen FEN y peligros en el cuerpo de la nota.
        if article.non_event:
            skipped_non_event += 1
            if article.direct_fen_primary:
                fen_non_event += 1
                label = "FEN-NO-EVENTO"
            else:
                label = "NO-EVENTO"
            logging.info("[%s] %s | %s", label, article.institucion, article.title[:140])
            continue

        # Nivel de evidencia FEN: causa explicita o declaratoria vigente => evento real.
        # Una declaratoria que no cubre la fecha del evento, o solo la alerta ENFEN, deja
        # el hecho como candidato.
        article.fen_level, article.fen_note = assess_fen_evidence(
            article, effective_event_date(article), enfen_alert_start
        )

        # Evento (real o candidato) por distrito: el vinculo con el FEN puede variar entre los
        # distritos de una misma publicacion (p. ej. solo algunos figuran en el anexo del decreto).
        handled = False
        if article.hazard_primary and article.event_occurred_primary:
            event_day = effective_event_date(article)
            article_level, article_note = article.fen_level, article.fen_note
            context_ok = None
            added_real = added_cand = 0
            for dep, prov, district, zone in locations:
                level, note = article_level, article_note
                if level not in FEN_REAL_LEVELS and district:
                    decree = decree_covering(location_ubigeo(article, dep, prov, district), event_day)
                    if decree:
                        level = FEN_LEVEL_DECREE
                        note = (
                            f"{decree['id']}, distrito incluido en el anexo oficial "
                            f"(vigente del {decree['_inicio']:%d/%m/%Y} al {decree['_fin']:%d/%m/%Y})"
                        )
                article.fen_level, article.fen_note = level, note
                is_real = level in FEN_REAL_LEVELS
                if not is_real:
                    if context_ok is None:
                        context_ok = candidate_event_context_ok(article.text, article.title)
                    if not context_ok:
                        continue
                handled = True

                key = article_key(article, dep, district)
                if key in existing_keys:
                    skipped_duplicate += 1
                    continue
                if inserted + candidates >= args.max_new:
                    skipped_cap += 1
                    logging.warning(
                        "[CAP-%s] %s | %s | distrito=%s",
                        "EVENTO" if is_real else "CANDIDATO",
                        article.institucion,
                        article.title[:140],
                        district or "SIN DISTRITO",
                    )
                    continue

                if is_real:
                    append_event_for_location(events_ws, article, (dep, prov, district, zone))
                    inserted += 1
                    added_real += 1
                    fen_level_counts["reales"][level] += 1
                    logging.info(
                        "[EVENTO] %s | %s | distrito=%s | zona=%s | nivel=%s",
                        article.institucion,
                        article.title[:120],
                        district or "SIN DISTRITO",
                        zone or "SIN ZONA",
                        level[:1],
                    )
                else:
                    append_candidate(
                        candidates_ws,
                        article,
                        dep,
                        district,
                        "Evento hidrometeorologico observado en fuente oficial, pero sin atribucion explicita al Fenomeno El Nino. Revisar antes de incorporarlo a Eventos_reales."
                        + (f" Nota: {note}." if note else ""),
                        prov,
                    )
                    candidates += 1
                    added_cand += 1
                    fen_level_counts["candidatos"][level] += 1
                    logging.info(
                        "[CANDIDATO] %s | %s | distrito=%s",
                        article.institucion,
                        article.title[:120],
                        district or "SIN DISTRITO",
                    )
                # Misma key para no repetir el evento+fecha+distrito dentro de esta ejecucion.
                existing_keys.add(key)

            if added_real:
                event_articles += 1
            if added_cand:
                candidate_articles += 1
            if added_real or added_cand:
                existing_urls.add(normalize_url(article.url))

        if handled:
            pass

        elif not handled and article.direct_fen_primary:
            # Estado/pronostico/preparacion relacionado al FEN: se registra solo
            # en el log, NO en Eventos_candidatos, porque no es un evento observado.
            fen_non_event += 1
            logging.info("[FEN-NO-EVENTO] %s | %s", article.institucion, article.title[:140])

        elif article.direct_fen:
            # FEN aparece solo como contexto secundario en el cuerpo. No se considera
            # una noticia FEN para efectos de la tabla.
            skipped_irrelevant += 1
            logging.info("[FEN-CONTEXTO] %s | %s", article.institucion, article.title[:140])

        else:
            skipped_irrelevant += 1


    logging.info("--- RESUMEN ---")
    logging.info("Filas agregadas a Eventos_reales: %s", inserted)
    logging.info("Publicaciones que generaron Eventos_reales: %s", event_articles)
    logging.info("Filas agregadas a Eventos_candidatos: %s", candidates)
    logging.info("Publicaciones que generaron candidatos: %s", candidate_articles)
    logging.info("FEN sin evento observado (solo log): %s", fen_non_event)
    logging.info("Omitidos por antiguedad: %s", skipped_old)
    logging.info("Omitidos por fecha no identificada: %s", skipped_no_date)
    logging.info("Omitidos por preparacion/institucional: %s", skipped_non_event)
    logging.info("Omitidos por irrelevantes: %s", skipped_irrelevant)
    logging.info("Omitidos por limite --max-new: %s", skipped_cap)
    logging.info("Omitidos por duplicados: %s", skipped_duplicate)
    if THROTTLE_STATS["reintentos"] or THROTTLE_STATS["no_recuperadas"] or THROTTLE_STATS["indices_vacios"]:
        logging.warning(
            "gob.pe limito las consultas: %s reintentos, %s noticias NO recuperadas, %s indices vacios. "
            "Esas noticias quedaron sin fecha y NO se evaluaron: repite la corrida con menos --threads.",
            THROTTLE_STATS["reintentos"], THROTTLE_STATS["no_recuperadas"], THROTTLE_STATS["indices_vacios"],
        )
    logging.info("Articulos no recuperados: %s", fetch_failures)
    logging.info("PDF COEN encontrados: %s", pdf_links_found)
    logging.info("PDF COEN leidos con texto: %s", pdfs_read)
    logging.info("PDF COEN sin texto / no legibles: %s", pdfs_without_text)
    elapsed_seconds = time.perf_counter() - process_started
    logging.info("Threads utilizados: %s", workers)
    logging.info("Tiempo total: %.1f s", elapsed_seconds)

    summary_payload = {
        "status": "OK",
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "workbook": str(workbook_path),
        "mode": mode,
        "dry_run": bool(args.dry_run),
        "lookback_days": args.lookback_days,
        "from_date": str(from_date) if from_date else None,
        "to_date": str(to_date) if to_date else None,
        "output_workbook": str(output_path) if output_path else None,
        "enlaces_no_vistos": len(discovered),
        "eventos_reales_agregados": inserted,
        "articulos_evento_agregados": event_articles,
        "eventos_candidatos_agregados": candidates,
        "articulos_candidato_agregados": candidate_articles,
        "fen_sin_evento": fen_non_event,
        "omitidos_antiguedad": skipped_old,
        "omitidos_sin_fecha": skipped_no_date,
        "omitidos_preparacion_institucional": skipped_non_event,
        "omitidos_irrelevantes": skipped_irrelevant,
        "omitidos_limite": skipped_cap,
        "duplicados": skipped_duplicate,
        "articulos_no_recuperados": fetch_failures,
        "pdf_coen_encontrados": pdf_links_found,
        "pdf_coen_leidos": pdfs_read,
        "pdf_coen_sin_texto": pdfs_without_text,
        "threads": workers,
        "tiempo_total_segundos": round(elapsed_seconds, 1),
        "fuentes": source_stats,
        "omitidos_por_fuente": {k: dict(v) for k, v in skipped_by_source.items()},
        "enfen_alerta_desde": str(enfen_alert_start) if enfen_alert_start else None,
        "nivel_evidencia_fen": {k: dict(v) for k, v in fen_level_counts.items()},
        "omitidos_antiguedad_por_url": prefiltered_by_url,
        "omitidos_irrelevantes_por_url": prefiltered_by_type,
        "gobpe_limitacion": dict(THROTTLE_STATS),
        "log": str(log_path),
    }

    try:
        summary_path = write_pad_summary(Path(args.log_dir), summary_payload)
    except Exception:
        summary_path = None
        logging.exception("No se pudo generar resumen JSON para PAD.")

    if args.dry_run:
        logging.info("DRY RUN: no se guardaron cambios.")
        if summary_path:
            logging.info("Resumen PAD JSON: %s", summary_path)
        logging.info("Log: %s", log_path)
        logging.info("=== FIN OK ===")
        emit_pad_summary(summary_payload, summary_path)
        if not args.no_popup:
            mostrar_resultado(
                "Actualizacion FEN - Dry Run",
                resumen_dry_run(summary_payload, log_path),
            )
        return 0

    # Si no hubo novedades, no tocamos el archivo; evita generar versiones inutiles
    # en SharePoint/OneDrive.
    if inserted == 0 and candidates == 0:
        logging.info("Sin novedades. El Excel no fue modificado.")
        if summary_path:
            logging.info("Resumen PAD JSON: %s", summary_path)
        logging.info("Log: %s", log_path)
        logging.info("=== FIN OK ===")
        emit_pad_summary(summary_payload, summary_path)
        if not args.no_popup:
            mostrar_resultado(
                "Actualizacion FEN - Sin novedades",
                resumen_exito(summary_payload, log_path, sin_novedades=True),
            )
        return 0

    try:
        save_workbook_safely(wb, workbook_path)
    except Exception as exc:
        logging.exception("Error al guardar el Excel: %s", exc)
        logging.info("Log: %s", log_path)
        if not args.no_popup:
            mostrar_resultado(
                "Actualizacion FEN - Error",
                "No se pudo guardar el Excel.\n\n"
                f"Detalle: {exc}\n\n"
                "Verifica que Excel no este abierto/bloqueado y que OneDrive este disponible.\n\n"
                f"Log:\n{log_path}",
                es_error=True,
            )
        return 4

    logging.info("Excel actualizado correctamente: %s", workbook_path)
    logging.info("OneDrive sincronizara el cambio con SharePoint.")
    if summary_path:
        logging.info("Resumen PAD JSON: %s", summary_path)
    logging.info("Log: %s", log_path)
    logging.info("=== FIN OK ===")
    emit_pad_summary(summary_payload, summary_path)
    if not args.no_popup:
        mostrar_resultado(
            "Actualizacion FEN - Exito",
            resumen_exito(summary_payload, log_path),
        )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        if sys.stderr is not None:
            print("Proceso cancelado por el usuario.", file=sys.stderr)
        mostrar_resultado(
            "Actualizacion FEN - Cancelada",
            "El proceso fue cancelado por el usuario.",
            es_error=True,
        )
        raise SystemExit(1)
    except Exception as exc:
        logging.exception("Error general no controlado: %s", exc)
        mostrar_resultado(
            "Actualizacion FEN - Error",
            f"Ocurrio un error general no controlado.\n\nDetalle: {exc}",
            es_error=True,
        )
        raise SystemExit(1)
