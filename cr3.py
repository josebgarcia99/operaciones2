"""Extractor de comprobantes de transferencia para el CEP por lotes de Banxico.

Lee PDFs de comprobantes (con texto o escaneados/imagen) y arma el archivo .TXT
que pide Banxico: fecha,clave_rastreo,clave_emisora,clave_receptora,cuenta,monto
"""

import difflib
import hashlib
import io
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import unicodedata
import zipfile
from datetime import date

import pandas as pd
import sio_tema
import streamlit as st

try:
    import pymupdf as fitz
except Exception:  # pragma: no cover
    try:
        import fitz
    except Exception:
        fitz = None

try:
    import numpy as np
except Exception:  # pragma: no cover
    np = None

try:
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None

try:
    import pdfplumber
except Exception:  # pragma: no cover
    pdfplumber = None

try:
    import pytesseract
except Exception:  # pragma: no cover
    pytesseract = None

try:
    from rapidocr_onnxruntime import RapidOCR
except Exception:  # pragma: no cover
    RapidOCR = None

# Consulta del estado del pago en el portal del CEP (Playwright).
try:
    try:
        _AQUI = os.path.dirname(os.path.abspath(__file__))
    except NameError:  # si el módulo se ejecuta sin __file__
        _AQUI = os.getcwd()
    if _AQUI not in sys.path:
        sys.path.insert(0, _AQUI)
    import cep_banxico
except Exception as _exc:  # pragma: no cover
    cep_banxico = None
    CEP_IMPORT_ERROR = str(_exc)
else:
    CEP_IMPORT_ERROR = ""

try:
    from excel_conv import (leer_excel_convenia, leer_base_tarjetas, normalizar_nombre,
                            a_numero, escribir_acumulado_xlsx, escribir_operaciones_xlsx,
                            escribir_carga_masiva)
except Exception as _exc:  # pragma: no cover
    leer_excel_convenia = None
    EXCEL_IMPORT_ERROR = str(_exc)
else:
    EXCEL_IMPORT_ERROR = ""


# --------------------------------------------------------------------------- #
# Catálogo de instituciones (claves oficiales Banxico)
# --------------------------------------------------------------------------- #

VALID_BANKS = {
    "40133": "ACTINVER",
    "40062": "AFIRME",
    "90721": "albo",
    "90706": "ARCUS FI",
    "90659": "ASP INTEGRA OPC",
    "40127": "AZTECA",
    "37166": "BaBien",
    "40030": "BAJIO",
    "40002": "BANAMEX",
    "40154": "BANCO COVALTO",
    "37006": "BANCOMEXT",
    "40137": "BANCOPPEL",
    "40160": "BANCO S3",
    "40152": "BANCREA",
    "37019": "BANJERCITO",
    "40147": "BANKAOOL",
    "40106": "BANK OF AMERICA",
    "40159": "BANK OF CHINA",
    "37009": "BANOBRAS",
    "40072": "BANORTE",
    "40058": "BANREGIO",
    "40060": "BANSI",
    "2001": "BANXICO",
    "40129": "BARCLAYS",
    "40145": "BBASE",
    "40012": "BBVA MEXICO",
    "40112": "BMONEX",
    "90677": "CAJA POP MEXICA",
    "90683": "CAJA TELEFONIST",
    "90715": "CASHI CUENTA",
    "40124": "CITI MEXICO",
    "90730": "Clip",
    "90901": "CLS",
    "90903": "CoDi Valida",
    "40130": "COMPARTAMOS",
    "40140": "CONSUBANCO",
    "90725": "COOPDESARROLLO",
    "90652": "CREDICAPITAL",
    "90688": "CREDICLUB",
    "90680": "CRISTOBAL COLON",
    "90723": "Cuenca",
    "90729": "Dep y Pag Dig",
    "40151": "DONDE",
    "90616": "FINAMEX",
    "90634": "FINCOMUN",
    "90734": "FINCO PAY",
    "90738": "FINTOC",
    "90699": "FONDEADORA",
    "90685": "FONDO (FIRA)",
    "90601": "GBM",
    "40167": "HEY BANCO",
    "37168": "HIPOTECARIA FED",
    "40021": "HSBC",
    "40155": "ICBC",
    "40036": "INBURSA",
    "90902": "INDEVAL",
    "40150": "INMOBILIARIO",
    "40136": "INTERCAM BANCO",
    "40059": "INVEX",
    "40110": "JP MORGAN",
    "40128": "KAPITAL",
    "90661": "KLAR",
    "90653": "KUSPIT",
    "90670": "LIBERTAD",
    "90602": "MASARI",
    "90722": "Mercado Pago W",
    "90720": "MexPago",
    "40042": "MIFEL",
    "40158": "MIZUHO BANK",
    "90600": "MONEXCB",
    "40108": "MUFG",
    "40132": "MULTIVA BANCO",
    "37135": "NAFIN",
    "40638": "NUBANK",
    "90710": "NVIO",
    "40148": "PAGATODO",
    "90732": "Peibo",
    "90714": "PPBALANCEMX",
    "90620": "PROFUTURO",
    "40156": "SABADELL",
    "40014": "SANTANDER",
    "40044": "SCOTIABANK",
    "40157": "SHINHAN",
    "90728": "SPIN BY OXXO",
    "90646": "STP",
    "90703": "TESORED",
    "90684": "TRANSFER",
    "90727": "TRANSFER DIRECT",
    "90631": "TRF",
    "40138": "UALA",
    "90656": "UNAGRA",
    "90617": "VALMEX",
    "90605": "VALUE",
    "40113": "VE POR MAS",
    "40141": "VOLKSWAGEN",
}

# Nombres comerciales que aparecen en los comprobantes -> nombre del catálogo.
BANK_ALIASES = {
    "BBVA": "BBVA MEXICO",
    "BBVA BANCOMER": "BBVA MEXICO",
    "BANCOMER": "BBVA MEXICO",
    "BBVA NET CASH": "BBVA MEXICO",
    "CITIBANAMEX": "BANAMEX",
    "BANCO NACIONAL DE MEXICO": "BANAMEX",
    "CITI BANAMEX": "BANAMEX",
    "BANCO SANTANDER": "SANTANDER",
    "SANTANDER MEXICO": "SANTANDER",
    "BANCO AZTECA": "AZTECA",
    "BANCO DEL BAJIO": "BAJIO",
    "BANBAJIO": "BAJIO",
    "BANCA AFIRME": "AFIRME",
    "BANCO REGIONAL": "BANREGIO",
    "BANCO REGIONAL DE MONTERREY": "BANREGIO",
    "BANCO INBURSA": "INBURSA",
    "BANCO MIFEL": "MIFEL",
    "BANCO INVEX": "INVEX",
    "BANCO MULTIVA": "MULTIVA BANCO",
    "BANCO MONEX": "BMONEX",
    "MONEX": "BMONEX",
    "BANCO VE POR MAS": "VE POR MAS",
    "BX+": "VE POR MAS",
    "BANCO INTERCAM": "INTERCAM BANCO",
    "INTERCAM": "INTERCAM BANCO",
    "BANCO ACTINVER": "ACTINVER",
    "SCOTIABANK INVERLAT": "SCOTIABANK",
    "HSBC MEXICO": "HSBC",
    "NU MEXICO": "NUBANK",
    "NU": "NUBANK",
    "HEY": "HEY BANCO",
    "MERCADO PAGO": "Mercado Pago W",
    "SPIN": "SPIN BY OXXO",
    "SPIN BY OXXO": "SPIN BY OXXO",
    "STP": "STP",
    "SISTEMA DE TRANSFERENCIAS Y PAGOS": "STP",
    "BANCO COMPARTAMOS": "COMPARTAMOS",
    "BANCOPPEL": "BANCOPPEL",
    "BANCO BANORTE": "BANORTE",
    "BANORTE IXE": "BANORTE",
}

# Marcas / dominios que identifican al banco EMISOR por el membrete de la hoja.
ISSUER_MARKERS = [
    # SuperLínea es la banca telefónica de SANTANDER, no de Scotiabank.
    ("superlinea", "40014"),
    ("enlacenegocios", "40072"),  # "Enlace Negocios" es el producto empresarial de BANORTE
    ("bbvanetcash", "40012"),
    ("bbvabancomer", "40012"),
    ("bbvamexico", "40012"),
    ("bbva", "40012"),
    ("citibanamex", "40002"),
    ("bancanet", "40002"),
    ("banamex", "40002"),
    ("supernetsantander", "40014"),
    ("santander", "40014"),
    ("banorte", "40072"),
    ("hsbcnet", "40021"),
    ("hsbc", "40021"),
    ("scotiabank", "40044"),
    ("scotiaenlinea", "40044"),
    ("inbursa", "40036"),
    ("banregio", "40058"),
    ("afirme", "40062"),
    ("bancodelbajio", "40030"),
    ("banbajio", "40030"),
    ("mifel", "40042"),
    ("actinver", "40133"),
    ("intercam", "40136"),
    ("multiva", "40132"),
    ("invex", "40059"),
    ("vepormas", "40113"),
    ("monex", "40112"),
    ("bancoazteca", "40127"),
    ("bancoppel", "40137"),
    ("compartamos", "40130"),
    ("stpmex", "90646"),
    ("nubank", "40638"),
    ("numexico", "40638"),
    ("klar", "90661"),
    ("heybanco", "40167"),
    ("mercadopago", "90722"),
    ("bancoconsubanco", "40140"),
    ("consubanco", "40140"),
    ("bansi", "40060"),
    ("bancamifel", "40042"),
    # Membretes y dominios que sólo puede traer la hoja del banco que emite:
    ("scotiabankinverlat", "40044"),
    ("invernet", "40044"),          # see.sbi.com.mx/invernet2000, portal de Scotiabank
    ("sbicommx", "40044"),
    ("bancanetempresarial", "40002"),
    ("santandercommx", "40014"),
    ("banortecommx", "40072"),
    ("hsbccommx", "40021"),
    ("bbvacommx", "40012"),
    ("bbvanetcashmx", "40012"),
    # Razones sociales del membrete: Banorte encabeza su reporte SPEI como
    # "BANCO MERCANTIL DEL NORTE S.A.", donde no aparece la palabra "Banorte".
    # Van como pista normal, no fuerte, porque algunos comprobantes también
    # imprimen la razón social del banco DESTINO.
    ("bancomercantildelnorte", "40072"),
    ("bancanacionaldemexico", "40002"),
    ("bancosantandermexico", "40014"),
    ("hsbcmexico", "40021"),
    ("bancoinbursa", "40036"),
    ("bancoregionaldemonterrey", "40058"),
]

# Pistas que valen más que el nombre pelón del banco.
#
# "BANORTE" aparece tanto en el membrete de Banorte como en el renglón "Banco:
# BANORTE" del beneficiario de CUALQUIER otro banco; en cambio "Scotia en Línea"
# o "BBVA Net Cash" sólo los imprime quien emite la hoja. Por eso estas mandan
# aunque aparezcan hasta el final del documento: el comprobante de Scotiabank
# trae el membrete al final del texto extraído y el BANORTE del beneficiario
# antes, así que ganar "por posición" daba emisor = BANORTE.
ISSUER_MARKERS_FUERTES = frozenset({
    "superlinea",
    "enlacenegocios",
    "bbvanetcash",
    "bbvanetcashmx",
    "bbvacommx",
    "bancanet",
    "bancanetempresarial",
    "supernetsantander",
    "santandercommx",
    "hsbcnet",
    "hsbccommx",
    "scotiaenlinea",
    "scotiabankinverlat",
    "invernet",
    "sbicommx",
    "banortecommx",
    "stpmex",
})


def _build_clabe_prefix_map():
    """CLABE: los 3 primeros dígitos identifican al banco (012 -> BBVA MEXICO)."""
    counts = {}
    for code in VALID_BANKS:
        counts[code[-3:]] = counts.get(code[-3:], 0) + 1
    return {code[-3:]: code for code in VALID_BANKS if counts[code[-3:]] == 1}


CLABE_PREFIX_TO_CODE = _build_clabe_prefix_map()


# --------------------------------------------------------------------------- #
# Utilidades de texto
# --------------------------------------------------------------------------- #

def _fold_char(ch: str):
    """Normaliza un caracter a ASCII minúsculo conservando el mapeo 1 a 1."""
    decomposed = unicodedata.normalize("NFKD", ch)
    decomposed = "".join(c for c in decomposed if not unicodedata.combining(c))
    if not decomposed:
        return None
    folded = decomposed[0].lower()
    # El OCR confunde la "o" final de las etiquetas con un cero ("rastre0").
    return "o" if folded == "0" else folded


# Puntuación que los bancos meten dentro de las etiquetas y que no debe estorbar
# al comparar: "CLABE, Plastico o Celular:" tiene que casar con
# "clabeplasticoocelular", y "Cuenta/ CLABE Ordenante" con "cuentaclabeordenante".
_PUNTUACION = set(",.;:()[]{}/\\-_–—*|\"'¿?¡!&#°º+")


def compact_with_map(text: str):
    """Devuelve el texto sin espacios/acentos/puntuación y el índice original."""
    chars, positions = [], []
    for i, ch in enumerate(text):
        if ch.isspace() or ch in _PUNTUACION:
            continue
        folded = _fold_char(ch)
        if folded is None:
            continue
        chars.append(folded)
        positions.append(i)
    return "".join(chars), positions


def compact(text: str) -> str:
    return compact_with_map(text)[0]


def label_value(line: str, variants) -> str | None:
    """Si la línea contiene alguna etiqueta, regresa lo que viene después de ella."""
    comp, positions = compact_with_map(line)
    best = None
    for variant in variants:
        pos = comp.find(variant)
        if pos < 0:
            continue
        end = pos + len(variant)
        start = positions[end] if end < len(positions) else len(line)
        value = line[start:].lstrip()
        value = value.lstrip(":;-–—.").strip()
        if best is None or pos < best[0]:
            best = (pos, value)
    return best[1] if best else None


def clean_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "")).strip()


# --------------------------------------------------------------------------- #
# Lectura de PDF (texto nativo + OCR)
# --------------------------------------------------------------------------- #

MIN_CHARS_PER_PAGE = 40  # menos que esto = página sin texto real -> OCR
OCR_ZOOM = 3.0
OCR_DPI = 216  # equivalente al zoom, para el render sin PyMuPDF


def log_ocr_error(message: str) -> None:
    try:
        st.session_state.setdefault("ocr_errors", []).append(message)
    except Exception:
        print(f"[OCR] {message}")


@st.cache_resource(show_spinner=False)
def get_rapidocr():
    if RapidOCR is None:
        return None
    try:
        return RapidOCR()
    except Exception:
        return None


def configure_tesseract_path() -> bool:
    if pytesseract is None:
        return False
    candidates = [
        shutil.which("tesseract"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            pytesseract.pytesseract.tesseract_cmd = candidate
            return True
    return False


TESSERACT_READY = configure_tesseract_path()


def _render_page(page):
    """Renderiza una página del PDF como imagen PIL."""
    if Image is None:
        return None
    pix = page.get_pixmap(matrix=fitz.Matrix(OCR_ZOOM, OCR_ZOOM), alpha=False)
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)


def _ocr_result_to_lines(result) -> list[str]:
    """Agrupa las cajas del OCR por renglón para no perder la relación etiqueta/valor."""
    items = []
    for entry in result or []:
        if not entry or len(entry) < 2:
            continue
        box, text = entry[0], entry[1]
        if not text:
            continue
        try:
            ys = [float(point[1]) for point in box]
            xs = [float(point[0]) for point in box]
        except Exception:
            continue
        items.append({
            "y": sum(ys) / len(ys),
            "x": min(xs),
            "h": max(ys) - min(ys),
            "text": str(text).strip(),
        })

    items.sort(key=lambda item: (item["y"], item["x"]))

    rows = []
    for item in items:
        if rows:
            row = rows[-1]
            tolerance = max(item["h"], row["h"], 8.0) * 0.6
            if abs(item["y"] - row["y"]) <= tolerance:
                row["items"].append(item)
                row["h"] = max(row["h"], item["h"])
                continue
        rows.append({"y": item["y"], "h": item["h"], "items": [item]})

    lines = []
    for row in rows:
        segments = [seg["text"] for seg in sorted(row["items"], key=lambda seg: seg["x"])]
        lines.extend(segments)
        if len(segments) > 1:
            # El renglón completo ayuda cuando la etiqueta y el valor quedaron separados.
            lines.append("   ".join(segments))
    return lines


def _blank_pages(raw_bytes: bytes) -> tuple[list[str], list[int]]:
    """Cuántas páginas tiene el PDF cuando ningún motor devolvió texto."""
    total = 0
    if pdfplumber is not None:
        try:
            with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
                total = len(pdf.pages)
        except Exception:
            total = 0
    if not total and PdfReader is not None:
        try:
            total = len(PdfReader(io.BytesIO(raw_bytes)).pages)
        except Exception:
            total = 0
    return [""] * total, list(range(total))


def _render_pages_plumber(raw_bytes: bytes, indexes) -> dict:
    """Render sin PyMuPDF: pdfplumber (pypdfium2) para los PDFs escaneados."""
    if pdfplumber is None:
        return {}
    images = {}
    try:
        with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
            for index in indexes:
                if index >= len(pdf.pages):
                    continue
                try:
                    render = pdf.pages[index].to_image(resolution=OCR_DPI)
                    images[index] = render.original.convert("RGB")
                except Exception as exc:  # noqa: BLE001
                    log_ocr_error(f"Render página {index + 1}: {exc}")
    except Exception as exc:  # noqa: BLE001
        log_ocr_error(f"Render PDF: {exc}")
    return images


def _ocr_image(image) -> str:
    if image is None:
        return ""

    ocr = get_rapidocr()
    if ocr is None and RapidOCR is not None:
        log_ocr_error("RapidOCR está instalado pero no se pudo inicializar el motor.")
    if ocr is not None and np is not None:
        try:
            result, _ = ocr(np.array(image))
            lines = _ocr_result_to_lines(result)
            if lines:
                return "\n".join(lines)
        except Exception as exc:  # noqa: BLE001
            log_ocr_error(f"RapidOCR: {exc}")

    if pytesseract is not None and TESSERACT_READY:
        try:
            return pytesseract.image_to_string(image, lang="spa+eng", config="--psm 6")
        except Exception:
            try:
                return pytesseract.image_to_string(image, config="--psm 6")
            except Exception as exc:  # noqa: BLE001
                log_ocr_error(f"Tesseract: {exc}")

    if get_rapidocr() is None and not (pytesseract is not None and TESSERACT_READY):
        log_ocr_error("No hay ningún motor de OCR disponible para leer las páginas escaneadas.")
    return ""


def _native_pages(raw_bytes: bytes) -> list[str]:
    """Texto nativo por página, probando varios motores."""
    pages = []
    if fitz is not None:
        try:
            with fitz.open(stream=raw_bytes, filetype="pdf") as doc:
                pages = [page.get_text("text") or "" for page in doc]
        except Exception:
            pages = []

    if not any(len(p.strip()) >= MIN_CHARS_PER_PAGE for p in pages) and PdfReader is not None:
        try:
            reader = PdfReader(io.BytesIO(raw_bytes))
            alt = [(page.extract_text() or "") for page in reader.pages]
            if sum(len(p.strip()) for p in alt) > sum(len(p.strip()) for p in pages):
                pages = alt
        except Exception:
            pass

    if not any(len(p.strip()) >= MIN_CHARS_PER_PAGE for p in pages) and pdfplumber is not None:
        try:
            with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
                alt = [(page.extract_text() or "") for page in pdf.pages]
            if sum(len(p.strip()) for p in alt) > sum(len(p.strip()) for p in pages):
                pages = alt
        except Exception:
            pass

    return pages


def extract_pdf_text(raw_bytes: bytes) -> tuple[str, str]:
    """Regresa (texto, origen) donde origen es 'texto', 'ocr' o 'mixto'."""
    pages = _native_pages(raw_bytes)
    used_native = any(len(p.strip()) >= MIN_CHARS_PER_PAGE for p in pages)

    needs_ocr = [i for i, p in enumerate(pages) if len(p.strip()) < MIN_CHARS_PER_PAGE]
    used_ocr = False

    if (needs_ocr or not pages) and fitz is not None:
        try:
            with fitz.open(stream=raw_bytes, filetype="pdf") as doc:
                if not pages:
                    pages = [""] * len(doc)
                    needs_ocr = list(range(len(doc)))
                for index in needs_ocr:
                    if index >= len(doc):
                        continue
                    text = _ocr_image(_render_page(doc[index]))
                    if text.strip():
                        pages[index] = text
                        used_ocr = True
        except Exception as exc:  # noqa: BLE001
            log_ocr_error(f"Render PDF: {exc}")
    elif needs_ocr or not pages:
        # Sin PyMuPDF: el render se hace con pdfplumber/pypdfium2.
        if not pages:
            pages, needs_ocr = _blank_pages(raw_bytes)
        for index, image in _render_pages_plumber(raw_bytes, needs_ocr).items():
            text = _ocr_image(image)
            if text.strip():
                pages[index] = text
                used_ocr = True

    if used_native and used_ocr:
        origin = "mixto"
    elif used_ocr:
        origin = "ocr"
    elif used_native:
        origin = "texto"
    else:
        origin = "sin texto"

    return "\n".join(pages), origin


# --------------------------------------------------------------------------- #
# Parseo de campos
# --------------------------------------------------------------------------- #

DATE_RE = re.compile(r"(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})")

DATE_LABELS_LAST = ("fecha",)          # Scotiabank y UnalanaPAY solo ponen "Fecha"
TRACKING_LABELS_LAST = ("referencia",)  # Scotiabank llama "Referencia" a la clave

DATE_LABELS_PRIMARY = (
    "fechadeaplicacion", "fechaaplicacion", "fechadeoperacion", "fechaoperacion",
)
DATE_LABELS_SECONDARY = (
    "fechadecreacion",
    "fechacreacion",
    "fechadelatransferencia",
    "fechadelpago",
    "fechadeenvio",
    "fechadeabono",
    "fechadecaptura",
    "fechayhoradelaoperacion",
)

TRACKING_LABELS = (
    "clavederastreo",
    "clavederastreospei",
    "claverastreo",
    "clavedeseguimiento",
    "claverastreospei",
    "trackingkey",
)

RECEIVER_ACCOUNT_LABELS = (
    "cuentaclabebeneficiario", "cuentaclabecelular", "clabereceptor",
    "clabeplasticoocelular", "clabeplastico",
    "alacuentadestino", "cuentadeabono", "cuentaabono", "cuentadedeposito",
    "cuentadeposito",
    "cuentabeneficiaria",
    "cuentadelbeneficiario",
    "cuentadestino",
    "cuentaabono",
    "cuentadeabono",
    "clabebeneficiario",
    "clabedestino",
    "clabeinterbancaria",
    "cuentareceptora",
)

SENDER_ACCOUNT_LABELS = (
    "cuentaclabeordenante", "delacuentaorigen", "cuentaorigen",
    "cuentadecargo", "cuentacargo", "cuentaderetiro",
    "cuentaretiro",
    "cuentacargo",
    "cuentadecargo",
    "cuentaordenante",
    "cuentaorigen",
)

AMOUNT_LABELS = (
    "importeatransferir", "montoatransferir", "importemxn", "importe",
    "montodelpago", "montodelatransferencia", "monto", "importetotal",
)

RECEIVER_BANK_LABELS = (
    "bancobeneficiario",
    "bancodelbeneficiario",
    "bancoreceptor",
    "bancodestino",
    "institucionreceptora",
    "institucionfinancierareceptora",
    "institucionbeneficiaria",
    "bancodeabono",
)

ISSUER_BANK_LABELS = (
    "institucionemisora",
    "institucionfinancieraemisora",
    "bancoemisor",
    "bancoordenante",
    "institucionordenante",
    "bancoorigen",
)

# Líneas que hablan del beneficiario: no sirven para detectar al banco emisor.
BENEFICIARY_LINE_MARKERS = (
    "beneficiario",
    "institucionreceptora",
    "titulardelacuenta",
    "cuentadedeposito",
    "cuentadeposito",
    "cuentadestino",
    "datonoverificado",
    "bancoreceptor",
    # Sin estas, un "Banco Destino: BBVA MEXICO" se tomaba como banco EMISOR:
    # el comprobante de Banamex salía reportado como BBVA.
    "bancodestino",
    "bancobeneficiario",
    "alacuentadestino",
    "cuentaabono",
    "cuentadeabono",
    "clabereceptor",
    "anombrede",
)


def split_lines(raw_text: str) -> list[str]:
    return [line for line in (l.strip() for l in raw_text.splitlines()) if line]


def find_by_labels(lines, labels, value_re, lookahead=1, group=0):
    """Busca etiqueta -> valor en la misma línea o en la(s) siguiente(s)."""
    for i, line in enumerate(lines):
        value = label_value(line, labels)
        if value is None:
            continue
        match = value_re.search(value)
        if match:
            return match.group(group)
        for nxt in lines[i + 1 : i + 1 + lookahead]:
            match = value_re.search(nxt)
            if match and len(nxt) <= 60:
                return match.group(group)
    return ""


def _es_solo_etiqueta(linea: str) -> bool:
    """¿El renglón es una etiqueta sin valor? ('Nombre del Ordenante', 'Importe:')"""
    limpio = clean_spaces(linea).rstrip(":").strip()
    if not limpio or len(limpio) > 60:
        return False
    comp = compact(limpio)
    return any(comp.startswith(etiqueta) for etiqueta in ETIQUETAS_TODAS)


def valor_etiquetado(lines, labels, validador=None, salto: int = 3, excluir=()) -> str:
    """Valor de una etiqueta, en el mismo renglón o en los siguientes.

    Los comprobantes se reparten en dos layouts. BBVA y Banamex escriben
    'Etiqueta: valor' en el mismo renglón; Banorte, PEIBO, Scotiabank y Monex
    ponen la etiqueta sola y el valor debajo. Esta función cubre ambos, y se
    detiene si el renglón de abajo resulta ser otra etiqueta.
    """
    for i, line in enumerate(lines):
        comp = compact(line)
        # `excluir` evita que una etiqueta se cuele dentro de otra más larga:
        # "razonsocial" hace match dentro de "nombrebeneficiariorazonsocial", y
        # sin esto el beneficiario terminaba registrado como ordenante.
        if any(marca in comp for marca in excluir):
            continue
        crudo = label_value(line, labels)
        if crudo is None:
            continue
        valor = clean_spaces(crudo)
        if valor and (validador is None or validador(valor)):
            return valor
        for siguiente in lines[i + 1: i + 1 + salto]:
            candidato = clean_spaces(siguiente)
            if not candidato:
                continue
            if _es_solo_etiqueta(candidato):
                break
            if validador is None or validador(candidato):
                return candidato
    return ""


MESES_ES = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12,
}
# "24/Jul./2026", "30-Jul-2026", "12 de Junio del 2026"
DATE_TEXTO_RE = re.compile(
    r"(\d{1,2})\s*(?:de\s+)?[/\-. ]\s*([A-Za-zÁÉÍÓÚÑáéíóúñ]{3,12})\.?\s*(?:de[l]?\s+)?[/\-. ]\s*(\d{2,4})"
)


def _fecha_con_mes_en_letra(value: str) -> str:
    match = DATE_TEXTO_RE.search(clean_spaces(value))
    if not match:
        return ""
    dia, mes_txt, anio = match.groups()
    clave = unicodedata.normalize("NFKD", mes_txt.lower())[:3]
    clave = "".join(c for c in clave if not unicodedata.combining(c))
    mes = MESES_ES.get(clave)
    if not mes:
        return ""
    try:
        dia_i, anio_i = int(dia), int(anio)
    except ValueError:
        return ""
    if anio_i < 100:
        anio_i += 2000
    if not (1 <= dia_i <= 31 and 1900 <= anio_i <= 2999):
        return ""
    return f"{anio_i:04d}-{mes:02d}-{dia_i:02d}"


def to_iso_date(value: str) -> str:
    texto = _fecha_con_mes_en_letra(value)
    if texto:
        return texto
    # Recorta primero la fecha del texto: varios comprobantes la traen pegada
    # a la hora ("2025-07-31 18:18:55") y el parseo crudo fallaba.
    encontrada = DATE_RE.search(clean_spaces(value))
    if encontrada:
        value = encontrada.group(0)
    value = clean_spaces(value).replace(".", "/").replace("-", "/")
    parts = [p for p in value.split("/") if p]
    if len(parts) != 3:
        return ""
    if len(parts[0]) == 4:
        year, month, day = parts
    else:
        day, month, year = parts
        if len(year) == 2:
            year = f"20{year}"
    try:
        day_i, month_i, year_i = int(day), int(month), int(year)
    except ValueError:
        return ""
    if not (1 <= day_i <= 31 and 1 <= month_i <= 12 and 1900 <= year_i <= 2999):
        return ""
    return f"{year_i:04d}-{month_i:02d}-{day_i:02d}"


def to_display_date(value: str) -> str:
    """AAAA-MM-DD -> DD-MM-YYYY para mostrar en la tabla."""
    iso = to_iso_date(value)
    if not iso:
        return clean_spaces(value)
    year, month, day = iso.split("-")
    return f"{day}-{month}-{year}"


def _valida_fecha(valor: str) -> bool:
    return bool(to_iso_date(valor))


def find_date(lines) -> str:
    for labels in (DATE_LABELS_PRIMARY, DATE_LABELS_SECONDARY, DATE_LABELS_LAST):
        iso = to_iso_date(valor_etiquetado(lines, labels, _valida_fecha))
        if iso:
            return iso
    for line in lines:
        if "fecha" in compact(line):
            iso = to_iso_date(line)
            if iso:
                return iso
    return ""


def _valida_rastreo(valor: str) -> bool:
    limpio = re.sub(r"[^A-Za-z0-9]", "", valor)
    # Descarta enmascarados y referencias cortas tipo "030826".
    return "*" not in valor and 8 <= len(limpio) <= 40 and any(c.isdigit() for c in limpio)


def _valida_referencia(valor: str) -> bool:
    """Más laxo que la clave de rastreo: acepta desde 6 caracteres."""
    limpio = re.sub(r"[^A-Za-z0-9]", "", valor)
    return "*" not in valor and 6 <= len(limpio) <= 40 and any(c.isdigit() for c in limpio)


CRITERIO_RASTREO = "T"      # el portal del CEP busca por clave de rastreo
CRITERIO_REFERENCIA = "R"   # ...o por número de referencia


def find_tracking_key(lines) -> tuple[str, str]:
    """Devuelve (valor, criterio) para la consulta en el portal del CEP.

    No todos los comprobantes traen clave de rastreo. Cuando solo hay un número
    de referencia —seis dígitos como 230426— el portal tiene que consultarse con
    el criterio "Número de referencia"; buscarlo como clave de rastreo no
    encuentra nada.
    """
    valor = valor_etiquetado(lines, TRACKING_LABELS, _valida_rastreo)
    criterio = CRITERIO_RASTREO
    if not valor:
        valor = valor_etiquetado(lines, TRACKING_LABELS_LAST, _valida_rastreo,
                                 excluir=("numerica",))
    if not valor:
        # Monex no imprime clave de rastreo, solo REFERENCIA_NUMERICA, y el
        # Banorte de agosto trae un "30570" de 5 dígitos que no es válido.
        valor = valor_etiquetado(lines, ("referencianumerica", "referencia"),
                                 _valida_referencia)
        criterio = CRITERIO_REFERENCIA
    if not valor:
        return "", CRITERIO_RASTREO
    primera = valor.split()[0] if valor.split() else valor
    candidato = re.sub(r"[^A-Za-z0-9]", "", primera)
    if len(candidato) < 8:
        candidato = re.sub(r"[^A-Za-z0-9]", "", valor)
    candidato = candidato[:30]

    # Seis dígitos son un número de referencia, no una clave de rastreo, aunque
    # el comprobante los haya puesto bajo esa etiqueta.
    if re.fullmatch(r"\d{6}", candidato):
        criterio = CRITERIO_REFERENCIA
    return candidato, criterio


def _digitos_de_cuenta(valor: str) -> str:
    """Extrae la cuenta del texto, rechazando lo enmascarado.

    Varios comprobantes escriben 'CLABE - NOMBRE' o 'BANCO - MXN ****184'. Los
    asteriscos significan que el banco ocultó la cuenta: es preferible dejar el
    campo vacío a guardar un número incompleto que Banxico va a rechazar.
    """
    if "*" in valor:
        return ""
    for trozo in re.findall(r"\d[\d\s.\-]{8,26}\d", valor):
        digitos = re.sub(r"\D", "", trozo)
        if len(digitos) == 18:          # CLABE completa, la mejor opción
            return digitos
    for trozo in re.findall(r"\d[\d\s.\-]{8,26}\d", valor):
        digitos = re.sub(r"\D", "", trozo)
        if 10 <= len(digitos) <= 19:
            return digitos
    return ""


def find_account(lines, labels) -> str:
    valor = valor_etiquetado(lines, labels, lambda v: bool(_digitos_de_cuenta(v)))
    return _digitos_de_cuenta(valor) if valor else ""


def fragmento_de_cuenta(lines, labels) -> str:
    """Dígitos visibles cuando el banco enmascaró la cuenta.

    Monex imprime '**************3529' y BNX 'BANORTE - MXN ***********184'.
    El fragmento no sirve para Banxico, pero sí para reconocer la cuenta si
    aparece completa en otro comprobante del mismo lote.
    """
    for line in lines:
        valor = label_value(line, labels)
        candidatos = [valor] if valor else []
        if valor is not None:
            indice = lines.index(line)
            candidatos += [l for l in lines[indice + 1: indice + 3]]
        for candidato in candidatos:
            if not candidato or "*" not in candidato:
                continue
            visibles = re.findall(r"\d{3,}", candidato.split("*")[-1])
            if visibles:
                return visibles[0]
    return ""


# Cuentas que ya conocemos. Sirven para reconstruir las que algún banco imprime
# enmascaradas, sin depender de que en la misma carga venga otro comprobante que
# las traiga completas. Se editan desde la app; el archivo JSON manda sobre esto.
CUENTAS_POR_DEFECTO = [
    {"clabe": "072650002119455184", "banco": "BANORTE",
     "titular": "ISLAS GOWER Y COMPANIA SUCESORES S EN C"},
    {"clabe": "044650256048143529", "banco": "SCOTIABANK",
     "titular": "ISLAS GOWER Y COMPANIA SUCESORES S EN C"},
    {"clabe": "012650001071268085", "banco": "BBVA MEXICO",
     "titular": "J2C SERVICIOS CORPORATIVOS SA"},
]

ARCHIVO_CUENTAS = "cuentas_conocidas.json"


def _ruta_cuentas() -> str:
    return os.path.join(_AQUI, ARCHIVO_CUENTAS)


def cargar_cuentas_conocidas() -> list:
    """Lee el catálogo del JSON; si no existe, arranca con los valores por defecto."""
    try:
        with open(_ruta_cuentas(), encoding="utf-8") as handle:
            datos = json.load(handle)
        if isinstance(datos, list) and datos:
            return datos
    except (OSError, ValueError):
        pass
    return [dict(c) for c in CUENTAS_POR_DEFECTO]


def guardar_cuentas_conocidas(cuentas: list) -> None:
    limpias = []
    for cuenta in cuentas:
        clabe = re.sub(r"\D", "", str(cuenta.get("clabe") or ""))
        if not clabe:
            continue
        limpias.append({
            "clabe": clabe,
            "banco": str(cuenta.get("banco") or "").strip(),
            "titular": str(cuenta.get("titular") or "").strip(),
        })
    with open(_ruta_cuentas(), "w", encoding="utf-8") as handle:
        json.dump(limpias, handle, ensure_ascii=False, indent=2)


def completar_cuentas_enmascaradas(records: list, avisos_por_archivo: dict) -> None:
    """Reconstruye las cuentas ocultas usando los demás comprobantes del lote.

    Solo actúa si el fragmento visible coincide con UNA sola cuenta completa
    entre los otros archivos. Si hay varias candidatas no adivina, porque un
    número de cuenta mal deducido es peor que un campo vacío. Lo que complete
    queda marcado como deducido para que se verifique.
    """
    # Dos fuentes: las cuentas completas que vienen en esta misma carga y el
    # catálogo de cuentas conocidas, que no depende de qué archivos se subieron.
    del_lote = {str(r.get("cuenta_beneficiaria") or "")
                for r in records if len(str(r.get("cuenta_beneficiaria") or "")) == 18}
    catalogo = {c["clabe"] for c in cargar_cuentas_conocidas()
                if len(re.sub(r"\D", "", str(c.get("clabe") or ""))) == 18}
    completas = del_lote | catalogo
    if not completas:
        return

    # Catálogo indexado por clave de banco: si el comprobante dice a qué banco
    # fue el pago, esa es la vía más directa y no depende de fragmentos.
    por_banco = {}
    for cuenta in cargar_cuentas_conocidas():
        clabe = re.sub(r"\D", "", str(cuenta.get("clabe") or ""))
        if len(clabe) != 18:
            continue
        code, _ = bank_code_from_name(str(cuenta.get("banco") or ""))
        if not code:
            code, _ = bank_from_clabe(clabe)
        if code:
            por_banco.setdefault(code, []).append(clabe)

    for record in records:
        actual = str(record.get("cuenta_beneficiaria") or "")
        if len(actual) == 18:
            continue

        # Vía 1: por banco receptor. "BANCO DESTINO: SCOTIA BANK INVERLAT" ya
        # dice qué cuenta toca, sin adivinar por los dígitos visibles.
        receptora = str(record.get("clave_receptora") or "")
        candidatas_banco = por_banco.get(receptora, [])
        if receptora and len(candidatas_banco) == 1:
            record["cuenta_beneficiaria"] = candidatas_banco[0]
            record["_cuenta_deducida"] = True
            avisos_por_archivo.setdefault(record.get("archivo", "?"), []).append(
                f"Cuenta beneficiaria DEDUCIDA: el comprobante indica banco receptor "
                f"{record.get('institucion_receptora') or receptora} y el catálogo tiene "
                f"{candidatas_banco[0]} para ese banco. Verifícala."
            )
            continue
        # Sirve tanto para cuentas enmascaradas ("****184") como para las que el
        # banco imprime incompletas: Banorte pone el número de cuenta de 10
        # dígitos en vez de la CLABE, y ese número vive dentro de la CLABE.
        fragmento = actual or str(record.get("_fragmento_cuenta") or "")
        if len(fragmento) < 3:
            continue
        candidatas = [c for c in completas if c.endswith(fragmento) or fragmento in c]
        if len(candidatas) != 1:
            if candidatas:
                avisos_por_archivo.setdefault(record.get("archivo", "?"), []).append(
                    f"La cuenta viene enmascarada (…{fragmento}) y hay {len(candidatas)} "
                    "cuentas del lote que podrían serlo. Captúrala a mano."
                )
            continue
        record["cuenta_beneficiaria"] = candidatas[0]
        record["_cuenta_deducida"] = True
        origen = "el catálogo de cuentas conocidas" if candidatas[0] in catalogo else \
                 "otro comprobante de esta carga"
        avisos_por_archivo.setdefault(record.get("archivo", "?"), []).append(
            f"Cuenta beneficiaria DEDUCIDA: el PDF solo trae {fragmento} y coincide "
            f"con {candidatas[0]}, tomada {origen}. Verifícala."
        )
        if not record.get("clave_receptora"):
            code, name = bank_from_clabe(candidatas[0])
            if code:
                record["clave_receptora"] = code
                record["institucion_receptora"] = name


def parse_amount(text: str) -> str:
    text = clean_spaces(text).replace("$", "").replace("MXN", "").replace("MXP", "")
    match = re.search(r"\d[\d.,]*", text)
    if not match:
        return ""
    number = match.group(0).rstrip(".,")
    last_dot, last_comma = number.rfind("."), number.rfind(",")
    if last_dot >= 0 and last_comma >= 0:
        # Conviven ambos: el separador de más a la derecha es el decimal.
        if last_dot > last_comma:
            number = number.replace(",", "")
        else:
            number = number.replace(".", "").replace(",", ".")
    elif last_dot >= 0 or last_comma >= 0:
        sep = "." if last_dot >= 0 else ","
        head, _, tail = number.rpartition(sep)
        # "1,200" o "1.200" son miles; "1200.50" o "1200,50" son decimales.
        if number.count(sep) > 1 or len(tail) == 3:
            number = number.replace(sep, "")
        else:
            number = f"{head.replace(sep, '')}.{tail}"
    try:
        return f"{float(number):.2f}"
    except ValueError:
        return ""


def _valida_monto(valor: str) -> bool:
    monto = parse_amount(valor)
    return bool(monto) and float(monto) > 0


def find_amount(lines) -> str:
    # Excluye comisión e IVA, que aparecen junto al importe y son montos válidos
    # pero no son el de la transferencia.
    valor = valor_etiquetado(
        lines, AMOUNT_LABELS, _valida_monto,
        excluir=("comision", "iva", "costodetransmision", "importeporconcepto"),
    )
    return parse_amount(valor) if valor else ""


def normalize_bank_name(name: str) -> str:
    name = clean_spaces(name).upper()
    name = re.sub(r"\b(S\.?A\.?|DE\s+C\.?V\.?|INSTITUCION\s+DE\s+BANCA\s+MULTIPLE|GRUPO\s+FINANCIERO)\b", " ", name)
    name = re.sub(r"[^A-Z0-9+ ]", " ", name)
    return clean_spaces(name)


def bank_code_from_name(name: str) -> tuple[str, str]:
    if not name:
        return "", ""
    clean = normalize_bank_name(name)
    if not clean:
        return "", ""

    alias = BANK_ALIASES.get(clean)
    if alias:
        for code, bank in VALID_BANKS.items():
            if bank.upper() == alias.upper():
                return code, bank

    for code, bank in VALID_BANKS.items():
        if clean == bank.upper():
            return code, bank

    # Sin espacios: el OCR pega las palabras y "SCOTIA BANK INVERLAT" llega como
    # "SCOTIABANKINVERLAT", que no casaba con ningún alias.
    compacto = re.sub(r"[^A-Z0-9]", "", clean)
    if compacto:
        for alias_name, target in BANK_ALIASES.items():
            if re.sub(r"[^A-Z0-9]", "", alias_name.upper()) == compacto:
                for code, bank in VALID_BANKS.items():
                    if bank.upper() == target.upper():
                        return code, bank
        for code, bank in VALID_BANKS.items():
            if re.sub(r"[^A-Z0-9]", "", bank.upper()) == compacto:
                return code, bank

    # Coincidencia por palabra completa: gana el nombre más largo encontrado.
    best = ("", "", 0)
    for alias_name, target in BANK_ALIASES.items():
        if re.search(rf"\b{re.escape(alias_name)}\b", clean) and len(alias_name) > best[2]:
            for code, bank in VALID_BANKS.items():
                if bank.upper() == target.upper():
                    best = (code, bank, len(alias_name))
    for code, bank in VALID_BANKS.items():
        upper = bank.upper()
        if re.search(rf"\b{re.escape(upper)}\b", clean) and len(upper) > best[2]:
            best = (code, bank, len(upper))
    return best[0], best[1]


def bank_from_clabe(account: str) -> tuple[str, str]:
    """Una CLABE de 18 dígitos trae el banco en los 3 primeros dígitos."""
    if not account or len(account) != 18 or not account.isdigit():
        return "", ""
    code = CLABE_PREFIX_TO_CODE.get(account[:3])
    if code:
        return code, VALID_BANKS[code]
    return "", ""


def find_bank_by_labels(lines, labels, etiquetas_de_cuenta=()) -> tuple[str, str]:
    """Banco a partir de sus etiquetas ("Banco Destino", "Institución Receptora").

    `etiquetas_de_cuenta` es un respaldo opcional para cuando el banco no viene
    etiquetado sino escrito en la misma línea de la cuenta ("BANORTE - MXN
    ****184"). Va como parámetro y no fijo dentro de la función porque si se
    usara siempre, al buscar el EMISOR tomaría el banco de la cuenta de abono,
    que es el receptor. Eso hacía que un comprobante de Santander se reportara
    como Scotiabank.
    """
    valor = valor_etiquetado(lines, labels, lambda v: bool(bank_code_from_name(v)[0]))
    if valor:
        code, name = bank_code_from_name(valor)
        if code:
            return code, name

    if etiquetas_de_cuenta:
        valor = valor_etiquetado(lines, etiquetas_de_cuenta,
                                 lambda v: bool(bank_code_from_name(v)[0]))
        if valor:
            return bank_code_from_name(valor)
    return "", ""


def bancos_por_orden(lines, etiqueta: str) -> list:
    """Bancos de una etiqueta que se repite: primero el origen, luego el destino.

    Hay comprobantes que usan el MISMO texto para las dos partes, como
    "De la institucion bancaria: ASPINTEGRAOPC" arriba y
    "De la institucion bancaria: BANORTE" abajo. Buscar por nombre de etiqueta
    devuelve siempre la primera y confunde emisor con receptor; lo único que los
    distingue es el orden.
    """
    codigos = []
    for i, line in enumerate(lines):
        if compact(line).count(etiqueta) != 1:
            continue
        valor = clean_spaces(label_value(line, (etiqueta,)) or "")
        code, _ = bank_code_from_name(valor)
        if not code:
            for siguiente in lines[i + 1: i + 2]:
                code, _ = bank_code_from_name(siguiente)
                if code:
                    break
        if code and (not codigos or codigos[-1] != code):
            codigos.append(code)
    return codigos


# Cuántos renglones seguidos puede tapar el bloque del beneficiario. Sin tope,
# una hoja con puras etiquetas sueltas se tragaba también el membrete del pie.
MAX_LINEAS_BENEFICIARIO = 6


def texto_de_membrete(lines) -> str:
    """Texto del comprobante sin los renglones que hablan del beneficiario.

    El descarte arrastra a los renglones siguientes porque en los layouts
    verticales la etiqueta y el valor están separados: Banamex escribe
    "Cuenta de depósito o beneficiario" y abajo "BANORTE - MXN ****184". Sin
    el arrastre se leía ese BANORTE como banco EMISOR, y como está antes en la
    hoja que la URL de bancanetempresarial.citibanamex.com.mx, le ganaba.
    """
    brand_lines = []
    arrastre = 0
    seguidos = 0
    for line in lines:
        comp = compact(line)
        if any(marker in comp for marker in BENEFICIARY_LINE_MARKERS):
            arrastre, seguidos = 2, 1
            continue
        if arrastre and seguidos < MAX_LINEAS_BENEFICIARIO:
            seguidos += 1
            # Una etiqueta sola vuelve a abrir la ventana para tapar SU valor,
            # que va en el renglón de abajo. Scotiabank imprime "BENEFICIARIO /
            # Cuenta de Abono / 0726... / Banco / BANORTE": sin esto, la
            # etiqueta pelona "Banco" cerraba el descarte y el BANORTE del
            # beneficiario entraba al membrete como si fuera el emisor.
            arrastre = 2 if _es_solo_etiqueta(line) else arrastre - 1
            continue
        arrastre, seguidos = 0, 0
        brand_lines.append(comp)
    return "\n".join(brand_lines)


def find_issuer_bank(lines, receptor_code: str = "") -> tuple[str, str]:
    code, name = find_bank_by_labels(lines, ISSUER_BANK_LABELS)
    if code:
        return code, name

    brand_text = texto_de_membrete(lines)

    # Se califica cada pista y gana la mejor, no la primera: manda la marca de
    # plataforma sobre el nombre pelón, luego el banco que NO es el receptor y,
    # ya en empate, el que aparece antes en la hoja.
    best = None
    for marker, code in ISSUER_MARKERS:
        pos = brand_text.find(marker)
        if pos < 0:
            continue
        rank = (
            marker in ISSUER_MARKERS_FUERTES,
            bool(receptor_code) and code != receptor_code,
            -pos,
            len(marker),
        )
        if best is None or rank > best[0]:
            best = (rank, code)
    if best:
        return best[1], VALID_BANKS[best[1]]
    return "", ""


HOLDER_LABELS = ("titulardelacuenta", "titularcuenta", "nombredeltitular", "titular")

# Etiquetas explícitas de cada parte. Casi todos los bancos que no son BBVA
# nombran directamente al ordenante y al beneficiario, lo cual es mucho más
# confiable que deducirlo del orden en que aparecen dos "Titular de la cuenta".
SENDER_NAME_LABELS = (
    "nombredelordenante", "clienteordenante", "nombreordenante", "razonsocial",
)
RECEIVER_NAME_LABELS = (
    "nombredelbeneficiario", "nombrebeneficiariorazonsocial", "nombrebeneficiario",
    "beneficiario",
)

# Todas las etiquetas conocidas, para saber si un renglón es etiqueta o valor.
ETIQUETAS_TODAS = (
    DATE_LABELS_PRIMARY + DATE_LABELS_SECONDARY + TRACKING_LABELS + AMOUNT_LABELS
    + RECEIVER_ACCOUNT_LABELS + SENDER_ACCOUNT_LABELS + RECEIVER_BANK_LABELS
    + HOLDER_LABELS + SENDER_NAME_LABELS + RECEIVER_NAME_LABELS
    + ("rfc", "rfcordenante", "rfcbeneficiario", "rfcocurpdelordenante", "moneda",
       "divisa", "concepto", "conceptodepago", "referencia", "referencianumerica",
       "bancodestino", "institucionordenante", "institucionreceptora", "estado",
       "tipodeoperacion", "tipodeenvio", "comision", "iva", "folio", "localizacion",
       "tipodecuenta", "tipodepersona", "tipodebeneficiario", "instrucciondepago")
)


def _valida_nombre(valor: str) -> bool:
    limpio = clean_spaces(valor)
    if not (3 < len(limpio) <= 80) or "*" in limpio:
        return False
    letras = sum(c.isalpha() for c in limpio)
    return letras >= 3 and letras >= len(limpio) * 0.5

# Etiquetas que marcan que ya se acabó el nombre del titular.
HOLDER_STOP_MARKERS = (
    "datonoverificado",
    "bancobeneficiario",
    "divisadelacuenta",
    "disponibilidaddepago",
    "conceptodepago",
    "fechade",
    "referencia",
    "cuentade",
    "importe",
    "banco",
)


def _es_continuacion_titular(line: str) -> bool:
    """Un renglón suelto en mayúsculas que continúa el nombre ('NORTE SA DE CV')."""
    limpio = clean_spaces(line)
    if not limpio or len(limpio) > 60:
        return False
    comp = compact(limpio)
    if any(marker in comp for marker in HOLDER_STOP_MARKERS):
        return False
    if label_value(limpio, HOLDER_LABELS) is not None:
        return False
    # Los nombres vienen en mayúsculas; descarta renglones con dígitos.
    return bool(re.fullmatch(r"[A-ZÁÉÍÓÚÑÜ&.,\- ]+", limpio)) and not re.search(r"\d", limpio)


def find_account_holders(lines) -> tuple[str, str]:
    """Devuelve (titular ordenante, titular beneficiario).

    En los comprobantes el primer 'Titular de la cuenta' es el de la cuenta de
    retiro y el segundo el de la cuenta de depósito. El nombre casi siempre se
    parte en dos o tres renglones.

    Ojo con el layout: en los PDF de texto los dos titulares vienen uno tras
    otro, pero en los escaneados el OCR los entrega en DOS COLUMNAS
    intercaladas — la continuación del primer nombre aparece ANTES del segundo
    nombre. Leerlos de corrido pega los nombres al revés, así que cuando se
    detecta el renglón combinado (el que trae las dos etiquetas) se reparten
    los segmentos por columna.
    """
    # Dos estrategias que se complementan. Primero las etiquetas explícitas
    # ("Nombre del Ordenante", "Nombre del Beneficiario"), que usan casi todos
    # los bancos salvo BBVA y no dependen del orden de aparición. Después, si
    # falta alguna, la lógica de los dos "Titular de la cuenta" de BBVA.
    # Se combinan campo por campo en vez de elegir una: un comprobante puede
    # traer el ordenante etiquetado y el beneficiario solo como titular.
    ordenante = valor_etiquetado(
        lines, SENDER_NAME_LABELS, _valida_nombre,
        excluir=("beneficiario", "receptor", "destino", "alacuenta", "banco", "institucion"),
    )
    beneficiario = valor_etiquetado(
        lines, RECEIVER_NAME_LABELS, _valida_nombre,
        excluir=("ordenante", "origen", "cargo", "retiro", "delacuenta", "banco", "institucion"),
    )
    if not (ordenante and beneficiario):
        # Algunas plantillas repiten la MISMA etiqueta para las dos partes
        # ("A nombre de :" arriba el origen, abajo el destino). Ahí no sirve
        # buscar por nombre de etiqueta: hay que ir por orden de aparición.
        valores = []
        for i, line in enumerate(lines):
            if compact(line).count("anombrede") != 1:
                continue
            valor = clean_spaces(label_value(line, ("anombrede",)) or "")
            if not _valida_nombre(valor):
                for siguiente in lines[i + 1: i + 2]:
                    if _valida_nombre(siguiente):
                        valor = clean_spaces(siguiente)
            if _valida_nombre(valor) and (not valores or valores[-1] != valor):
                valores.append(valor)
        if len(valores) >= 2:
            ordenante = ordenante or valores[0]
            beneficiario = valores[1]

    if not ordenante:
        # Layout por secciones: un encabezado "ORDENANTE" y más abajo "Nombre"
        # con el valor en el renglón siguiente (Scotiabank en vertical).
        for i, line in enumerate(lines):
            if compact(line).rstrip(":") != "ordenante":
                continue
            for j in range(i + 1, min(i + 12, len(lines))):
                if compact(lines[j]).startswith("beneficiario"):
                    break
                if compact(lines[j]).rstrip(":") == "nombre" and j + 1 < len(lines):
                    if _valida_nombre(lines[j + 1]):
                        ordenante = clean_spaces(lines[j + 1])
                    break
            if ordenante:
                break

    if not ordenante:
        # Scotiabank escribe "Cuenta Cargo: 65509483172 - LAM SOLUCIONES SC":
        # el nombre va después del guion, en la misma línea de la cuenta.
        crudo = valor_etiquetado(lines, SENDER_ACCOUNT_LABELS,
                                 lambda v: " - " in v and _valida_nombre(v.split(" - ", 1)[1]))
        if crudo:
            ordenante = clean_spaces(crudo.split(" - ", 1)[1])

    if not ordenante and any("institucionordenante" in compact(l) for l in lines):
        # PEIBO y UnalanaPAY no etiquetan al ordenante: lo ponen como
        # encabezado del documento, en el primer renglón.
        for line in lines[:3]:
            if _valida_nombre(line) and not label_value(line, ETIQUETAS_TODAS):
                ordenante = clean_spaces(line)
                break

    if ordenante and beneficiario:
        return clean_spaces(ordenante), clean_spaces(beneficiario)

    etiquetados = []          # (índice, valor) de renglones con UNA sola etiqueta
    combinado = -1            # renglón del OCR con las DOS etiquetas juntas

    for i, line in enumerate(lines):
        comp = compact(line)
        apariciones = comp.count("titulardelacuenta")
        if apariciones >= 2:
            if combinado < 0:
                combinado = i
            continue
        if apariciones == 1:
            valor = label_value(line, HOLDER_LABELS)
            if valor:
                etiquetados.append((i, clean_spaces(valor)))

    if not etiquetados:
        return clean_spaces(ordenante), clean_spaces(beneficiario)

    # Lo que ya se obtuvo por etiqueta explícita manda sobre el orden de los titulares.
    ordenante = ordenante or etiquetados[0][1]
    beneficiario = beneficiario or (etiquetados[1][1] if len(etiquetados) > 1 else "")

    if combinado >= 0 and beneficiario:
        # OCR en dos columnas: la continuación viene en un renglón combinado
        # posterior, con los segmentos unidos por tres espacios.
        for line in lines[combinado + 1 : combinado + 6]:
            partes = [p for p in re.split(r"\s{3,}", line.strip()) if p]
            if len(partes) == 2 and all(_es_continuacion_titular(p) for p in partes):
                ordenante = f"{ordenante} {clean_spaces(partes[0])}".strip()
                beneficiario = f"{beneficiario} {clean_spaces(partes[1])}".strip()
                break
        return clean_spaces(ordenante), clean_spaces(beneficiario)

    # PDF de texto: la continuación son los renglones que siguen a cada etiqueta.
    def con_continuacion(indice: int, valor: str) -> str:
        for line in lines[indice + 1 :]:
            if not _es_continuacion_titular(line):
                break
            valor = f"{valor} {clean_spaces(line)}"
        return clean_spaces(valor)

    ordenante = con_continuacion(etiquetados[0][0], ordenante)
    if len(etiquetados) > 1:
        beneficiario = con_continuacion(etiquetados[1][0], beneficiario)
    return ordenante, beneficiario


def parse_record(raw_text: str) -> dict:
    lines = split_lines(raw_text)

    receiver_account = find_account(lines, RECEIVER_ACCOUNT_LABELS)
    sender_account = find_account(lines, SENDER_ACCOUNT_LABELS)
    holder_sender, holder_receiver = find_account_holders(lines)
    clave_rastreo, criterio_busqueda = find_tracking_key(lines)

    # La CLABE del beneficiario dice sin ambigüedad quién RECIBE, y eso sirve de
    # pista para el emisor: entre dos nombres de banco sueltos en la hoja, el
    # que ya es el receptor es el candidato débil.
    clabe_code, clabe_name = bank_from_clabe(receiver_account)

    issuer_code, issuer_name = find_issuer_bank(lines, clabe_code)
    if not issuer_code:
        issuer_code, issuer_name = bank_from_clabe(sender_account)

    # Etiqueta repetida para origen y destino: el orden manda sobre lo anterior.
    repetidos = bancos_por_orden(lines, "delainstitucionbancaria")
    if len(repetidos) >= 2:
        issuer_code, issuer_name = repetidos[0], VALID_BANKS[repetidos[0]]

    receiver_code, receiver_name = find_bank_by_labels(
        lines, RECEIVER_BANK_LABELS, RECEIVER_ACCOUNT_LABELS)
    if len(repetidos) >= 2:
        receiver_code, receiver_name = repetidos[1], VALID_BANKS[repetidos[1]]
    if not receiver_code:
        receiver_code, receiver_name = clabe_code, clabe_name

    record = {
        "fecha": find_date(lines),
        "clave_rastreo": clave_rastreo,
        "criterio_busqueda": criterio_busqueda,
        "clave_emisora": issuer_code,
        "institucion_emisora": issuer_name,
        "clave_receptora": receiver_code,
        "institucion_receptora": receiver_name,
        "titular_ordenante": holder_sender,
        "cuenta_beneficiaria": receiver_account,
        "monto": find_amount(lines),
    }

    avisos = []
    if clabe_code and receiver_code and clabe_code != receiver_code:
        avisos.append(
            f"El banco beneficiario dice {receiver_name} pero la CLABE apunta a {clabe_name}"
        )
    record["_avisos"] = avisos
    record["_cuenta_ordenante_detectada"] = sender_account
    record["_fragmento_cuenta"] = (
        "" if receiver_account else fragmento_de_cuenta(lines, RECEIVER_ACCOUNT_LABELS))
    record["_titular_beneficiario"] = holder_receiver
    return record


FIELD_LABELS = {
    "fecha": "Fecha (DD-MM-YYYY)",
    "clave_rastreo": "Clave de rastreo",
    "clave_emisora": "Clave institución emisora",
    "clave_receptora": "Clave institución receptora",
    "cuenta_beneficiaria": "Cuenta beneficiaria",
    "monto": "Monto",
    "institucion_emisora": "Institución emisora",
    "institucion_receptora": "Institución receptora",
    "titular_ordenante": "Titular (quien deposita)",
    "criterio_busqueda": "Criterio de búsqueda",
}

# Columnas de nombre que se derivan de su clave (solo informativas, no van al TXT).
BANK_NAME_COLS = {
    "clave_emisora": "institucion_emisora",
    "clave_receptora": "institucion_receptora",
}

TABLE_COLS = [
    "archivo",
    "fecha",
    "clave_rastreo",
    "criterio_busqueda",
    "clave_emisora",
    "institucion_emisora",
    "titular_ordenante",
    "clave_receptora",
    "institucion_receptora",
    "cuenta_beneficiaria",
    "monto",
]

# Columnas que se muestran y editan pero NO van al TXT de Banxico.
EXTRA_FIELDS = ["titular_ordenante", "criterio_busqueda"]

TXT_FIELDS = [
    "fecha",
    "clave_rastreo",
    "clave_emisora",
    "clave_receptora",
    "cuenta_beneficiaria",
    "monto",
]


def validate_record(record: dict) -> list[str]:
    problems = []
    fecha = str(record.get("fecha", "") or "")
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", fecha):
        problems.append("fecha")
    if not re.fullmatch(r"[A-Za-z0-9]{6,30}", str(record.get("clave_rastreo", "") or "")):
        problems.append("clave de rastreo")
    if str(record.get("clave_emisora", "") or "") not in VALID_BANKS:
        problems.append("clave emisora")
    if str(record.get("clave_receptora", "") or "") not in VALID_BANKS:
        problems.append("clave receptora")
    cuenta = re.sub(r"\D", "", str(record.get("cuenta_beneficiaria", "") or ""))
    if not (10 <= len(cuenta) <= 19):
        problems.append("cuenta beneficiaria")
    try:
        if float(str(record.get("monto", "") or 0)) <= 0:
            problems.append("monto")
    except ValueError:
        problems.append("monto")
    return problems


def motivo_no_verificable(record: dict) -> str:
    """Por qué este pago no se puede consultar en el CEP, aunque esté completo.

    Un traspaso entre cuentas del MISMO banco no viaja por SPEI: lo liquida el
    banco por dentro, Banxico nunca lo ve y no existe un CEP que consultar. El
    portal contestaría "no se encontró el pago" aunque todos los datos estén
    bien, así que más vale decirlo antes de mandarlo al lote.
    """
    emisora = str(record.get("clave_emisora", "") or "").strip()
    receptora = str(record.get("clave_receptora", "") or "").strip()
    if emisora and emisora == receptora:
        banco = VALID_BANKS.get(emisora, emisora)
        return (
            f"el banco emisor y el receptor son el mismo ({emisora} {banco}), "
            "o sea un traspaso interno: no pasa por SPEI y Banxico no emite CEP"
        )
    return ""


def build_txt(records) -> str:
    lines = []
    for record in records:
        if validate_record(record) or motivo_no_verificable(record):
            continue
        lines.append(",".join(str(record.get(field, "")).strip() for field in TXT_FIELDS))
    return "\n".join(lines) + ("\n" if lines else "")


# Lo que `cep_banxico.py` antepone a cada línea de avance.
PREFIJO_PROGRESO = "CEP_PROGRESS"


def _bombear(flujo, recibir):
    """Vacía un pipe línea por línea hasta que el otro lado lo cierre.

    Los dos pipes SE TIENEN que ir vaciando aunque no nos interesen: el búfer
    del sistema operativo es chico, y si se llena el subproceso se queda
    bloqueado escribiendo y la consulta no avanza nunca.
    """
    try:
        for linea in flujo:
            recibir(linea)
    except Exception:  # noqa: BLE001
        # El pipe se cierra de golpe cuando matamos el proceso; es lo esperado.
        pass


def _correr_consulta(comando, limite, al_avanzar):
    """Corre cep_banxico.py leyendo su avance en vivo.

    Devuelve (expiro, codigo, stderr). No lee resultados: de eso se encarga
    quien llama, a partir del JSON.
    """
    entorno = dict(os.environ)
    # Los dos lados del pipe hablando UTF-8: sin esto, en Windows el subproceso
    # escribe en la codificación local y los estados con acento llegan rotos.
    entorno["PYTHONIOENCODING"] = "utf-8"

    proceso = subprocess.Popen(
        comando, cwd=_AQUI,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="replace", bufsize=1,
        env=entorno,
    )

    lineas = queue.Queue()
    stderr_partes = []
    FIN = object()   # centinela: stdout se cerró

    def recoger_stdout(linea):
        lineas.put(linea)

    def leer_stdout():
        try:
            _bombear(proceso.stdout, recoger_stdout)
        finally:
            lineas.put(FIN)

    hilo_out = threading.Thread(target=leer_stdout, daemon=True)
    hilo_err = threading.Thread(
        target=_bombear, args=(proceso.stderr, stderr_partes.append), daemon=True)
    hilo_out.start()
    hilo_err.start()

    expiro = False
    vence = time.monotonic() + limite
    try:
        while True:
            restante = vence - time.monotonic()
            if restante <= 0:
                expiro = True
                break
            try:
                # El sondeo corto es lo que permite respetar el timeout aunque
                # el subproceso se quede callado y colgado.
                linea = lineas.get(timeout=min(restante, 1.0))
            except queue.Empty:
                continue
            if linea is FIN:
                break
            _reportar_avance(linea, al_avanzar)

        if not expiro:
            try:
                proceso.wait(timeout=max(0.0, vence - time.monotonic()))
            except subprocess.TimeoutExpired:
                expiro = True
    finally:
        # Matar, cosechar y sólo entonces cerrar: si se cierran los pipes con
        # los hilos todavía leyendo, revientan; y sin el wait() el proceso queda
        # de zombi. Los hilos son daemon, así que ni aunque se atoren impiden
        # que Python termine.
        if proceso.poll() is None:
            proceso.kill()
        proceso.wait()
        hilo_out.join(timeout=5)
        hilo_err.join(timeout=5)
        for flujo in (proceso.stdout, proceso.stderr):
            if flujo is not None:
                try:
                    flujo.close()
                except Exception:  # noqa: BLE001
                    pass

    detalle = "".join(stderr_partes).strip()[-1500:]
    return expiro, proceso.returncode, detalle


def _reportar_avance(linea, al_avanzar):
    """Traduce una línea de stdout a una llamada de avance, si es que lo es.

    Todo lo que no sea exactamente `CEP_PROGRESS|hechos|total|estado` se ignora
    en silencio: por stdout puede salir cualquier otra cosa y no queremos que un
    renglón suelto se confunda con avance.
    """
    if al_avanzar is None:
        return
    linea = (linea or "").strip()
    if not linea.startswith(PREFIJO_PROGRESO + "|"):
        return
    partes = linea.split("|")
    if len(partes) != 4:
        return
    try:
        hechos, total = int(partes[1]), int(partes[2])
    except ValueError:
        return
    try:
        al_avanzar(hechos, total, partes[3])
    except Exception:  # noqa: BLE001
        # Que falle el pintado de la barra no puede cancelar la consulta.
        pass


def consultar_cep(records, al_avanzar=None):
    """Consulta en el portal del CEP los renglones que ya están completos.

    `al_avanzar(hechos, total, estado)` se llama conforme `cep_banxico.py` va
    cerrando pagos, para poder mover una barra de progreso de verdad. El avance
    llega por stdout del subproceso; los resultados siguen viniendo del JSON.

    Siempre headless: no se abre ninguna ventana de Chrome ni se le pide nada al
    usuario. Si Banxico saca el captcha, `cep_banxico.py` marca ese pago y sigue
    con los demás, así que aquí basta con leer los resultados.

    Corre `cep_banxico.py` como proceso aparte en vez de llamarlo aquí. Motivo:
    Streamlit deja la política de asyncio en WindowsSelectorEventLoopPolicy, y
    Playwright arma su loop con asyncio.new_event_loop(), que hereda esa
    política. En Windows el SelectorEventLoop no implementa subprocesos, y
    Playwright necesita lanzar su driver de Node como subproceso: revienta con
    NotImplementedError. La política es global al proceso, así que un hilo
    aparte tampoco lo salva; un proceso nuevo arranca con la política por
    defecto (Proactor) y funciona.
    """
    pagos = [
        {
            "fecha": cep_banxico.iso_a_ddmmyyyy(record["fecha"]),
            "clave_rastreo": record["clave_rastreo"],
            "criterio": (record.get("criterio_busqueda") or "T").strip().upper()[:1] or "T",
            "clave_emisora": record["clave_emisora"],
            "clave_receptora": record["clave_receptora"],
            "cuenta": record["cuenta_beneficiaria"],
            "monto": record["monto"],
            "etiqueta": record.get("archivo", ""),
        }
        for record in records
        if not validate_record(record) and not motivo_no_verificable(record)
    ]
    if not pagos:
        return []

    carpeta = tempfile.mkdtemp(prefix="cep_")
    ruta_entrada = os.path.join(carpeta, "pagos.json")
    ruta_salida = os.path.join(carpeta, "resultados.json")
    with open(ruta_entrada, "w", encoding="utf-8") as handle:
        json.dump(pagos, handle, ensure_ascii=False)

    comando = [
        sys.executable,
        os.path.join(_AQUI, "cep_banxico.py"),
        "--json-in", ruta_entrada,
        "--json-out", ruta_salida,
    ]

    # El margen fijo cubre el arranque de Playwright y, en el peor caso, las
    # pausas escalonadas que `cep_banxico.py` toma cuando sale el captcha (unos
    # tres minutos en total antes de anotar el resto del lote y salir).
    limite = 300 + 90 * len(pagos)

    # Se usa Popen y no subprocess.run porque hay que leer stdout MIENTRAS el
    # proceso corre: con run() el avance llegaría cuando ya no sirve de nada.
    # El log del subproceso se va por stderr, así que stdout trae el avance sin
    # mezclarse; y los resultados no salen por ninguno de los dos, sino del JSON
    # que el subproceso deja en disco.
    expiro, codigo, detalle = _correr_consulta(comando, limite, al_avanzar)

    if not os.path.exists(ruta_salida):
        if expiro:
            raise RuntimeError(
                "La consulta pasó del tiempo máximo sin alcanzar a resolver ningún "
                "pago. Vuelve a intentarlo con menos comprobantes, o usa la consulta "
                "por lotes de Banxico."
            )
        raise RuntimeError(
            f"cep_banxico.py terminó con código {codigo} sin escribir "
            f"resultados.\n\n{detalle}"
        )

    with open(ruta_salida, encoding="utf-8") as handle:
        return [cep_banxico.Resultado(**dato) for dato in json.load(handle)]


def _version_parser() -> float:
    """Marca de tiempo de este archivo, para invalidar la caché al editarlo."""
    try:
        return os.path.getmtime(_AQUI + os.sep + "cr3.py")
    except OSError:
        return 0.0


# Umbral de parecido para distinguir "otro nombre" de "el OCR se equivocó".
SIMILITUD_MINIMA = 0.90


def _parecido(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def comparar_nombres(excel: str, pdf: str) -> tuple[str, str]:
    """(estado, explicación). Estado: 'ok', 'parecido' o 'distinto'."""
    a, b = normalizar_nombre(excel), normalizar_nombre(pdf)
    if not a or not b:
        return "distinto", "falta el nombre en alguno de los dos"
    if a == b:
        return "ok", ""
    if a in b or b in a:
        return "ok", "coincide por contención (uno es nombre corto del otro)"
    razon = _parecido(a, b)
    if razon >= SIMILITUD_MINIMA:
        return "parecido", f"se parecen {razon:.0%}, probable ruido de OCR"
    return "distinto", f"solo se parecen {razon:.0%}"


def validar_excel_contra_pdf(datos_excel: dict, record: dict, completo: dict) -> dict:
    """Compara un Excel de dispersión contra el comprobante ya extraído.

    Bloquean el emparejamiento el monto y el pagador; el proveedor y la cuenta
    solo generan observación. La cuenta de los Excel viene sistemáticamente
    distinta a la del pago, así que si bloqueara no pasaría nunca ningún renglón.
    """
    detalles, duro, blando = [], 0, 0

    monto_pdf = a_numero(record.get("monto"))
    monto_xls = datos_excel.get("total")
    if monto_pdf is None or monto_xls is None:
        duro += 1
        detalles.append(f"{sio_tema.ICONO['error']} monto: falta en alguno de los dos")
    elif round(monto_pdf, 2) == round(monto_xls, 2):
        detalles.append(f"{sio_tema.ICONO['ok']} monto: {round(monto_pdf, 2):,.2f}")
    else:
        duro += 1
        detalles.append(f"{sio_tema.ICONO['error']} monto: Excel {monto_xls:,.2f} vs comprobante {monto_pdf:,.2f}")

    estado, nota = comparar_nombres(datos_excel.get("cliente", ""), record.get("titular_ordenante", ""))
    if estado == "ok":
        detalles.append(f"{sio_tema.ICONO['ok']} pagador: " + (nota or "coincide"))
    elif estado == "parecido":
        blando += 1
        detalles.append(f"{sio_tema.ICONO['aviso']} pagador: {nota}")
    else:
        duro += 1
        detalles.append(f"{sio_tema.ICONO['error']} pagador: Excel «{datos_excel.get('cliente','')}» vs "
                        f"comprobante «{record.get('titular_ordenante','')}» ({nota})")

    estado, nota = comparar_nombres(datos_excel.get("proveedor", ""), completo.get("_titular_beneficiario", ""))
    if estado == "ok":
        detalles.append(f"{sio_tema.ICONO['ok']} proveedor: " + (nota or "coincide"))
    else:
        blando += 1
        detalles.append(f"{sio_tema.ICONO['aviso']} proveedor: Excel «{datos_excel.get('proveedor','')}» vs "
                        f"comprobante «{completo.get('_titular_beneficiario','')}»")

    cuenta_xls = re.sub(r"\D", "", datos_excel.get("cuenta_bancaria", "") or "")
    cuenta_pdf = re.sub(r"\D", "", record.get("cuenta_beneficiaria", "") or "")
    if cuenta_xls and cuenta_pdf and cuenta_xls == cuenta_pdf:
        detalles.append(f"{sio_tema.ICONO['ok']} cuenta de depósito coincide")
    else:
        blando += 1
        detalles.append(f"{sio_tema.ICONO['aviso']} cuenta: Excel {cuenta_xls or '(vacía)'} vs pago {cuenta_pdf or '(vacía)'}")

    if duro:
        estado_final = "NO CUADRA"
    elif blando:
        estado_final = "CON OBSERVACIONES"
    else:
        estado_final = "CUADRA"
    return {"estado": estado_final, "detalles": detalles, "duro": duro, "blando": blando}


def emparejar_excels(excels, records, debug) -> tuple[list, list]:
    """Empareja por contenido (pagador + monto); el nombre de archivo solo desempata.

    Devuelve (emparejados, sueltos). Cada emparejado trae el Excel, el renglón
    del comprobante y el resultado de la validación.
    """
    emparejados, sueltos = [], []
    usados = set()

    for nombre_archivo, datos in excels:
        candidatos = []
        for indice, record in enumerate(records):
            if indice in usados or validate_record(record):
                continue
            completo = debug.get(record.get("archivo", ""), {})
            veredicto = validar_excel_contra_pdf(datos, record, completo)
            if veredicto["duro"] == 0:
                candidatos.append((indice, record, veredicto))

        if not candidatos:
            sueltos.append((nombre_archivo, datos, "ningún comprobante cuadra en pagador y monto"))
            continue

        if len(candidatos) > 1:
            # Desempate por nombre de archivo; si sigue habiendo varios, no adivino.
            base = normalizar_nombre(os.path.splitext(nombre_archivo)[0])
            afinados = [c for c in candidatos
                        if normalizar_nombre(os.path.splitext(c[1].get("archivo", ""))[0]) == base]
            if len(afinados) == 1:
                candidatos = afinados
            else:
                sueltos.append((nombre_archivo, datos,
                                f"ambiguo: {len(candidatos)} comprobantes cuadran igual, no lo empareja solo"))
                continue

        indice, record, veredicto = candidatos[0]
        usados.add(indice)
        emparejados.append({"excel": nombre_archivo, "datos": datos,
                            "record": record, "veredicto": veredicto})

    return emparejados, sueltos


def _ancho_tabla() -> dict:
    """Streamlit cambió el parámetro de ancho en 1.49.

    Antes `width` era un entero de píxeles y el ancho completo se pedía con
    `use_container_width`; después `width="stretch"` lo reemplazó. Esta máquina
    tiene 1.45 en el Python del sistema y 1.61 en el .venv, así que el archivo
    tiene que servir en los dos.
    """
    try:
        mayor, menor = (int(parte) for parte in st.__version__.split(".")[:2])
    except Exception:  # noqa: BLE001
        return {}
    return {"width": "stretch"} if (mayor, menor) >= (1, 49) else {"use_container_width": True}


ANCHO_TABLA = _ancho_tabla()

# Cómo se lee cada estado del portal dentro de la tabla de resultados.
#
# Va en texto y no en icono a propósito: las celdas de `st.dataframe` se pintan
# en plano, así que una directiva de icono se vería tal cual está escrita. Es
# además lo que hacen las tablas de SIO, que marcan el estado con una etiqueta
# de texto y no con un símbolo.
ESTADO_LEGIBLE = {
    "LIQUIDADO": "Liquidado",
    "NO_ENCONTRADO": "No encontrado",
    # Dos renglones distintos: al primero Banxico le pidió el captcha; al
    # segundo ni se le preguntó, porque la tanda venía trabada.
    "CAPTCHA": "Captcha de Banxico",
    "OMITIDO_CAPTCHA": "No consultado por captcha",
    "ERROR": "Error",
}

def estado_legible(estado: str) -> str:
    """Cómo se le muestra un estado a quien usa la app.

    Los estados que vienen del portal (DEVUELTO, CANCELADO...) no están en el
    diccionario, así que se acomodan a mano en vez de salir gritando.
    """
    estado = (estado or "").strip()
    if not estado:
        return ""
    return ESTADO_LEGIBLE.get(estado, estado.replace("_", " ").capitalize())


# Lo que se pone en la columna «nombre» del cruce cuando la tarjeta no aparece
# en la base. Es una constante y no un texto suelto porque, además de leerse,
# sirve para reconocer esos renglones al armar la dispersión.
SIN_FICHA = "(no está en la base)"


def render_bank_reference() -> str:
    lines = ["Clave de la institución,Nombre de la institución"]
    lines.extend(f"{code},{name}" for code, name in VALID_BANKS.items())
    return "\n".join(lines) + "\n"


INSTRUCCIONES = """Para poder utilizar el servicio de consulta de Comprobante Electrónico de Pago (CEP) por lotes es necesario generar un archivo en formato de texto ".TXT" con la información necesaria para obtener los CEP desde dos y hasta 500 transferencias.
Cada línea del archivo debe corresponder a los datos de una de las transferencias y debe contener la siguiente información:

- Fecha en la que realizó la transferencia (AAAA-MM-DD)
- Clave de rastreo
- Clave de la institución financiera emisora de la transferencia
- Clave de la institución financiera receptora de la transferencia
- Cuenta beneficiaria (CLABE/Tarjeta de débito/Número celular)
- Monto de la transferencia

Cada dato debe estar separado por una coma dentro del archivo, como se muestra en el siguiente ejemplo:

2018-09-06,11FEF28A36F,40058,40102,5533302929,1200.50

"""


# --------------------------------------------------------------------------- #
# App
# --------------------------------------------------------------------------- #

def process_pdf(raw_bytes: bytes, key: str) -> dict:
    """Cachea solo las lecturas exitosas: un 'sin texto' se vuelve a intentar.

    (Con st.cache_data un fallo quedaba pegado en memoria aunque después se
    instalara el motor de OCR, porque Streamlit recarga en el mismo proceso.)
    """
    cache = st.session_state.setdefault("_pdf_cache", {})
    cached = cache.get(key)
    # Se invalida por dos motivos: que al parser le hayan agregado campos, o que
    # el archivo del parser haya cambiado. Lo segundo importa porque Streamlit
    # recarga el código en caliente pero la caché sobrevive, y entonces se
    # siguen viendo los valores que calculó la versión anterior.
    if (cached is not None
            and cached.get("_version_parser") == _version_parser()
            and all(campo in cached for campo in TABLE_COLS[1:])):
        return cached

    raw_text, origin = extract_pdf_text(raw_bytes)
    record = parse_record(raw_text)
    record["_texto"] = raw_text
    record["_origen"] = origin
    record["_version_parser"] = _version_parser()
    if origin != "sin texto":
        cache[key] = record
    return record


# Clave y nombre de institución quedan sincronizados en los dos sentidos: se
# puede corregir por clave o por nombre, y el otro campo sigue. Hace falta
# porque cuando el banco solo se identifica por su logo, el parser no tiene de
# dónde leerlo y la corrección tiene que ser a mano.
NOMBRE_A_CLAVE = {nombre.upper(): codigo for codigo, nombre in VALID_BANKS.items()}


def _config_columnas(columnas):
    """Cómo se ve y se edita cada columna de la tabla."""
    todas = {
        "archivo": st.column_config.TextColumn("Archivo", disabled=True),
        "fecha": st.column_config.TextColumn(FIELD_LABELS["fecha"]),
        "clave_rastreo": st.column_config.TextColumn(FIELD_LABELS["clave_rastreo"]),
        "criterio_busqueda": st.column_config.SelectboxColumn(
            FIELD_LABELS["criterio_busqueda"], options=["T", "R"], required=False,
            help="T = clave de rastreo, R = número de referencia. Es lo que se "
                 "elige en el portal del CEP; 6 dígitos suelen ser referencia.",
        ),
        "clave_emisora": st.column_config.SelectboxColumn(
            FIELD_LABELS["clave_emisora"], options=list(VALID_BANKS), required=False
        ),
        "institucion_emisora": st.column_config.SelectboxColumn(
            FIELD_LABELS["institucion_emisora"], options=list(VALID_BANKS.values()),
            required=False,
            help="Editable: si el comprobante solo identifica al banco por su logo, "
                 "corrígelo aquí y la clave se ajusta sola.",
        ),
        "clave_receptora": st.column_config.SelectboxColumn(
            FIELD_LABELS["clave_receptora"], options=list(VALID_BANKS), required=False
        ),
        "institucion_receptora": st.column_config.SelectboxColumn(
            FIELD_LABELS["institucion_receptora"], options=list(VALID_BANKS.values()),
            required=False,
            help="Editable: al cambiarlo se ajusta la clave receptora.",
        ),
        "titular_ordenante": st.column_config.TextColumn(
            FIELD_LABELS["titular_ordenante"],
            help="Titular de la cuenta de retiro, o sea quien envía el dinero. "
                 "No va al TXT; es para validar.",
        ),
        "cuenta_beneficiaria": st.column_config.TextColumn(FIELD_LABELS["cuenta_beneficiaria"]),
        "monto": st.column_config.TextColumn(FIELD_LABELS["monto"]),
    }
    return {col: config for col, config in todas.items() if col in columnas}


def tabla_editable(rows, key: str, columnas=None):
    """Muestra los comprobantes en una tabla editable y devuelve los registros.

    `key` tiene que ser distinta en cada pestaña: Streamlit identifica los
    widgets por esa llave y dos tablas con la misma se pisarían las ediciones.
    """
    columnas = list(columnas or TABLE_COLS)
    # Todo se maneja como texto. Las columnas de nombre de institución no vienen
    # en `rows` (se derivan de la clave), así que pandas 3 las crea vacías con
    # dtype float64 y luego revienta al escribirles texto:
    #   TypeError: Invalid value '<ArrowStringArray> [...]' for dtype 'float64'
    df = pd.DataFrame(rows, columns=columnas).fillna("").astype(str)

    # Si el usuario ya cambió una clave en el editor, adelanta ese cambio para que
    # el nombre de la institución se muestre al parejo y no un paso atrás.
    editor_state = st.session_state.get(key)
    if isinstance(editor_state, dict):
        for index, changes in (editor_state.get("edited_rows") or {}).items():
            if int(index) >= len(df):
                continue
            for clave_col, nombre_col in BANK_NAME_COLS.items():
                if nombre_col not in columnas or clave_col not in columnas:
                    continue
                if nombre_col in changes:
                    # Editaron el nombre: manda el nombre y la clave lo sigue.
                    nombre = str(changes[nombre_col] or "").strip()
                    df.at[int(index), nombre_col] = nombre
                    df.at[int(index), clave_col] = NOMBRE_A_CLAVE.get(nombre.upper(), "")
                    changes.pop(clave_col, None)
                elif clave_col in changes:
                    df.at[int(index), clave_col] = changes[clave_col]

    for clave_col, nombre_col in BANK_NAME_COLS.items():
        if nombre_col not in columnas or clave_col not in columnas:
            continue
        faltan = df[nombre_col].isna() | (df[nombre_col].astype(str).str.strip() == "")
        df.loc[faltan, nombre_col] = df.loc[faltan, clave_col].map(
            lambda code: VALID_BANKS.get(str(code or "").strip(), ""))

    # En la tabla la fecha se ve y se edita como DD-MM-YYYY; al TXT va como AAAA-MM-DD.
    df["fecha"] = df["fecha"].map(to_display_date)

    edited = st.data_editor(
        df,
        key=key,
        **ANCHO_TABLA,
        hide_index=True,
        num_rows="dynamic",
        column_config=_config_columnas(columnas),
    )

    records = edited.fillna("").astype(str).to_dict("records")
    for record in records:
        record["fecha"] = to_iso_date(record.get("fecha", ""))
    return records


def bloque_cep(records, valid_lines: int, prefijo: str):
    """Sección "Estado del pago en Banxico": botón, consulta y resultados.

    `prefijo` separa los widgets y los resultados de cada pestaña, para que
    verificar en una no le pise la tabla de resultados a la otra.
    """
    st.subheader("Estado del pago en Banxico")

    llave_resultados = f"cep_resultados_{prefijo}"

    if cep_banxico is None:
        st.warning(
            f"No se pudo cargar `cep_banxico.py` ({CEP_IMPORT_ERROR}). "
            "Instálalo con `pip install playwright` y luego `playwright install chromium`."
        )
    else:
        st.caption(
            "Consulta cada pago en el portal del CEP y te dice lo que reporta SPEI. "
            "El portal atiende de 09:30 a 23:00 hrs y pide captcha después de unas pocas "
            "consultas seguidas; los pagos que caigan en el captcha se marcan y la "
            "consulta sigue con los demás."
        )
        verificar = st.button(
            "Verificar en Banxico", type="primary", disabled=valid_lines == 0,
            key=f"cep_verificar_{prefijo}", icon=sio_tema.ICONO["verificar"],
        )

        if verificar:
            # La barra arranca en cero y sólo se mueve cuando `cep_banxico.py`
            # avisa que cerró un pago: nada de progreso inventado por tiempo.
            barra = st.progress(0.0)
            aviso = st.empty()
            aviso.caption(f"Consultando pago 0 de {valid_lines}...")

            def avanzar(procesados, total, estado):
                total = total or valid_lines or 1
                barra.progress(min(procesados / total, 1.0))
                aviso.caption(
                    f"Consultando pago {procesados} de {total} — {estado_legible(estado)}"
                )

            with st.spinner(f"Consultando {valid_lines} pago(s) en el portal del CEP..."):
                try:
                    resultados = consultar_cep(records, al_avanzar=avanzar)
                    st.session_state[llave_resultados] = resultados

                    if len(resultados) >= valid_lines:
                        # Todos tienen estado final, aunque varios se hayan
                        # omitido por captcha: el lote está cerrado.
                        barra.progress(1.0)
                        aviso.caption(
                            f"Consulta finalizada: {len(resultados)} de {valid_lines} "
                            "pagos procesados."
                        )
                    else:
                        # Se cortó a medias. La barra se queda donde llegó: no
                        # tiene por qué decir que terminó cuando no terminó.
                        aviso.caption(
                            f"Consulta interrumpida: se procesaron {len(resultados)} "
                            f"de {valid_lines} pagos."
                        )
                except Exception as exc:  # noqa: BLE001
                    st.session_state[llave_resultados] = []
                    # No se rescató nada: la barra sobra y estorba.
                    barra.empty()
                    aviso.empty()
                    detalle = traceback.format_exc()
                    st.error(f"Falló la consulta: {type(exc).__name__}: {exc}")
                    with st.expander("Detalle técnico del error"):
                        st.code(detalle, language="text")
                    try:
                        ruta_log = os.path.join(_AQUI, "cep_error.log")
                        with open(ruta_log, "w", encoding="utf-8") as log:
                            log.write(detalle)
                        st.caption(f"Traza guardada en {ruta_log}")
                    except Exception:  # noqa: BLE001
                        pass

    resultados_cep = st.session_state.get(llave_resultados) or []
    if resultados_cep:
        filas = []
        for resultado in resultados_cep:
            # Lo que dice el portal manda; el diccionario sólo cubre los casos
            # que resuelve la app sin llegar a él.
            estado = resultado.estado_portal or estado_legible(resultado.estado)
            filas.append({
                "Archivo": resultado.etiqueta,
                "Estado": estado,
                "Clave de rastreo": resultado.clave_rastreo,
                "Recibido en SPEI": resultado.recepcion,
                "Procesado": resultado.procesamiento,
                "Monto según Banxico": resultado.monto_portal,
                "Detalle": resultado.detalle,
            })
        st.dataframe(pd.DataFrame(filas), hide_index=True, **ANCHO_TABLA)

        liquidados = sum(1 for r in resultados_cep if r.estado == "LIQUIDADO")
        st.metric("Pagos liquidados", f"{liquidados} de {len(resultados_cep)}")

        # Menos resultados que pagos mandados = la consulta se cortó a medias
        # (por tiempo o porque el proceso se cayó). Lo de arriba es lo que sí se
        # alcanzó a rescatar, y vale.
        faltantes = valid_lines - len(resultados_cep)
        if faltantes > 0:
            st.warning(
                f"Quedaron {faltantes} pago(s) sin consultar de {valid_lines}: la "
                "consulta no alcanzó a terminar. Lo que aparece arriba ya está "
                "verificado; vuelve a darle a Verificar para intentar el resto."
            )

        con_captcha = sum(1 for r in resultados_cep if r.estado == "CAPTCHA")
        omitidos = sum(1 for r in resultados_cep if r.estado == "OMITIDO_CAPTCHA")
        if con_captcha or omitidos:
            partes = []
            if con_captcha:
                partes.append(
                    f"Banxico solicitó CAPTCHA en {con_captcha} pago(s): esa consulta se "
                    "omitió y el proceso continuó."
                )
            if omitidos:
                partes.append(
                    f"Otros {omitidos} pago(s) ya no se consultaron, porque el captcha "
                    "siguió apareciendo en varias consultas seguidas."
                )
            partes.append(
                "Los demás resultados de arriba son válidos. Vuelve a darle a Verificar "
                "más tarde para los que faltan, o usa la consulta por lotes, que es lo "
                "que conviene en tandas grandes: https://www.banxico.org.mx/cep-scl/"
            )
            st.info(" ".join(partes))

        for resultado in resultados_cep:
            if resultado.mensaje_portal:
                with st.expander(f"Respuesta del portal — {resultado.etiqueta}"):
                    st.text(resultado.mensaje_portal)

    return resultados_cep


def revisar_registros(records):
    """Avisa lo que falta y lo que no se puede consultar. Devuelve los no verificables."""
    incomplete = [(r.get("archivo", "?"), validate_record(r)) for r in records]
    for name, problems in incomplete:
        if problems:
            st.warning(f"'{name}': falta o es inválido -> {', '.join(problems)}")

    # Se evalúa sobre la tabla YA editada: si corriges a mano el banco emisor o el
    # receptor, el aviso se quita solo y el renglón vuelve al TXT.
    sin_cep = [(r.get("archivo", "?"), motivo_no_verificable(r)) for r in records]
    sin_cep = [(name, motivo) for name, motivo in sin_cep if motivo]
    for name, motivo in sin_cep:
        st.error(
            f"'{name}': no se puede verificar en Banxico porque {motivo}. "
            "Queda fuera del TXT y de la consulta del CEP."
        )
    return sin_cep


def resumen_lote(txt_content: str, sin_cep) -> int:
    """Contador de transferencias listas y los avisos de límites de Banxico."""
    valid_lines = len(txt_content.strip().splitlines()) if txt_content.strip() else 0
    col_a, col_b = st.columns([1, 2])
    with col_a:
        st.metric("Transferencias listas", valid_lines)
    with col_b:
        if sin_cep:
            st.info(
                f"{len(sin_cep)} traspaso(s) interno(s) fuera del lote: mismo banco de "
                "origen y destino, no hay CEP que consultar."
            )
        if valid_lines == 1:
            st.warning("Banxico requiere de 2 a 500 transferencias por archivo.")
        elif valid_lines > 500:
            st.error("Banxico acepta máximo 500 transferencias por archivo.")
    return valid_lines


def leer_comprobantes(uploaded_files):
    """Lee los PDFs subidos y avisa lo que salió raro.

    Devuelve (renglones para la tabla, detalle completo por archivo). Vive
    aparte de la vista porque las dos pestañas de app_cep.py hacen exactamente
    lo mismo para empezar: leer, avisar y armar los renglones.
    """
    st.session_state["ocr_errors"] = []

    rows, debug = [], {}
    progress = st.progress(0.0, text="Leyendo PDFs...")
    for i, uploaded_file in enumerate(uploaded_files, start=1):
        raw_bytes = uploaded_file.getvalue()
        key = hashlib.md5(raw_bytes).hexdigest()
        progress.progress(i / len(uploaded_files), text=f"Leyendo {uploaded_file.name}...")
        record = process_pdf(raw_bytes, key)

        if record["_origen"] == "sin texto":
            st.error(
                f"'{uploaded_file.name}': no se pudo leer texto ni con OCR. "
                "Revisa que el PDF no esté protegido o dañado."
            )
        elif record["_origen"] in ("ocr", "mixto"):
            st.info(f"'{uploaded_file.name}': leído con OCR ({record['_origen']}). Verifica los datos.")

        for aviso in record.get("_avisos", []):
            st.warning(f"'{uploaded_file.name}': {aviso}")

        rows.append({
            "archivo": uploaded_file.name,
            **{f: record.get(f, "") for f in TXT_FIELDS + EXTRA_FIELDS},
        })
        debug[uploaded_file.name] = record
    progress.empty()

    # Con todos los comprobantes leídos ya se pueden reconstruir las cuentas que
    # algún banco imprimió enmascaradas o incompletas.
    avisos_cruce = {}
    completar_cuentas_enmascaradas(rows, avisos_cruce)
    for archivo, mensajes in avisos_cruce.items():
        for mensaje in mensajes:
            st.warning(f"'{archivo}': {mensaje}")

    for error in st.session_state.get("ocr_errors", []):
        st.warning(f"OCR: {error}")

    return rows, debug


COMO_SE_LLENAN = """
| Campo Banxico | De dónde sale del comprobante |
| --- | --- |
| Fecha del pago | Fecha de aplicación (si no está, fecha de creación) |
| Clave de rastreo | Clave de rastreo |
| Institución emisora | Membrete/marca del banco que emite la hoja, o la CLABE de la cuenta de retiro |
| Institución receptora | Banco beneficiario, validado contra los 3 primeros dígitos de la CLABE |
| Cuenta beneficiaria | Cuenta de depósito del comprobante. En el portal del CEP se captura en el campo "Cuenta Beneficiaria", con "Pago a Banco" **desmarcado** |
| Monto del pago | Importe |
"""


COLS_DISPERSION = ["nombre", "tarjeta", "importe", "retencion", "pago_final"]
COLS_CRUCE = ["nombre", "cuenta", "importe", "retencion", "validacion"]


def con_totales(tabla, columnas_suma, etiqueta="TOTAL DISPERSIÓN", columna_etiqueta="nombre"):
    """Copia de la tabla con un renglón de totales al final.

    Sólo para mostrar y exportar: las métricas y los cruces se calculan siempre
    sobre la tabla sin el total, para no contarlo como un renglón más.
    """
    if tabla.empty:
        return tabla
    total = {}
    for columna in tabla.columns:
        if columna in columnas_suma:
            total[columna] = pd.to_numeric(tabla[columna], errors="coerce").sum()
        elif columna == columna_etiqueta:
            total[columna] = etiqueta
        else:
            # None y no "" para no volver texto una columna de números.
            total[columna] = None
    return pd.concat([tabla, pd.DataFrame([total])], ignore_index=True)


# Dos resguardos, dos cosas distintas:
#   cliente  -> lo que se le retuvo a la gente; ese dinero se le queda al
#               cliente, no se deposita.
#   convenia -> lo que sí se le debe a alguien pero no tiene tarjeta a dónde
#               mandarlo, así que se queda en Convenia hasta que la haya.
COLS_RESGUARDO_CLIENTE = ["nombre", "tarjeta", "importe", "retencion"]
COLS_RESGUARDO_CONVENIA = ["nombre", "importe", "retencion", "en_resguardo"]


def tablas_resguardo(datos):
    """Los dos resguardos de un Excel. Devuelve (cliente, convenia)."""
    # La retención cuenta venga de donde venga: también de un renglón sin
    # tarjeta, y también cuando se llevó el pago completo y el neto quedó en
    # cero. Ese último caso es justo el que hoy se perdía, porque el acumulado
    # descarta los renglones que no dejan nada que depositar.
    cliente = [
        {
            "nombre": renglon["nombre"],
            "tarjeta": renglon.get("tarjeta", ""),
            "importe": renglon.get("importe"),
            "retencion": renglon.get("retencion"),
        }
        for renglon in list(datos.get("detalle", [])) + list(datos.get("sin_tarjeta", []))
        if (renglon.get("retencion") or 0)
    ]

    convenia = []
    for renglon in datos.get("sin_tarjeta", []):
        neto = renglon.get("pago_final")
        if neto is None:
            # Sin columna de neto, lo que queda en resguardo es el bruto menos
            # lo retenido, que es la misma cuenta que hace el resto de la app.
            neto = (renglon.get("importe") or 0) - (renglon.get("retencion") or 0)
        convenia.append({
            "nombre": renglon["nombre"],
            "importe": renglon.get("importe"),
            "retencion": renglon.get("retencion") or 0.0,
            "en_resguardo": round(neto, 2),
        })

    return (pd.DataFrame(cliente, columns=COLS_RESGUARDO_CLIENTE),
            pd.DataFrame(convenia, columns=COLS_RESGUARDO_CONVENIA))


def _filas_para_excel(tabla, columnas):
    """La tabla como lista de listas, con NaN convertido a vacío."""
    filas = []
    for renglon in tabla[columnas].to_dict("records"):
        fila = []
        for columna in columnas:
            valor = renglon[columna]
            if valor is None or (isinstance(valor, float) and pd.isna(valor)):
                fila.append("")
            elif isinstance(valor, (int, float)):
                fila.append(float(valor))
            else:
                fila.append(str(valor))
        filas.append(fila)
    return filas


def armar_bloques(excels, acumulados, cruces):
    """Un bloque por cliente para el Excel de descarga."""
    bloques = []
    for nombre_excel, datos in excels:
        dispersion = acumulados.get(nombre_excel)
        if dispersion is None:
            dispersion = pd.DataFrame(columns=COLS_DISPERSION)
        resguardo_cliente, resguardo_convenia = tablas_resguardo(datos)
        # Se exporta si hay algo que exportar. Antes bastaba con que la
        # dispersión viniera vacía para saltarse el cliente entero; ahora un
        # periodo donde todo se retuvo sigue apareciendo, con su resguardo.
        if dispersion.empty and resguardo_cliente.empty and resguardo_convenia.empty:
            continue
        cruce = cruces.get(nombre_excel)
        bloques.append({
            "cliente": datos.get("cliente") or nombre_excel,
            "excel": nombre_excel,
            "dispersion": {
                "encabezados": ["NOMBRE", "TARJETA", "IMPORTE", "RETENCIÓN", "PAGO FINAL"],
                "filas": _filas_para_excel(dispersion, COLS_DISPERSION),
                "numericas": [2, 3, 4],
                "total_en": [2, 3, 4],
                "columnas_guion": [3],
                "columna_etiqueta": 1,
                "etiqueta_total": "TOTAL DISPERSIÓN",
            },
            "cruce": {
                "encabezados": ["NOMBRE", "CUENTA", "IMPORTE", "RETENCIÓN", "VALIDACIÓN"],
                "filas": _filas_para_excel(cruce, COLS_CRUCE) if cruce is not None else [],
                "numericas": [2, 3],
                "total_en": [2],
                "columnas_guion": [3],
                "columna_etiqueta": 1,
                "etiqueta_total": "TOTAL",
            },
            "resguardo_cliente": {
                "titulo": "RESGUARDO CLIENTE",
                "encabezados": ["NOMBRE", "TARJETA", "IMPORTE", "RETENIDO"],
                "filas": _filas_para_excel(resguardo_cliente, COLS_RESGUARDO_CLIENTE),
                "numericas": [2, 3],
                "total_en": [3],
                "columna_etiqueta": 0,
                "etiqueta_total": "TOTAL RESGUARDO CLIENTE",
            },
            "resguardo_convenia": {
                "titulo": "RESGUARDO CONVENIA",
                "encabezados": ["NOMBRE", "IMPORTE", "RETENCIÓN", "EN RESGUARDO"],
                "filas": _filas_para_excel(resguardo_convenia, COLS_RESGUARDO_CONVENIA),
                "numericas": [1, 2, 3],
                "total_en": [3],
                "columnas_guion": [2],
                "columna_etiqueta": 0,
                "etiqueta_total": "TOTAL RESGUARDO CONVENIA",
            },
        })
    return bloques


def bloque_resguardos(datos):
    """Pinta los resguardos de un Excel, si es que tiene.

    Se dibujan aparte de la dispersión a propósito: ese dinero NO se deposita,
    y mezclarlo con la tabla de pagos es justo lo que confunde a la hora de
    cuadrar contra el comprobante.
    """
    resguardo_cliente, resguardo_convenia = tablas_resguardo(datos)

    if not resguardo_cliente.empty:
        retenido = resguardo_cliente["retencion"].sum()
        st.markdown("**RESGUARDO CLIENTE**")
        st.caption(
            f"{len(resguardo_cliente)} renglón(es) con retención o descuento. "
            f"Esos {retenido:,.2f} no se dispersan: se le quedan al cliente."
        )
        st.dataframe(
            con_totales(resguardo_cliente, ("importe", "retencion"),
                        etiqueta="TOTAL RESGUARDO CLIENTE"),
            hide_index=True, **ANCHO_TABLA,
            column_config={
                "nombre": st.column_config.TextColumn("Nombre"),
                "tarjeta": st.column_config.TextColumn("Tarjeta"),
                "importe": st.column_config.NumberColumn("Importe", format="%.2f"),
                "retencion": st.column_config.NumberColumn("Retenido", format="%.2f"),
            },
        )

    if not resguardo_convenia.empty:
        en_resguardo = resguardo_convenia["en_resguardo"].sum()
        st.markdown("**RESGUARDO CONVENIA**")
        st.caption(
            f"{len(resguardo_convenia)} renglón(es) sin número de tarjeta. "
            f"Esos {en_resguardo:,.2f} no se pueden depositar y quedan en resguardo "
            "hasta que haya cuenta a dónde mandarlos."
        )
        st.dataframe(
            con_totales(resguardo_convenia, ("importe", "retencion", "en_resguardo"),
                        etiqueta="TOTAL RESGUARDO CONVENIA"),
            hide_index=True, **ANCHO_TABLA,
            column_config={
                "nombre": st.column_config.TextColumn("Nombre"),
                "importe": st.column_config.NumberColumn("Importe", format="%.2f"),
                "retencion": st.column_config.NumberColumn("Retención", format="%.2f"),
                "en_resguardo": st.column_config.NumberColumn("En resguardo", format="%.2f"),
            },
        )


def seccion_cruce(acumulados):
    """Cruce del acumulado contra la base de tarjetas. Devuelve {excel: tabla}."""
    st.subheader("Cruce con la base de datos de tarjetas")

    if not acumulados or all(tabla.empty for tabla in acumulados.values()):
        st.info("Primero sube los Excel de dispersión para armar el acumulado.")
        return {}

    st.caption(
        "Sube la base general (STOCK LISTADO). Cada tarjeta del acumulado se busca en la "
        "columna **CUENTA** y de ahí se trae el nombre. El importe sale de redondear el "
        "pago final a 2 decimales, y se valida que coincida con el importe original."
    )
    archivo_bd = st.file_uploader(
        "Base de datos de tarjetas (.xls o .xlsx)", type=["xls", "xlsx"], key="base_tarjetas"
    )

    if archivo_bd is None:
        st.info("Sube la base para hacer el cruce.")
        return {}

    try:
        base = leer_base_tarjetas(archivo_bd.getvalue(), archivo_bd.name)
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo leer la base: {exc}")
        return {}

    if not (base and base["por_cuenta"]):
        st.error("No encontré en ese archivo ninguna hoja con columnas CUENTA y NOMBRE COMPLETO.")
        return {}

    st.caption(
        f"Hoja **{base['hoja']}** — {base['filas']:,} cuentas indexadas"
        + (f", {base['duplicadas']} duplicadas ignoradas" if base["duplicadas"] else "")
    )

    cruces = {}
    for nombre_excel, tabla in acumulados.items():
        filas_cruce = []
        for renglon in tabla.to_dict("records"):
            ficha = base["por_cuenta"].get(str(renglon["tarjeta"]))
            pago_final = renglon.get("pago_final")
            importe_original = renglon.get("importe")
            # El importe es el pago final redondeado a centavos, igual
            # que =REDONDEAR(F,2) en el Excel.
            importe = round(pago_final, 2) if pago_final is not None else None
            cuadra = (
                importe_original is not None and importe is not None
                and round(importe_original, 2) == importe
            )
            filas_cruce.append({
                "nombre": ficha["nombre"] if ficha else SIN_FICHA,
                "cuenta": renglon["tarjeta"],
                "importe": importe,
                "retencion": renglon.get("retencion") or 0.0,
                "validacion": "VERDADERO" if cuadra else "FALSO",
                "en_base": bool(ficha),
            })

        cruce = pd.DataFrame(filas_cruce)
        cruces[nombre_excel] = cruce[COLS_CRUCE]

        st.markdown(f"##### {nombre_excel}")
        st.dataframe(
            con_totales(cruce[COLS_CRUCE], ("importe", "retencion"), "TOTAL", "nombre"),
            hide_index=True, **ANCHO_TABLA,
            column_config={
                "nombre": st.column_config.TextColumn("Nombre (base de datos)"),
                "cuenta": st.column_config.TextColumn("Cuenta"),
                "importe": st.column_config.NumberColumn("Importe", format="%.2f"),
                "retencion": st.column_config.NumberColumn("Retención", format="%.2f"),
                "validacion": st.column_config.TextColumn("Pago final = importe"),
            },
        )

        # Las cuentas se sacan de la tabla sin el renglón de totales.
        sin_base = int((~cruce["en_base"]).sum())
        en_falso = int((cruce["validacion"] == "FALSO").sum())
        col_p, col_q, col_r = st.columns(3)
        with col_p:
            st.metric("Renglones", len(cruce))
        with col_q:
            st.metric("Sin match en la base", sin_base)
        with col_r:
            st.metric("Importe ≠ pago final", en_falso)

        if sin_base:
            st.warning(
                f"{sin_base} tarjeta(s) de este archivo no aparecen en la base. Revisa "
                "que la base esté actualizada o la tarjeta bien capturada en el Excel."
            )
        if en_falso:
            st.warning(
                f"{en_falso} renglón(es) de este archivo donde el importe no coincide "
                "con el pago final. Suele ser por una retención aplicada."
            )

        st.download_button(
            "Descargar este cruce en CSV",
            data=cruce[COLS_CRUCE].to_csv(index=False).encode("utf_8_sig"),
            file_name=f"cruce_{os.path.splitext(nombre_excel)[0]}.csv",
            mime="text/csv",
            key=f"csv_{nombre_excel}",
            icon=sio_tema.ICONO["descargar"],
        )

    return cruces


def seccion_excel(records, debug):
    """Todo lo del Excel de dispersión: acumulado, cruce y descarga.

    Corre con o sin comprobantes. El PDF sólo sirve para verificar que los
    totales del Excel cuadren contra lo que de verdad se pagó; no hace falta
    para procesar la dispersión.
    """
    st.subheader("Validación contra el Excel de dispersión")

    if leer_excel_convenia is None:
        st.warning(f"No se pudo cargar `excel_conv.py` ({EXCEL_IMPORT_ERROR}).")
        return

    st.caption(
        "Sube los Excel de dispersión. Si además cargaste comprobantes, se emparejan "
        "por **pagador y monto** para verificar los totales; el comprobante no es "
        "requisito para procesar el Excel."
    )
    excel_files = st.file_uploader(
        "Sube uno o varios Excel de dispersión", type=["xlsx"],
        accept_multiple_files=True, key="excels_dispersion",
    )

    excels = []
    for archivo in excel_files or []:
        try:
            excels.append((archivo.name, leer_excel_convenia(archivo.getvalue())))
        except Exception as exc:  # noqa: BLE001
            st.error(f"'{archivo.name}': no se pudo leer ({exc})")

    if not excels:
        st.info("Sube los Excel de dispersión para armar el acumulado.")
        st.session_state["acumuladoconv"] = pd.DataFrame(columns=COLS_DISPERSION)
        return

    if records:
        emparejados, sueltos = emparejar_excels(excels, records, debug)
        for emparejado in emparejados:
            veredicto = emparejado["veredicto"]
            icono = {
                "CUADRA": sio_tema.ICONO["ok"],
                "CON OBSERVACIONES": sio_tema.ICONO["aviso"],
                "NO CUADRA": sio_tema.ICONO["error"],
            }[veredicto["estado"]]
            with st.expander(
                f"{emparejado['excel']}  ↔  {emparejado['record'].get('archivo','?')}  "
                f"— {veredicto['estado']}",
                expanded=veredicto["estado"] != "CUADRA",
                icon=icono,
            ):
                for linea in veredicto["detalles"]:
                    st.markdown(f"- {linea}")

        for nombre_archivo, _datos, motivo in sueltos:
            st.warning(f"'{nombre_archivo}': sin emparejar — {motivo}")

        st.metric("Excel emparejados", f"{len(emparejados)} de {len(excels)}")
    else:
        emparejados = []
        st.info(
            "Sin comprobantes cargados: el acumulado y el cruce se arman igual. "
            "Lo único que falta es contra qué verificar los totales."
        )

    por_excel = {emparejado["excel"]: emparejado for emparejado in emparejados}

    st.subheader("acumuladoconv")

    # Una tabla por Excel, no todo revuelto: cada operación se revisa por
    # separado y sus totales tienen que cuadrar contra su propio comprobante.
    acumulados = {}
    for nombre_excel, datos in excels:
        filas_acumulado = []
        for renglon in datos["detalle"]:
            if not (renglon.get("pago_final") or 0):
                continue  # empleados sin dispersión este periodo
            filas_acumulado.append({
                "nombre": renglon["nombre"],
                "tarjeta": renglon["tarjeta"],
                "importe": renglon.get("importe"),
                # Sin retención capturada es cero, no un hueco.
                "retencion": renglon.get("retencion") or 0.0,
                "pago_final": renglon.get("pago_final"),
            })
        tabla = pd.DataFrame(filas_acumulado, columns=COLS_DISPERSION)
        acumulados[nombre_excel] = tabla

        emparejado = por_excel.get(nombre_excel)
        st.markdown(f"##### {nombre_excel}")
        st.caption(
            f"{datos.get('cliente', '')} — "
            + (f"comprobante: {emparejado['record'].get('archivo', '?')}"
               if emparejado else "sin comprobante emparejado")
        )
        st.dataframe(
            con_totales(tabla, ("importe", "retencion", "pago_final")),
            hide_index=True, **ANCHO_TABLA,
            column_config={
                "nombre": st.column_config.TextColumn("Nombre"),
                "tarjeta": st.column_config.TextColumn("Tarjeta"),
                "importe": st.column_config.NumberColumn("Importe", format="%.2f"),
                "retencion": st.column_config.NumberColumn("Retención", format="%.2f"),
                "pago_final": st.column_config.NumberColumn("Pago final", format="%.2f"),
            },
        )

        suma = tabla["pago_final"].sum()
        retenido = tabla["retencion"].sum()
        esperado = datos.get("total_dispersion")
        col_x, col_y, col_z = st.columns(3)
        with col_x:
            st.metric("Renglones", len(tabla))
        with col_y:
            st.metric("Suma pago final", f"{suma:,.2f}",
                      delta=f"retención {retenido:,.2f}" if retenido else None,
                      delta_color="off")
        with col_z:
            if esperado is None:
                st.metric("Total dispersión", "—")
            # Con retenciones los pagos finales salen abajo del total a propósito:
            # lo que tiene que cuadrar es pagado + retenido. Sin esto marcaba
            # "faltan 850.00" en rojo, como si se hubiera perdido dinero.
            elif round(suma + retenido, 2) == round(esperado, 2):
                st.metric("Total dispersión", f"{esperado:,.2f}",
                          delta="cuadra" + (" con retención" if retenido else ""))
            else:
                st.metric("Total dispersión", f"{esperado:,.2f}",
                          delta=f"faltan {esperado - suma - retenido:,.2f}",
                          delta_color="inverse")

        bloque_resguardos(datos)

    acumulado = (pd.concat(acumulados.values(), ignore_index=True)
                 if acumulados else pd.DataFrame(columns=COLS_DISPERSION))
    st.session_state["acumuladoconv"] = acumulado

    cruces = seccion_cruce(acumulados)

    # Cada descarga va por su cuenta: si una truena, las otras dos se siguen
    # ofreciendo en vez de que se caiga toda la sección.
    bloque_concentrado(excels, acumulados, cruces)
    bloque_operaciones(excels)
    bloque_carga_masiva(excels, acumulados, cruces)


def bloque_concentrado(excels, acumulados, cruces):
    """Descarga del acumulado por cliente, dispersión y cruce lado a lado."""
    st.subheader("Concentrado en Excel")
    bloques = armar_bloques(excels, acumulados, cruces)
    if not bloques:
        st.info("No hay renglones que exportar.")
        return

    if not cruces:
        st.caption(
            "Se exporta sólo la dispersión: sube la base de tarjetas para que el "
            "concentrado incluya también el cruce del lado derecho."
        )
    try:
        contenido = escribir_acumulado_xlsx(bloques)
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo armar el Excel: {exc}")
        return

    st.download_button(
        f"Descargar concentrado ({len(bloques)} cliente(s))",
        data=contenido,
        file_name=f"acumulado_conv_{date.today():%d%m%Y}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descarga_concentrado",
        icon=sio_tema.ICONO["descargar"],
    )
    st.caption(
        "Un bloque por cliente: la dispersión del Excel a la izquierda y el cruce "
        "contra la base a la derecha, cada tabla con su renglón de totales."
    )


def nombre_archivo_seguro(texto: str) -> str:
    """'CAMIONES BRONCOS DEL NORTE, S.A. DE C.V.' -> 'CAMIONES_BRONCOS_DEL_NORTE_SA_DE_CV'."""
    limpio = unicodedata.normalize("NFKD", str(texto or ""))
    limpio = "".join(c for c in limpio if not unicodedata.combining(c))
    limpio = re.sub(r"[^A-Za-z0-9 ]", "", limpio)
    limpio = re.sub(r"\s+", "_", limpio.strip())
    return limpio[:80] or "cliente"


def filas_de_carga(tabla_dispersion, tabla_cruce):
    """Renglones para el formato del banco: cuenta, importe y beneficiario.

    El importe es el PAGO FINAL, no el importe bruto: si hubo retención, al banco
    se le manda lo que de verdad se deposita. El nombre sale de la base de
    tarjetas cuando está disponible, porque es el que el banco tiene registrado;
    si no, se usa el del Excel de dispersión.
    """
    nombres_base = {}
    if tabla_cruce is not None and not tabla_cruce.empty:
        for renglon in tabla_cruce.to_dict("records"):
            nombre = str(renglon.get("nombre") or "")
            if nombre and nombre != SIN_FICHA:
                nombres_base[str(renglon.get("cuenta"))] = nombre

    filas = []
    for renglon in tabla_dispersion.to_dict("records"):
        pago = renglon.get("pago_final")
        if not pago:
            continue
        cuenta = str(renglon.get("tarjeta") or "").strip()
        filas.append({
            "cuenta": cuenta,
            "importe": round(float(pago), 2),
            "nombre": nombres_base.get(cuenta) or str(renglon.get("nombre") or ""),
        })
    return filas


def bloque_carga_masiva(excels, acumulados, cruces):
    """Un formato de carga masiva por cliente, listo para subir al banco."""
    st.subheader("Formato de carga masiva (uno por cliente)")
    st.caption(
        "Un archivo por cliente con las columnas que pide el banco: cuenta destino, "
        "importe, nombre del beneficiario y concepto. El importe es el **pago final**, "
        "o sea ya con la retención descontada."
    )

    concepto = st.text_input(
        "Concepto (se repite en todos los renglones)",
        value="", placeholder="Déjalo vacío para capturarlo tú en el archivo",
        key="carga_concepto",
    )

    paquete = []
    usados = set()
    for nombre_excel, datos in excels:
        dispersion = acumulados.get(nombre_excel)
        if dispersion is None or dispersion.empty:
            continue
        filas = filas_de_carga(dispersion, cruces.get(nombre_excel))
        if not filas:
            continue
        cliente = datos.get("cliente") or nombre_excel

        # Puede haber dos Excel del mismo cliente en la misma carga (dos
        # dispersiones distintas). Sin desempate, el segundo archivo pisaría al
        # primero dentro del ZIP; se distingue con el nombre del Excel de origen.
        etiqueta = nombre_archivo_seguro(cliente)
        if etiqueta in usados:
            etiqueta = f"{etiqueta}_{nombre_archivo_seguro(os.path.splitext(nombre_excel)[0])}"
            st.caption(
                f"Hay más de un Excel de **{cliente}**; este formato lleva el nombre "
                f"del archivo de origen para no encimarse con el otro."
            )
        usados.add(etiqueta)

        paquete.append({
            "cliente": cliente,
            "excel": nombre_excel,
            "archivo": f"formato_carga_{etiqueta}.xlsx",
            "filas": filas,
            "contenido": escribir_carga_masiva(filas, concepto),
        })

    if not paquete:
        st.info("No hay renglones con pago final para armar los formatos.")
        return

    for indice, item in enumerate(paquete):
        total = sum(f["importe"] for f in item["filas"])
        col_a, col_b = st.columns([3, 1])
        with col_a:
            st.markdown(
                f"**{item['cliente']}** — {len(item['filas'])} transferencia(s) "
                f"por {total:,.2f}  \n`{item['archivo']}`"
            )
        with col_b:
            st.download_button(
                "Descargar", data=item["contenido"], file_name=item["archivo"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"carga_{indice}",
                icon=sio_tema.ICONO["descargar"],
            )

    if len(paquete) > 1:
        memoria = io.BytesIO()
        with zipfile.ZipFile(memoria, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in paquete:
                zf.writestr(item["archivo"], item["contenido"])
        st.download_button(
            f"Descargar los {len(paquete)} formatos en un ZIP",
            data=memoria.getvalue(),
            file_name=f"formatos_carga_{date.today():%d%m%Y}.zip",
            mime="application/zip",
            key="carga_zip",
            icon=sio_tema.ICONO["descargar"],
        )


def operaciones_desde_excels(excels):
    """Un renglón del concentrado por Excel. Devuelve (operaciones, avisos).

    La diferencia entre lo que se reparte y lo que de verdad se dispersa son las
    retenciones del propio Excel:

      Importe a Repartir  = total de dispersión (el bruto, suma de importes)
      REAL A DISPERSAR    = suma de los pagos finales (ya con la retención hecha)
      DEVOLUCIONES        = la resta de los dos, que la calcula Excel

    Sin retenciones capturadas los dos números son el mismo y el renglón queda
    igual que antes; el caso interesante es cuando el Excel sí trae algo.
    """
    operaciones, avisos = [], []
    for nombre_excel, datos in excels:
        detalle = [r for r in datos.get("detalle", []) if (r.get("pago_final") or 0)]
        if not detalle and datos.get("total_dispersion") is None:
            continue

        importes = sum(r.get("importe") or 0 for r in detalle)
        retenciones = sum(r.get("retencion") or 0 for r in detalle)
        pagado = sum(r.get("pago_final") or 0 for r in detalle)

        total = datos.get("total_dispersion")
        a_repartir = round(float(total) if total is not None else importes, 2)
        real = round(pagado, 2) if detalle else a_repartir
        etiqueta = datos.get("cliente") or nombre_excel

        if retenciones:
            avisos.append(
                f"'{etiqueta}': el Excel trae retenciones por {retenciones:,.2f}. "
                f"REAL A DISPERSAR queda en {real:,.2f} y esa diferencia es la que "
                "aparece como DEVOLUCIONES."
            )
        elif abs(a_repartir - real) >= 0.01:
            # Sin retenciones los dos deberían coincidir; si no, algo no cuadra
            # en el Excel y más vale decirlo que escribir un número raro.
            avisos.append(
                f"'{etiqueta}': sin retenciones capturadas, pero el total de dispersión "
                f"({a_repartir:,.2f}) no coincide con la suma de pagos finales "
                f"({real:,.2f}). Revisa ese archivo."
            )

        operaciones.append({
            "operacion": etiqueta,
            "importe_a_repartir": a_repartir,
            "retenciones": round(retenciones, 2),
            "real_a_dispersar": real,
        })
    return operaciones, avisos


def bloque_operaciones(excels):
    """Hoja CR3 CONV: un renglón por operación, con las fórmulas vivas."""
    st.subheader("Concentrado de operaciones (CR3 CONV)")
    st.caption(
        "Un renglón por Excel cargado: **Operación** es el cliente e **Importe a "
        "Repartir** es su total de dispersión. Lo demás son las fórmulas de la "
        "plantilla, que Excel recalcula al abrir el archivo."
    )

    col_a, col_b = st.columns([3, 1])
    with col_a:
        cliente = st.text_input(
            "CLIENTE (encabezado de la hoja)", value="JAIME MOLINA - CONVENIA",
            key="cr3_cliente",
        )
    with col_b:
        ejercicio = st.number_input(
            "EJERCICIO", min_value=2000, max_value=2100, value=date.today().year,
            step=1, key="cr3_ejercicio",
        )

    operaciones, avisos = operaciones_desde_excels(excels)

    for aviso in avisos:
        st.warning(aviso)

    if not operaciones:
        st.info("Ningún Excel trae total de dispersión para armar el concentrado.")
        return

    st.dataframe(
        pd.DataFrame(operaciones)[
            ["operacion", "importe_a_repartir", "retenciones", "real_a_dispersar"]
        ].rename(columns={
            "operacion": "Operación",
            "importe_a_repartir": "Importe a Repartir",
            "retenciones": "Retenciones (→ DEVOLUCIONES)",
            "real_a_dispersar": "REAL A DISPERSAR",
        }),
        hide_index=True, **ANCHO_TABLA,
        column_config={
            "Importe a Repartir": st.column_config.NumberColumn(format="%.2f"),
            "Retenciones (→ DEVOLUCIONES)": st.column_config.NumberColumn(format="%.2f"),
            "REAL A DISPERSAR": st.column_config.NumberColumn(format="%.2f"),
        },
    )

    try:
        contenido = escribir_operaciones_xlsx(operaciones, cliente, int(ejercicio))
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo armar el concentrado de operaciones: {exc}")
        return

    st.download_button(
        f"Descargar CR3 CONV ({len(operaciones)} operación(es))",
        data=contenido,
        file_name=f"cr3_conv_{date.today():%d%m%Y}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descarga_operaciones",
        icon=sio_tema.ICONO["descargar"],
    )


def vista_completa():
    """El proceso completo: extracción, TXT, consulta al CEP y cruce con los Excel.

    Vive en una función, y ya no suelta a nivel de módulo, para poder llamarla
    desde app_cep.py como una pestaña más sin que importar este archivo dispare
    la interfaz."""
    sio_tema.encabezado(
        "Extractor de comprobantes de transferencia para CEP Banxico",
        "Sube uno o varios PDFs (con texto o escaneados) y genera el archivo TXT para "
        "la consulta de CEP por lotes.",
    )

    # El aviso de motores solo aparece si de verdad falta uno; ya no ocupa la barra
    # lateral cuando todo está en orden.
    if not (RapidOCR and np) and not TESSERACT_READY:
        st.warning(
            "No hay motor de OCR instalado: los PDFs escaneados no se van a poder leer. "
            "Instálalo con `pip install rapidocr-onnxruntime`."
        )

    uploaded_files = st.file_uploader(
        "Sube uno o varios archivos PDF", type=["pdf"], accept_multiple_files=True
    )

    if uploaded_files:
        rows, debug = leer_comprobantes(uploaded_files)

        st.subheader("Datos extraídos")
        st.caption(
            "Puedes corregir cualquier celda antes de generar el archivo. "
            "La fecha se captura como DD-MM-YYYY y se convierte a AAAA-MM-DD en el TXT, "
            "que es el formato que pide Banxico."
        )

        records = tabla_editable(rows, "tabla_cep")

        sin_cep = revisar_registros(records)

        txt_content = build_txt(records)
        valid_lines = resumen_lote(txt_content, sin_cep)

        bloque_cep(records, valid_lines, "completo")
    else:
        # El comprobante es el paso de verificación, no el requisito de entrada:
        # sin PDFs se sigue derecho al Excel de dispersión.
        rows, debug, records = [], {}, []
        txt_content, valid_lines = "", 0
        st.info(
            "No hay comprobantes cargados. Puedes seguir con los Excel de dispersión "
            "aquí abajo; el comprobante sólo sirve para verificar que los totales cuadren."
        )

    seccion_excel(records, debug)

    # Todo el material de consulta en un solo expander. Streamlit no permite anidar
    # expanders, así que adentro va en pestañas en vez de uno por sección.
    with st.expander("Vista previa, referencia y diagnóstico"):
        tab_txt, tab_cuentas, tab_campos, tab_diag, tab_bancos = st.tabs(
            ["Vista previa del TXT", "Cuentas conocidas", "Cómo se llenan los campos",
             "Texto leído de cada PDF", "Catálogo de instituciones"]
        )

        with tab_cuentas:
            st.caption(
                "Cuentas que la app da por conocidas. Se usan para reconstruir la CLABE "
                "cuando un banco la imprime enmascarada (`****184`) o incompleta, sin "
                "depender de que en la misma carga venga otro comprobante que la traiga "
                "completa. Edita, agrega o borra renglones y guarda."
            )
            cuentas_editadas = st.data_editor(
                pd.DataFrame(cargar_cuentas_conocidas(), columns=["clabe", "banco", "titular"]),
                hide_index=True, num_rows="dynamic", key="editor_cuentas", **ANCHO_TABLA,
                column_config={
                    "clabe": st.column_config.TextColumn("CLABE (18 dígitos)"),
                    "banco": st.column_config.TextColumn("Banco"),
                    "titular": st.column_config.TextColumn("Titular"),
                },
            )
            col_g, col_h = st.columns([1, 3])
            with col_g:
                if st.button("Guardar cuentas", icon=sio_tema.ICONO["guardar"]):
                    guardar_cuentas_conocidas(cuentas_editadas.fillna("").to_dict("records"))
                    st.success("Guardadas. Vuelve a subir los PDFs para aplicarlas.")
            with col_h:
                st.caption(f"Se guardan en `{ARCHIVO_CUENTAS}`, junto al script.")

        with tab_txt:
            st.download_button(
                "Descargar archivo .txt para Banxico",
                data=txt_content,
                file_name="transferencias_cep.txt",
                mime="text/plain",
                disabled=valid_lines == 0,
                icon=sio_tema.ICONO["descargar"],
            )
            st.code(txt_content or "Aún no hay renglones válidos.", language="text")

        with tab_campos:
            st.markdown(COMO_SE_LLENAN)

        with tab_diag:
            for name, record in debug.items():
                st.markdown(f"**{name}** — origen: `{record['_origen']}`")
                if record.get("_cuenta_ordenante_detectada"):
                    st.caption(f"Cuenta de retiro detectada: {record['_cuenta_ordenante_detectada']}")
                st.text(record["_texto"][:6000] or "(vacío)")

        with tab_bancos:
            st.download_button(
                "Descargar instrucciones + catálogo",
                data=INSTRUCCIONES + render_bank_reference(),
                file_name="instrucciones_banxico_cep.txt",
                mime="text/plain",
                icon=sio_tema.ICONO["descargar"],
            )
            st.code(render_bank_reference(), language="text")


if __name__ == "__main__":
    sio_tema.config_pagina("Extractor de CEP")
    sio_tema.aplicar()
    vista_completa()
    sio_tema.pie()
