"""Consulta automatizada del CEP en el portal de Banxico.

Toma los pagos del .TXT que genera cr3.py (o de argumentos sueltos), los captura
en https://www.banxico.org.mx/cep/ y reporta si el pago aparece liquidado.

    python cep_banxico.py --txt transferencias_cep.txt
    python cep_banxico.py --fecha 12-08-2026 --rastreo 0026... --emisor 40012 \
                          --receptor 40072 --cuenta 0726... --monto 55540.80

Notas:
  - El portal solo atiende de 09:30 a 23:00 hrs.
  - Chromium corre siempre headless: no hace falta escritorio, DISPLAY ni VNC.
  - El formulario tiene un captcha que normalmente viene oculto. Si Banxico lo
    activa, el script NO lo intenta resolver ni espera a nadie: marca ese pago
    como CAPTCHA, tira la sesión, abre una limpia y sigue con los que faltan.
    El lote siempre termina escribiendo los resultados que alcanzó a juntar.
"""

import argparse
import csv
import json
import logging
import os
import platform
import re
import sys
import tempfile
import time
from dataclasses import dataclass, asdict, field
from datetime import datetime
from typing import Callable, List, Optional

from playwright.sync_api import sync_playwright, Page

CEP_URL = "https://www.banxico.org.mx/cep/"
CHROME_WIN = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Espera entre consultas. El portal es un servicio público y gratuito; no hay
# razón para atropellarlo, y disparar el captcha solo nos deja peor.
PAUSA_ENTRE_CONSULTAS = 4.0

# Cuando sale el captcha, la sesión ya quedó marcada: se tira y se abre otra
# limpia. Antes de volver a tocar el portal se espera cada vez más, para no
# insistirle a un servicio que justo nos está pidiendo calma.
PAUSA_TRAS_CAPTCHA = (20.0, 45.0, 90.0)

# Si ni con sesiones nuevas se quita, dejamos de consultar: los pagos que
# faltan se anotan sin pedirle nada más al portal. Así el lote termina en un
# tiempo acotado (importante: cr3.py lo corre con timeout) en vez de arrastrar
# una pausa larga por cada pago restante.
MAX_CAPTCHAS_SEGUIDOS = 3

# Dos estados distintos a propósito. CAPTCHA es un hecho sobre ese pago: se
# mandó al portal y Banxico pidió el código. OMITIDO_CAPTCHA no dice nada del
# pago —nunca se consultó—, solo que la tanda venía trabada. Mezclarlos haría
# creer que Banxico se quejó de pagos que jamás vio.
CAPTCHA = "CAPTCHA"
OMITIDO_CAPTCHA = "OMITIDO_CAPTCHA"

# Marca de las líneas de avance que cr3.py lee de stdout mientras esto corre.
# El log normal se va por stderr, así que stdout queda casi limpio; aun así el
# prefijo es lo que distingue el avance de cualquier otra cosa que se imprima.
PREFIJO_PROGRESO = "CEP_PROGRESS"

MSG_CAPTCHA = ("Banxico solicitó CAPTCHA. La consulta de este pago se omitió y "
               "el proceso continuó.")
MSG_OMITIDO = ("No se intentó consultar este pago porque Banxico solicitó CAPTCHA "
               "en varias consultas consecutivas.")


@dataclass
class Pago:
    fecha: str            # DD-MM-YYYY (como lo pide el formulario)
    clave_rastreo: str
    clave_emisora: str
    clave_receptora: str
    cuenta: str
    monto: str
    etiqueta: str = ""
    pago_a_banco: bool = False   # marca "Pago a Banco" -> la cuenta es la ordenante
    criterio: str = "T"          # T = clave de rastreo, R = número de referencia


@dataclass
class Resultado:
    etiqueta: str = ""
    clave_rastreo: str = ""
    estado: str = ""            # LIQUIDADO / NO_ENCONTRADO / CAPTCHA /
                                # OMITIDO_CAPTCHA / ERROR
    estado_portal: str = ""     # lo que dice Banxico: "Liquidado", "Devuelto", ...
    recepcion: str = ""         # fecha y hora en que SPEI recibió el pago
    procesamiento: str = ""     # fecha y hora en que SPEI lo procesó
    monto_portal: str = ""
    detalle: str = ""
    mensaje_portal: str = ""
    consultado_en: str = ""


def setup_logging(verbose: bool = False):
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )


# --------------------------------------------------------------------------- #
# Entrada
# --------------------------------------------------------------------------- #

def iso_a_ddmmyyyy(valor: str) -> str:
    valor = (valor or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", valor):
        year, month, day = valor.split("-")
        return f"{day}-{month}-{year}"
    return valor


def pagos_desde_txt(ruta: str) -> List[Pago]:
    """Lee el TXT de Banxico: fecha,rastreo,emisora,receptora,cuenta,monto."""
    pagos = []
    with open(ruta, encoding="utf-8") as handle:
        for numero, linea in enumerate(handle, start=1):
            linea = linea.strip()
            if not linea:
                continue
            partes = [p.strip() for p in linea.split(",")]
            if len(partes) != 6:
                logging.warning(f"Línea {numero} ignorada, esperaba 6 campos y tiene {len(partes)}.")
                continue
            fecha, rastreo, emisora, receptora, cuenta, monto = partes
            pagos.append(Pago(
                fecha=iso_a_ddmmyyyy(fecha),
                clave_rastreo=rastreo,
                clave_emisora=emisora,
                clave_receptora=receptora,
                cuenta=cuenta,
                monto=monto,
                etiqueta=f"línea {numero}",
            ))
    return pagos


# --------------------------------------------------------------------------- #
# Portal
# --------------------------------------------------------------------------- #

MESES_JQUERY = {  # el datepicker usa mes 0-indexado
    "01": "0", "02": "1", "03": "2", "04": "3", "05": "4", "06": "5",
    "07": "6", "08": "7", "09": "8", "10": "9", "11": "10", "12": "11",
}


def _elegir_fecha(page: Page, fecha_ddmmyyyy: str) -> None:
    """El input de fecha es readonly: hay que pasar por el datepicker de jQuery UI.

    Escribir el value por JS no sirve — el formulario no marca el campo como
    válido y el botón Consultar se queda deshabilitado.
    """
    dia, mes, anio = fecha_ddmmyyyy.split("-")
    page.click("#input_fecha")
    page.wait_for_selector("#ui-datepicker-div", state="visible", timeout=15000)

    page.select_option("#ui-datepicker-div .ui-datepicker-year", anio)
    page.wait_for_timeout(300)
    page.select_option("#ui-datepicker-div .ui-datepicker-month", MESES_JQUERY[mes])
    page.wait_for_timeout(300)

    page.click(f"#ui-datepicker-div a.ui-state-default >> text=/^{int(dia)}$/")
    page.wait_for_timeout(500)

    puesta = page.input_value("#input_fecha")
    if puesta != fecha_ddmmyyyy:
        raise RuntimeError(f"El portal quedó con la fecha {puesta!r}, esperaba {fecha_ddmmyyyy!r}.")


def _teclear(page: Page, selector: str, texto: str) -> None:
    """Teclea carácter por carácter.

    Los campos del portal cuelgan validaciones de onkeyup (el monto se
    autoformatea ahí). page.fill() asigna el value de golpe y esos handlers
    nunca corren, así que el portal recibe el campo a medio inicializar.
    """
    campo = page.locator(selector)
    campo.click()
    campo.fill("")
    try:
        campo.press_sequentially(texto, delay=60)
    except AttributeError:  # Playwright viejo
        campo.type(texto, delay=60)
    page.wait_for_timeout(200)


def _llenar_formulario(page: Page, pago: Pago) -> None:
    _elegir_fecha(page, pago.fecha)
    # El portal busca por clave de rastreo (T) o por número de referencia (R).
    # Mandar una referencia de 6 dígitos como clave de rastreo no encuentra nada.
    page.select_option("#input_tipoCriterio", pago.criterio or "T")
    _teclear(page, "#input_criterio", pago.clave_rastreo)
    page.select_option("#input_emisor", pago.clave_emisora)
    page.select_option("#input_receptor", pago.clave_receptora)

    # Desmarcado -> el campo pide "Cuenta beneficiaria".
    # Marcado ("Pago a Banco") -> pide "Cuenta Ordenante".
    casilla = page.locator("#input_benef_es_part")
    if pago.pago_a_banco:
        casilla.check()
    else:
        casilla.uncheck()
    page.wait_for_timeout(300)

    _teclear(page, "#input_cuenta", pago.cuenta)
    _teclear(page, "#input_monto", pago.monto)
    page.wait_for_timeout(500)
    _esperar_overlay(page)


def _esperar_overlay(page: Page, timeout_ms: int = 30000) -> None:
    """Al capturar cuenta y monto el portal lanza una validación con spinner.

    Ese spinner (#divValidacionPertenencia) queda encima del botón Consultar e
    intercepta el clic, así que hay que esperar a que se vaya.
    """
    limite = time.time() + timeout_ms / 1000
    while time.time() < limite:
        tapado = page.evaluate("""() => {
            const el = document.querySelector('#divValidacionPertenencia');
            if (!el) return false;
            const cs = getComputedStyle(el);
            return cs.display !== 'none' && cs.visibility !== 'hidden' && el.offsetParent !== null;
        }""")
        if not tapado:
            return
        page.wait_for_timeout(300)
    logging.warning("El overlay de validación no se quitó; intento el clic de todos modos.")


def _captcha_visible(page: Page) -> bool:
    try:
        return page.locator("#captchaInput").is_visible()
    except Exception:
        return False


def _resultado_captcha(pago: Pago, detalle: str, estado: str = CAPTCHA) -> "Resultado":
    return Resultado(
        etiqueta=pago.etiqueta or pago.clave_rastreo,
        clave_rastreo=pago.clave_rastreo,
        estado=estado,
        detalle=detalle,
        consultado_en=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


def _avisar_avance(hechos: int, total: int, estado: str) -> None:
    """Escupe una línea de avance por stdout: CEP_PROGRESS|3|10|LIQUIDADO.

    Va con flush porque cuando cr3.py lo corre, stdout es un pipe y Python lo
    almacena en bloques: sin esto el avance llegaría de golpe al final, que es
    justo lo que no queremos.
    """
    try:
        print(f"{PREFIJO_PROGRESO}|{hechos}|{total}|{estado}", flush=True)
    except Exception:  # noqa: BLE001
        # Un stdout cerrado o roto no es motivo para tumbar el lote.
        pass


def guardar_json(resultados: List[Resultado], ruta: str) -> None:
    """Escribe el JSON de resultados sin dejarlo nunca a medias.

    Se escribe en un temporal de la MISMA carpeta, se cierra —con fsync, para
    que el contenido esté en disco y no solo en el buffer del sistema— y apenas
    entonces se reemplaza el definitivo con os.replace(), que es atómico. Quien
    lea `ruta` ve el archivo entero de antes o el entero de ahora; nunca uno
    truncado. Importa porque esto se llama después de cada pago y a este proceso
    lo puede matar un timeout justo a media escritura.
    """
    carpeta = os.path.dirname(os.path.abspath(ruta))
    temporal = None
    try:
        with tempfile.NamedTemporaryFile(
            "w", encoding="utf-8", dir=carpeta, prefix=".resultados_",
            suffix=".tmp", delete=False,
        ) as handle:
            temporal = handle.name
            json.dump([asdict(r) for r in resultados], handle, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporal, ruta)
        temporal = None
    finally:
        # Si algo falló antes del replace, no dejamos basura en la carpeta.
        if temporal and os.path.exists(temporal):
            try:
                os.remove(temporal)
            except OSError:
                pass


def _opciones_navegador() -> dict:
    """Cómo se lanza Chromium. Siempre headless, aquí y en el servidor.

    En Windows aprovecha el Chrome instalado si está; en Linux usa el Chromium
    que baja Playwright. Ni uno ni otro necesitan escritorio: nada de DISPLAY,
    X Server ni VNC.
    """
    opciones = {"headless": True}
    if platform.system() == "Windows" and os.path.exists(CHROME_WIN):
        opciones["executable_path"] = CHROME_WIN
        return opciones

    # En contenedores /dev/shm suele venir en 64 MB y Chromium se cae solo;
    # y corriendo como root el sandbox se niega a arrancar.
    args = ["--disable-dev-shm-usage"]
    if hasattr(os, "geteuid") and os.geteuid() == 0:
        args.append("--no-sandbox")
    opciones["args"] = args
    return opciones


class _Sesion:
    """Navegador, contexto y página del lote, con cierre y renovación limpios.

    Se abre sola la primera vez que se le pide la página, así que `renovar()`
    solo tiene que cerrar: si el lote ya se acabó no se levanta un Chromium de
    más, y si abrir falla el error cae en el pago que toca, no en todo el lote.
    """

    def __init__(self, playwright):
        self._playwright = playwright
        self.browser = None
        self.contexto = None
        self.page = None

    def pagina(self) -> Page:
        if self.page is None or self.page.is_closed():
            self.abrir()
        return self.page

    def abrir(self) -> Page:
        self.cerrar()
        try:
            self.browser = self._playwright.chromium.launch(**_opciones_navegador())
            self.contexto = self.browser.new_context(accept_downloads=True)
            self.page = self.contexto.new_page()
        except Exception:
            # Si se cayó a media construcción no dejamos el Chromium colgado.
            self.cerrar()
            raise
        return self.page

    def cerrar(self) -> None:
        """Cierra página, contexto y navegador, en ese orden y sin quejarse.

        Va uno por uno —y no solo browser.close()— para que un objeto ya muerto
        no impida cerrar los demás, que es justo como se fugan los Chromium.
        """
        for nombre in ("page", "contexto", "browser"):
            objeto = getattr(self, nombre)
            if objeto is None:
                continue
            try:
                objeto.close()
            except Exception as exc:  # noqa: BLE001
                logging.debug(f"No se pudo cerrar {nombre}: {exc}")
            finally:
                setattr(self, nombre, None)

    def renovar(self) -> None:
        logging.info("    Sesión descartada; la siguiente consulta abre una limpia.")
        self.cerrar()


def _leer_resultado(page: Page, pago: Pago) -> Resultado:
    resultado = Resultado(
        etiqueta=pago.etiqueta,
        clave_rastreo=pago.clave_rastreo,
        consultado_en=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    if _captcha_visible(page):
        resultado.estado = CAPTCHA
        resultado.detalle = MSG_CAPTCHA
        return resultado

    # El mensaje del portal vive en #consultaMISPEI / div.cuerpo-msg.
    aviso = page.evaluate("""() => {
        const el = document.querySelector('#consultaMISPEI .cuerpo-msg')
               || document.querySelector('#consultaMISPEI')
               || document.querySelector('.cuerpo-msg');
        return el ? el.innerText.trim().replace(/\\s+/g, ' ') : '';
    }""") or ""
    resultado.mensaje_portal = aviso[:600]

    if re.search(r"no encontrada|no ha recibido", aviso, re.I):
        resultado.estado = "NO_ENCONTRADO"
        resultado.detalle = "El SPEI no tiene una orden de pago con esos datos."
        return resultado

    # Cuando lo encuentra, el portal devuelve el detalle en pares etiqueta/valor.
    def campo(etiqueta: str, patron: str) -> str:
        hallazgo = re.search(etiqueta + r"\s*" + patron, aviso, re.I)
        return " ".join(hallazgo.group(1).split()) if hallazgo else ""

    resultado.estado_portal = campo(r"Estado del pago en Banxico", r"(\w+)")
    resultado.recepcion = campo(r"Fecha y hora de recepci[óo]n", r"([\d/]+ [\d:]+)")
    resultado.procesamiento = campo(r"Fecha y hora de procesamiento", r"([\d/]+ [\d:]+)")
    resultado.monto_portal = campo(r"Monto", r"([\d,]+\.\d{2})")

    if re.search(r"\bLiquidado\b", resultado.estado_portal, re.I):
        resultado.estado = "LIQUIDADO"
        resultado.detalle = f"Liquidado en SPEI el {resultado.recepcion}."
    elif resultado.estado_portal:
        resultado.estado = resultado.estado_portal.upper()
        resultado.detalle = f"Banxico reporta el pago como '{resultado.estado_portal}'."
    else:
        resultado.estado = "NO_ENCONTRADO"
        resultado.detalle = aviso[:200] or "El portal no devolvió detalle del pago."
    return resultado


def consultar(pagos: List[Pago], descargar_en: Optional[str] = None,
              resultados: Optional[List[Resultado]] = None,
              al_anotar: Optional[Callable[[List[Resultado]], None]] = None) -> List[Resultado]:
    """Consulta el lote completo. Ningún pago suelto tumba a los demás.

    Un captcha, un error del portal o una página caída se anotan en el pago que
    toca y el `for` sigue. `resultados` se puede pasar desde fuera para que
    quien llama conserve lo procesado aunque esto reviente a media tanda.

    `al_anotar` se llama con la lista completa cada vez que se cierra un pago;
    main() la usa para dejar el JSON al día pago por pago, en vez de esperar al
    final. Si falla, se avisa y el lote sigue: no poder guardar el avance no es
    motivo para dejar de consultar.
    """
    resultados = [] if resultados is None else resultados
    total = len(pagos)
    captchas_seguidos = 0

    procesados = 0

    def anotar(resultado: Resultado) -> None:
        """Único punto por donde entra un resultado.

        Aquí se cierra el pago: se guarda en la lista, se persiste el parcial y
        recién entonces se avisa del avance. En ese orden, para que quien vea el
        contador ya pueda encontrar ese resultado en el JSON.
        """
        nonlocal procesados
        resultados.append(resultado)
        procesados += 1
        if al_anotar is not None:
            try:
                al_anotar(resultados)
            except Exception as exc:  # noqa: BLE001
                logging.warning(f"    No se pudo guardar el avance: {exc}")
        _avisar_avance(procesados, total, resultado.estado)

    with sync_playwright() as p:
        sesion = _Sesion(p)
        try:
            for indice, pago in enumerate(pagos, start=1):
                etiqueta = pago.etiqueta or pago.clave_rastreo
                logging.info(f"=== [{indice}/{total}] {etiqueta} — {pago.clave_rastreo} ===")

                # Ya se probó con sesiones nuevas y el captcha sigue: lo que
                # falta se anota sin volver a tocar el portal. Va como
                # OMITIDO_CAPTCHA, no como CAPTCHA: a este pago nunca se le
                # preguntó nada a Banxico.
                if captchas_seguidos >= MAX_CAPTCHAS_SEGUIDOS:
                    anotar(_resultado_captcha(pago, MSG_OMITIDO, OMITIDO_CAPTCHA))
                    logging.warning(f"    -> {OMITIDO_CAPTCHA}: {MSG_OMITIDO}")
                    continue

                # "ok" / "captcha" / "error": de esto dependen la pausa que
                # sigue y si la sesión se tira.
                desenlace = "ok"
                try:
                    page = sesion.pagina()
                    page.goto(CEP_URL, timeout=60000)
                    page.wait_for_selector("#input_fecha", timeout=30000)
                    page.wait_for_timeout(1500)

                    if _captcha_visible(page):
                        # Sale ya en la carga: ni se llena el formulario.
                        anotar(_resultado_captcha(pago, MSG_CAPTCHA))
                        desenlace = "captcha"
                        logging.warning(f"    -> CAPTCHA: {MSG_CAPTCHA}")
                    else:
                        _llenar_formulario(page, pago)

                        clase_boton = page.locator("#btn_Consultar").get_attribute("class") or ""
                        if "disabled" in clase_boton:
                            raise RuntimeError(
                                "El botón Consultar sigue deshabilitado: algún campo no pasó "
                                "la validación del portal."
                            )

                        _esperar_overlay(page)
                        page.click("#btn_Consultar")
                        page.wait_for_timeout(7000)
                        _esperar_overlay(page)

                        resultado = _leer_resultado(page, pago)
                        anotar(resultado)
                        if resultado.estado == CAPTCHA:
                            desenlace = "captcha"
                        logging.info(f"    -> {resultado.estado}: {resultado.detalle}")

                        if resultado.estado == "LIQUIDADO" and descargar_en:
                            try:
                                with page.expect_download(timeout=30000) as espera:
                                    page.click("#btn_Descargar")
                                descarga = espera.value
                                os.makedirs(descargar_en, exist_ok=True)
                                destino = os.path.join(
                                    descargar_en, f"CEP_{pago.clave_rastreo}.pdf")
                                descarga.save_as(destino)
                                logging.info(f"    CEP guardado en {destino}")
                            except Exception as exc:  # noqa: BLE001
                                logging.warning(f"    No se pudo descargar el CEP: {exc}")

                except Exception as exc:  # noqa: BLE001
                    desenlace = "error"
                    logging.error(f"    Falló la consulta: {exc}")
                    anotar(Resultado(
                        etiqueta=etiqueta, clave_rastreo=pago.clave_rastreo,
                        estado="ERROR", detalle=str(exc)[:300],
                        consultado_en=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    ))

                if desenlace == "captcha":
                    captchas_seguidos += 1
                    espera = PAUSA_TRAS_CAPTCHA[
                        min(captchas_seguidos, len(PAUSA_TRAS_CAPTCHA)) - 1]
                else:
                    # El contador solo se reinicia con una consulta que sí salió:
                    # un error no prueba que el captcha se haya quitado, y si no
                    # fuera así una tanda de captchas y errores alternados nunca
                    # llegaría al tope.
                    if desenlace == "ok":
                        captchas_seguidos = 0
                    espera = PAUSA_ENTRE_CONSULTAS

                if desenlace != "ok":
                    # El captcha se le pega a la sesión, y un error pudo dejar la
                    # página a medias: en los dos casos el siguiente pago arranca
                    # con una sesión limpia.
                    sesion.renovar()

                if indice < total:
                    time.sleep(espera)
        finally:
            # Pase lo que pase, no se queda ningún Chromium suelto.
            sesion.cerrar()

    return resultados


def guardar_csv(resultados: List[Resultado], ruta: str) -> None:
    with open(ruta, "w", newline="", encoding="utf_8_sig") as handle:
        escritor = csv.DictWriter(handle, fieldnames=list(asdict(Resultado()).keys()))
        escritor.writeheader()
        for resultado in resultados:
            escritor.writerow(asdict(resultado))
    logging.info(f"Resultados en {ruta}")


def main():
    parser = argparse.ArgumentParser(description="Consulta el CEP de Banxico para uno o varios pagos.")
    parser.add_argument("--txt", help="TXT generado por cr3.py (fecha,rastreo,emisora,receptora,cuenta,monto).")
    parser.add_argument("--json-in", dest="json_in",
                        help="JSON con la lista de pagos (lo usa cr3.py para llamar a este script).")
    parser.add_argument("--json-out", dest="json_out",
                        help="Escribe los resultados como JSON en esta ruta.")
    parser.add_argument("--fecha", help="DD-MM-YYYY o AAAA-MM-DD.")
    parser.add_argument("--rastreo")
    parser.add_argument("--emisor", help="Clave de 5 dígitos, ej. 40012.")
    parser.add_argument("--receptor", help="Clave de 5 dígitos, ej. 40072.")
    parser.add_argument("--cuenta")
    parser.add_argument("--monto")
    parser.add_argument("--pago-a-banco", action="store_true",
                        help="Marca 'Pago a Banco': la cuenta pasa a ser la ordenante.")
    # Se sigue aceptando para no romper llamadas viejas, pero ya no hace nada:
    # la consulta corre headless siempre, aquí y en el servidor.
    parser.add_argument("--visible", action="store_true",
                        help="Obsoleto: se ignora, la consulta siempre corre headless.")
    parser.add_argument("--descargar-en", help="Carpeta donde guardar los CEP en PDF.")
    parser.add_argument("-o", "--output", default="cep_resultados.csv")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    setup_logging(args.verbose)

    if args.json_in:
        with open(args.json_in, encoding="utf-8") as handle:
            pagos = [Pago(**dato) for dato in json.load(handle)]
    elif args.txt:
        pagos = pagos_desde_txt(args.txt)
    elif all([args.fecha, args.rastreo, args.emisor, args.receptor, args.cuenta, args.monto]):
        pagos = [Pago(
            fecha=iso_a_ddmmyyyy(args.fecha), clave_rastreo=args.rastreo,
            clave_emisora=args.emisor, clave_receptora=args.receptor,
            cuenta=args.cuenta, monto=args.monto,
            etiqueta="manual", pago_a_banco=args.pago_a_banco,
        )]
    else:
        parser.error("Pasa --txt o los seis campos del pago.")

    if not pagos:
        logging.error("No hay pagos que consultar.")
        sys.exit(1)

    ahora = datetime.now().time()
    if not (ahora.replace(hour=9, minute=30) <= ahora <= ahora.replace(hour=23, minute=0)):
        logging.warning("El portal atiende de 09:30 a 23:00; fuera de ese horario suele fallar.")

    if args.visible:
        logging.warning("--visible ya no hace nada: la consulta siempre corre headless.")

    logging.info(f"{len(pagos)} pago(s) por consultar.")

    # El JSON se reescribe entero después de cada pago, no al final: si a este
    # proceso lo matan por timeout o se cae de golpe, cr3.py todavía encuentra
    # todo lo que ya se había consultado. El `finally` de abajo es la red de
    # abajo (y el único que escribe el CSV), no la única escritura.
    resultados: List[Resultado] = []
    al_anotar = None
    if args.json_out:
        def al_anotar(parciales: List[Resultado]) -> None:
            guardar_json(parciales, args.json_out)

    fallo: Optional[BaseException] = None
    try:
        consultar(pagos, descargar_en=args.descargar_en, resultados=resultados,
                  al_anotar=al_anotar)
    except Exception as exc:  # noqa: BLE001
        fallo = exc
        logging.exception("El lote se interrumpió; guardo lo que alcancé a consultar.")
    finally:
        guardar_csv(resultados, args.output)
        if args.json_out:
            guardar_json(resultados, args.json_out)
            logging.info(f"Resultados en {args.json_out}")

    print("\n" + "=" * 78)
    for resultado in resultados:
        print(f"{resultado.estado:<14} {resultado.clave_rastreo:<26} {resultado.detalle}")
    print("=" * 78)

    if fallo is not None:
        sys.exit(1)


if __name__ == "__main__":
    main()
