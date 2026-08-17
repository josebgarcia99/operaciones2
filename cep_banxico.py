"""Consulta automatizada del CEP en el portal de Banxico.

Toma los pagos del .TXT que genera cr3.py (o de argumentos sueltos), los captura
en https://www.banxico.org.mx/cep/ y reporta si el pago aparece liquidado.

    python cep_banxico.py --txt transferencias_cep.txt
    python cep_banxico.py --fecha 12-08-2026 --rastreo 0026... --emisor 40012 \
                          --receptor 40072 --cuenta 0726... --monto 55540.80

Notas:
  - El portal solo atiende de 09:30 a 23:00 hrs.
  - El formulario tiene un captcha que normalmente viene oculto. Si Banxico lo
    activa, el script NO lo intenta resolver: se detiene y te dice que corras
    con --visible para capturarlo tú.
"""

import argparse
import csv
import json
import logging
import os
import platform
import re
import sys
import time
from dataclasses import dataclass, asdict, field
from datetime import datetime
from typing import List, Optional

from playwright.sync_api import sync_playwright, Page

CEP_URL = "https://www.banxico.org.mx/cep/"
CHROME_WIN = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Espera entre consultas. El portal es un servicio público y gratuito; no hay
# razón para atropellarlo, y disparar el captcha solo nos deja peor.
PAUSA_ENTRE_CONSULTAS = 4.0


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
    estado: str = ""            # LIQUIDADO / NO_ENCONTRADO / CAPTCHA / ERROR
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


def _leer_resultado(page: Page, pago: Pago) -> Resultado:
    resultado = Resultado(
        etiqueta=pago.etiqueta,
        clave_rastreo=pago.clave_rastreo,
        consultado_en=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    if _captcha_visible(page):
        resultado.estado = "CAPTCHA"
        resultado.detalle = "El portal pidió código de seguridad. Corre con --visible y captúralo tú."
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


def consultar(pagos: List[Pago], visible: bool = False,
              descargar_en: Optional[str] = None) -> List[Resultado]:
    resultados: List[Resultado] = []

    with sync_playwright() as p:
        if platform.system() == "Windows" and os.path.exists(CHROME_WIN):
            browser = p.chromium.launch(executable_path=CHROME_WIN, headless=not visible)
        else:
            browser = p.chromium.launch(headless=not visible)
        contexto = browser.new_context(accept_downloads=True)
        page = contexto.new_page()

        try:
            for indice, pago in enumerate(pagos, start=1):
                etiqueta = pago.etiqueta or pago.clave_rastreo
                logging.info(f"=== [{indice}/{len(pagos)}] {etiqueta} — {pago.clave_rastreo} ===")
                try:
                    page.goto(CEP_URL, timeout=60000)
                    page.wait_for_selector("#input_fecha", timeout=30000)
                    page.wait_for_timeout(1500)

                    if _captcha_visible(page) and not visible:
                        resultados.append(Resultado(
                            etiqueta=etiqueta, clave_rastreo=pago.clave_rastreo,
                            estado="CAPTCHA",
                            detalle="Captcha activo al cargar. Corre con --visible y captúralo tú.",
                            consultado_en=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        ))
                        logging.warning("Captcha activo; me detengo aquí.")
                        break

                    _llenar_formulario(page, pago)

                    clase_boton = page.locator("#btn_Consultar").get_attribute("class") or ""
                    if "disabled" in clase_boton:
                        raise RuntimeError(
                            "El botón Consultar sigue deshabilitado: algún campo no pasó "
                            "la validación del portal."
                        )

                    if visible and _captcha_visible(page):
                        logging.warning("Captcha visible: captúralo en la ventana. Espero 60 s.")
                        page.wait_for_timeout(60000)

                    _esperar_overlay(page)
                    page.click("#btn_Consultar")
                    page.wait_for_timeout(7000)
                    _esperar_overlay(page)

                    resultado = _leer_resultado(page, pago)
                    resultados.append(resultado)
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

                    if resultado.estado == "CAPTCHA":
                        logging.warning("Captcha activado; detengo el resto.")
                        break

                except Exception as exc:  # noqa: BLE001
                    logging.error(f"    Falló la consulta: {exc}")
                    resultados.append(Resultado(
                        etiqueta=etiqueta, clave_rastreo=pago.clave_rastreo,
                        estado="ERROR", detalle=str(exc)[:300],
                        consultado_en=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    ))

                if indice < len(pagos):
                    time.sleep(PAUSA_ENTRE_CONSULTAS)
        finally:
            contexto.close()
            browser.close()

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
    parser.add_argument("--visible", action="store_true",
                        help="Abre Chrome a la vista (necesario si aparece el captcha).")
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

    logging.info(f"{len(pagos)} pago(s) por consultar.")
    resultados = consultar(pagos, visible=args.visible, descargar_en=args.descargar_en)
    guardar_csv(resultados, args.output)

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as handle:
            json.dump([asdict(r) for r in resultados], handle, ensure_ascii=False)
        logging.info(f"Resultados en {args.json_out}")

    print("\n" + "=" * 78)
    for resultado in resultados:
        print(f"{resultado.estado:<14} {resultado.clave_rastreo:<26} {resultado.detalle}")
    print("=" * 78)


if __name__ == "__main__":
    main()
