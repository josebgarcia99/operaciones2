"""Vista y validaciones del archivo de Litografía."""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from io import BytesIO
import re
import unicodedata

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, PatternFill


@dataclass
class ImporteDetectado:
    valor: Decimal | None
    referencia: str


@dataclass
class FilaResumen:
    empresa: str
    importe: ImporteDetectado
    celda_empresa: object


def _texto_normalizado(valor: object) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    return re.sub(r"[^a-z0-9]+", " ", texto.lower()).strip()


def _decimal(valor: object) -> Decimal | None:
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float, Decimal)):
        try:
            return Decimal(str(valor))
        except InvalidOperation:
            return None

    texto = str(valor).strip().replace("$", "").replace(" ", "")
    if not texto or texto.startswith("="):
        return None
    if "," in texto and "." in texto:
        texto = texto.replace(",", "") if texto.rfind(".") > texto.rfind(",") else texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        partes = texto.split(",")
        texto = "".join(partes) if len(partes[-1]) == 3 else texto.replace(",", ".")
    try:
        return Decimal(texto)
    except InvalidOperation:
        return None


def _es_suma(valor: object) -> bool:
    return _texto_normalizado(valor) == "suma"


def _valor_cercano(hoja, fila: int, columna: int) -> ImporteDetectado:
    # Primero se buscan las posiciones usuales de un total: derecha y abajo.
    desplazamientos = [(0, paso) for paso in range(1, 6)]
    desplazamientos += [(paso, 0) for paso in range(1, 6)]
    desplazamientos += [(0, -paso) for paso in range(1, 4)]
    for delta_fila, delta_columna in desplazamientos:
        nueva_fila, nueva_columna = fila + delta_fila, columna + delta_columna
        if nueva_fila < 1 or nueva_columna < 1:
            continue
        celda = hoja.cell(nueva_fila, nueva_columna)
        valor = _decimal(celda.value)
        if valor is not None:
            return ImporteDetectado(valor, f"{hoja.title}!{celda.coordinate}")
    return ImporteDetectado(None, f"{hoja.title}!{hoja.cell(fila, columna).coordinate}")


def _suma_de_hoja(hoja) -> ImporteDetectado:
    candidatos: list[ImporteDetectado] = []
    for fila in hoja.iter_rows():
        for celda in fila:
            if not isinstance(celda, MergedCell) and _es_suma(celda.value):
                candidatos.append(_valor_cercano(hoja, celda.row, celda.column))
    for candidato in candidatos:
        if candidato.valor is not None:
            return candidato
    return candidatos[0] if candidatos else ImporteDetectado(None, "No se encontró la celda Suma")


def _coincidencia_hoja(valor: object, nombre_hoja: str) -> float:
    izquierda = _texto_normalizado(valor)
    derecha = _texto_normalizado(nombre_hoja)
    if not izquierda or not derecha:
        return 0.0
    if izquierda == derecha:
        return 1.0
    if izquierda in derecha or derecha in izquierda:
        return 0.9
    return SequenceMatcher(None, izquierda, derecha).ratio()


def _filas_del_resumen(hoja_resumen) -> list[FilaResumen]:
    candidatos = [
        celda
        for fila in hoja_resumen.iter_rows()
        for celda in fila
        if not isinstance(celda, MergedCell) and _es_suma(celda.value)
    ]
    if not candidatos:
        return []

    # La celda "Suma" que funciona como encabezado es la que tiene más importes
    # debajo. Así se distingue de la fila final que también se llama "Suma".
    def puntuacion_encabezado(celda) -> int:
        limite = min(hoja_resumen.max_row, celda.row + 50)
        return sum(
            _decimal(hoja_resumen.cell(fila, celda.column).value) is not None
            for fila in range(celda.row + 1, limite + 1)
        )

    encabezado = max(candidatos, key=puntuacion_encabezado)
    filas: list[FilaResumen] = []
    limite = min(hoja_resumen.max_row, encabezado.row + 50)
    for numero_fila in range(encabezado.row + 1, limite + 1):
        celdas_izquierda = [
            hoja_resumen.cell(numero_fila, columna)
            for columna in range(1, encabezado.column)
        ]
        if any(_es_suma(celda.value) for celda in celdas_izquierda):
            break

        celda_importe = hoja_resumen.cell(numero_fila, encabezado.column)
        importe = _decimal(celda_importe.value)
        if importe is None:
            continue

        celda_empresa = next(
            (
                celda
                for celda in celdas_izquierda
                if _texto_normalizado(celda.value)
                and _decimal(celda.value) is None
                and _texto_normalizado(celda.value) not in {"empresa", "suma"}
            ),
            None,
        )
        if celda_empresa is None:
            continue
        filas.append(
            FilaResumen(
                empresa=str(celda_empresa.value).strip(),
                importe=ImporteDetectado(
                    importe, f"{hoja_resumen.title}!{celda_importe.coordinate}"
                ),
                celda_empresa=celda_empresa,
            )
        )
    return filas


def _textos_de_hoja(hoja) -> list[str]:
    return [
        str(celda.value)
        for fila in hoja.iter_rows()
        for celda in fila
        if not isinstance(celda, MergedCell)
        and isinstance(celda.value, str)
        and _texto_normalizado(celda.value)
    ]


def _relacionar_resumen(
    hoja, detalle: ImporteDetectado, filas_resumen: list[FilaResumen]
) -> FilaResumen | None:
    # Un importe igual y único es la evidencia más fuerte. Además evita que
    # nombres mencionados dentro de otra hoja provoquen una asociación incorrecta.
    if detalle.valor is not None:
        coincidencias = [
            fila
            for fila in filas_resumen
            if fila.importe.valor is not None
            and fila.importe.valor == detalle.valor
        ]
        if len(coincidencias) == 1:
            return coincidencias[0]

    textos = [hoja.title, *_textos_de_hoja(hoja)]
    mejor: tuple[float, FilaResumen] | None = None
    for fila_resumen in filas_resumen:
        puntuacion = max(
            _coincidencia_hoja(fila_resumen.empresa, texto) for texto in textos
        )
        if mejor is None or puntuacion > mejor[0]:
            mejor = (puntuacion, fila_resumen)

    # ENA, por ejemplo, se relaciona con Nahuatl porque ese nombre aparece en
    # el encabezado y en el contenido de la hoja, aunque no aparezca en la pestaña.
    if mejor is not None and mejor[0] >= 0.82:
        return mejor[1]

    return None


def analizar_excel(contenido: bytes) -> tuple[pd.DataFrame, str]:
    libro = load_workbook(BytesIO(contenido), data_only=True, read_only=False)
    hojas_resumen = [hoja for hoja in libro.worksheets if _texto_normalizado(hoja.title) == "resumen"]
    if not hojas_resumen:
        raise ValueError("El archivo no contiene una hoja llamada 'Resumen'.")

    hoja_resumen = hojas_resumen[0]
    hojas_excluidas = {"resumen", "cuentas"}
    hojas_detalle = [
        hoja
        for hoja in libro.worksheets
        if _texto_normalizado(hoja.title) not in hojas_excluidas
    ]
    if not hojas_detalle:
        raise ValueError("El archivo no contiene hojas para comparar con Resumen.")

    filas_resumen = _filas_del_resumen(hoja_resumen)
    if not filas_resumen:
        raise ValueError("No se encontraron empresas e importes bajo la columna 'Suma' de Resumen.")

    filas = []
    for hoja in hojas_detalle:
        detalle = _suma_de_hoja(hoja)
        fila_resumen = _relacionar_resumen(hoja, detalle, filas_resumen)
        resumen = (
            fila_resumen.importe
            if fila_resumen is not None
            else ImporteDetectado(None, "No se encontró la empresa en Resumen")
        )
        diferencia = None
        estado = "Falta dato"
        if resumen.valor is not None and detalle.valor is not None:
            diferencia = detalle.valor - resumen.valor
            estado = "Coincide" if diferencia == Decimal("0.00") else "No coincide"
        filas.append(
            {
                "Hoja": hoja.title,
                "Empresa en Resumen": fila_resumen.empresa if fila_resumen else None,
                "Suma en Resumen": float(resumen.valor) if resumen.valor is not None else None,
                "Suma en hoja": float(detalle.valor) if detalle.valor is not None else None,
                "Diferencia": float(diferencia) if diferencia is not None else None,
                "Estado": estado,
                "Celda Resumen": resumen.referencia,
                "Celda hoja": detalle.referencia,
            }
        )
    return pd.DataFrame(filas), hoja_resumen.title


def _valor_como_texto(valor: object, quitar_apostrofe: bool = False) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        texto = str(int(valor))
    else:
        texto = str(valor)
    texto = texto.strip()
    if quitar_apostrofe:
        texto = texto.lstrip("'").strip()
    return texto


def _limpiar_clave(valor: object) -> str:
    texto = _valor_como_texto(valor)
    return re.sub(r"[.\s]+", "", texto)


def _columnas_de_personas(hoja) -> tuple[int, dict[str, int]] | None:
    for fila in hoja.iter_rows():
        encabezados = {
            _texto_normalizado(celda.value): celda.column
            for celda in fila
            if not isinstance(celda, MergedCell) and celda.value is not None
        }
        columna_cuenta = next(
            (
                columna
                for texto, columna in encabezados.items()
                if texto.startswith("no") and ("cuenta" in texto or "clabe" in texto)
            ),
            None,
        )
        requeridas = {
            "clave": encabezados.get("clave"),
            "nombre": encabezados.get("nombre"),
            "total": encabezados.get("total a depositar"),
            "cuenta": columna_cuenta,
            "banco": encabezados.get("banco"),
        }
        if all(requeridas.values()):
            return fila[0].row, requeridas
    return None


def extraer_tabla_comparacion(contenido: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    libro = load_workbook(BytesIO(contenido), data_only=True, read_only=False)
    hojas_excluidas = {"resumen", "cuentas"}
    registros: list[dict[str, object]] = []
    registros_efectivo: list[dict[str, object]] = []

    for hoja in libro.worksheets:
        if _texto_normalizado(hoja.title) in hojas_excluidas:
            continue
        estructura = _columnas_de_personas(hoja)
        if estructura is None:
            continue
        fila_encabezado, columnas = estructura

        for numero_fila in range(fila_encabezado + 1, hoja.max_row + 1):
            clave = _limpiar_clave(hoja.cell(numero_fila, columnas["clave"]).value)
            nombre = _valor_como_texto(hoja.cell(numero_fila, columnas["nombre"]).value)
            total = _decimal(hoja.cell(numero_fila, columnas["total"]).value)
            cuenta = _valor_como_texto(
                hoja.cell(numero_fila, columnas["cuenta"]).value,
                quitar_apostrofe=True,
            )
            banco = _valor_como_texto(hoja.cell(numero_fila, columnas["banco"]).value)

            # Excluye títulos de bloque, filas de totales y el resumen al pie.
            if not nombre or total is None or not banco:
                continue

            registro = {
                "EMPRESA": hoja.title.strip(),
                "BLOQUE": "",
                "CLAVE": clave,
                "Nombre": nombre,
                "Total a Depositar": float(total),
                "No. Cuenta": cuenta,
                "Banco": banco,
            }
            if _texto_normalizado(banco) == "efectivo":
                registros_efectivo.append(registro)
            elif clave:
                registros.append(registro)

    columnas_salida = [
        "EMPRESA",
        "BLOQUE",
        "CLAVE",
        "Nombre",
        "Total a Depositar",
        "No. Cuenta",
        "Banco",
    ]
    return (
        pd.DataFrame(registros, columns=columnas_salida),
        pd.DataFrame(registros_efectivo, columns=columnas_salida),
    )


def validar_catalogo(contenido: bytes) -> int:
    libro = load_workbook(BytesIO(contenido), data_only=True, read_only=True)
    hojas = [
        hoja for hoja in libro.worksheets if _texto_normalizado(hoja.title) == "catalogo"
    ]
    if not hojas:
        raise ValueError("El archivo no contiene una hoja llamada 'Catálogo'.")
    hoja = hojas[0]
    encabezados = {
        _texto_normalizado(celda.value) for celda in next(hoja.iter_rows(min_row=1, max_row=1))
    }
    faltantes = {
        "empresa",
        "clave",
        "nombre",
        "clabe interbancaria",
        "banco",
        "id islas",
        "id sind",
    } - encabezados
    if faltantes:
        raise ValueError(
            "A la hoja Catálogo le faltan encabezados requeridos: "
            + ", ".join(sorted(faltantes))
        )
    return max(hoja.max_row - 1, 0)


def _normalizar_cuenta(valor: object) -> str:
    return re.sub(r"[\s']+", "", _valor_como_texto(valor))


def _texto_igual_excel(izquierda: object, derecha: object) -> bool:
    if izquierda is None or derecha is None:
        return False
    return _valor_como_texto(izquierda).casefold() == _valor_como_texto(derecha).casefold()


def crear_tabla_con_catalogo(
    tabla_origen: pd.DataFrame, contenido_catalogo: bytes
) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    libro = load_workbook(BytesIO(contenido_catalogo), data_only=True, read_only=True)
    hoja = next(
        hoja
        for hoja in libro.worksheets
        if _texto_normalizado(hoja.title) == "catalogo"
    )
    encabezados: dict[str, int] = {}
    for celda in next(hoja.iter_rows(min_row=1, max_row=1)):
        texto = _texto_normalizado(celda.value)
        if texto:
            # BUSCARV usa las primeras columnas del Catálogo (B:H). Si un
            # encabezado se repite más adelante, se conserva el primero.
            encabezados.setdefault(texto, celda.column)
    columnas = {
        "clave": encabezados["clave"],
        "nombre": encabezados["nombre"],
        "clabe": encabezados["clabe interbancaria"],
        "banco": encabezados["banco"],
        "id_islas": encabezados["id islas"],
        "id_sindicato": encabezados["id sind"],
    }

    # Conserva la primera coincidencia, igual que BUSCARV con coincidencia exacta.
    catalogo_por_clave: dict[str, dict[str, object]] = {}
    catalogo_por_cuenta: dict[str, dict[str, object]] = {}
    for valores in hoja.iter_rows(min_row=2, values_only=True):
        clave = _limpiar_clave(valores[columnas["clave"] - 1])
        if not clave:
            continue
        registro_catalogo = {
            "Clave Catálogo": clave,
            "Nombre Catálogo": _valor_como_texto(valores[columnas["nombre"] - 1]),
            "CLABE": _valor_como_texto(
                valores[columnas["clabe"] - 1],
                quitar_apostrofe=True,
            ),
            "Banco Catálogo": _valor_como_texto(valores[columnas["banco"] - 1]),
            "ID Islas": valores[columnas["id_islas"] - 1],
            "ID Sindicato": valores[columnas["id_sindicato"] - 1],
        }
        catalogo_por_clave.setdefault(clave, registro_catalogo)
        cuenta_catalogo = _normalizar_cuenta(registro_catalogo["CLABE"])
        if cuenta_catalogo:
            catalogo_por_cuenta.setdefault(cuenta_catalogo, registro_catalogo)

    filas: list[dict[str, object]] = []
    incidencias: list[dict[str, object]] = []
    cuentas_verificadas = 0
    for registro in tabla_origen.to_dict("records"):
        clave_origen = str(registro["CLAVE"])
        cuenta_origen = _normalizar_cuenta(registro["No. Cuenta"])
        clave_temporal = _texto_normalizado(clave_origen) == "s n"
        if clave_temporal:
            encontrado = catalogo_por_cuenta.get(cuenta_origen, {})
        else:
            encontrado = catalogo_por_clave.get(clave_origen, {})
        fila = {
            **registro,
            "CLAVE": (
                encontrado.get("Clave Catálogo", clave_origen)
                if clave_temporal
                else clave_origen
            ),
            "REPT": "-",
            "Nombre Catálogo": encontrado.get("Nombre Catálogo"),
            "CLABE": encontrado.get("CLABE"),
            "Banco Catálogo": encontrado.get("Banco Catálogo"),
            "ID Islas": encontrado.get("ID Islas"),
            "ID Sindicato": encontrado.get("ID Sindicato"),
        }
        clabe_catalogo = _normalizar_cuenta(encontrado.get("CLABE"))
        cuenta_coincide = bool(encontrado) and cuenta_origen == clabe_catalogo
        fila.update(
            {
                "NOM": bool(encontrado)
                and _texto_igual_excel(registro["Nombre"], encontrado.get("Nombre Catálogo")),
                "CUEN": cuenta_coincide,
                "BAN": bool(encontrado)
                and _texto_igual_excel(registro["Banco"], encontrado.get("Banco Catálogo")),
            }
        )
        filas.append(fila)

        if cuenta_coincide:
            cuentas_verificadas += 1
        else:
            incidencias.append(
                {
                    "EMPRESA": registro["EMPRESA"],
                    "CLAVE": registro["CLAVE"],
                    "Nombre": registro["Nombre"],
                    "No. Cuenta": registro["No. Cuenta"],
                    "CLABE Catálogo": encontrado.get("CLABE"),
                    "Resultado": (
                        "Cuenta no coincide" if encontrado else "Clave no encontrada"
                    ),
                }
            )

    columnas_salida = [
        "EMPRESA",
        "BLOQUE",
        "CLAVE",
        "Nombre",
        "Total a Depositar",
        "No. Cuenta",
        "Banco",
        "REPT",
        "Nombre Catálogo",
        "CLABE",
        "Banco Catálogo",
        "ID Islas",
        "ID Sindicato",
        "NOM",
        "CUEN",
        "BAN",
    ]
    return (
        pd.DataFrame(filas, columns=columnas_salida),
        pd.DataFrame(incidencias),
        cuentas_verificadas,
    )


def obtener_no_coincidencias(tabla: pd.DataFrame) -> pd.DataFrame:
    comparaciones = {
        "NOM": ("Nombre", "Nombre Catálogo"),
        "CUEN": ("No. Cuenta", "CLABE"),
        "BAN": ("Banco", "Banco Catálogo"),
    }
    filas: list[dict[str, object]] = []
    for registro in tabla.to_dict("records"):
        for validacion, (columna_origen, columna_catalogo) in comparaciones.items():
            if bool(registro[validacion]):
                continue
            filas.append(
                {
                    "EMPRESA": registro["EMPRESA"],
                    "CLAVE": registro["CLAVE"],
                    "Nombre": registro["Nombre"],
                    "Validación": validacion,
                    "Valor del archivo": registro[columna_origen],
                    "Valor del Catálogo": registro[columna_catalogo],
                }
            )
    return pd.DataFrame(
        filas,
        columns=[
            "EMPRESA",
            "CLAVE",
            "Nombre",
            "Validación",
            "Valor del archivo",
            "Valor del Catálogo",
        ],
    )


def crear_tabla_corregida(tabla: pd.DataFrame) -> pd.DataFrame:
    corregida = tabla.copy()
    for indice, registro in corregida.iterrows():
        nombre_catalogo = registro["Nombre Catálogo"]
        clabe_catalogo = registro["CLABE"]
        banco_catalogo = registro["Banco Catálogo"]

        if pd.notna(nombre_catalogo) and _valor_como_texto(nombre_catalogo):
            corregida.at[indice, "Nombre"] = nombre_catalogo
        if pd.notna(clabe_catalogo) and _valor_como_texto(clabe_catalogo):
            corregida.at[indice, "No. Cuenta"] = clabe_catalogo
        if pd.notna(banco_catalogo) and _valor_como_texto(banco_catalogo):
            corregida.at[indice, "Banco"] = banco_catalogo

        corregida.at[indice, "NOM"] = _texto_igual_excel(
            corregida.at[indice, "Nombre"], nombre_catalogo
        )
        corregida.at[indice, "CUEN"] = _normalizar_cuenta(
            corregida.at[indice, "No. Cuenta"]
        ) == _normalizar_cuenta(clabe_catalogo)
        corregida.at[indice, "BAN"] = _texto_igual_excel(
            corregida.at[indice, "Banco"], banco_catalogo
        )
    return corregida


def filtrar_tabla(tabla: pd.DataFrame, busqueda: str) -> pd.DataFrame:
    termino = busqueda.strip()
    if not termino:
        return tabla
    patron = re.escape(termino)
    mascara = tabla.fillna("").astype(str).apply(
        lambda columna: columna.str.contains(patron, case=False, regex=True)
    ).any(axis=1)
    return tabla.loc[mascara]


def exportar_comparacion_excel(
    tabla: pd.DataFrame, contenido_catalogo: bytes, nombre_hoja: str
) -> bytes:
    libro_plantilla = load_workbook(
        BytesIO(contenido_catalogo), data_only=False, read_only=False
    )
    hojas_comparacion = [
        hoja
        for hoja in libro_plantilla.worksheets
        if _texto_normalizado(hoja.title) == "comparacion"
    ]
    if not hojas_comparacion:
        raise ValueError("El Catálogo no contiene la hoja 'Comparación' para tomar el formato.")
    plantilla = hojas_comparacion[0]

    libro_salida = Workbook()
    libro_salida.loaded_theme = libro_plantilla.loaded_theme
    hoja_salida = libro_salida.active
    hoja_salida.title = nombre_hoja[:31]

    encabezados_excel = [
        "EMPRESA",
        "BLOQUE",
        "CLAVE",
        "Nombre",
        "Total a Depositar",
        "No. Cuenta",
        "Banco",
        "REPT",
        "Nombre",
        "CLABE",
        "Banco",
        "ID Islas",
        "ID Sindicato",
        "NOM",
        "CUEN",
        "BAN",
    ]
    for numero_columna, encabezado in enumerate(encabezados_excel, 1):
        destino = hoja_salida.cell(1, numero_columna, encabezado)
        origen = plantilla.cell(1, numero_columna)
        destino.font = copy(origen.font)
        destino.fill = copy(origen.fill)
        destino.border = copy(origen.border)
        destino.alignment = copy(origen.alignment)
        destino.protection = copy(origen.protection)
        destino.number_format = origen.number_format

        letra = origen.column_letter
        dimension_origen = plantilla.column_dimensions[letra]
        dimension_destino = hoja_salida.column_dimensions[letra]
        dimension_destino.width = dimension_origen.width
        dimension_destino.hidden = dimension_origen.hidden

    for numero_fila, valores in enumerate(tabla.itertuples(index=False, name=None), 2):
        for numero_columna, valor in enumerate(valores, 1):
            if pd.isna(valor):
                valor = None
            elif hasattr(valor, "item"):
                valor = valor.item()
            destino = hoja_salida.cell(numero_fila, numero_columna, valor)
            origen = plantilla.cell(2, numero_columna)
            destino.font = copy(origen.font)
            destino.fill = copy(origen.fill)
            destino.border = copy(origen.border)
            destino.alignment = copy(origen.alignment)
            destino.protection = copy(origen.protection)
            destino.number_format = origen.number_format

    # Garantiza que las cuentas y CLABE se mantengan como texto y no pierdan ceros.
    for numero_fila in range(2, len(tabla) + 2):
        hoja_salida.cell(numero_fila, 6).number_format = "@"
        hoja_salida.cell(numero_fila, 10).number_format = "@"

    hoja_salida.row_dimensions[1].height = plantilla.row_dimensions[1].height
    hoja_salida.freeze_panes = "A2"
    hoja_salida.auto_filter.ref = f"A1:P{len(tabla) + 1}"
    hoja_salida.sheet_view.showGridLines = plantilla.sheet_view.showGridLines
    hoja_salida.sheet_properties.tabColor = copy(plantilla.sheet_properties.tabColor)

    salida = BytesIO()
    libro_salida.save(salida)
    return salida.getvalue()


def preparar_tabla_efectivo(tabla_efectivo: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Empresa": tabla_efectivo["EMPRESA"],
            "Clave": tabla_efectivo["CLAVE"],
            "Nombre": tabla_efectivo["Nombre"],
            "Importe": tabla_efectivo["Total a Depositar"],
            "CUENTA": "",
            "OBSERVACION": "",
        }
    )


def exportar_efectivo_excel(tabla: pd.DataFrame) -> bytes:
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Efectivo"

    total = float(tabla["Importe"].sum()) if not tabla.empty else 0.0
    hoja["D1"] = "EFECTIVO REQUERIDO"
    hoja["E1"] = total
    for celda in (hoja["D1"], hoja["E1"]):
        celda.font = Font(name="Arial", size=9, bold=True, color="FFC00000")
    hoja["D1"].alignment = Alignment(horizontal="center")
    hoja["E1"].number_format = '$#,##0.00'

    encabezados = ["Empresa", "Clave", "Nombre", "Importe", "CUENTA", "OBSERVACION"]
    relleno_encabezado = PatternFill("solid", fgColor="FF0D0D0D")
    for columna, encabezado in enumerate(encabezados, 2):
        celda = hoja.cell(2, columna, encabezado)
        celda.font = Font(name="Arial", size=8, bold=True, color="FFFFFFFF")
        celda.fill = relleno_encabezado
        celda.alignment = Alignment(horizontal="center")

    for numero_fila, valores in enumerate(tabla.itertuples(index=False, name=None), 3):
        for numero_columna, valor in enumerate(valores, 2):
            if pd.isna(valor):
                valor = None
            celda = hoja.cell(numero_fila, numero_columna, valor)
            celda.font = Font(name="Arial", size=8)
            if numero_columna == 2:
                celda.alignment = Alignment(horizontal="center")
            elif numero_columna in {4, 6}:
                celda.alignment = Alignment(horizontal="left")
        hoja.cell(numero_fila, 3).number_format = "@"
        hoja.cell(numero_fila, 5).number_format = '$#,##0.00'

    anchos = {"A": 13, "B": 13, "C": 13, "D": 38, "E": 13, "F": 13, "G": 18}
    for letra, ancho in anchos.items():
        hoja.column_dimensions[letra].width = ancho
    hoja.freeze_panes = "B3"
    hoja.auto_filter.ref = f"B2:G{len(tabla) + 2}"

    salida = BytesIO()
    libro.save(salida)
    return salida.getvalue()


def _catalogo_bancos(contenido: bytes) -> tuple[dict[str, str], dict[str, str]]:
    filas: list[tuple[object, ...]]
    if contenido.startswith(b"\xd0\xcf\x11\xe0"):
        import xlrd

        libro = xlrd.open_workbook(file_contents=contenido, on_demand=True)
        nombre_hoja = next(
            (
                nombre
                for nombre in libro.sheet_names()
                if _texto_normalizado(nombre) == "anexo catalogo de bancos"
            ),
            None,
        )
        if nombre_hoja is None:
            raise ValueError("No se encontró la hoja 'Anexo - Catálogo de Bancos'.")
        hoja = libro.sheet_by_name(nombre_hoja)
        filas = [tuple(hoja.row_values(fila)) for fila in range(hoja.nrows)]
    else:
        libro = load_workbook(BytesIO(contenido), data_only=True, read_only=True)
        hoja = next(
            (
                hoja
                for hoja in libro.worksheets
                if _texto_normalizado(hoja.title) == "anexo catalogo de bancos"
            ),
            None,
        )
        if hoja is None:
            raise ValueError("No se encontró la hoja 'Anexo - Catálogo de Bancos'.")
        filas = list(hoja.iter_rows(values_only=True))

    encabezado = None
    for indice, fila in enumerate(filas):
        normalizados = [_texto_normalizado(valor) for valor in fila]
        if "nombre de institucion" in normalizados and "clave transfer" in normalizados:
            encabezado = (indice, normalizados)
            break
    if encabezado is None:
        raise ValueError("No se encontraron NOMBRE DE INSTITUCION y CLAVE TRANSFER.")

    indice_fila, encabezados = encabezado
    columna_nombre = encabezados.index("nombre de institucion")
    columna_transfer = encabezados.index("clave transfer")
    columna_institucion = (
        encabezados.index("no inst") if "no inst" in encabezados else None
    )
    por_nombre: dict[str, str] = {}
    por_prefijo: dict[str, str] = {}
    for fila in filas[indice_fila + 1 :]:
        if max(columna_nombre, columna_transfer) >= len(fila):
            continue
        nombre = _texto_normalizado(fila[columna_nombre])
        clave_transfer = _valor_como_texto(fila[columna_transfer])
        if not nombre or not clave_transfer:
            continue
        por_nombre.setdefault(nombre, clave_transfer)
        if columna_institucion is not None and columna_institucion < len(fila):
            numero = re.sub(r"\D", "", _valor_como_texto(fila[columna_institucion]))
            if numero:
                por_prefijo.setdefault(numero[-3:].zfill(3), clave_transfer)
    return por_nombre, por_prefijo


def _nombre_banco_catalogo(nombre: object) -> str:
    normalizado = _texto_normalizado(nombre)
    alias = {
        "banbajio": "bajio",
        "bancomer": "bbva mexico",
        "bbva": "bbva mexico",
        "santader": "banco santander",
        "santander": "banco santander",
        "banorte": "banorte ixe",
    }
    return alias.get(normalizado, normalizado)


def _cuentas_csi_despues_del_corte(contenido: bytes) -> set[str]:
    libro = load_workbook(BytesIO(contenido), data_only=True, read_only=False)
    hoja = next(
        (hoja for hoja in libro.worksheets if _texto_normalizado(hoja.title) == "csi"),
        None,
    )
    if hoja is None:
        return set()
    estructura = _columnas_de_personas(hoja)
    if estructura is None:
        return set()
    fila_encabezado, columnas = estructura
    personas_vistas = False
    despues_del_corte = False
    cuentas: set[str] = set()

    for numero_fila in range(fila_encabezado + 1, hoja.max_row + 1):
        fila_vacia = all(
            hoja.cell(numero_fila, columna).value is None
            or str(hoja.cell(numero_fila, columna).value).strip() == ""
            for columna in range(1, hoja.max_column + 1)
        )
        if fila_vacia and personas_vistas and not despues_del_corte:
            despues_del_corte = True
            continue

        nombre = _valor_como_texto(hoja.cell(numero_fila, columnas["nombre"]).value)
        total = _decimal(hoja.cell(numero_fila, columnas["total"]).value)
        banco = _valor_como_texto(hoja.cell(numero_fila, columnas["banco"]).value)
        cuenta = _normalizar_cuenta(hoja.cell(numero_fila, columnas["cuenta"]).value)
        if not nombre or total is None or not banco or not cuenta:
            continue
        personas_vistas = True
        if despues_del_corte and _texto_normalizado(banco) != "efectivo":
            cuentas.add(cuenta)
    return cuentas


def crear_tabla_transferencias(
    tabla_corregida: pd.DataFrame,
    contenido_bancos: bytes,
    contenido_litografia: bytes,
) -> tuple[pd.DataFrame, list[str]]:
    bancos_por_nombre, bancos_por_prefijo = _catalogo_bancos(contenido_bancos)
    cuentas_con_recibo = _cuentas_csi_despues_del_corte(contenido_litografia)
    filas: list[dict[str, object]] = []
    bancos_sin_clave: set[str] = set()

    csi = tabla_corregida[tabla_corregida["EMPRESA"].map(_texto_normalizado) == "csi"]
    csi_no_santander = 0
    for _, registro in csi.iterrows():
        clabe = _normalizar_cuenta(registro["CLABE"])
        nombre_banco = _nombre_banco_catalogo(registro["Banco"])
        es_santander = nombre_banco == "banco santander" or clabe.startswith("014")
        if not es_santander:
            csi_no_santander += 1
    total_bloques = max(1, (csi_no_santander + 49) // 50)

    posicion_csi = 0
    for registro in tabla_corregida.to_dict("records"):
        empresa = _valor_como_texto(registro["EMPRESA"])
        clabe = _normalizar_cuenta(registro["CLABE"])
        nombre_banco = _nombre_banco_catalogo(registro["Banco"])
        es_santander = nombre_banco == "banco santander" or clabe.startswith("014")

        if es_santander:
            banco_transferencia = "SANTANDER"
        else:
            banco_transferencia = bancos_por_nombre.get(nombre_banco)
            if banco_transferencia is None and len(clabe) >= 3:
                banco_transferencia = bancos_por_prefijo.get(clabe[:3])
            if banco_transferencia is None:
                bancos_sin_clave.add(_valor_como_texto(registro["Banco"]))

        numero_bloque = ""
        if _texto_normalizado(empresa) == "csi":
            if es_santander:
                numero_bloque = "SANTANDER"
            else:
                posicion_csi += 1
                numero_bloque = f"{(posicion_csi - 1) // 50 + 1} DE {total_bloques}"

        recibo = ""
        if (
            _texto_normalizado(empresa) == "csi"
            and not es_santander
            and _normalizar_cuenta(registro["No. Cuenta"]) in cuentas_con_recibo
        ):
            recibo = "X"

        filas.append(
            {
                "CIA": empresa,
                "No Bloque": numero_bloque,
                "Clave": registro["CLAVE"],
                "Nombre": registro["Nombre"],
                "Banco": banco_transferencia or "",
                "CLABE": clabe,
                "CUENTA": clabe[6:-1] if es_santander and len(clabe) > 7 else "",
                "RECIBO": recibo,
                "Importe": registro["Total a Depositar"],
                "Com Banc": "",
            }
        )
    return pd.DataFrame(filas), sorted(bancos_sin_clave)


def exportar_transferencias_excel(tabla: pd.DataFrame) -> bytes:
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Transferencias"
    hoja["J1"] = float(tabla["Importe"].sum()) if not tabla.empty else 0.0
    hoja["J1"].font = Font(name="Arial", size=8, bold=True)
    hoja["J1"].number_format = '$#,##0.00'

    encabezados = list(tabla.columns)
    relleno = PatternFill("solid", fgColor="FF0D0D0D")
    for columna, encabezado in enumerate(encabezados, 2):
        celda = hoja.cell(2, columna, encabezado)
        celda.font = Font(name="Arial", size=8, bold=True, color="FFFFFFFF")
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center")

    for numero_fila, valores in enumerate(tabla.itertuples(index=False, name=None), 3):
        for numero_columna, valor in enumerate(valores, 2):
            if pd.isna(valor):
                valor = None
            celda = hoja.cell(numero_fila, numero_columna, valor)
            celda.font = Font(name="Arial", size=8)
        hoja.cell(numero_fila, 4).number_format = "@"
        hoja.cell(numero_fila, 7).number_format = "@"
        hoja.cell(numero_fila, 8).number_format = "@"
        hoja.cell(numero_fila, 10).number_format = '$#,##0.00'

    anchos = {
        "A": 3.55,
        "B": 7.44,
        "C": 11.44,
        "D": 9,
        "E": 41.89,
        "F": 11.11,
        "G": 19.11,
        "H": 12,
        "I": 11.55,
        "J": 12.33,
        "K": 11.55,
    }
    for letra, ancho in anchos.items():
        hoja.column_dimensions[letra].width = ancho
    hoja.freeze_panes = "B3"
    hoja.auto_filter.ref = f"B2:K{len(tabla) + 2}"

    salida = BytesIO()
    libro.save(salida)
    return salida.getvalue()


def vista_litografia() -> None:
    fecha_descarga = date.today().strftime("%d%m%Y")
    st.header("Litografía")
    st.write(
        "Carga el libro de Excel para comparar la **Suma** indicada en la hoja "
        "**Resumen** contra la **Suma** de cada una de las demás hojas."
    )

    archivo = st.file_uploader(
        "Cargar archivo de Litografía",
        type=["xlsx", "xlsm"],
        key="excel_litografia",
        help="Formatos admitidos: .xlsx y .xlsm",
    )
    if archivo is None:
        st.info("Carga el archivo para ejecutar la primera verificación.")
        return

    try:
        resultados, nombre_resumen = analizar_excel(archivo.getvalue())
    except Exception as error:
        st.error(f"No se pudo analizar el archivo: {error}")
        return

    coinciden = int((resultados["Estado"] == "Coincide").sum())
    no_coinciden = int((resultados["Estado"] == "No coincide").sum())
    faltantes = int((resultados["Estado"] == "Falta dato").sum())
    columna_1, columna_2, columna_3 = st.columns(3)
    columna_1.metric("Coinciden", coinciden)
    columna_2.metric("No coinciden", no_coinciden)
    columna_3.metric("Faltan datos", faltantes)

    st.subheader("Primera verificación")
    st.dataframe(
        resultados,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Suma en Resumen": st.column_config.NumberColumn(format="$ %.2f"),
            "Suma en hoja": st.column_config.NumberColumn(format="$ %.2f"),
            "Diferencia": st.column_config.NumberColumn(format="$ %.2f"),
        },
    )

    if no_coinciden:
        st.warning(f"Hay {no_coinciden} hoja(s) cuya suma no coincide con Resumen.")
    elif faltantes:
        st.warning(
            "No fue posible localizar todos los importes. Revisa las referencias de celda; "
            "si las sumas son fórmulas sin resultado guardado, abre y guarda el archivo en Excel."
        )
    else:
        st.success("Todas las sumas coinciden con el Resumen.")

    st.divider()
    st.subheader("Base de datos de personas")
    archivo_catalogo = st.file_uploader(
        "Cargar Catálogo de personas",
        type=["xlsx", "xlsm"],
        key="catalogo_litografia",
        help="El archivo debe contener una hoja llamada Catálogo.",
    )
    if archivo_catalogo is None:
        st.info("Carga el Catálogo para generar la tabla de comparación.")
        return

    try:
        contenido_catalogo = archivo_catalogo.getvalue()
        personas_catalogo = validar_catalogo(contenido_catalogo)
        tabla_origen, tabla_efectivo = extraer_tabla_comparacion(archivo.getvalue())
        tabla, _incidencias_cuenta, _cuentas_verificadas = crear_tabla_con_catalogo(
            tabla_origen, contenido_catalogo
        )
        no_coincidencias = obtener_no_coincidencias(tabla)
        tabla_corregida = crear_tabla_corregida(tabla)
        excel_comparacion = exportar_comparacion_excel(
            tabla, contenido_catalogo, "Comparación"
        )
        excel_corregido = exportar_comparacion_excel(
            tabla_corregida, contenido_catalogo, "Comparación corregida"
        )
        tabla_efectivo_preparada = preparar_tabla_efectivo(tabla_efectivo)
        excel_efectivo = exportar_efectivo_excel(tabla_efectivo_preparada)
    except Exception as error:
        st.error(f"No se pudo preparar la comparación: {error}")
        return

    st.success(f"Catálogo cargado correctamente: {personas_catalogo:,} registros.")
    st.subheader("Tabla para comparación")
    st.caption(
        "La clave no contiene puntos ni espacios. En Nombre sólo se quitaron "
        "espacios al principio y al final. Las columnas de Catálogo se obtienen "
        "por CLAVE. NOM, CUEN y BAN validan Nombre, Cuenta y Banco, respectivamente."
    )
    metrica_nom, metrica_cuen, metrica_ban = st.columns(3)
    metrica_nom.metric("NOM correctos", int(tabla["NOM"].sum()))
    metrica_cuen.metric("CUEN correctas", int(tabla["CUEN"].sum()))
    metrica_ban.metric("BAN correctos", int(tabla["BAN"].sum()))

    busqueda_original = st.text_input(
        "Buscar en la tabla de comparación",
        placeholder="Empresa, clave, nombre, cuenta, banco...",
        key="buscar_tabla_comparacion_litografia",
    )
    tabla_visible = filtrar_tabla(tabla, busqueda_original)
    if busqueda_original.strip():
        st.caption(f"Se encontraron {len(tabla_visible)} registro(s).")

    st.dataframe(
        tabla_visible,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Total a Depositar": st.column_config.NumberColumn(format="$ %.2f"),
            "No. Cuenta": st.column_config.TextColumn(),
            "Nombre Catálogo": st.column_config.TextColumn("Nombre"),
            "CLABE": st.column_config.TextColumn(),
            "Banco Catálogo": st.column_config.TextColumn("Banco"),
        },
    )
    st.download_button(
        "Descargar tabla de comparación",
        data=excel_comparacion,
        file_name=f"comparacion_litografia_{fecha_descarga}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descargar_comparacion_litografia",
    )
    st.subheader("No coincidencias")
    if no_coincidencias.empty:
        st.success("Nombre, cuenta y banco coinciden en todos los registros.")
    else:
        st.warning(f"Se encontraron {len(no_coincidencias)} validación(es) sin coincidencia.")
        st.dataframe(no_coincidencias, hide_index=True, use_container_width=True)

    st.subheader("Tabla de comparación corregida")
    st.caption(
        "Nombre, No. Cuenta y Banco se sustituyen por los valores del Catálogo. "
        "Los datos originales sólo se muestran en la tabla anterior y no se modifica "
        "ninguno de los archivos cargados."
    )
    busqueda_corregida = st.text_input(
        "Buscar en la tabla corregida",
        placeholder="Empresa, clave, nombre, cuenta, banco...",
        key="buscar_tabla_corregida_litografia",
    )
    tabla_corregida_visible = filtrar_tabla(tabla_corregida, busqueda_corregida)
    if busqueda_corregida.strip():
        st.caption(f"Se encontraron {len(tabla_corregida_visible)} registro(s).")
    st.dataframe(
        tabla_corregida_visible,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Total a Depositar": st.column_config.NumberColumn(format="$ %.2f"),
            "No. Cuenta": st.column_config.TextColumn(),
            "Nombre Catálogo": st.column_config.TextColumn("Nombre"),
            "CLABE": st.column_config.TextColumn(),
            "Banco Catálogo": st.column_config.TextColumn("Banco"),
        },
    )
    st.download_button(
        "Descargar tabla corregida",
        data=excel_corregido,
        file_name=f"comparacion_litografia_corregida_{fecha_descarga}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descargar_comparacion_litografia_corregida",
    )
    st.divider()
    st.subheader("Tabla de efectivo")
    if tabla_efectivo_preparada.empty:
        st.info("No se encontraron registros con Banco igual a Efectivo.")
    else:
        efectivo_requerido = float(tabla_efectivo_preparada["Importe"].sum())
        st.metric("EFECTIVO REQUERIDO", f"$ {efectivo_requerido:,.2f}")
        st.caption(
            "Incluye todos los registros cuyo Banco es Efectivo, incluso filas sin "
            "clave como EFECTIVO *** CSIP ****."
        )
        busqueda_efectivo = st.text_input(
            "Buscar en la tabla de efectivo",
            placeholder="Empresa, clave, nombre, importe...",
            key="buscar_tabla_efectivo_litografia",
        )
        efectivo_visible = filtrar_tabla(tabla_efectivo_preparada, busqueda_efectivo)
        if busqueda_efectivo.strip():
            st.caption(f"Se encontraron {len(efectivo_visible)} registro(s).")
        st.dataframe(
            efectivo_visible,
            hide_index=True,
            use_container_width=True,
            column_config={
                "Importe": st.column_config.NumberColumn(format="$ %.2f"),
                "Clave": st.column_config.TextColumn(),
            },
        )
        st.download_button(
            "Descargar tabla de efectivo",
            data=excel_efectivo,
            file_name=f"efectivo_litografia_{fecha_descarga}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="descargar_efectivo_litografia",
        )

    st.divider()
    st.subheader("Tabla de transferencias")
    archivo_bancos = st.file_uploader(
        "Cargar Catálogo de Bancos",
        type=["xls", "xlsx", "xlsm"],
        key="catalogo_bancos_litografia",
        help="Debe contener la hoja Anexo - Catálogo de Bancos.",
    )
    if archivo_bancos is None:
        st.info("Carga el Catálogo de Bancos para generar las transferencias.")
        return

    try:
        tabla_transferencias, bancos_sin_clave = crear_tabla_transferencias(
            tabla_corregida,
            archivo_bancos.getvalue(),
            archivo.getvalue(),
        )
        excel_transferencias = exportar_transferencias_excel(tabla_transferencias)
    except Exception as error:
        st.error(f"No se pudo preparar la tabla de transferencias: {error}")
        return

    total_transferencias = float(tabla_transferencias["Importe"].sum())
    bloques_csi = tabla_transferencias.loc[
        (tabla_transferencias["CIA"].map(_texto_normalizado) == "csi")
        & tabla_transferencias["No Bloque"].str.contains(" DE ", na=False),
        "No Bloque",
    ].nunique()
    metrica_transferencias_1, metrica_transferencias_2, metrica_transferencias_3 = st.columns(3)
    metrica_transferencias_1.metric("Transferencias", len(tabla_transferencias))
    metrica_transferencias_2.metric("Total", f"$ {total_transferencias:,.2f}")
    metrica_transferencias_3.metric("Bloques CSI", bloques_csi)

    if bancos_sin_clave:
        st.warning(
            "No se encontró CLAVE TRANSFER para: " + ", ".join(bancos_sin_clave)
        )

    busqueda_transferencias = st.text_input(
        "Buscar en la tabla de transferencias",
        placeholder="Empresa, bloque, clave, nombre, banco, CLABE...",
        key="buscar_tabla_transferencias_litografia",
    )
    transferencias_visibles = filtrar_tabla(
        tabla_transferencias, busqueda_transferencias
    )
    if busqueda_transferencias.strip():
        st.caption(f"Se encontraron {len(transferencias_visibles)} registro(s).")
    st.dataframe(
        transferencias_visibles,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Clave": st.column_config.TextColumn(),
            "CLABE": st.column_config.TextColumn(),
            "CUENTA": st.column_config.TextColumn(),
            "Importe": st.column_config.NumberColumn(format="$ %.2f"),
        },
    )
    st.download_button(
        "Descargar tabla de transferencias",
        data=excel_transferencias,
        file_name=f"transferencias_litografia_{fecha_descarga}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descargar_transferencias_litografia",
    )
