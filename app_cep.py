"""App de comprobantes SPEI con dos vistas.

    streamlit run app_cep.py

Pestaña 1 — Verificar en Banxico: lo del día a día. Subes los comprobantes,
revisas lo que Banxico reporta de cada pago y, si son varios, te llevas el TXT
para la consulta por lotes. Nada más.

Pestaña 2 — cr3: la app de siempre, sin recortes (`cr3.py`). El Excel de
dispersión de esa pestaña funciona con o sin comprobantes: el PDF sólo sirve
para verificar que los totales cuadren.

Las dos pestañas son independientes: cada una tiene su propio cargador de
archivos, su propia tabla y sus propios resultados. Lo único que comparten es
el código, que vive en `cr3.py` y se llama desde aquí: la verificación en
Banxico aparece en las dos vistas, pero está escrita una sola vez.
"""

import cr3
import streamlit as st

st.set_page_config(page_title="Comprobantes SPEI - CEP Banxico", page_icon="💳", layout="wide")

# En la vista rápida sobra el titular del ordenante, que es para cuadrar contra
# los Excel de dispersión y no lo pide el portal del CEP.
COLUMNAS_BANXICO = [col for col in cr3.TABLE_COLS if col != "titular_ordenante"]


def vista_banxico():
    """Sólo verificar: leer los comprobantes y preguntarle a Banxico por cada pago."""
    st.title("Verificar pagos en Banxico")
    st.caption(
        "Sube los comprobantes y consulta el estado de cada pago en el portal del CEP. "
        "Para el proceso completo (Excel de dispersión, cruce con la base de tarjetas) "
        "usa la otra pestaña."
    )

    if not (cr3.RapidOCR and cr3.np) and not cr3.TESSERACT_READY:
        st.warning(
            "No hay motor de OCR instalado: los PDFs escaneados no se van a poder leer. "
            "Instálalo con `pip install rapidocr-onnxruntime`."
        )

    uploaded_files = st.file_uploader(
        "Sube uno o varios comprobantes en PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdfs_banxico",
    )

    if not uploaded_files:
        st.info("No hay comprobantes cargados. Sube un PDF para comenzar.")
        return

    rows, _debug = cr3.leer_comprobantes(uploaded_files)

    st.subheader("Datos que se van a consultar")
    st.caption(
        "Corrige aquí cualquier dato antes de consultar; si un dígito viene mal, "
        "Banxico no encuentra el pago. La fecha se captura como DD-MM-YYYY."
    )
    records = cr3.tabla_editable(rows, "tabla_banxico", COLUMNAS_BANXICO)

    sin_cep = cr3.revisar_registros(records)
    txt_content = cr3.build_txt(records)
    valid_lines = cr3.resumen_lote(txt_content, sin_cep)

    cr3.bloque_cep(records, valid_lines, "banxico")

    st.subheader("Consulta por lotes")
    st.caption(
        "Si son muchos pagos, conviene más el servicio por lotes de Banxico "
        "(de 2 a 500 por archivo): descarga el TXT y súbelo en "
        "https://www.banxico.org.mx/cep-scl/"
    )
    st.download_button(
        "⬇️ Descargar archivo .txt para Banxico",
        data=txt_content,
        file_name="transferencias_cep.txt",
        mime="text/plain",
        disabled=valid_lines == 0,
        key="txt_banxico",
    )
    st.code(txt_content or "Aún no hay renglones válidos.", language="text")


tab_banxico, tab_completo = st.tabs(["🔎 Verificar en Banxico", "🧾 cr3"])

with tab_banxico:
    vista_banxico()

with tab_completo:
    cr3.vista_completa()
