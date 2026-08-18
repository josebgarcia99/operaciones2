"""Apariencia SIO para la app de comprobantes.

Lo que el tema de Streamlit ya cubre —colores, tipografías, radios— vive en
`.streamlit/config.toml`. Aquí queda lo que ese tema no alcanza y que sí
distingue al panel de SIO:

  * la barra de marca azul marino con el logo,
  * las pestañas con la forma de los enlaces del menú lateral,
  * el encabezado de página (`.sio-heading`) y las tarjetas,
  * el pie con la razón social.

La paleta es la de `sio_islas/app/globals.css`, que a su vez sale de
`sio-tracking`: navy #1b2958 y azul de acción #003da6. Igual que allá, el
modo oscuro se decide con `prefers-color-scheme` y no con un interruptor.
"""

from __future__ import annotations

import base64
from functools import lru_cache
from pathlib import Path

import streamlit as st

ASSETS = Path(__file__).parent / "assets"

NOMBRE_SISTEMA = "Sistema Integral de Operaciones"
RAZON_SOCIAL = "Islas Gower & Compañía Sucesores S. en C. de C.V."

# Iconos. El panel de SIO dibuja los suyos como SVG de un solo trazo
# (`nav-secciones.tsx`: 24×24, `stroke-width` 1.7, sin relleno). En Streamlit
# no se puede meter HTML en la etiqueta de una pestaña ni de un botón, así que
# aquí se usan los Material Symbols que el propio Streamlit trae: también son
# de contorno y toman el color del texto. El CSS los adelgaza para que peguen
# con el grosor de los de allá.
#
# Nombrados por lo que significan, no por el glifo: si mañana cambia el juego
# de iconos, se cambia aquí y no en cada llamada.
ICONO = {
    "verificar": ":material/search:",
    "recibo": ":material/receipt_long:",
    "descargar": ":material/download:",
    "guardar": ":material/save:",
    "ok": ":material/check_circle:",
    "error": ":material/cancel:",
    "aviso": ":material/warning:",
    "bloqueado": ":material/lock:",
    "pendiente": ":material/schedule:",
}


@lru_cache(maxsize=None)
def _data_uri(nombre: str) -> str:
    """El logo incrustado en el CSS: así no depende de que se sirva un archivo."""
    datos = base64.b64encode((ASSETS / nombre).read_bytes()).decode("ascii")
    return f"data:image/png;base64,{datos}"


def config_pagina(titulo: str) -> None:
    """`set_page_config` con el logo de la casa como favicon."""
    icono = ASSETS / "logo.png"
    st.set_page_config(
        page_title=f"{titulo} · SIO",
        page_icon=str(icono) if icono.exists() else ICONO["recibo"],
        layout="wide",
    )


def _css() -> str:
    return f"""
<style>
/* ==========================================================================
   Sistema de diseño SIO — tokens de marca
   ========================================================================== */
:root {{
    --sio-navy: #1b2958;
    --sio-navy-700: #223466;
    --sio-blue: #003da6;
    --sio-blue-600: #0a4fc4;
    --sio-blue-100: #d6e4ff;
    --sio-blue-050: #eef4ff;

    --sio-surface: #f5f9ff;
    --sio-card: #ffffff;
    --sio-line: #e4e9f2;
    --sio-line-soft: #eef1f7;

    --sio-ink: #16233f;
    --sio-ink-soft: #5b6781;
    --sio-ink-muted: #8f98ab;

    --sio-shadow-sm: 0 1px 2px rgba(27, 41, 88, 0.06);
    --sio-shadow: 0 2px 8px rgba(27, 41, 88, 0.08);
    --sio-shadow-lg: 0 12px 28px rgba(27, 41, 88, 0.12);

    --sio-radius: 0.75rem;
    --sio-radius-lg: 1rem;

    /* La cabecera de SIO mide 60px, pero allá el logo vive en la barra
       lateral y aquí comparte renglón: a 60px el nombre de la casa deja de
       leerse, así que la barra crece lo justo para que quepa. */
    --sio-header-h: 72px;
    --sio-logo-w: 132px;
    --sio-logo-h: 52px;
}}

@media (prefers-color-scheme: dark) {{
    :root {{
        --sio-navy: #101a33;
        --sio-navy-700: #1b2a4d;

        --sio-surface: #0b1020;
        --sio-card: #141b2d;
        --sio-line: #26304a;
        --sio-line-soft: #1d2539;

        --sio-ink: #e8ecf6;
        --sio-ink-soft: #a7b0c4;
        --sio-ink-muted: #7b859c;

        --sio-blue-050: #16224a;
        --sio-blue-100: #1d2f63;

        --sio-shadow-sm: 0 1px 2px rgba(0, 0, 0, 0.4);
        --sio-shadow: 0 2px 8px rgba(0, 0, 0, 0.45);
        --sio-shadow-lg: 0 12px 28px rgba(0, 0, 0, 0.5);
    }}
}}

/* ==========================================================================
   Barra de marca. Streamlit ya trae una cabecera fija de 60px —la misma
   altura que la de SIO—, así que se aprovecha ésa en lugar de añadir otra:
   una segunda barra dejaría el logo desplazándose con la página.
   ========================================================================== */

header[data-testid="stHeader"] {{
    height: var(--sio-header-h);
    background: var(--sio-navy) !important;
    border-bottom: 1px solid var(--sio-navy-700);
    box-shadow: var(--sio-shadow-sm);
}}

/* El logo va en un pseudo-elemento del contenedor de la cabecera: el hueco
   de la izquierda está libre porque esta app no usa barra lateral. */
header[data-testid="stHeader"]::before {{
    content: "";
    position: absolute;
    left: 1.5rem;
    top: 50%;
    transform: translateY(-50%);
    width: var(--sio-logo-w);
    height: var(--sio-logo-h);
    background-image: url("{_data_uri('logo-blanco.png')}");
    background-repeat: no-repeat;
    background-position: left center;
    background-size: contain;
}}

/* El nombre del sistema, al lado del logo. Se calla en pantallas angostas,
   donde el logo ya identifica de sobra. */
header[data-testid="stHeader"]::after {{
    content: "{NOMBRE_SISTEMA}";
    position: absolute;
    left: calc(1.5rem + var(--sio-logo-w) + 0.875rem);
    top: 50%;
    transform: translateY(-50%);
    padding-left: 0.875rem;
    border-left: 1px solid rgba(255, 255, 255, 0.18);
    font-size: 0.8125rem;
    font-weight: 600;
    line-height: 1.15;
    color: rgba(255, 255, 255, 0.62);
    white-space: nowrap;
}}

@media (max-width: 720px) {{
    header[data-testid="stHeader"]::after {{ content: none; }}
    header[data-testid="stHeader"]::before {{ left: 1rem; }}
}}

/* Los controles propios de Streamlit, legibles sobre el azul marino */
header[data-testid="stHeader"] [data-testid="stToolbar"] button,
header[data-testid="stHeader"] [data-testid="stToolbar"] a,
header[data-testid="stHeader"] [data-testid="stToolbar"] span {{
    color: rgba(255, 255, 255, 0.82) !important;
    fill: rgba(255, 255, 255, 0.82) !important;
}}

header[data-testid="stHeader"] [data-testid="stToolbar"] button:hover {{
    background: var(--sio-navy-700) !important;
    color: #fff !important;
}}

/* ==========================================================================
   Lienzo y encabezado de página
   ========================================================================== */

[data-testid="stMainBlockContainer"] {{
    padding-top: calc(var(--sio-header-h) + 1.5rem);
    padding-bottom: 1.5rem;
    max-width: 1400px;

    /* La columna ocupa la ventana entera aunque la página quede corta. Sin
       esto el pie se queda pegado al último elemento, a media pantalla. */
    display: flex;
    flex-direction: column;
    min-height: 100dvh;
}}

/* El bloque de contenido se lleva el alto sobrante; lo que venga después
   —el pie— queda abajo. `1 0 auto` y no `1`: encogerlo haría que una página
   larga se comprimiera en lugar de desplazarse. */
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] {{
    flex: 1 0 auto;
}}

.sio-heading {{
    margin-bottom: 1.25rem;
}}

.sio-heading__title {{
    font-size: 1.375rem;
    font-weight: 700;
    letter-spacing: -0.01em;
    line-height: 1.25;
    color: var(--sio-navy);
}}

@media (prefers-color-scheme: dark) {{
    .sio-heading__title {{ color: var(--sio-ink); }}
}}

.sio-heading__subtitle {{
    margin-top: 0.1875rem;
    font-size: 0.875rem;
    line-height: 1.5;
    color: var(--sio-ink-soft);
}}

/* Los títulos de sección de la app (`st.subheader`) toman el navy de marca */
[data-testid="stHeading"] h2,
[data-testid="stHeading"] h3 {{
    color: var(--sio-navy);
    letter-spacing: -0.01em;
}}

@media (prefers-color-scheme: dark) {{
    [data-testid="stHeading"] h2,
    [data-testid="stHeading"] h3 {{ color: var(--sio-ink); }}
}}

[data-testid="stCaptionContainer"] {{
    color: var(--sio-ink-soft);
}}

/* ==========================================================================
   Iconos. Los Material Symbols de Streamlit vienen con el trazo grueso y
   parte del juego relleno; los de SIO son de contorno y están dibujados a
   1.7, así que se adelgazan y se vacían para que peguen con ellos.
   ========================================================================== */

[data-testid="stIconMaterial"] {{
    font-variation-settings: "FILL" 0, "wght" 300, "GRAD" 0, "opsz" 24;
}}

/* ==========================================================================
   Pestañas: la forma de los enlaces del menú de SIO (.sio-navlink).
   Streamlit las arma con react-aria, así que son `div[data-testid="stTab"]`
   dentro de un `[role="tablist"]`, no botones.
   ========================================================================== */

[data-testid="stTabs"] [role="tablist"] {{
    gap: 0.25rem;
    padding: 0.25rem;
    margin-bottom: 1.25rem;
    border: 1px solid var(--sio-line);
    border-bottom: 1px solid var(--sio-line);
    border-radius: var(--sio-radius);
    background: var(--sio-card);
    box-shadow: var(--sio-shadow-sm);
}}

[data-testid="stTab"] {{
    padding: 0.5625rem 0.875rem;
    border-radius: 0.625rem;
    color: var(--sio-ink-soft);
    transition: background 0.15s ease, color 0.15s ease;
}}

[data-testid="stTab"] [data-testid="stMarkdownContainer"] p {{
    font-size: 0.875rem;
    font-weight: 600;
}}

[data-testid="stTab"]:hover {{
    background: var(--sio-blue-050);
    color: var(--sio-blue);
}}

[data-testid="stTab"][aria-selected="true"] {{
    background: var(--sio-blue);
    color: #fff;
    box-shadow: var(--sio-shadow-sm);
}}

/* El color del rótulo lo pone el markdown de adentro, no el contenedor */
[data-testid="stTab"][aria-selected="true"] [data-testid="stMarkdownContainer"] p {{
    color: #fff;
}}

[data-testid="stTab"]:hover:not([aria-selected="true"]) [data-testid="stMarkdownContainer"] p {{
    color: var(--sio-blue);
}}

/* El subrayado deslizante sobra: aquí el estado activo es el bloque azul */
[data-testid="stTabs"] .react-aria-SelectionIndicator {{
    display: none;
}}

/* ==========================================================================
   Tarjetas: métricas y desplegables
   ========================================================================== */

[data-testid="stMetric"] {{
    padding: 1rem 1.125rem;
    border: 1px solid var(--sio-line);
    border-radius: var(--sio-radius-lg);
    background: var(--sio-card);
    box-shadow: var(--sio-shadow-sm);
}}

[data-testid="stMetricLabel"] {{
    font-size: 0.75rem;
    font-weight: 700;
    color: var(--sio-ink-soft);
}}

[data-testid="stMetricValue"] {{
    color: var(--sio-navy);
    letter-spacing: -0.02em;
}}

@media (prefers-color-scheme: dark) {{
    [data-testid="stMetricValue"] {{ color: var(--sio-ink); }}
}}

[data-testid="stExpander"] details {{
    border: 1px solid var(--sio-line);
    border-radius: var(--sio-radius-lg);
    background: var(--sio-card);
    box-shadow: var(--sio-shadow-sm);
    overflow: hidden;
}}

[data-testid="stExpander"] summary {{
    font-size: 0.9375rem;
    font-weight: 700;
    color: var(--sio-navy);
}}

[data-testid="stExpander"] summary:hover {{
    background: var(--sio-blue-050);
    color: var(--sio-blue);
}}

@media (prefers-color-scheme: dark) {{
    [data-testid="stExpander"] summary {{ color: var(--sio-ink); }}
}}

/* ==========================================================================
   Cargador de archivos: el marco punteado del panel
   ========================================================================== */

[data-testid="stFileUploaderDropzone"] {{
    border: 1px dashed var(--sio-line);
    border-radius: var(--sio-radius);
    background: var(--sio-card);
    transition: border-color 0.15s ease, background 0.15s ease;
}}

[data-testid="stFileUploaderDropzone"]:hover {{
    border-color: var(--sio-blue);
    background: var(--sio-blue-050);
}}

[data-testid="stFileUploaderDropzoneInstructions"] span {{
    color: var(--sio-ink-muted);
}}

/* ==========================================================================
   Botones (.sio-button) y avisos
   ========================================================================== */

[data-testid="stDownloadButton"] button,
[data-testid="stButton"] button {{
    font-weight: 600;
    box-shadow: var(--sio-shadow-sm);
}}

[data-testid="stAlert"] {{
    border-radius: var(--sio-radius);
}}

/* El bloque de texto plano del TXT de Banxico: ficha, no párrafo suelto */
[data-testid="stCode"] {{
    border: 1px solid var(--sio-line);
    border-radius: var(--sio-radius);
}}

[data-testid="stDataFrame"] {{
    border-radius: var(--sio-radius);
    overflow: hidden;
}}

/* ==========================================================================
   Pie de página (.sio-sidebar__footer)
   ========================================================================== */

/* `margin-top: auto` se come el hueco que sobra y empuja el pie al fondo.
   Cuando la página es larga no sobra nada, así que el `padding-top` es el que
   garantiza la separación con el último elemento en ese caso. */
[data-testid="stElementContainer"]:has(.sio-pie) {{
    margin-top: auto;
    padding-top: 2.5rem;
}}

.sio-pie {{
    padding-top: 1rem;
    border-top: 1px solid var(--sio-line);
    font-size: 0.6875rem;
    line-height: 1.6;
    color: var(--sio-ink-muted);
}}
</style>
"""


def aplicar() -> None:
    """Inyecta el CSS de marca. Se llama una sola vez, al inicio del script."""
    st.html(_css())


def encabezado(titulo: str, subtitulo: str = "") -> None:
    """El encabezado de página del panel: título en navy y una línea de apoyo."""
    sub = f'<p class="sio-heading__subtitle">{subtitulo}</p>' if subtitulo else ""
    st.html(
        f'<div class="sio-heading">'
        f'<h1 class="sio-heading__title">{titulo}</h1>'
        f"{sub}"
        f"</div>"
    )


def pie() -> None:
    """Cierra la página con la razón social, como el pie del menú de SIO."""
    st.html(f'<p class="sio-pie">{NOMBRE_SISTEMA} · {RAZON_SOCIAL}</p>')
