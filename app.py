from __future__ import annotations

import streamlit as st

from clientes.estados_financieros_niif import render_niif_financial_statements
from clientes.mensualizados import render_monthly_reports
from utils.ui_components import (
    feature_grid,
    inject_app_styles,
    section_header,
    sidebar_brand,
    top_bar,
)


st.set_page_config(
    page_title="Informes financieros | Julio Salazar",
    layout="wide",
    initial_sidebar_state="expanded",
)
inject_app_styles()


def render_home() -> None:
    section_header(
        "Informes financieros",
        "Herramienta de trabajo para preparar comparativos mensuales y estados financieros bajo NIIF a partir de balances contables.",
        eyebrow="Panel ejecutivo",
        badge="Version inicial",
    )
    feature_grid(
        [
            (
                "Comparativo mensual",
                "Modulo basado en la estructura de TLG para cargar saldos iniciales, balances mensuales, indicadores, terceros y descargas.",
            ),
            (
                "Estados financieros NIIF",
                "Borrador estructurado con situacion financiera, resultado integral, flujo de efectivo, cambios en patrimonio, notas y control.",
            ),
            (
                "Validacion contable",
                "Revision de cuadre, clasificacion por cuenta y alertas para detectar diferencias antes de entregar.",
            ),
            (
                "Trazabilidad",
                "Cada informe conserva archivo fuente, periodo detectado, cuentas procesadas y detalle clasificado.",
            ),
            (
                "Descargas",
                "Exportacion a Excel y PDF para revision interna, ajustes y presentacion al cliente.",
            ),
            (
                "Escalable",
                "La base queda lista para agregar clientes, plantillas o reglas contables especificas mas adelante.",
            ),
        ]
    )


sidebar_brand()
st.sidebar.caption("NAVEGACION")
section = st.sidebar.radio(
    "Menu principal",
    ["Inicio", "Comparativo mensual", "Estados financieros NIIF"],
    label_visibility="collapsed",
)
st.sidebar.divider()
st.sidebar.caption("Version operativa 2026")

top_bar(section)
if section == "Inicio":
    render_home()
elif section == "Comparativo mensual":
    render_monthly_reports()
else:
    render_niif_financial_statements()
