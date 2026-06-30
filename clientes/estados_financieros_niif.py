from __future__ import annotations

import pandas as pd
import streamlit as st

from utils.niif_reports import build_niif_report, export_niif_excel, export_niif_pdf, money
from utils.tlg_data_cleaning import load_tlg_trial_balance
from utils.ui_components import info_panel, process_steps, section_header


def _metric_grid(metrics: dict[str, float]) -> None:
    rows = [
        ("Activo", metrics.get("total_assets", 0)),
        ("Pasivo", metrics.get("total_liabilities", 0)),
        ("Patrimonio", metrics.get("total_equity", 0)),
        ("Resultado", metrics.get("profit", 0)),
        ("Ingresos", metrics.get("total_income", 0)),
        ("Costos", metrics.get("total_costs", 0)),
        ("Gastos", metrics.get("total_expenses", 0)),
        ("Diferencia cuadre", metrics.get("check_difference", 0)),
    ]
    for start in range(0, len(rows), 4):
        cols = st.columns(4)
        for col, (label, value) in zip(cols, rows[start : start + 4]):
            col.metric(label, money(value))


def _money_columns(df: pd.DataFrame) -> dict[str, object]:
    return {
        column: st.column_config.NumberColumn(format="$ %.0f")
        for column in df.columns
        if column.lower() in {"valor", "saldo_inicial", "movimiento_debito", "movimiento_credito", "saldo_final"}
    }


def render_niif_financial_statements() -> None:
    section_header(
        "Estados financieros bajo NIIF",
        "Construye un borrador estructurado con estado de situacion financiera, resultado integral, flujo de efectivo, cambios en patrimonio, notas y control de preparacion.",
        eyebrow="Cumplimiento financiero",
        badge="NIIF Colombia",
    )
    process_steps(
        [
            "Carga el balance de prueba",
            "Clasifica y valida el cuadre",
            "Revisa estados, notas y descarga",
        ]
    )

    left, right = st.columns([1, 1], gap="large")
    with left:
        company_name = st.text_input("Nombre de la entidad", value="Entidad reportante")
        standard_group = st.selectbox(
            "Marco NIIF aplicable",
            ["Grupo 2 - NIIF para Pymes", "Grupo 1 - NIIF plenas", "Grupo 3 - Microempresas"],
        )
    with right:
        balance_file = st.file_uploader(
            "Balance de prueba por tercero",
            type=["xlsx"],
            key="niif_trial_balance_uploader",
            help="Debe contener codigo, nombre de cuenta y saldos finales. Puede venir del mismo formato usado para TLG.",
        )
        info_panel(
            "Alcance del borrador",
            "El sistema arma la estructura y los controles principales. Antes de emitir se deben completar politicas, revelaciones, firmas y autorizaciones.",
        )

    if balance_file is None:
        st.info("Carga un balance de prueba para iniciar la preparacion NIIF.")
        return

    try:
        raw_df, metadata = load_tlg_trial_balance(balance_file)
        report = build_niif_report(
            raw_df,
            metadata,
            company_name=company_name.strip() or "Entidad reportante",
            standard_group=standard_group,
        )
    except Exception as exc:
        st.error(f"No fue posible preparar los estados financieros: {exc}")
        return

    metrics = report["metrics"]
    tolerance = max(1000.0, abs(metrics.get("total_assets", 0)) * 0.00001)
    st.success("Balance leido y clasificado correctamente.")
    _metric_grid(metrics)
    if abs(metrics.get("check_difference", 0)) <= tolerance:
        st.success("Validacion de cuadre aprobada dentro de la tolerancia tecnica.")
    else:
        st.warning("Hay diferencia de cuadre. Revisa cierre de resultado, patrimonio y saldos contables antes de emitir.")

    tab_esf, tab_result, tab_cash, tab_equity, tab_notes, tab_control, tab_download = st.tabs(
        [
            "Situacion financiera",
            "Resultado integral",
            "Flujo efectivo",
            "Cambios patrimonio",
            "Notas",
            "Control NIIF",
            "Descargas",
        ]
    )

    with tab_esf:
        st.dataframe(
            report["financial_position"],
            width="stretch",
            hide_index=True,
            column_config=_money_columns(report["financial_position"]),
        )

    with tab_result:
        st.dataframe(
            report["comprehensive_income"],
            width="stretch",
            hide_index=True,
            column_config=_money_columns(report["comprehensive_income"]),
        )

    with tab_cash:
        st.caption("Metodo indirecto preliminar. Para estados definitivos se requiere balance comparativo y conciliacion de variaciones.")
        st.dataframe(
            report["cash_flow"],
            width="stretch",
            hide_index=True,
            column_config=_money_columns(report["cash_flow"]),
        )

    with tab_equity:
        st.dataframe(
            report["equity_changes"],
            width="stretch",
            hide_index=True,
            column_config=_money_columns(report["equity_changes"]),
        )

    with tab_notes:
        st.dataframe(report["notes"], width="stretch", hide_index=True)

    with tab_control:
        st.dataframe(report["checklist"], width="stretch", hide_index=True)
        with st.expander("Balance limpio y clasificado", expanded=False):
            st.dataframe(
                report["detail"],
                width="stretch",
                hide_index=True,
                column_config=_money_columns(report["detail"]),
            )

    with tab_download:
        st.download_button(
            "Descargar estados financieros NIIF en Excel",
            data=export_niif_excel(report),
            file_name="Estados_financieros_NIIF.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
        )
        st.download_button(
            "Descargar resumen NIIF en PDF",
            data=export_niif_pdf(report),
            file_name="Resumen_estados_financieros_NIIF.pdf",
            mime="application/pdf",
            width="stretch",
        )
