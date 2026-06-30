from __future__ import annotations

from io import BytesIO

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from utils.tlg_financial_statements import prepare_tlg_detail


NIIF_REFERENCES = [
    "Marco tecnico contable colombiano compilado en el Decreto Unico Reglamentario 2420 de 2015 y sus modificaciones.",
    "Presentacion de estados financieros conforme a NIIF/NIIF para Pymes, segun el grupo aplicable.",
    "Estados minimos: situacion financiera, resultado integral, cambios en patrimonio, flujos de efectivo y notas.",
]


def money(value: float) -> str:
    amount = float(value or 0)
    formatted = "$" + f"{abs(amount):,.0f}".replace(",", ".")
    return f"-{formatted}" if amount < 0 else formatted


def _sum_by_prefix(detail: pd.DataFrame, prefixes: tuple[str, ...], *, absolute: bool = True) -> float:
    values = detail.loc[detail["CODIGO_CUENTA"].astype(str).str.startswith(prefixes), "SALDO_FINAL"]
    total = float(values.sum())
    return abs(total) if absolute else total


def _section_rows(detail: pd.DataFrame, sections: list[tuple[str, tuple[str, ...]]]) -> pd.DataFrame:
    rows = []
    for section, prefixes in sections:
        section_detail = detail[detail["CODIGO_CUENTA"].astype(str).str.startswith(prefixes)].copy()
        if section_detail.empty:
            rows.append({"Rubro": section, "Codigo": "", "Cuenta": "Sin movimiento", "Valor": 0.0})
            continue
        grouped = (
            section_detail.groupby(["CODIGO_CUENTA", "NOMBRE_CUENTA"], as_index=False)["SALDO_FINAL"]
            .sum()
            .sort_values("CODIGO_CUENTA")
        )
        for _, row in grouped.iterrows():
            rows.append(
                {
                    "Rubro": section,
                    "Codigo": str(row["CODIGO_CUENTA"]),
                    "Cuenta": row["NOMBRE_CUENTA"],
                    "Valor": abs(float(row["SALDO_FINAL"])),
                }
            )
    return pd.DataFrame(rows)


def _liability_rows(detail: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    liabilities = detail[detail["CODIGO_CUENTA"].astype(str).str.startswith("2")].copy()
    if liabilities.empty:
        empty = pd.DataFrame(columns=["Rubro", "Codigo", "Cuenta", "Valor"])
        return empty, empty
    concept = liabilities["NOMBRE_CUENTA"].astype(str).str.upper()
    non_current_mask = concept.str.contains("LARGO PLAZO|NO CORRIENTE|LP", regex=True, na=False)
    current = liabilities[~non_current_mask]
    non_current = liabilities[non_current_mask]

    def grouped(frame: pd.DataFrame, label: str) -> pd.DataFrame:
        if frame.empty:
            return pd.DataFrame([{"Rubro": label, "Codigo": "", "Cuenta": "Sin movimiento", "Valor": 0.0}])
        out = frame.groupby(["CODIGO_CUENTA", "NOMBRE_CUENTA"], as_index=False)["SALDO_FINAL"].sum()
        out = out.sort_values("CODIGO_CUENTA")
        return pd.DataFrame(
            {
                "Rubro": label,
                "Codigo": out["CODIGO_CUENTA"].astype(str),
                "Cuenta": out["NOMBRE_CUENTA"],
                "Valor": out["SALDO_FINAL"].abs(),
            }
        )

    return grouped(current, "Pasivo corriente"), grouped(non_current, "Pasivo no corriente")


def build_niif_report(
    raw_df: pd.DataFrame,
    metadata: dict[str, str | None],
    *,
    company_name: str,
    standard_group: str,
) -> dict[str, object]:
    detail = prepare_tlg_detail(raw_df)
    current_assets = _section_rows(detail, [("Activo corriente", ("11", "12", "13", "14"))])
    non_current_assets = _section_rows(detail, [("Activo no corriente", ("15", "16", "17", "18", "19"))])
    current_liabilities, non_current_liabilities = _liability_rows(detail)
    equity = _section_rows(detail, [("Patrimonio", ("3",))])
    income = _section_rows(detail, [("Ingresos de actividades ordinarias", ("41",))])
    other_income = _section_rows(detail, [("Otros ingresos", ("42", "47"))])
    costs = _section_rows(detail, [("Costos", ("6", "7"))])
    expenses = _section_rows(detail, [("Gastos", ("5",))])

    total_assets = _sum_by_prefix(detail, ("1",))
    total_liabilities = _sum_by_prefix(detail, ("2",))
    total_equity = _sum_by_prefix(detail, ("3",))
    total_income = _sum_by_prefix(detail, ("4",))
    total_costs = _sum_by_prefix(detail, ("6", "7"))
    total_expenses = _sum_by_prefix(detail, ("5",))
    profit = total_income - total_costs - total_expenses
    check_difference = total_assets - (total_liabilities + total_equity + profit)

    financial_position = pd.concat(
        [current_assets, non_current_assets, current_liabilities, non_current_liabilities, equity],
        ignore_index=True,
    )
    comprehensive_income = pd.concat([income, other_income, costs, expenses], ignore_index=True)

    cash_flow = pd.DataFrame(
        [
            ("Resultado del periodo", profit),
            ("Variacion neta estimada de activos corrientes", -_sum_by_prefix(detail, ("11", "12", "13", "14"))),
            ("Variacion neta estimada de pasivos corrientes", _sum_by_prefix(detail, ("21", "22", "23", "24", "25", "26", "27", "28"))),
            ("Flujo neto estimado de operacion", profit - _sum_by_prefix(detail, ("11", "12", "13", "14")) + _sum_by_prefix(detail, ("21", "22", "23", "24", "25", "26", "27", "28"))),
            ("Flujo de inversion por clasificar", -_sum_by_prefix(detail, ("15", "16", "17", "18", "19"))),
            ("Flujo de financiacion por clasificar", _sum_by_prefix(detail, ("2", "3"))),
        ],
        columns=["Concepto", "Valor"],
    )

    equity_changes = pd.DataFrame(
        [
            ("Patrimonio contable de cierre", total_equity),
            ("Resultado del periodo", profit),
            ("Patrimonio ajustado con resultado", total_equity + profit),
        ],
        columns=["Concepto", "Valor"],
    )

    notes = pd.DataFrame(
        [
            ("Base de preparacion", f"Estados preparados para {company_name} bajo {standard_group}."),
            ("Moneda funcional", "Cifras tomadas del balance de prueba cargado; validar moneda y redondeo antes de emitir."),
            ("Politicas contables", "Completar politicas de reconocimiento, medicion inicial, medicion posterior y deterioro."),
            ("Juicios y estimaciones", "Documentar deterioro de cartera, vida util de activos, provisiones, impuestos y contingencias."),
            ("Hechos posteriores", "Confirmar hechos relevantes entre la fecha de corte y la fecha de autorizacion."),
            ("Revelaciones", "Cruzar saldos materiales con terceros, anexos, contratos, impuestos y conciliaciones auxiliares."),
        ],
        columns=["Nota", "Contenido"],
    )

    checklist = pd.DataFrame(
        [
            ("Estado de situacion financiera", bool(abs(check_difference) <= max(1000, total_assets * 0.00001)), "Validar cuadre contable."),
            ("Estado de resultado integral", bool(not comprehensive_income.empty), "Clasificar ingresos, costos, gastos y otros resultados integrales si aplican."),
            ("Estado de flujos de efectivo", True, "La version generada es indirecta y requiere validar variaciones reales de balance comparativo."),
            ("Estado de cambios en patrimonio", True, "Completar aportes, dividendos, reservas, ajustes de adopcion y otros movimientos."),
            ("Notas a los estados financieros", True, "Completar revelaciones cualitativas y cuantitativas antes de firmar."),
            ("Certificacion y aprobacion", False, "Pendiente firmas, autorizacion de emision y documentos societarios."),
        ],
        columns=["Requisito", "Preparado", "Observacion"],
    )

    return {
        "metadata": metadata,
        "company_name": company_name,
        "standard_group": standard_group,
        "detail": detail,
        "financial_position": financial_position,
        "comprehensive_income": comprehensive_income,
        "cash_flow": cash_flow,
        "equity_changes": equity_changes,
        "notes": notes,
        "checklist": checklist,
        "metrics": {
            "total_assets": total_assets,
            "total_liabilities": total_liabilities,
            "total_equity": total_equity,
            "total_income": total_income,
            "total_costs": total_costs,
            "total_expenses": total_expenses,
            "profit": profit,
            "check_difference": check_difference,
        },
    }


def export_niif_excel(report: dict[str, object]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame(
            [
                {
                    "Empresa": report["company_name"],
                    "Grupo NIIF": report["standard_group"],
                    **report["metadata"],
                }
            ]
        ).to_excel(writer, sheet_name="Portada", index=False)
        pd.DataFrame([report["metrics"]]).to_excel(writer, sheet_name="Indicadores", index=False)
        report["financial_position"].to_excel(writer, sheet_name="ESF", index=False)
        report["comprehensive_income"].to_excel(writer, sheet_name="Resultado integral", index=False)
        report["cash_flow"].to_excel(writer, sheet_name="Flujo efectivo", index=False)
        report["equity_changes"].to_excel(writer, sheet_name="Cambios patrimonio", index=False)
        report["notes"].to_excel(writer, sheet_name="Notas", index=False)
        report["checklist"].to_excel(writer, sheet_name="Control NIIF", index=False)
        report["detail"].to_excel(writer, sheet_name="Balance limpio", index=False)
    return output.getvalue()


def export_niif_pdf(report: dict[str, object]) -> bytes:
    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=letter,
        leftMargin=0.55 * inch,
        rightMargin=0.55 * inch,
        topMargin=0.55 * inch,
        bottomMargin=0.55 * inch,
    )
    styles = getSampleStyleSheet()
    metrics = report["metrics"]
    story = [
        Paragraph("Estados financieros bajo NIIF", styles["Title"]),
        Paragraph(str(report["company_name"]), styles["Heading2"]),
        Paragraph(f"Grupo aplicado: {report['standard_group']}", styles["Normal"]),
        Spacer(1, 0.15 * inch),
    ]
    table_data = [
        ["Indicador", "Valor"],
        ["Activo", money(metrics["total_assets"])],
        ["Pasivo", money(metrics["total_liabilities"])],
        ["Patrimonio", money(metrics["total_equity"])],
        ["Resultado", money(metrics["profit"])],
        ["Diferencia de cuadre", money(metrics["check_difference"])],
    ]
    table = Table(table_data, colWidths=[2.6 * inch, 2.0 * inch])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B6B57")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D0D5DD")),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(table)
    story.append(Spacer(1, 0.18 * inch))
    story.append(Paragraph("Referencias de preparacion", styles["Heading2"]))
    for item in NIIF_REFERENCES:
        story.append(Paragraph(f"- {item}", styles["BodyText"]))
    story.append(Spacer(1, 0.12 * inch))
    story.append(Paragraph("Este borrador requiere revision profesional, revelaciones completas y autorizaciones antes de emision.", styles["BodyText"]))
    doc.build(story)
    output.seek(0)
    return output.getvalue()
