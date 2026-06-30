from __future__ import annotations

import pandas as pd


CLASS_NAMES = {
    "1": "Activo",
    "2": "Pasivo",
    "3": "Patrimonio",
    "4": "Ingresos",
    "5": "Gastos",
    "6": "Costos de ventas",
    "7": "Costos de produccion u operacion",
    "8": "Cuentas de orden deudoras",
    "9": "Cuentas de orden acreedoras",
}


def classify_tlg_accounts(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["CLASE"] = out["CODIGO_CUENTA"].astype(str).str[0]
    out["CLASIFICACION"] = out["CLASE"].map(CLASS_NAMES).fillna("Sin clasificar")
    return out


def _detail_level(df: pd.DataFrame) -> pd.DataFrame:
    if "TRANSACCIONAL" in df.columns:
        trans = df["TRANSACCIONAL"].astype(str).str.upper().str.strip()
        filtered = df[trans.isin(["SI", "SÍ", "TRUE", "1", "X"])].copy()
        if not filtered.empty:
            return filtered
    if "NIVEL" in df.columns:
        level = df["NIVEL"].astype(str).str.upper()
        filtered = df[level.str.contains("AUXILIAR|SUBCUENTA|CUENTA", regex=True, na=False)].copy()
        if not filtered.empty:
            max_len = filtered["CODIGO_CUENTA"].astype(str).str.len().max()
            return filtered[filtered["CODIGO_CUENTA"].astype(str).str.len() == max_len].copy()
    max_len = df["CODIGO_CUENTA"].astype(str).str.len().max()
    return df[df["CODIGO_CUENTA"].astype(str).str.len() == max_len].copy()


def prepare_tlg_detail(df: pd.DataFrame) -> pd.DataFrame:
    classified = classify_tlg_accounts(df)
    detail = _detail_level(classified)
    grouped = (
        detail.groupby(["CODIGO_CUENTA", "NOMBRE_CUENTA", "CLASE", "CLASIFICACION"], as_index=False)
        .agg(
            SALDO_INICIAL=("SALDO_INICIAL", "sum"),
            MOVIMIENTO_DEBITO=("MOVIMIENTO_DEBITO", "sum"),
            MOVIMIENTO_CREDITO=("MOVIMIENTO_CREDITO", "sum"),
            SALDO_FINAL=("SALDO_FINAL", "sum"),
        )
        .sort_values("CODIGO_CUENTA")
    )
    return grouped


def build_tlg_financial_summary(df: pd.DataFrame) -> dict[str, object]:
    detail = prepare_tlg_detail(df)
    by_class = detail.groupby(["CLASE", "CLASIFICACION"], as_index=False)["SALDO_FINAL"].sum()

    total_activo = by_class.loc[by_class["CLASE"] == "1", "SALDO_FINAL"].sum()
    total_pasivo = abs(by_class.loc[by_class["CLASE"] == "2", "SALDO_FINAL"].sum())
    total_patrimonio = abs(by_class.loc[by_class["CLASE"] == "3", "SALDO_FINAL"].sum())
    total_ingresos = abs(by_class.loc[by_class["CLASE"] == "4", "SALDO_FINAL"].sum())
    total_gastos = abs(by_class.loc[by_class["CLASE"] == "5", "SALDO_FINAL"].sum())
    total_costos = abs(by_class.loc[by_class["CLASE"].isin(["6", "7"]), "SALDO_FINAL"].sum())
    resultado = total_ingresos - total_costos - total_gastos
    diferencia_cuadre = total_activo - (total_pasivo + total_patrimonio + resultado)

    balance = detail[detail["CLASE"].isin(["1", "2", "3"])].copy()
    income_statement = detail[detail["CLASE"].isin(["4", "5", "6", "7"])].copy()

    return {
        "detail": detail,
        "by_class": by_class,
        "balance": balance,
        "income_statement": income_statement,
        "metrics": {
            "total_activo": total_activo,
            "total_pasivo": total_pasivo,
            "total_patrimonio": total_patrimonio,
            "total_ingresos": total_ingresos,
            "total_costos": total_costos,
            "total_gastos": total_gastos,
            "resultado": resultado,
            "diferencia_cuadre": diferencia_cuadre,
            "total_saldo_inicial": detail["SALDO_INICIAL"].sum(),
            "total_debito": detail["MOVIMIENTO_DEBITO"].sum(),
            "total_credito": detail["MOVIMIENTO_CREDITO"].sum(),
            "total_saldo_final": detail["SALDO_FINAL"].sum(),
            "cuentas": detail["CODIGO_CUENTA"].nunique(),
        },
    }


def build_tlg_management_text(metrics: dict[str, float], metadata: dict[str, str | None]) -> str:
    result = metrics.get("resultado", 0)
    tone = "utilidad" if result >= 0 else "perdida"
    return (
        f"Para el periodo {metadata.get('periodo') or 'cargado'}, TLG presenta una {tone} estimada de "
        f"${abs(result):,.0f}. Los activos suman ${metrics.get('total_activo', 0):,.0f}, frente a pasivos por "
        f"${metrics.get('total_pasivo', 0):,.0f} y patrimonio por ${metrics.get('total_patrimonio', 0):,.0f}. "
        f"Los ingresos del periodo ascienden a ${metrics.get('total_ingresos', 0):,.0f}, con costos por "
        f"${metrics.get('total_costos', 0):,.0f} y gastos por ${metrics.get('total_gastos', 0):,.0f}. "
        "Se recomienda revisar las cuentas con mayores saldos, validar terceros relevantes y confirmar el cuadre antes de emitir estados financieros definitivos."
    ).replace(",", ".")
