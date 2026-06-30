from __future__ import annotations

import re
import unicodedata
from io import BytesIO

import pandas as pd


REQUIRED_COLUMNS = {
    "NIVEL": ["NIVEL"],
    "TRANSACCIONAL": ["TRANSACCIONAL"],
    "CODIGO_CUENTA": ["CODIGO CUENTA CONTABLE", "CODIGO CUENTA", "CUENTA CONTABLE", "CODIGO"],
    "NOMBRE_CUENTA": ["NOMBRE CUENTA CONTABLE", "NOMBRE CUENTA", "CUENTA"],
    "IDENTIFICACION": ["IDENTIFICACION", "NIT", "DOCUMENTO"],
    "SUCURSAL": ["SUCURSAL"],
    "NOMBRE_TERCERO": ["NOMBRE TERCERO", "TERCERO"],
    "SALDO_INICIAL": ["SALDO INICIAL"],
    "MOVIMIENTO_DEBITO": ["MOVIMIENTO DEBITO", "MOVIMIENTO DÉBITO", "DEBITO", "DÉBITO"],
    "MOVIMIENTO_CREDITO": ["MOVIMIENTO CREDITO", "MOVIMIENTO CRÉDITO", "CREDITO", "CRÉDITO"],
    "SALDO_FINAL": ["SALDO FINAL"],
}

MONTHS = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}


def normalize_text(value: object) -> str:
    text = "" if pd.isna(value) else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^A-Za-z0-9]+", " ", text.upper()).strip()
    return re.sub(r"\s+", " ", text)


def clean_money(value: object) -> float:
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or normalize_text(text) in {"NAN", "NONE", "NULL", "-"}:
        return 0.0
    negative = text.startswith("-") or bool(re.match(r"^\(.*\)$", text))
    text = re.sub(r"[^0-9,.\-]", "", text).replace("-", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        text = "".join(parts) if len(parts[-1]) == 3 else text.replace(",", ".")
    elif "." in text:
        parts = text.split(".")
        if len(parts) > 1 and all(len(part) == 3 for part in parts[1:]):
            text = "".join(parts)
    try:
        number = float(text)
    except ValueError:
        return 0.0
    return -number if negative else number


def _read_raw_excel(file: BytesIO) -> tuple[str, pd.DataFrame, dict[str, pd.DataFrame]]:
    sheets = pd.read_excel(file, sheet_name=None, header=None, dtype=object, engine="openpyxl")
    best_name = ""
    best_df = pd.DataFrame()
    best_score = -1
    expected = {alias for aliases in REQUIRED_COLUMNS.values() for alias in aliases}
    for name, raw in sheets.items():
        if raw.dropna(how="all").empty:
            continue
        for idx in range(min(30, len(raw))):
            values = [normalize_text(value) for value in raw.iloc[idx].tolist()]
            score = sum(value in {normalize_text(alias) for alias in expected} for value in values)
            if score > best_score:
                best_score = score
                best_name = name
                best_df = raw
    if best_score < 4:
        raise ValueError("No se encontro una fila de encabezados valida para el balance de prueba.")
    return best_name, best_df, sheets


def _detect_header_row(raw: pd.DataFrame) -> int:
    expected = {normalize_text(alias) for aliases in REQUIRED_COLUMNS.values() for alias in aliases}
    best_idx = 0
    best_score = -1
    for idx in range(min(30, len(raw))):
        values = [normalize_text(value) for value in raw.iloc[idx].tolist()]
        score = sum(value in expected for value in values)
        if score > best_score:
            best_idx = idx
            best_score = score
    return best_idx


def _standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    normalized_cols = {normalize_text(col): col for col in df.columns}
    rename = {}
    for standard, aliases in REQUIRED_COLUMNS.items():
        for alias in aliases:
            key = normalize_text(alias)
            if key in normalized_cols:
                rename[normalized_cols[key]] = standard
                break
    return df.rename(columns=rename)


def extract_tlg_metadata(file: BytesIO) -> dict[str, str | None]:
    file.seek(0)
    sheets = pd.read_excel(file, sheet_name=None, header=None, dtype=object, nrows=30, engine="openpyxl")
    texts: list[str] = []
    for raw in sheets.values():
        for value in raw.to_numpy().flatten():
            if not pd.isna(value):
                texts.append(str(value))
    joined = "\n".join(texts)
    normalized = normalize_text(joined)
    nit_match = re.search(r"(?:NIT|N I T|IDENTIFICACION)\D*([0-9][0-9.\- ]{5,})", joined, re.IGNORECASE)
    period_match = re.search(
        r"(?:DE|DEL)\s+([A-Za-zÁÉÍÓÚáéíóú]+)\s+(\d{4})\s+(?:A|AL)\s+([A-Za-zÁÉÍÓÚáéíóú]+)\s+(\d{4})",
        joined,
        re.IGNORECASE,
    )
    company = "THE LATAM GROUP" if "THE LATAM GROUP" in normalized else None
    period = period_match.group(0) if period_match else None
    month = period_match.group(3).title() if period_match else None
    year = period_match.group(4) if period_match else None
    return {
        "empresa": company,
        "nit": nit_match.group(1).strip() if nit_match else None,
        "periodo": period,
        "mes": month,
        "anio": year,
    }


def load_tlg_trial_balance(file: BytesIO) -> tuple[pd.DataFrame, dict[str, str | None]]:
    metadata = extract_tlg_metadata(file)
    file.seek(0)
    sheet_name, raw, _ = _read_raw_excel(file)
    header_idx = _detect_header_row(raw)
    df = raw.iloc[header_idx + 1 :].copy()
    df.columns = raw.iloc[header_idx].fillna("").astype(str).tolist()
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df = _standardize_columns(df)
    missing = [col for col in ["CODIGO_CUENTA", "NOMBRE_CUENTA", "SALDO_FINAL"] if col not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias en el balance: {', '.join(missing)}.")
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    for col in ["SALDO_INICIAL", "MOVIMIENTO_DEBITO", "MOVIMIENTO_CREDITO", "SALDO_FINAL"]:
        df[col] = df[col].map(clean_money)
    df["CODIGO_CUENTA"] = df["CODIGO_CUENTA"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["CLASE"] = df["CODIGO_CUENTA"].str[0]
    df["HOJA_ORIGEN"] = sheet_name
    df = df[df["CODIGO_CUENTA"].str.len() > 0].copy()
    return df.reset_index(drop=True), metadata


def validate_tlg_company(metadata: dict[str, str | None], confirmed_by_user: bool = False) -> tuple[bool, str]:
    if metadata.get("empresa") and "THE LATAM GROUP" in normalize_text(metadata.get("empresa")):
        return True, "Empresa detectada como THE LATAM GROUP."
    if confirmed_by_user:
        return True, "El usuario confirmo visualmente que el archivo pertenece a TLG."
    return False, "El archivo cargado no parece corresponder a TLG. Por seguridad, no se actualizaran los Estados Financieros."
