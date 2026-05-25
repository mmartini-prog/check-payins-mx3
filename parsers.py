"""
parsers.py — Lectura y parsing de archivos Kyriba y Payins.

Responsabilidades:
  - Detección automática de header row
  - Extracción de columnas canónicas
  - Limpieza de montos
  - Identificación de Company/Account desde metadatos Kyriba
  - Sin dependencias de Streamlit (testeable aislado)
"""

from __future__ import annotations

import io
import re
import logging
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)


# ── COLUMNAS CANÓNICAS DE SALIDA ───────────────────────────────────────────────

KYRIBA_COLS = [
    "Account code", "Account ID", "Company",
    "Transaction date", "Description", "Complementary info",
    "Credit", "Debit", "Text to match",
    "Processor", "Day", "Source file",
]

PAYINS_COLS = [
    "Date", "Amount", "Processor", "Original Processor",
    "Collection Agent", "Payment Method Code", "Country",
    "Day", "Source file",
]


# ── RESULTADO DE PARSING ───────────────────────────────────────────────────────

@dataclass
class ParseResult:
    df: pd.DataFrame
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors and not self.df.empty


# ── UTILIDADES DE TEXTO ────────────────────────────────────────────────────────

def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().lower().replace("_", " ").replace("-", " ")


def clean_amount(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("MXN", "", regex=False)
        .str.replace("USD", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip(),
        errors="coerce",
    ).fillna(0)


# ── BÚSQUEDA DE COLUMNAS ───────────────────────────────────────────────────────

def find_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    """Busca la primera columna que coincida (exacta primero, luego parcial)."""
    norm_cols = {normalize_text(c): c for c in df.columns}

    for name in candidates:
        n = normalize_text(name)
        if n in norm_cols:
            return norm_cols[n]

    for name in candidates:
        n = normalize_text(name)
        for col_norm, original in norm_cols.items():
            if n in col_norm or col_norm in n:
                return original

    return None


# ── DETECCIÓN DE HEADER ROW ────────────────────────────────────────────────────

_KYRIBA_STRONG = {"account code", "transaction date"}
_PAYINS_STRONG = {"payment date", "total local amount"}
_GP_STRONG     = {"accounting date", "processor name"}

_HEADER_KEYWORDS = [
    "transaction date", "value date", "booking date", "accounting date",
    "payment date", "creation date", "date", "account code", "account id",
    "description", "complementary info", "credit", "debit", "amount",
    "total local amount", "total usd amount", "amount approved",
    "name", "processor", "processor name", "collection agent", "code",
]


def _read_raw(file_bytes: bytes, file_name: str, nrows: int = 60) -> pd.DataFrame:
    buf = io.BytesIO(file_bytes)
    if file_name.lower().endswith(".csv"):
        return pd.read_csv(buf, header=None, nrows=nrows, encoding="utf-8-sig")
    return pd.read_excel(buf, header=None, nrows=nrows)


def find_header_row(file_bytes: bytes, file_name: str) -> int:
    """Devuelve el índice de fila (0-based) donde están los headers."""
    raw = _read_raw(file_bytes, file_name)
    best_row, best_score = 0, 0

    for idx, row in raw.iterrows():
        cells = {normalize_text(v) for v in row.values}

        if _KYRIBA_STRONG.issubset(cells):
            return idx
        if _PAYINS_STRONG.issubset(cells):
            return idx
        if _GP_STRONG.issubset(cells):
            return idx

        score = sum(1 for kw in _HEADER_KEYWORDS if any(kw in c for c in cells))
        if score > best_score:
            best_score, best_row = score, idx

    return best_row if best_score >= 2 else 0


# ── LECTURA GENERAL ────────────────────────────────────────────────────────────

def read_table(file_bytes: bytes, file_name: str, dtype=None) -> pd.DataFrame:
    """Lee un archivo detectando automáticamente la fila de encabezado."""
    header_row = find_header_row(file_bytes, file_name)
    buf = io.BytesIO(file_bytes)

    if file_name.lower().endswith(".csv"):
        df = pd.read_csv(buf, header=header_row, encoding="utf-8-sig", dtype=dtype)
    else:
        df = pd.read_excel(buf, header=header_row, dtype=dtype)

    # Limpiar nombres de columnas
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    df = df.dropna(how="all")
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    return df


# ── METADATOS KYRIBA (Company, Account) ───────────────────────────────────────

_COMPANY_RE  = re.compile(r"Company:\s*(.+?)(?:\s*$)", re.IGNORECASE)
_ACCOUNT_RE  = re.compile(r"Account:\s*(.+?)(?:\s*$)", re.IGNORECASE)
_ACCT_ID_RE  = re.compile(r"Account Id:\s*(\S+)", re.IGNORECASE)
_CURRENCY_RE = re.compile(r"-\s*(MXN|USD|EUR)\s*$", re.IGNORECASE)


def _extract_kyriba_metadata(file_bytes: bytes, file_name: str) -> dict:
    """Extrae Company, Account y Currency de las primeras filas de un Kyriba."""
    raw = _read_raw(file_bytes, file_name, nrows=10)
    meta: dict = {"company": "", "account_label": "", "account_id_meta": "", "currency": "MXN"}

    for _, row in raw.iterrows():
        for cell in row.values:
            text = str(cell).strip()
            m = _COMPANY_RE.match(text)
            if m:
                meta["company"] = m.group(1).strip()
            m = _ACCOUNT_RE.match(text)
            if m:
                meta["account_label"] = m.group(1).strip()
                cm = _CURRENCY_RE.search(text)
                if cm:
                    meta["currency"] = cm.group(1).upper()
            m = _ACCT_ID_RE.match(text)
            if m:
                meta["account_id_meta"] = m.group(1).strip()

    return meta


# ── PARSE KYRIBA ───────────────────────────────────────────────────────────────

def parse_kyriba(
    file_bytes: bytes,
    file_name: str,
    treasury_keywords: list[str],
    treasury_accounts: set[str],
    manual_account_map: dict[str, str],  # {"AA370": "Banorte", ...}
    processor_mapper,                    # callable(row) -> str | None
) -> ParseResult:
    """
    Parsea un archivo Kyriba y devuelve transacciones limpias de payins.

    Args:
        file_bytes: Contenido binario del archivo.
        file_name: Nombre del archivo (para logs).
        treasury_keywords: Palabras clave que identifican movimientos internos.
        treasury_accounts: Account codes que son 100% treasury (excluir).
        manual_account_map: Mapeo manual Account code -> Processor name.
        processor_mapper: Función que recibe una row y devuelve el processor.

    Returns:
        ParseResult con df limpio y listas de warnings/errors.
    """
    result = ParseResult(df=pd.DataFrame(columns=KYRIBA_COLS))

    try:
        meta = _extract_kyriba_metadata(file_bytes, file_name)
        df = read_table(file_bytes, file_name, dtype=str)
    except Exception as exc:
        result.errors.append(f"{file_name}: No se pudo leer el archivo — {exc}")
        return result

    # ── Detectar columnas ─────────────────────────────────────────────────────
    account_code_col = find_column(df, ["Account code", "Account Code", "Account"])
    account_id_col   = find_column(df, ["Account ID", "Account Id", "Bank account"])
    date_col         = find_column(df, ["Transaction date", "Value date", "Booking date", "Date", "Fecha"])
    desc_col         = find_column(df, ["Description", "Description + Complementary info", "Concepto", "Descripcion"])
    compl_col        = find_column(df, ["Complementary info", "Complementary Info", "Additional info", "Reference"])
    credit_col       = find_column(df, ["Credit (MXN)", "Credit (USD)", "Credit", "Credito", "Deposit"])
    debit_col        = find_column(df, ["Debit (MXN)", "Debit (USD)", "Debit", "Debito"])
    amount_col       = find_column(df, ["Amount", "Amount (MXN)", "Importe"])

    missing = []
    if not date_col:   missing.append("Transaction date")
    if not desc_col:   missing.append("Description")
    if not credit_col and not amount_col: missing.append("Credit / Amount")

    if missing:
        result.errors.append(f"{file_name}: Columnas faltantes — {missing}. Disponibles: {list(df.columns)}")
        return result

    # ── Construir DataFrame canónico ──────────────────────────────────────────
    work = pd.DataFrame()
    work["Account code"]      = df[account_code_col].astype(str).str.strip() if account_code_col else ""
    work["Account ID"]        = df[account_id_col].astype(str).str.strip() if account_id_col else meta.get("account_id_meta", "")
    work["Company"]           = meta.get("company", "")
    work["Transaction date"]  = df[date_col]
    work["Description"]       = df[desc_col].astype(str)
    work["Complementary info"]= df[compl_col].astype(str) if compl_col else ""

    if credit_col:
        work["Credit"] = clean_amount(df[credit_col])
    else:
        raw_amt = clean_amount(df[amount_col])
        work["Credit"] = raw_amt.clip(lower=0)

    work["Debit"] = clean_amount(df[debit_col]) if debit_col else 0

    # ── Filtros básicos ───────────────────────────────────────────────────────
    # Eliminar filas de balance y encabezados replicados
    skip_descs = {"Opening balance", "Closing balance", "Description", "nan"}
    work = work[~work["Description"].isin(skip_descs)].copy()

    # Parsear fechas y eliminar inválidas
    work["Transaction date"] = pd.to_datetime(work["Transaction date"], errors="coerce")
    work = work[work["Transaction date"].notna()].copy()

    # Solo créditos positivos
    work = work[work["Credit"] > 0].copy()

    if work.empty:
        result.warnings.append(f"{file_name}: Sin movimientos de crédito positivos.")
        return result

    # ── Excluir cuentas treasury ──────────────────────────────────────────────
    n_before = len(work)
    work = work[~work["Account code"].isin(treasury_accounts)].copy()
    excluded_treasury_accts = n_before - len(work)
    if excluded_treasury_accts:
        result.warnings.append(
            f"{file_name}: {excluded_treasury_accts} filas excluidas por cuenta treasury "
            f"({treasury_accounts & set(work['Account code'] if not work.empty else [])})."
        )

    # ── Excluir movimientos internos por keyword ──────────────────────────────
    work["Text to match"] = (
        work["Description"].astype(str) + " " + work["Complementary info"].astype(str)
    ).str.upper()

    treasury_pattern = "|".join(re.escape(k) for k in treasury_keywords)
    is_treasury = work["Text to match"].str.contains(treasury_pattern, na=False)
    n_treasury = is_treasury.sum()
    if n_treasury:
        result.warnings.append(
            f"{file_name}: {n_treasury} movimientos treasury/internos excluidos."
        )
    work = work[~is_treasury].copy()

    if work.empty:
        result.warnings.append(f"{file_name}: Sin movimientos tras filtro treasury.")
        return result

    # ── Asignar Processor ─────────────────────────────────────────────────────
    def _assign(row):
        code = str(row["Account code"]).strip().upper()
        # 1. Mapping manual override
        if code in manual_account_map:
            return manual_account_map[code]
        # 2. Reglas keyword + account
        p = processor_mapper(row)
        if p:
            return p
        return None

    work["Processor"] = work.apply(_assign, axis=1)

    unmapped = work["Processor"].isna().sum()
    if unmapped:
        result.warnings.append(
            f"{file_name}: {unmapped} filas sin processor asignado (quedarán en 'No mapeado')."
        )

    # ── Columnas finales ──────────────────────────────────────────────────────
    work["Day"]         = work["Transaction date"].dt.strftime("%Y-%m-%d")
    work["Source file"] = file_name

    result.df = work[KYRIBA_COLS].copy()
    return result


# ── PARSE PAYINS / GROSS PROFIT ────────────────────────────────────────────────

def parse_payins(
    file_bytes: bytes,
    file_name: str,
    valid_agents: list[str],
    processor_normalizer,        # callable(str) -> str
    col_date: str = "",
    col_amount: str = "",
    col_processor: str = "",
) -> ParseResult:
    """
    Parsea un archivo de estimaciones Payins / Gross Profit.

    Args:
        file_bytes: Contenido binario del archivo.
        file_name: Nombre del archivo.
        valid_agents: Lista de collection agents válidos para la entidad.
        processor_normalizer: Función que normaliza nombres de processor.
        col_date/col_amount/col_processor: Override manual de columnas.

    Returns:
        ParseResult con df limpio.
    """
    result = ParseResult(df=pd.DataFrame(columns=PAYINS_COLS))

    try:
        df = read_table(file_bytes, file_name)
    except Exception as exc:
        result.errors.append(f"{file_name}: No se pudo leer — {exc}")
        return result

    # ── Detectar columnas ─────────────────────────────────────────────────────
    date_col = (
        (col_date if col_date in df.columns else find_column(df, [col_date]))
        if col_date else None
    ) or find_column(df, [
        "Payment Date", "Approved Date", "Payins Creation Date",
        "Creation Date", "Created Date", "Accounting Date", "Date", "Day", "Fecha",
    ])

    amount_col = (
        (col_amount if col_amount in df.columns else find_column(df, [col_amount]))
        if col_amount else None
    ) or find_column(df, [
        "Total Local Amount", "Approved Amount Local", "Local Amount",
        "LC Amount", "PI | Amount Approved | LC", "Amount Approved LC",
        "Amount Local", "Amount", "Approved Amount", "Monto",
    ])

    processor_col = (
        (col_processor if col_processor in df.columns else find_column(df, [col_processor]))
        if col_processor else None
    ) or find_column(df, [
        "Name", "Procesador", "Processor", "Processor Name",
        "Payins Processor", "Payment Processor", "Acquirer", "Gateway",
    ])

    missing = []
    if not date_col:      missing.append("Fecha")
    if not amount_col:    missing.append("Monto")
    if not processor_col: missing.append("Procesador")

    if missing:
        result.errors.append(
            f"{file_name}: Columnas faltantes — {missing}. Disponibles: {list(df.columns)}"
        )
        return result

    # ── Construir DataFrame canónico ──────────────────────────────────────────
    work = pd.DataFrame()
    work["Date"]               = df[date_col]
    work["Amount"]             = clean_amount(df[amount_col])
    work["Original Processor"] = df[processor_col].astype(str)
    work["Processor"]          = work["Original Processor"].apply(processor_normalizer)

    # Collection agent — filtro de entidad
    agent_col = find_column(df, ["Collection Agent", "Legal Entity", "Entity", "Company"])
    if agent_col:
        work["Collection Agent"] = df[agent_col].astype(str).str.strip()
       mask = work["Collection Agent"].str.lower().apply(
    lambda v: any(a in v for a in valid_agents) if isinstance(v, str) else False
)
        excluded = (~mask).sum()
        if excluded:
            result.warnings.append(
                f"{file_name}: {excluded} filas de otras entidades excluidas."
            )
        work = work[mask].copy()
    else:
        work["Collection Agent"] = ""

    # País — solo México si hay columna
    country_col = find_column(df, ["Country", "Country Transaction", "Pais", "País"])
    if country_col:
        work["Country"] = df[country_col].astype(str)
        n_before = len(work)
        work = work[work["Country"].str.lower().str.contains("mex", na=False)].copy()
        result.warnings.append(
            f"{file_name}: Filtro país — {n_before - len(work)} filas excluidas (no México)."
        )
    else:
        work["Country"] = ""

    # Payment method code
    code_col = find_column(df, ["Code", "Payment Method Code", "Payment Method", "PM Code"])
    work["Payment Method Code"] = df[code_col].astype(str) if code_col else ""

    # ── Limpieza final ────────────────────────────────────────────────────────
    work["Date"] = pd.to_datetime(work["Date"], errors="coerce")
    work = work[work["Date"].notna()].copy()
    work = work[work["Processor"].notna()].copy()
    work = work[work["Amount"] > 0].copy()

    if work.empty:
        result.warnings.append(f"{file_name}: Sin registros válidos tras limpieza.")
        return result

    work["Day"]         = work["Date"].dt.strftime("%Y-%m-%d")
    work["Source file"] = file_name

    result.df = work[PAYINS_COLS].copy()
    return result
