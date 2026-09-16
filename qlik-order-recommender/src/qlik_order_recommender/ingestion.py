"""CSV/Excel okuma ve Qlik disa aktarim sutunlarini ic semaya normalize etme."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Union

import pandas as pd
import yaml

PathLike = Union[str, Path]

REQUIRED_FIELDS = {
    "dealers": ["dealer_id", "dealer_name", "region", "segment"],
    "sales_history": ["dealer_id", "date", "product_category", "sales_amount", "quantity"],
    "invoices": ["dealer_id", "invoice_date", "invoice_amount", "due_date", "paid_date", "payment_status"],
    "targets_history": ["dealer_id", "month", "target_amount", "achieved_amount"],
}

OPTIONAL_FIELDS = {
    "invoices": {"paid_date"},
}

NUMERIC_FIELDS = {
    "sales_history": ["sales_amount", "quantity"],
    "invoices": ["invoice_amount"],
    "targets_history": ["target_amount", "achieved_amount"],
}

DATE_FIELDS = {
    "sales_history": ["date"],
    "invoices": ["invoice_date", "due_date", "paid_date"],
}


def load_mapping(mapping_path: PathLike) -> dict:
    with open(mapping_path, "r", encoding="utf-8") as f:
        mapping = yaml.safe_load(f)
    return mapping or {}


def _read_raw(path: PathLike) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Girdi dosyasi bulunamadi: {path}")
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path)
    return pd.read_csv(path)


def load_table(path: PathLike, table_name: str, mapping: dict) -> pd.DataFrame:
    """Bir Qlik disa aktarim dosyasini okuyup ic semaya gore normalize eder."""
    if table_name not in REQUIRED_FIELDS:
        raise ValueError(f"Bilinmeyen tablo adi: {table_name}")

    raw = _read_raw(path)
    table_mapping = mapping.get(table_name, {}) if mapping else {}
    required = REQUIRED_FIELDS[table_name]
    optional = OPTIONAL_FIELDS.get(table_name, set())

    rename = {}
    missing = []
    for field in required:
        source_col = table_mapping.get(field, field)
        if source_col in raw.columns:
            rename[source_col] = field
        elif field in raw.columns:
            rename[field] = field
        elif field in optional:
            continue
        else:
            missing.append(field)

    if missing:
        raise ValueError(
            f"'{table_name}' tablosunda beklenen alanlar bulunamadi: {missing}. "
            f"Sutun eslemesi icin config/column_mapping.yaml dosyasini kontrol edin "
            f"(mevcut sutunlar: {list(raw.columns)})."
        )

    df = raw.rename(columns=rename)
    keep_cols = [c for c in required if c in df.columns]
    df = df[keep_cols].copy()

    for col in NUMERIC_FIELDS.get(table_name, []):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in DATE_FIELDS.get(table_name, []):
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if table_name == "targets_history" and "month" in df.columns:
        df["month"] = df["month"].astype(str).str.slice(0, 7)

    df["dealer_id"] = df["dealer_id"].astype(str).str.strip()

    _validate(df, table_name)
    return df.reset_index(drop=True)


def _validate(df: pd.DataFrame, table_name: str) -> None:
    if df["dealer_id"].isna().any() or (df["dealer_id"] == "").any():
        raise ValueError(f"'{table_name}' tablosunda bos dealer_id degeri bulunan satirlar var.")

    for col in NUMERIC_FIELDS.get(table_name, []):
        if col in df.columns and (df[col] < 0).any():
            raise ValueError(f"'{table_name}.{col}' alaninda negatif deger bulundu.")

    if table_name in ("sales_history",) and df["date"].isna().any():
        raise ValueError(f"'{table_name}.date' alaninda gecersiz/eksik tarih bulundu.")

    if table_name == "invoices" and df["invoice_date"].isna().any():
        raise ValueError("'invoices.invoice_date' alaninda gecersiz/eksik tarih bulundu.")


@dataclass
class DataBundle:
    dealers: pd.DataFrame
    sales_history: pd.DataFrame
    invoices: pd.DataFrame
    targets_history: pd.DataFrame


def load_all(
    dealers_path: PathLike,
    sales_path: PathLike,
    invoices_path: PathLike,
    targets_path: PathLike,
    mapping_path: PathLike,
) -> DataBundle:
    mapping = load_mapping(mapping_path)
    return DataBundle(
        dealers=load_table(dealers_path, "dealers", mapping),
        sales_history=load_table(sales_path, "sales_history", mapping),
        invoices=load_table(invoices_path, "invoices", mapping),
        targets_history=load_table(targets_path, "targets_history", mapping),
    )
