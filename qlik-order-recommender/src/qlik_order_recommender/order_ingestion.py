"""Gercek is surecine ait 3 girdi dosyasini okur:
 - transactions: Qlik'ten alinan ham bayi x urun x ay islem gecmisi
   (Diva Bildirim/Stok/Ileri Teslim/Fatura/Montaj adetleri)
 - dealer_targets: bayi bazli donemsel (or. ceyreklik) hedef/fiili adet
 - product_quota: urun kodu bazli sirket kotasi (bu donem girilecek adet)

Sutun adlari gercek Qlik/Excel disa aktarimlarindaki haliyle varsayilan
olarak taninir; farkli bir rapor formatinda ise config/order_column_mapping
YAML dosyasi ile eslenebilir."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Union

import pandas as pd
import yaml

PathLike = Union[str, Path]

REQUIRED_FIELDS = {
    "transactions": [
        "dealer_name", "lph2", "lph3", "product_code", "month",
        "bildirim", "stok", "ileri_teslim", "fatura", "montaj",
    ],
    "dealer_targets": ["dealer_name", "lph2", "hedef", "fiili"],
    "product_quota": ["product_code", "quota"],
}

DEFAULT_COLUMNS = {
    "transactions": {
        "dealer_name": "Nihai Ünvan",
        "lph2": "LPH2",
        "lph3": "LPH3",
        "product_code": "UrunKodu",
        "month": "YılAy",
        "bildirim": "Diva Bildirim",
        "stok": "Diva Stok",
        "ileri_teslim": "Diva - İleri Teslim",
        "fatura": "Fatura",
        "montaj": "Montaj",
    },
    "dealer_targets": {
        "dealer_name": "Ünvan",
        "lph2": "LPH-2",
        "hedef": "Hedef Q4",
        "fiili": "Fiili Q4",
    },
    "product_quota": {
        "product_code": "ÜRÜNLER",
        "quota": "ADET",
    },
}

NUMERIC_FIELDS = {
    "transactions": ["bildirim", "stok", "ileri_teslim", "fatura", "montaj"],
    "dealer_targets": ["hedef", "fiili"],
    "product_quota": ["quota"],
}

# Bu alanlarda negatif deger gecerlidir (iade/duzeltme kayitlari).
ALLOW_NEGATIVE = {"bildirim", "stok", "ileri_teslim", "fatura", "montaj"}

MONTH_PATTERN = r"^\d{4}-\d{2}$"


def load_mapping(mapping_path: PathLike | None) -> dict:
    if mapping_path is None:
        return {}
    with open(mapping_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _read_raw(path: PathLike) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Girdi dosyasi bulunamadi: {path}")
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path)
    return pd.read_csv(path)


def load_table(path: PathLike, table_name: str, mapping: dict) -> pd.DataFrame:
    if table_name not in REQUIRED_FIELDS:
        raise ValueError(f"Bilinmeyen tablo adi: {table_name}")

    raw = _read_raw(path)
    table_mapping = mapping.get(table_name, {}) if mapping else {}
    required = REQUIRED_FIELDS[table_name]
    defaults = DEFAULT_COLUMNS[table_name]

    rename = {}
    missing = []
    for field_name in required:
        source_col = table_mapping.get(field_name, defaults.get(field_name, field_name))
        if source_col in raw.columns:
            rename[source_col] = field_name
        elif field_name in raw.columns:
            rename[field_name] = field_name
        else:
            missing.append(field_name)

    if missing:
        raise ValueError(
            f"'{table_name}' tablosunda beklenen alanlar bulunamadi: {missing}. "
            f"Sutun eslemesi icin bir mapping YAML dosyasi verin "
            f"(mevcut sutunlar: {list(raw.columns)})."
        )

    df = raw.rename(columns=rename)
    df = df[required].copy()

    for col in NUMERIC_FIELDS.get(table_name, []):
        df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in ("dealer_name", "product_code", "lph2", "lph3"):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    if table_name == "transactions":
        before = len(df)
        df = df[df["month"].astype(str).str.match(MONTH_PATTERN)].copy()
        dropped = before - len(df)
        if dropped:
            print(f"[order_ingestion] '{table_name}': {dropped} satir gecersiz/eksik ay (YılAy) "
                  f"degeri nedeniyle atlandi.")
        for col in NUMERIC_FIELDS["transactions"]:
            df[col] = df[col].fillna(0.0)

    _validate(df, table_name)
    return df.reset_index(drop=True)


def _validate(df: pd.DataFrame, table_name: str) -> None:
    for col in ("dealer_name", "product_code"):
        if col in df.columns and ((df[col].isna()) | (df[col] == "") | (df[col] == "nan")).any():
            raise ValueError(f"'{table_name}.{col}' alaninda bos deger bulundu.")

    for col in NUMERIC_FIELDS.get(table_name, []):
        if col in ALLOW_NEGATIVE:
            continue
        if df[col].isna().any():
            raise ValueError(f"'{table_name}.{col}' alaninda sayisal olmayan/eksik deger bulundu.")
        if (df[col] < 0).any():
            raise ValueError(f"'{table_name}.{col}' alaninda negatif deger bulundu.")


@dataclass
class OrderDataBundle:
    transactions: pd.DataFrame
    dealer_targets: pd.DataFrame
    product_quota: pd.DataFrame


def resolve_dealer_names(
    dealer_targets: pd.DataFrame, transactions: pd.DataFrame
) -> tuple[dict[str, str], list[str]]:
    """dealer_targets'taki bayi adlarini transactions'taki (Qlik) tam
    unvanlarla eslendirir. Once birebir, sonra buyuk/kucuk harf farkini
    yok sayarak, sonra tek anlamli bir on-ek/ic-ice gecme adayi varsa onu
    kullanir (ör. 'ASL LIVING DAYANIKLI TUKETIM MALLARI' -> '... TICARET
    LIMITED SIRKETI'). Eslendirilemeyen adlar ayrica dondurulur."""
    known = list(transactions["dealer_name"].unique())
    known_upper = {k.upper(): k for k in known}

    mapping: dict[str, str] = {}
    unresolved: list[str] = []
    for name in dealer_targets["dealer_name"].unique():
        if name in known:
            mapping[name] = name
            continue
        upper = name.upper()
        if upper in known_upper:
            mapping[name] = known_upper[upper]
            continue
        candidates = [k for k in known if k.upper().startswith(upper) or upper.startswith(k.upper())]
        if len(candidates) == 1:
            mapping[name] = candidates[0]
        else:
            unresolved.append(name)
    return mapping, unresolved


def load_order_data(
    transactions_path: PathLike,
    dealer_targets_path: PathLike,
    product_quota_path: PathLike,
    mapping_path: PathLike | None = None,
) -> OrderDataBundle:
    mapping = load_mapping(mapping_path)
    return OrderDataBundle(
        transactions=load_table(transactions_path, "transactions", mapping),
        dealer_targets=load_table(dealer_targets_path, "dealer_targets", mapping),
        product_quota=load_table(product_quota_path, "product_quota", mapping),
    )
