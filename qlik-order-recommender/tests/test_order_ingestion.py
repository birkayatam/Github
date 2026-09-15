import pandas as pd
import pytest

from qlik_order_recommender.order_ingestion import load_table, resolve_dealer_names


def _write_csv(tmp_path, name, df):
    path = tmp_path / name
    df.to_csv(path, index=False)
    return path


def test_load_transactions_with_default_columns(tmp_path):
    df = pd.DataFrame(
        [
            {
                "Nihai Ünvan": "ACME LTD",
                "LPH2": "Bosch-LDA",
                "LPH3": "Bosch-LDA-X",
                "UrunKodu": "P1",
                "YılAy": "2026-06",
                "Diva Bildirim": 3,
                "Diva Stok": 0,
                "Diva - İleri Teslim": 0,
                "Fatura": 2,
                "Montaj": 1,
            },
            {
                "Nihai Ünvan": "ACME LTD",
                "LPH2": "Bosch-LDA",
                "LPH3": "Bosch-LDA-X",
                "UrunKodu": "P1",
                "YılAy": "-",  # gecersiz ay, atlanmali
                "Diva Bildirim": 1,
                "Diva Stok": 0,
                "Diva - İleri Teslim": 0,
                "Fatura": 1,
                "Montaj": 0,
            },
        ]
    )
    path = _write_csv(tmp_path, "transactions.csv", df)

    result = load_table(path, "transactions", mapping={})

    assert len(result) == 1  # gecersiz ay satiri atlandi
    assert result.iloc[0]["dealer_name"] == "ACME LTD"
    assert result.iloc[0]["fatura"] == 2


def test_load_dealer_targets_with_default_columns(tmp_path):
    df = pd.DataFrame(
        [{"Ünvan": "ACME LTD", "LPH-2": "LDA", "Hedef Q4": 100, "Fiili Q4": 40}]
    )
    path = _write_csv(tmp_path, "targets.csv", df)

    result = load_table(path, "dealer_targets", mapping={})

    assert result.iloc[0]["hedef"] == 100
    assert result.iloc[0]["fiili"] == 40


def test_load_product_quota_missing_column_raises(tmp_path):
    df = pd.DataFrame([{"WRONG_COL": "P1", "ADET": 5}])
    path = _write_csv(tmp_path, "quota.csv", df)

    with pytest.raises(ValueError):
        load_table(path, "product_quota", mapping={})


def test_resolve_dealer_names_exact_and_fuzzy():
    dealer_targets = pd.DataFrame(
        {"dealer_name": ["ACME LTD", "ASL LIVING DAYANIKLI TUKETIM MALLARI", "UNKNOWN CO"]}
    )
    transactions = pd.DataFrame(
        {"dealer_name": ["ACME LTD", "ASL LIVING DAYANIKLI TUKETIM MALLARI TICARET LIMITED SIRKETI"]}
    )

    mapping, unresolved = resolve_dealer_names(dealer_targets, transactions)

    assert mapping["ACME LTD"] == "ACME LTD"
    assert mapping["ASL LIVING DAYANIKLI TUKETIM MALLARI"] == (
        "ASL LIVING DAYANIKLI TUKETIM MALLARI TICARET LIMITED SIRKETI"
    )
    assert unresolved == ["UNKNOWN CO"]
