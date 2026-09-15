import pandas as pd

from qlik_order_recommender.order_pipeline import run


def _write(tmp_path, name, df):
    path = tmp_path / name
    df.to_csv(path, index=False)
    return path


def test_order_pipeline_end_to_end(tmp_path, monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)

    transactions = pd.DataFrame(
        [
            {"Nihai Ünvan": "ACME LTD", "LPH2": "Bosch-LDA", "LPH3": "X", "UrunKodu": "P1",
             "YılAy": "2026-06", "Diva Bildirim": 0, "Diva Stok": 0, "Diva - İleri Teslim": 0,
             "Fatura": 10, "Montaj": 0},
            {"Nihai Ünvan": "BETA CO", "LPH2": "Bosch-LDA", "LPH3": "X", "UrunKodu": "P2",
             "YılAy": "2026-06", "Diva Bildirim": 0, "Diva Stok": 0, "Diva - İleri Teslim": 0,
             "Fatura": 8, "Montaj": 0},
        ]
    )
    dealer_targets = pd.DataFrame(
        [
            {"Ünvan": "ACME LTD", "LPH-2": "LDA", "Hedef Q4": 20, "Fiili Q4": 5},
            {"Ünvan": "BETA CO", "LPH-2": "LDA", "Hedef Q4": 10, "Fiili Q4": 8},
        ]
    )
    product_quota = pd.DataFrame(
        [
            {"ÜRÜNLER": "P1", "ADET": 15, "SİPARİŞ": None},
            {"ÜRÜNLER": "P2", "ADET": 15, "SİPARİŞ": None},
        ]
    )

    transactions_path = _write(tmp_path, "transactions.csv", transactions)
    dealer_targets_path = _write(tmp_path, "targets.csv", dealer_targets)
    product_quota_path = _write(tmp_path, "quota.csv", product_quota)
    out_dir = tmp_path / "out"

    excel_path = run(
        transactions_path=str(transactions_path),
        dealer_targets_path=str(dealer_targets_path),
        product_quota_path=str(product_quota_path),
        out_dir=str(out_dir),
    )

    assert excel_path.exists()
    summary = pd.read_excel(excel_path, sheet_name="Bayi Ozet")
    assert set(summary["Bayi"]) == {"ACME LTD", "BETA CO"}
    assert (summary["Onerilen Toplam"] <= summary["Kalan"] + 1e-6).all()

    quota_sheet = pd.read_excel(excel_path, sheet_name="Siparis Dosyasi")
    assert (quota_sheet["SIPARIS (tahsis edilen)"] <= quota_sheet["ADET"]).all()

    md_dir = out_dir / "bayi_raporlari"
    md_files = list(md_dir.glob("*.md"))
    assert len(md_files) == 2
    for path in md_files:
        assert "Siparis Onerisi" in path.read_text(encoding="utf-8")
