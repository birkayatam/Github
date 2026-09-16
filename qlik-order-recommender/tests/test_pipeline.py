import pandas as pd

from qlik_order_recommender.pipeline import run


def test_pipeline_end_to_end(tmp_path, sample_data_dir, monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)

    project_root = sample_data_dir.rsplit("/data/sample", 1)[0]
    mapping_path = f"{project_root}/config/column_mapping.example.yaml"
    out_dir = tmp_path / "reports"

    excel_path = run(
        dealers_path=f"{sample_data_dir}/dealers.csv",
        sales_path=f"{sample_data_dir}/sales_history.csv",
        invoices_path=f"{sample_data_dir}/invoices.csv",
        targets_path=f"{sample_data_dir}/targets_history.csv",
        mapping_path=mapping_path,
        target_month="2026-10",
        out_dir=str(out_dir),
    )

    assert excel_path.exists()
    df = pd.read_excel(excel_path)
    assert len(df) == 5
    assert {"Bayi Kodu", "Onerilen Hedef", "Min", "Max", "Guven"}.issubset(df.columns)
    assert (df["Min"] <= df["Onerilen Hedef"]).all()
    assert (df["Onerilen Hedef"] <= df["Max"]).all()

    md_files = list(out_dir.glob("*_2026-10.md"))
    assert len(md_files) == 5
    for path in md_files:
        content = path.read_text(encoding="utf-8")
        assert "Hedef Onerisi" in content
