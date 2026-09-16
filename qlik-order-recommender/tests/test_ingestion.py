import pytest

from qlik_order_recommender.ingestion import load_all, load_mapping, load_table


def test_load_all_sample_data(sample_data_dir):
    mapping_path = sample_data_dir.rsplit("/data/sample", 1)[0] + "/config/column_mapping.example.yaml"
    bundle = load_all(
        dealers_path=f"{sample_data_dir}/dealers.csv",
        sales_path=f"{sample_data_dir}/sales_history.csv",
        invoices_path=f"{sample_data_dir}/invoices.csv",
        targets_path=f"{sample_data_dir}/targets_history.csv",
        mapping_path=mapping_path,
    )
    assert set(bundle.dealers.columns) == {"dealer_id", "dealer_name", "region", "segment"}
    assert len(bundle.dealers) == 5
    assert set(bundle.sales_history.columns) == {"dealer_id", "date", "product_category", "sales_amount", "quantity"}
    assert bundle.sales_history["sales_amount"].min() >= 0
    assert bundle.invoices["invoice_date"].notna().all()
    assert bundle.targets_history["month"].str.match(r"^\d{4}-\d{2}$").all()


def test_missing_column_raises(tmp_path, sample_data_dir):
    mapping = load_mapping(sample_data_dir.rsplit("/data/sample", 1)[0] + "/config/column_mapping.example.yaml")
    bad_mapping = dict(mapping)
    bad_mapping["dealers"] = {**mapping["dealers"], "dealer_id": "Olmayan Sutun"}

    import yaml

    bad_mapping_path = tmp_path / "bad_mapping.yaml"
    bad_mapping_path.write_text(yaml.safe_dump(bad_mapping), encoding="utf-8")

    with pytest.raises(ValueError):
        load_table(f"{sample_data_dir}/dealers.csv", "dealers", load_mapping(bad_mapping_path))
