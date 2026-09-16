import pandas as pd
import pytest

from qlik_order_recommender.ingestion import DataBundle

SAMPLE_DIR = __file__.rsplit("/tests/", 1)[0] + "/data/sample"


@pytest.fixture
def synthetic_bundle() -> DataBundle:
    dealers = pd.DataFrame(
        [{"dealer_id": "D1", "dealer_name": "Test Bayi", "region": "Marmara", "segment": "Orta"}]
    )

    months = ["2025-01-01", "2025-02-01", "2025-03-01", "2025-04-01", "2025-05-01", "2025-06-01"]
    sales_amounts = [1000, 1100, 1200, 1300, 1400, 1500]
    sales_history = pd.DataFrame(
        [
            {
                "dealer_id": "D1",
                "date": pd.Timestamp(m),
                "product_category": "Elektronik",
                "sales_amount": amt,
                "quantity": amt / 100,
            }
            for m, amt in zip(months, sales_amounts)
        ]
    )

    targets_history = pd.DataFrame(
        [
            {"dealer_id": "D1", "month": m[:7], "target_amount": amt * 0.95, "achieved_amount": amt}
            for m, amt in zip(months, sales_amounts)
        ]
    )

    invoices = pd.DataFrame(
        [
            {
                "dealer_id": "D1",
                "invoice_date": pd.Timestamp(m),
                "invoice_amount": amt,
                "due_date": pd.Timestamp(m) + pd.Timedelta(days=30),
                "paid_date": pd.Timestamp(m) + pd.Timedelta(days=25),
                "payment_status": "Odendi",
            }
            for m, amt in zip(months, sales_amounts)
        ]
    )

    return DataBundle(
        dealers=dealers,
        sales_history=sales_history,
        invoices=invoices,
        targets_history=targets_history,
    )


@pytest.fixture(scope="session")
def sample_data_dir() -> str:
    return SAMPLE_DIR
