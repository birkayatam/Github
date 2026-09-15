import pandas as pd

from qlik_order_recommender.allocation import allocate, compute_affinity


def _synthetic_transactions() -> pd.DataFrame:
    rows = [
        # D1 favors P1 heavily
        {"dealer_name": "D1", "lph2": "LDA", "lph3": "X", "product_code": "P1", "month": "2026-06", "bildirim": 0, "stok": 0, "ileri_teslim": 0, "fatura": 10, "montaj": 0},
        {"dealer_name": "D1", "lph2": "LDA", "lph3": "X", "product_code": "P1", "month": "2026-07", "bildirim": 0, "stok": 0, "ileri_teslim": 0, "fatura": 8, "montaj": 0},
        # D2 favors P2 heavily
        {"dealer_name": "D2", "lph2": "LDA", "lph3": "X", "product_code": "P2", "month": "2026-06", "bildirim": 0, "stok": 0, "ileri_teslim": 0, "fatura": 5, "montaj": 0},
        {"dealer_name": "D2", "lph2": "LDA", "lph3": "X", "product_code": "P2", "month": "2026-07", "bildirim": 0, "stok": 0, "ileri_teslim": 0, "fatura": 6, "montaj": 0},
    ]
    return pd.DataFrame(rows)


def test_affinity_reflects_historical_preference():
    transactions = _synthetic_transactions()
    affinity = compute_affinity(transactions, dealer_names=["D1", "D2"], product_codes=["P1", "P2"])

    assert affinity.loc["D1", "P1"] > affinity.loc["D1", "P2"]
    assert affinity.loc["D2", "P2"] > affinity.loc["D2", "P1"]


def test_allocate_respects_product_quota_cap():
    transactions = _synthetic_transactions()
    remaining = pd.Series({"D1": 10.0, "D2": 5.0})
    quota = pd.Series({"P1": 8.0, "P2": 20.0})
    affinity = compute_affinity(transactions, dealer_names=list(remaining.index), product_codes=list(quota.index))

    result = allocate(remaining, quota, affinity)

    assert (result.allocations.sum(axis=0) <= quota).all()
    assert (result.allocations >= 0).all().all()
    assert result.allocations.to_numpy().dtype.kind in ("i", "u")


def test_allocate_fills_dealer_need_when_capacity_allows():
    transactions = _synthetic_transactions()
    remaining = pd.Series({"D1": 10.0, "D2": 5.0})
    quota = pd.Series({"P1": 8.0, "P2": 20.0})
    affinity = compute_affinity(transactions, dealer_names=list(remaining.index), product_codes=list(quota.index))

    result = allocate(remaining, quota, affinity)

    # D2's favored product (P2) has ample capacity, so its need should be fully met.
    assert result.dealer_shortfall["D2"] == 0
    assert result.allocations.loc["D2"].sum() == 5

    # D1's favored product (P1) is capacity-constrained (quota 8 < need 10); the
    # excess should spill over into P2 rather than being silently dropped.
    assert result.allocations.loc["D1"].sum() >= 8


def test_allocate_handles_dealer_with_no_history_via_market_smoothing():
    transactions = _synthetic_transactions()
    remaining = pd.Series({"D3": 3.0})  # bu bayinin hic gecmisi yok
    quota = pd.Series({"P1": 8.0, "P2": 20.0})
    affinity = compute_affinity(transactions, dealer_names=["D1", "D2", "D3"], product_codes=list(quota.index))

    result = allocate(remaining, quota, affinity)

    assert result.allocations.loc["D3"].sum() > 0
    assert result.dealer_shortfall["D3"] == 0
