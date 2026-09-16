import pandas as pd

from qlik_order_recommender.order_features import (
    compute_seasonality,
    compute_sell_through,
    demand_signal,
    detect_target_quarter,
    quarter_periods,
)


def _row(dealer, product, month, bildirim=0, ileri_teslim=0, montaj=0, fatura=0):
    return {
        "dealer_name": dealer,
        "lph2": "LDA",
        "lph3": "X",
        "product_code": product,
        "month": month,
        "bildirim": bildirim,
        "stok": 0,
        "ileri_teslim": ileri_teslim,
        "fatura": fatura,
        "montaj": montaj,
    }


def _seasonal_transactions() -> pd.DataFrame:
    rows = []
    # D1/P1: Eylul-Ekim-Kasim'da belirgin sekilde daha guclu, 2025 Q4 2024'ten daha yuksek (pozitif YoY)
    for year, q4_level in ((2024, 40), (2025, 60)):
        for month in range(1, 13):
            level = q4_level if month in (9, 10, 11) else 10
            rows.append(_row("D1", "P1", f"{year}-{month:02d}", bildirim=level))
    # D2/P2: yil boyunca sabit (mevsimsellik yok)
    for year in (2024, 2025):
        for month in range(1, 13):
            rows.append(_row("D2", "P2", f"{year}-{month:02d}", bildirim=20))
    # Hedef donemi (2026 Q4) tespit ettirmek icin en guncel ay
    rows.append(_row("D1", "P1", "2026-09", bildirim=50))
    return pd.DataFrame(rows)


def test_demand_signal_weights_sellout_over_sellin():
    df = pd.DataFrame(
        [_row("D1", "P1", "2026-01", bildirim=10, ileri_teslim=5, montaj=2, fatura=100)]
    )
    signal = demand_signal(df).iloc[0]
    # Bildirim+Ileri Teslim+1.5*Montaj = 10+5+3 = 18, Fatura sadece 0.3*100=30 katkida bulunur (baskin degil)
    assert signal == 10 + 5 + 1.5 * 2 + 0.3 * 100
    assert signal < 100  # ham fatura miktarindan kucuk olmali, sell-in tek basina baskin olmamali


def test_quarter_periods_handles_december_year_boundary():
    assert quarter_periods(2026, 1) == ["2025-12", "2026-01", "2026-02"]
    assert quarter_periods(2026, 4) == ["2026-09", "2026-10", "2026-11"]


def test_detect_target_quarter():
    df = pd.DataFrame({"month": ["2026-06", "2026-07", "2026-09"]})
    assert detect_target_quarter(df) == (2026, 4)

    df_dec = pd.DataFrame({"month": ["2025-11", "2025-12"]})
    assert detect_target_quarter(df_dec) == (2026, 1)


def test_seasonality_detects_strong_quarter_and_positive_yoy():
    transactions = _seasonal_transactions()
    result = compute_seasonality(transactions, "dealer_name", 2026, 4)

    assert result.seasonal_index["D1"] > 1.1  # D1 Eylul-Ekim-Kasim'da belirgin guclu
    assert result.yoy_growth["D1"] > 0  # 2025 Q4 (60) > 2024 Q4 (40)


def test_seasonality_neutral_for_flat_dealer():
    transactions = _seasonal_transactions()
    result = compute_seasonality(transactions, "dealer_name", 2026, 4)

    assert abs(result.seasonal_index["D2"] - 1.0) < 0.05  # D2 sabit, mevsimsellik yok


def test_seasonality_no_history_returns_neutral():
    transactions = pd.DataFrame([_row("D1", "P1", "2026-09", bildirim=10)])
    result = compute_seasonality(transactions, "dealer_name", 2026, 4)

    assert result.seasonal_index["D1"] == 1.0
    assert pd.isna(result.yoy_growth["D1"])


def test_sell_through_flags_oversupply_and_fast_movers():
    transactions = pd.DataFrame(
        [
            # P_OVERSUPPLY: cok fatura, hic sell-out yok -> dusuk oran
            _row("D1", "P_OVERSUPPLY", "2026-06", fatura=100, bildirim=0, montaj=0),
            # P_FASTMOVER: az fatura, cok sell-out -> yuksek oran
            _row("D1", "P_FASTMOVER", "2026-06", fatura=5, bildirim=20, montaj=5),
            # P_NO_FATURA: hic fatura yok -> oran hesaplanamaz (NaN)
            _row("D1", "P_NO_FATURA", "2026-06", fatura=0, bildirim=10, montaj=0),
        ]
    )
    result = compute_sell_through(transactions)

    assert result.loc[("D1", "P_OVERSUPPLY"), "sell_through_ratio"] < 0.3
    assert result.loc[("D1", "P_FASTMOVER"), "sell_through_ratio"] > 1.2
    assert pd.isna(result.loc[("D1", "P_NO_FATURA"), "sell_through_ratio"])
