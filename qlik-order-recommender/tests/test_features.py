from qlik_order_recommender.features import compute_dealer_features, monthly_totals


def test_trailing_averages_and_growth(synthetic_bundle):
    totals = monthly_totals(synthetic_bundle.sales_history)
    feats = compute_dealer_features("D1", "2025-07", synthetic_bundle, totals)

    assert feats.months_of_history == 6
    assert feats.trailing_avg_3m == (1300 + 1400 + 1500) / 3
    assert feats.trailing_avg_6m == sum([1000, 1100, 1200, 1300, 1400, 1500]) / 6
    assert feats.mom_growth_rate > 0  # sales are steadily increasing
    assert feats.target_accuracy_ratio is not None
    assert feats.target_accuracy_ratio > 1.0  # achieved consistently above target in fixture


def test_no_history_before_first_month(synthetic_bundle):
    totals = monthly_totals(synthetic_bundle.sales_history)
    feats = compute_dealer_features("D1", "2025-01", synthetic_bundle, totals)

    assert feats.months_of_history == 0
    assert feats.trailing_avg_3m == 0.0
    assert "Guvenilir tahmin icin yeterli gecmis veri yok" in feats.risk_flags
