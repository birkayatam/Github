from qlik_order_recommender.features import compute_dealer_features, monthly_totals
from qlik_order_recommender.target_model import recommend_target, recommend_targets_all


def test_recommend_target_range_is_sane(synthetic_bundle):
    totals = monthly_totals(synthetic_bundle.sales_history)
    feats = compute_dealer_features("D1", "2025-07", synthetic_bundle, totals)

    rec = recommend_target(feats)

    assert rec.target_min <= rec.recommended_target <= rec.target_max
    assert 0.2 <= rec.confidence <= 0.95
    assert rec.recommended_target > 0
    assert rec.method == "heuristic"  # no trained model passed


def test_recommend_targets_all_covers_every_dealer(synthetic_bundle):
    totals = monthly_totals(synthetic_bundle.sales_history)
    feats = {"D1": compute_dealer_features("D1", "2025-07", synthetic_bundle, totals)}

    recs = recommend_targets_all(synthetic_bundle, feats)

    assert set(recs.keys()) == {"D1"}
    assert recs["D1"].dealer_id == "D1"
