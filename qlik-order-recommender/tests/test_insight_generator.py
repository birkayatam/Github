from qlik_order_recommender.features import compute_dealer_features, monthly_totals
from qlik_order_recommender.insight_generator import generate_insight
from qlik_order_recommender.target_model import recommend_target


def test_template_fallback_used_without_api_key(monkeypatch, synthetic_bundle):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)

    totals = monthly_totals(synthetic_bundle.sales_history)
    feats = compute_dealer_features("D1", "2025-07", synthetic_bundle, totals)
    rec = recommend_target(feats)

    insight = generate_insight("Test Bayi", feats, rec)

    assert insight.source == "template"
    assert "Test Bayi" in insight.narrative
    assert "Hedef Onerisi" in insight.narrative
    assert "Ikna icin kullanilacak veri noktalari" in insight.narrative
    assert "Dikkat edilmesi gerekenler" in insight.narrative
