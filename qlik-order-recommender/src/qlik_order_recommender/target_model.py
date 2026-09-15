"""Hedef oneri modeli: heuristik taban tahmin + (yeterli veri varsa) ML ile
ince ayar. Cikti tek bir sayi degil, guven araligi ve risk bayraklari icerir
-- amac ne dusuk hedefle firsat kacirmak ne de gerceklesmesi zor, motivasyonu
kiran bir hedef koymak."""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from .features import DealerFeatures, compute_dealer_features, monthly_totals
from .ingestion import DataBundle

MIN_TRAINING_SAMPLES = 24
FEATURE_NAMES = [
    "trailing_avg_3m",
    "trailing_avg_6m",
    "mom_growth_rate",
    "yoy_growth_rate",
    "seasonality_factor",
    "target_accuracy_ratio",
    "late_payment_rate",
]


@dataclass
class TargetRecommendation:
    dealer_id: str
    target_month: str
    recommended_target: float
    target_min: float
    target_max: float
    confidence: float
    method: str
    risk_flags: list


def _features_to_vector(f: DealerFeatures) -> list[float]:
    return [
        f.trailing_avg_3m,
        f.trailing_avg_6m,
        f.mom_growth_rate,
        f.yoy_growth_rate if f.yoy_growth_rate is not None else 0.0,
        f.seasonality_factor,
        f.target_accuracy_ratio if f.target_accuracy_ratio is not None else 1.0,
        f.late_payment_rate if f.late_payment_rate is not None else 0.0,
    ]


def build_training_set(bundle: DataBundle) -> tuple[np.ndarray, np.ndarray]:
    """Gecmis her ay icin, o ana kadarki veriyle hesaplanan ozellikleri (X)
    ve o ayda gerceklesen toplam satisi (y) esleyerek egitim seti kurar."""
    totals = monthly_totals(bundle.sales_history)
    X, y = [], []
    for dealer_id in bundle.dealers["dealer_id"]:
        dealer_months = sorted(totals.loc[totals["dealer_id"] == dealer_id, "month"].unique())
        for month in dealer_months[3:]:
            feats = compute_dealer_features(dealer_id, str(month), bundle, totals)
            if feats.months_of_history < 3:
                continue
            label_row = totals[(totals["dealer_id"] == dealer_id) & (totals["month"] == month)]
            if label_row.empty:
                continue
            X.append(_features_to_vector(feats))
            y.append(float(label_row["sales_amount"].iloc[0]))
    return np.array(X), np.array(y)


def train_model(bundle: DataBundle):
    """Yeterli gecmis veri varsa bir GradientBoostingRegressor egitir, yoksa
    None doner (bu durumda tamamen heuristik moda dusulur)."""
    X, y = build_training_set(bundle)
    if len(X) < MIN_TRAINING_SAMPLES:
        return None
    from sklearn.ensemble import GradientBoostingRegressor

    model = GradientBoostingRegressor(random_state=42, n_estimators=150, max_depth=2, learning_rate=0.08)
    model.fit(X, y)
    return model


def _heuristic_estimate(f: DealerFeatures) -> float:
    trend_component = f.trailing_avg_3m * (1 + f.mom_growth_rate)
    seasonal_component = f.trailing_avg_6m * f.seasonality_factor if f.trailing_avg_6m else trend_component
    base = 0.5 * trend_component + 0.5 * seasonal_component

    if f.target_accuracy_ratio is not None:
        # Gecmiste hedefler surekli tutturulamamissa (ratio<1) tahmini asagi,
        # surekli asilmissa (ratio>1) hafifce yukari cek -- ama tam oranla
        # degil, yumusatilmis (damped) sekilde.
        calibration = 0.5 + 0.5 * min(max(f.target_accuracy_ratio, 0.5), 1.5)
        base *= calibration

    return max(base, 0.0)


def _volatility(f: DealerFeatures) -> float:
    vals = [v for _, v in f.recent_monthly_totals]
    if len(vals) < 2:
        return 0.25
    arr = np.array(vals, dtype=float)
    mean = arr.mean()
    if mean == 0:
        return 0.25
    return float(arr.std() / mean)


def recommend_target(f: DealerFeatures, model=None) -> TargetRecommendation:
    heuristic = _heuristic_estimate(f)

    method = "heuristic"
    base = heuristic
    if model is not None and f.months_of_history >= 3:
        try:
            ml_pred = float(model.predict([_features_to_vector(f)])[0])
            base = 0.5 * heuristic + 0.5 * max(ml_pred, 0.0)
            method = "ml+heuristic"
        except Exception:
            base = heuristic
            method = "heuristic"

    vol = _volatility(f)
    confidence = max(0.2, min(0.95, 1 - vol - 0.05 * len(f.risk_flags)))
    band = 0.05 + (1 - confidence) * 0.35

    return TargetRecommendation(
        dealer_id=f.dealer_id,
        target_month=f.target_month,
        recommended_target=round(base, 2),
        target_min=round(base * (1 - band), 2),
        target_max=round(base * (1 + band), 2),
        confidence=round(confidence, 2),
        method=method,
        risk_flags=list(f.risk_flags),
    )


def recommend_targets_all(
    bundle: DataBundle, features_by_dealer: dict[str, DealerFeatures]
) -> dict[str, TargetRecommendation]:
    model = train_model(bundle)
    return {dealer_id: recommend_target(f, model) for dealer_id, f in features_by_dealer.items()}
