"""Bayi bazli ozellik muhendisligi: trend, mevsimsellik, odeme sagligi,
gecmis hedef isabet orani, urun karmasi degisimi."""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from .ingestion import DataBundle

TRAILING_SHORT = 3
TRAILING_LONG = 6


@dataclass
class DealerFeatures:
    dealer_id: str
    target_month: str
    trailing_avg_3m: float
    trailing_avg_6m: float
    mom_growth_rate: float
    yoy_growth_rate: float | None
    seasonality_factor: float
    target_accuracy_ratio: float | None
    target_bias: float | None
    late_payment_rate: float | None
    pending_payment_rate: float | None
    avg_delay_days: float | None
    category_shares: dict = field(default_factory=dict)
    category_share_change: dict = field(default_factory=dict)
    recent_monthly_totals: list = field(default_factory=list)
    months_of_history: int = 0
    risk_flags: list = field(default_factory=list)


def monthly_totals(sales_history: pd.DataFrame) -> pd.DataFrame:
    df = sales_history.copy()
    df["month"] = df["date"].dt.to_period("M")
    return df.groupby(["dealer_id", "month"], as_index=False)["sales_amount"].sum()


def _linear_growth_rate(values: list[float]) -> float:
    if len(values) < 2:
        return 0.0
    y = np.array(values, dtype=float)
    x = np.arange(len(y))
    mean_y = y.mean()
    if mean_y == 0:
        return 0.0
    slope = np.polyfit(x, y, 1)[0]
    return float(slope / mean_y)


def _category_shares(sales_history: pd.DataFrame, dealer_id: str, months: list) -> dict:
    df = sales_history[(sales_history["dealer_id"] == dealer_id)].copy()
    df["month"] = df["date"].dt.to_period("M")
    df = df[df["month"].isin(months)]
    if df.empty:
        return {}
    totals = df.groupby("product_category")["sales_amount"].sum()
    grand_total = totals.sum()
    if grand_total == 0:
        return {}
    return (totals / grand_total).to_dict()


def compute_dealer_features(
    dealer_id: str,
    target_month: str,
    bundle: DataBundle,
    monthly_totals: pd.DataFrame,
) -> DealerFeatures:
    target_period = pd.Period(target_month, freq="M")

    dealer_totals = monthly_totals[monthly_totals["dealer_id"] == dealer_id].sort_values("month")
    history = dealer_totals[dealer_totals["month"] < target_period]
    months_available = list(history["month"])
    values_available = list(history["sales_amount"])

    months_of_history = len(months_available)

    last3 = values_available[-TRAILING_SHORT:]
    last6 = values_available[-TRAILING_LONG:]
    trailing_avg_3m = float(np.mean(last3)) if last3 else 0.0
    trailing_avg_6m = float(np.mean(last6)) if last6 else trailing_avg_3m
    mom_growth_rate = _linear_growth_rate(last6) if len(last6) >= 2 else 0.0

    yoy_growth_rate = None
    same_month_last_year = target_period - 12
    prior_year_row = history[history["month"] == same_month_last_year]
    this_year_recent = values_available[-1] if values_available else None
    if not prior_year_row.empty and this_year_recent is not None:
        prior_val = float(prior_year_row["sales_amount"].iloc[0])
        if prior_val > 0:
            yoy_growth_rate = (this_year_recent - prior_val) / prior_val

    calendar_month = target_period.month
    same_calendar = history[history["month"].apply(lambda p: p.month) == calendar_month]
    overall_avg = float(np.mean(values_available)) if values_available else 0.0
    if not same_calendar.empty and overall_avg > 0:
        seasonality_factor = float(same_calendar["sales_amount"].mean()) / overall_avg
    else:
        seasonality_factor = 1.0

    targets = bundle.targets_history[bundle.targets_history["dealer_id"] == dealer_id]
    targets = targets[targets["month"] < str(target_period)]
    target_accuracy_ratio = None
    target_bias = None
    if not targets.empty:
        ratios = targets["achieved_amount"] / targets["target_amount"].replace(0, np.nan)
        ratios = ratios.dropna()
        if not ratios.empty:
            target_accuracy_ratio = float(ratios.mean())
            target_bias = float(((targets["achieved_amount"] - targets["target_amount"]) / targets["target_amount"].replace(0, np.nan)).dropna().mean())

    window_start = (target_period - TRAILING_LONG).start_time
    window_end = target_period.start_time
    inv = bundle.invoices[
        (bundle.invoices["dealer_id"] == dealer_id)
        & (bundle.invoices["invoice_date"] >= window_start)
        & (bundle.invoices["invoice_date"] < window_end)
    ]
    late_payment_rate = None
    pending_payment_rate = None
    avg_delay_days = None
    if not inv.empty:
        total_n = len(inv)
        late_payment_rate = float((inv["payment_status"] == "Gecikti").sum() / total_n)
        pending_payment_rate = float((inv["payment_status"] == "Beklemede").sum() / total_n)
        late_paid = inv[(inv["payment_status"] == "Gecikti") & inv["paid_date"].notna()]
        if not late_paid.empty:
            delays = (late_paid["paid_date"] - late_paid["due_date"]).dt.days
            avg_delay_days = float(delays.mean())

    recent_months_3 = months_available[-TRAILING_SHORT:]
    prior_months_3 = months_available[-2 * TRAILING_SHORT:-TRAILING_SHORT]
    category_shares = _category_shares(bundle.sales_history, dealer_id, recent_months_3)
    prior_shares = _category_shares(bundle.sales_history, dealer_id, prior_months_3)
    category_share_change = {
        cat: round(category_shares.get(cat, 0.0) - prior_shares.get(cat, 0.0), 4)
        for cat in set(category_shares) | set(prior_shares)
    }

    risk_flags = []
    if mom_growth_rate < -0.03:
        risk_flags.append("Son aylarda satis trendi dususte")
    if late_payment_rate is not None and late_payment_rate > 0.2:
        risk_flags.append("Fatura odemelerinde belirgin gecikme riski")
    if pending_payment_rate is not None and pending_payment_rate > 0.3:
        risk_flags.append("Yuksek oranda bekleyen/odenmemis fatura var")
    if target_bias is not None and target_bias < -0.1:
        risk_flags.append("Gecmiste hedefler sistematik olarak tutturulamamis")
    if months_of_history < 3:
        risk_flags.append("Guvenilir tahmin icin yeterli gecmis veri yok")

    return DealerFeatures(
        dealer_id=dealer_id,
        target_month=target_month,
        trailing_avg_3m=trailing_avg_3m,
        trailing_avg_6m=trailing_avg_6m,
        mom_growth_rate=mom_growth_rate,
        yoy_growth_rate=yoy_growth_rate,
        seasonality_factor=seasonality_factor,
        target_accuracy_ratio=target_accuracy_ratio,
        target_bias=target_bias,
        late_payment_rate=late_payment_rate,
        pending_payment_rate=pending_payment_rate,
        avg_delay_days=avg_delay_days,
        category_shares=category_shares,
        category_share_change=category_share_change,
        recent_monthly_totals=list(zip([str(m) for m in months_available[-6:]], values_available[-6:])),
        months_of_history=months_of_history,
        risk_flags=risk_flags,
    )


def compute_features_all(target_month: str, bundle: DataBundle) -> dict[str, DealerFeatures]:
    totals = monthly_totals(bundle.sales_history)
    result = {}
    for dealer_id in bundle.dealers["dealer_id"]:
        result[dealer_id] = compute_dealer_features(dealer_id, target_month, bundle, totals)
    return result
