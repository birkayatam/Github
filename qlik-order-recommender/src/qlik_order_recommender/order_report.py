"""Siparis tahsis sonuclarini Excel (bayi ozeti + bayi x urun detayi +
doldurulmus siparis dosyasi) ve bayi basina Markdown oneri metni olarak
yazar."""
from __future__ import annotations

from pathlib import Path
from typing import Union

import pandas as pd

from .allocation import AllocationResult
from .order_insight import DealerOrderInsight, MarketContext

PathLike = Union[str, Path]


def _build_market_overview(product_quota: pd.DataFrame, market: MarketContext) -> pd.DataFrame:
    """Urun bazinda pazar geneli (tum bayiler) mevsimsellik/YoY/sell-through
    ozeti -- kullaniciyi sadece tek bir bayiyle sinirli birakmayan,
    genel duruma tek bakista cevap veren referans tablo."""
    avg_sell_through = market.sell_through["sell_through_ratio"].groupby("product_code").mean()

    rows = []
    for product in product_quota["product_code"]:
        seasonal = float(market.product_seasonality.seasonal_index.get(product, 1.0))
        yoy = market.product_seasonality.yoy_growth.get(product)
        rows.append(
            {
                "Urun Kodu": product,
                f"Mevsimsellik Endeksi ({market.period_label})": round(seasonal, 2),
                "YoY Buyume (Pazar)": None if (yoy is None or pd.isna(yoy)) else round(float(yoy), 3),
                "Ortalama Sell-Through (Pazar)": (
                    round(float(avg_sell_through.get(product)), 2) if product in avg_sell_through.index and not pd.isna(avg_sell_through.get(product)) else None
                ),
            }
        )
    return pd.DataFrame(rows).sort_values(f"Mevsimsellik Endeksi ({market.period_label})", ascending=False)


def build_excel_report(
    dealer_targets: pd.DataFrame,
    remaining: pd.Series,
    result: AllocationResult,
    product_quota: pd.DataFrame,
    out_path: PathLike,
    market: MarketContext | None = None,
) -> None:
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    onerilen_toplam = result.allocations.sum(axis=1)
    summary_rows = []
    for _, row in dealer_targets.iterrows():
        dealer = row["dealer_name"]
        summary_rows.append(
            {
                "Bayi": dealer,
                "LPH2": row["lph2"],
                "Hedef": row["hedef"],
                "Fiili": row["fiili"],
                "Kalan": remaining.get(dealer, 0.0),
                "Onerilen Toplam": float(onerilen_toplam.get(dealer, 0.0)),
                "Karsilanamayan": float(result.dealer_shortfall.get(dealer, 0.0)),
            }
        )
    summary_df = pd.DataFrame(summary_rows)

    detail_rows = []
    for dealer in result.allocations.index:
        for product in result.allocations.columns:
            qty = int(result.allocations.loc[dealer, product])
            if qty > 0:
                detail_rows.append({"Bayi": dealer, "Urun Kodu": product, "Onerilen Adet": qty})
    detail_df = pd.DataFrame(detail_rows, columns=["Bayi", "Urun Kodu", "Onerilen Adet"])

    filled_quota = product_quota.copy()
    filled_quota["SIPARIS (tahsis edilen)"] = filled_quota["product_code"].map(
        result.allocations.sum(axis=0)
    ).fillna(0).astype(int)
    filled_quota["Kalan Kota"] = filled_quota["product_code"].map(result.product_leftover).fillna(0).astype(int)
    filled_quota["Kota Asildi Mi"] = filled_quota["product_code"].map(result.product_oversubscribed).fillna(False)
    filled_quota = filled_quota.rename(columns={"product_code": "Urun Kodu", "quota": "ADET"})

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Bayi Ozet", index=False)
        detail_df.to_excel(writer, sheet_name="Bayi x Urun Tahsis", index=False)
        filled_quota.to_excel(writer, sheet_name="Siparis Dosyasi", index=False)
        if market is not None:
            _build_market_overview(product_quota, market).to_excel(writer, sheet_name="Pazar Genel Bakis", index=False)


def build_markdown_reports(insights: dict[str, DealerOrderInsight], out_dir: PathLike) -> list[Path]:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for dealer, insight in insights.items():
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in dealer)[:80].strip()
        path = out_dir / f"{safe_name}.md"
        path.write_text(insight.narrative + "\n", encoding="utf-8")
        written.append(path)
    return written
