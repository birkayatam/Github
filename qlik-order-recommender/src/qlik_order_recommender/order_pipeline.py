"""Siparis tahsis pipeline'i: gercek Qlik islem gecmisi + bayi ceyrek
hedefi + urun bazli sirket kotasindan, bayi basina 'hangi urunden kac
adet siparis toplanmali' onerisini ve buna dayanan ikna metnini uretir.
CLI giris noktasi."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .allocation import allocate, compute_affinity
from .order_features import compute_seasonality, compute_sell_through, detect_target_quarter
from .order_ingestion import load_order_data, resolve_dealer_names
from .order_insight import DEFAULT_MODEL, MarketContext, generate_order_insight
from .order_report import build_excel_report, build_markdown_reports


def run(
    transactions_path: str,
    dealer_targets_path: str,
    product_quota_path: str,
    out_dir: str,
    mapping_path: str | None = None,
    model: str = DEFAULT_MODEL,
) -> Path:
    bundle = load_order_data(transactions_path, dealer_targets_path, product_quota_path, mapping_path)

    if bundle.dealer_targets["lph2"].nunique() > 1:
        print(
            "[order_pipeline] Uyari: bayi hedef dosyasinda birden fazla LPH2 kategorisi var; "
            "bu surum tek bir urun havuzuna (siparis dosyasindaki urunler) gore tahsis yapar, "
            "kategoriler ayri ayri degerlendirilmiyor."
        )

    name_map, unresolved = resolve_dealer_names(bundle.dealer_targets, bundle.transactions)
    if unresolved:
        print(f"[order_pipeline] Uyari: {len(unresolved)} bayi Qlik gecmisiyle eslendirilemedi "
              f"(gecmis verisi olmadigi icin pazar ortalamasina gore tahsis edilecek): {unresolved}")

    reverse_map = {canonical: original for original, canonical in name_map.items()}
    transactions = bundle.transactions.copy()
    transactions["dealer_name"] = transactions["dealer_name"].replace(reverse_map)

    targets = bundle.dealer_targets.groupby("dealer_name", as_index=False).agg(
        lph2=("lph2", lambda s: ", ".join(sorted(set(s)))),
        hedef=("hedef", "sum"),
        fiili=("fiili", "sum"),
    )
    remaining = (targets.set_index("dealer_name")["hedef"] - targets.set_index("dealer_name")["fiili"]).clip(lower=0)

    quota = bundle.product_quota.set_index("product_code")["quota"]

    target_fiscal_year, target_quarter = detect_target_quarter(transactions)
    product_seasonality = compute_seasonality(transactions, "product_code", target_fiscal_year, target_quarter)
    dealer_seasonality = compute_seasonality(transactions, "dealer_name", target_fiscal_year, target_quarter)
    sell_through = compute_sell_through(transactions)
    market = MarketContext(
        sell_through=sell_through,
        product_seasonality=product_seasonality,
        dealer_seasonality=dealer_seasonality,
        target_fiscal_year=target_fiscal_year,
        target_quarter=target_quarter,
    )
    print(f"[order_pipeline] Hedef donem otomatik tespit edildi: {market.period_label}")

    affinity = compute_affinity(
        transactions,
        dealer_names=list(remaining.index),
        product_codes=list(quota.index),
        product_seasonal_index=product_seasonality.seasonal_index,
        sell_through=sell_through,
    )
    result = allocate(remaining, quota, affinity)

    insights = {}
    for dealer in remaining.index:
        insights[dealer] = generate_order_insight(
            dealer_name=dealer,
            remaining=float(remaining[dealer]),
            shortfall=float(result.dealer_shortfall.get(dealer, 0.0)),
            recommendations=result.allocations.loc[dealer],
            transactions=transactions,
            market=market,
            model=model,
        )

    out_dir_path = Path(out_dir)
    excel_path = out_dir_path / "siparis_onerileri.xlsx"
    build_excel_report(targets, remaining, result, bundle.product_quota, excel_path, market=market)
    build_markdown_reports(insights, out_dir_path / "bayi_raporlari")

    return excel_path


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Bayi kalan hedefini urun bazinda dagitip siparis onerisi uretir.")
    parser.add_argument("--transactions", required=True, help="Qlik ham islem gecmisi (bayi x urun x ay) dosyasi")
    parser.add_argument("--dealer-targets", required=True, help="Bayi bazli donemsel hedef/fiili dosyasi")
    parser.add_argument("--product-quota", required=True, help="Urun kodu bazli sirket kotasi dosyasi")
    parser.add_argument("--mapping", default=None, help="Sutun eslemesi YAML dosyasi (opsiyonel)")
    parser.add_argument("--out", required=True, help="Cikti klasoru")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="Ikna metni icin kullanilacak Claude modeli")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = _parse_args(argv if argv is not None else sys.argv[1:])
    excel_path = run(
        transactions_path=args.transactions,
        dealer_targets_path=args.dealer_targets,
        product_quota_path=args.product_quota,
        out_dir=args.out,
        mapping_path=args.mapping,
        model=args.model,
    )
    print(f"Ozet rapor: {excel_path}")
    print(f"Bayi bazli siparis oneri raporlari: {Path(args.out) / 'bayi_raporlari'}")


if __name__ == "__main__":
    main()
