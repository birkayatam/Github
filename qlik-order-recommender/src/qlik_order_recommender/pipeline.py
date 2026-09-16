"""Uctan uca orkestrasyon: ingestion -> features -> target_model ->
insight_generator -> report. CLI giris noktasi."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .features import compute_features_all
from .ingestion import load_all
from .insight_generator import DEFAULT_MODEL, generate_insight
from .report import build_excel_report, build_markdown_reports
from .target_model import recommend_targets_all


def run(
    dealers_path: str,
    sales_path: str,
    invoices_path: str,
    targets_path: str,
    mapping_path: str,
    target_month: str,
    out_dir: str,
    model: str = DEFAULT_MODEL,
) -> Path:
    bundle = load_all(dealers_path, sales_path, invoices_path, targets_path, mapping_path)

    features_by_dealer = compute_features_all(target_month, bundle)
    recommendations = recommend_targets_all(bundle, features_by_dealer)

    dealer_names = dict(zip(bundle.dealers["dealer_id"], bundle.dealers["dealer_name"]))
    insights = {}
    for dealer_id, rec in recommendations.items():
        insights[dealer_id] = generate_insight(
            dealer_names.get(dealer_id, dealer_id),
            features_by_dealer[dealer_id],
            rec,
            model=model,
        )

    out_dir_path = Path(out_dir)
    excel_path = out_dir_path / f"hedef_onerileri_{target_month}.xlsx"
    build_excel_report(bundle.dealers, recommendations, excel_path)
    build_markdown_reports(insights, out_dir_path, target_month)

    return excel_path


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Qlik verisinden bayi hedef/ikna onerisi uretir.")
    parser.add_argument("--dealers", required=True, help="Bayi listesi CSV/Excel dosyasi")
    parser.add_argument("--sales", required=True, help="Satis gecmisi CSV/Excel dosyasi")
    parser.add_argument("--invoices", required=True, help="Fatura gecmisi CSV/Excel dosyasi")
    parser.add_argument("--targets", required=True, help="Gecmis hedef/gerceklesme CSV/Excel dosyasi")
    parser.add_argument("--mapping", required=True, help="Sutun eslemesi YAML dosyasi")
    parser.add_argument("--month", required=True, help="Hedef ay, YYYY-MM formatinda (ornek: 2026-10)")
    parser.add_argument("--out", required=True, help="Cikti klasoru")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="Ikna metni icin kullanilacak Claude modeli")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = _parse_args(argv if argv is not None else sys.argv[1:])
    excel_path = run(
        dealers_path=args.dealers,
        sales_path=args.sales,
        invoices_path=args.invoices,
        targets_path=args.targets,
        mapping_path=args.mapping,
        target_month=args.month,
        out_dir=args.out,
        model=args.model,
    )
    print(f"Ozet rapor: {excel_path}")
    print(f"Bayi bazli ikna raporlari: {Path(args.out)}/<bayi_kodu>_{args.month}.md")


if __name__ == "__main__":
    main()
