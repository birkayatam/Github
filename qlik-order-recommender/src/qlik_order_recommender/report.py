"""Excel ozet raporu ve bayi basina Markdown ikna raporu uretimi."""
from __future__ import annotations

from pathlib import Path
from typing import Union

import pandas as pd

from .insight_generator import DealerInsight
from .target_model import TargetRecommendation

PathLike = Union[str, Path]


def build_excel_report(
    dealers: pd.DataFrame,
    recommendations: dict[str, TargetRecommendation],
    out_path: PathLike,
) -> None:
    rows = []
    for _, dealer in dealers.iterrows():
        rec = recommendations.get(dealer["dealer_id"])
        if rec is None:
            continue
        rows.append(
            {
                "Bayi Kodu": dealer["dealer_id"],
                "Bayi Adi": dealer["dealer_name"],
                "Bolge": dealer["region"],
                "Segment": dealer["segment"],
                "Hedef Ayi": rec.target_month,
                "Onerilen Hedef": rec.recommended_target,
                "Min": rec.target_min,
                "Max": rec.target_max,
                "Guven": rec.confidence,
                "Yontem": rec.method,
                "Risk Bayraklari": "; ".join(rec.risk_flags) if rec.risk_flags else "",
            }
        )
    df = pd.DataFrame(rows)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out_path, index=False)


def build_markdown_reports(
    insights: dict[str, DealerInsight],
    out_dir: PathLike,
    target_month: str,
) -> list[Path]:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for dealer_id, insight in insights.items():
        path = out_dir / f"{dealer_id}_{target_month}.md"
        path.write_text(insight.narrative + "\n", encoding="utf-8")
        written.append(path)
    return written
