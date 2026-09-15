"""LLM (Claude API) ile bayi ozelinde ikna edici, veriye dayali anlatim
uretir. ANTHROPIC_API_KEY tanimli degilse ya da API cagrisi basarisiz olursa
kural tabanli sablon fallback kullanilir -- boylece internet erisimi
olmadan da uctan uca calisir."""
from __future__ import annotations

import os
from dataclasses import dataclass

from .features import DealerFeatures
from .target_model import TargetRecommendation

DEFAULT_MODEL = "claude-sonnet-5"


@dataclass
class DealerInsight:
    dealer_id: str
    narrative: str
    source: str  # "llm" veya "template"


def _format_supporting_facts(dealer_name: str, f: DealerFeatures, rec: TargetRecommendation) -> str:
    lines = [
        f"Bayi: {dealer_name} ({f.dealer_id})",
        f"Hedef ayi: {rec.target_month}",
        f"Onerilen hedef: {rec.recommended_target:,.0f} "
        f"(aralik: {rec.target_min:,.0f} - {rec.target_max:,.0f}, guven: %{rec.confidence * 100:.0f})",
        f"Son 3 ay ortalama satis: {f.trailing_avg_3m:,.0f}",
        f"Son 6 ay ortalama satis: {f.trailing_avg_6m:,.0f}",
        f"Aylik trend buyume orani: %{f.mom_growth_rate * 100:.1f}",
    ]
    if f.yoy_growth_rate is not None:
        lines.append(f"Yillik (YoY) buyume orani: %{f.yoy_growth_rate * 100:.1f}")
    lines.append(f"Mevsimsellik faktoru (hedef ay icin): {f.seasonality_factor:.2f}")
    if f.target_accuracy_ratio is not None:
        lines.append(f"Gecmis hedef isabet orani (gerceklesen/hedef ortalamasi): %{f.target_accuracy_ratio * 100:.0f}")
    if f.late_payment_rate is not None:
        lines.append(f"Son 6 ayda geciken fatura orani: %{f.late_payment_rate * 100:.0f}")
    if f.pending_payment_rate is not None:
        lines.append(f"Bekleyen/odenmemis fatura orani: %{f.pending_payment_rate * 100:.0f}")
    if f.category_share_change:
        growers = sorted(f.category_share_change.items(), key=lambda kv: kv[1], reverse=True)
        top = growers[0]
        if top[1] > 0:
            lines.append(
                f"En hizli buyuyen urun kategorisi payi: {top[0]} "
                f"(son 3 ayda {top[1] * 100:+.1f} puan degisim)"
            )
    if rec.risk_flags:
        lines.append("Risk bayraklari: " + "; ".join(rec.risk_flags))
    return "\n".join(lines)


def _template_narrative(dealer_name: str, f: DealerFeatures, rec: TargetRecommendation) -> str:
    facts = _format_supporting_facts(dealer_name, f, rec)
    growth_txt = "artis" if f.mom_growth_rate >= 0 else "azalis"
    parts = [
        f"## {dealer_name} - {rec.target_month} Hedef Onerisi",
        "",
        f"**Onerilen hedef:** {rec.recommended_target:,.0f} "
        f"(gerceklesebilir aralik: {rec.target_min:,.0f} - {rec.target_max:,.0f}, guven: %{rec.confidence * 100:.0f})",
        "",
        "### Ikna icin kullanilacak veri noktalari",
        f"- Son 3 aylik ortalama satis {f.trailing_avg_3m:,.0f}, trendde aylik "
        f"%{abs(f.mom_growth_rate) * 100:.1f} {growth_txt} var.",
    ]
    if f.yoy_growth_rate is not None:
        yoy_txt = "buyume" if f.yoy_growth_rate >= 0 else "gerileme"
        parts.append(f"- Gecen yilin ayni ayina gore %{abs(f.yoy_growth_rate) * 100:.1f} {yoy_txt} gozlemleniyor.")
    if f.target_accuracy_ratio is not None:
        parts.append(
            f"- Bayi gecmiste hedeflerin ortalama %{f.target_accuracy_ratio * 100:.0f}'ini gerceklestirmis; "
            "hedef bu gerceklesme oranina gore kalibre edildi."
        )
    if f.category_share_change:
        growers = sorted(f.category_share_change.items(), key=lambda kv: kv[1], reverse=True)
        top = growers[0]
        if top[1] > 0:
            parts.append(
                f"- {top[0]} kategorisindeki payi son 3 ayda {top[1] * 100:+.1f} puan artti; "
                "bu buyumeyi one cikararak ek siparis tesvik edilebilir."
            )
    if f.late_payment_rate is not None and f.late_payment_rate > 0:
        parts.append(
            f"- Son 6 ayda faturalarin %{f.late_payment_rate * 100:.0f}'i gecikmeli odenmis; "
            "hedef gorusmesinde odeme planinin da konusulmasi onerilir."
        )
    parts.append("")
    parts.append("### Dikkat edilmesi gerekenler")
    if rec.risk_flags:
        for flag in rec.risk_flags:
            parts.append(f"- {flag}")
    else:
        parts.append("- Belirgin bir risk sinyali tespit edilmedi.")
    parts.append("")
    parts.append("### Destekleyici veriler")
    parts.append("```")
    parts.append(facts)
    parts.append("```")
    return "\n".join(parts)


def _llm_narrative(dealer_name: str, f: DealerFeatures, rec: TargetRecommendation, model: str) -> str:
    import anthropic

    client = anthropic.Anthropic()
    facts = _format_supporting_facts(dealer_name, f, rec)
    prompt = (
        "Asagida bir bayiye ait satis, fatura ve hedef verisinden turetilmis ozet bilgiler var. "
        "Bu bilgileri kullanarak saha satis temsilcisinin bayi ziyaretinde kullanacagi, "
        "Turkce, kisa ve somut veriye dayali bir hedef gerekcesi ve ikna konusma noktalari metni yaz. "
        "Uydurma veri veya rakam ekleme, sadece verilen bilgileri kullan. "
        "Cikti Markdown formatinda olsun ve su bolumleri icersin: "
        "'## <Bayi Adi> - Hedef Onerisi', '### Ikna icin kullanilacak veri noktalari' (madde madde), "
        "'### Dikkat edilmesi gerekenler' (risk bayraklarina dayali, yoksa bunu belirt).\n\n"
        f"Veriler:\n{facts}"
    )
    response = client.messages.create(
        model=model,
        max_tokens=800,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(block.text for block in response.content if getattr(block, "type", None) == "text").strip()


def generate_insight(
    dealer_name: str,
    features: DealerFeatures,
    recommendation: TargetRecommendation,
    model: str = DEFAULT_MODEL,
) -> DealerInsight:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if api_key:
        try:
            narrative = _llm_narrative(dealer_name, features, recommendation, model)
            if narrative:
                return DealerInsight(dealer_id=features.dealer_id, narrative=narrative, source="llm")
        except Exception:
            pass  # API erisimi basarisizsa sablon fallback'e dusulur
    narrative = _template_narrative(dealer_name, features, recommendation)
    return DealerInsight(dealer_id=features.dealer_id, narrative=narrative, source="template")
