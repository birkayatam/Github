"""Bayi bazinda urun-adet onerisini, gecmis satin alma verisine dayanan
ikna edici bir metne donusturur. ANTHROPIC_API_KEY tanimli degilse ya da
API cagrisi basarisiz olursa kural tabanli sablon fallback kullanilir."""
from __future__ import annotations

import os
from dataclasses import dataclass

import pandas as pd

DEFAULT_MODEL = "claude-sonnet-5"
HISTORY_MONTHS = 12
MAX_DISPLAYED_PRODUCTS = 15


@dataclass
class DealerOrderInsight:
    dealer_name: str
    narrative: str
    source: str  # "llm" veya "template"


def _dealer_product_history(transactions: pd.DataFrame, dealer_name: str, product_code: str) -> dict:
    df = transactions[(transactions["dealer_name"] == dealer_name) & (transactions["product_code"] == product_code)]
    if df.empty:
        return {"total_fatura": 0.0, "months_bought": 0, "last_month": None}
    df = df.sort_values("month").tail(HISTORY_MONTHS)
    bought = df[df["fatura"] > 0]
    return {
        "total_fatura": float(df["fatura"].clip(lower=0).sum()),
        "months_bought": int((df["fatura"] > 0).sum()),
        "last_month": bought["month"].max() if not bought.empty else None,
    }


def _reason_sentence(product_code: str, qty: int, hist: dict) -> str:
    if hist["months_bought"] >= 3:
        return (
            f"- **{product_code}**: {qty} adet — bu bayi son {HISTORY_MONTHS} ayda {hist['months_bought']} "
            f"farkli ayda bu urunden toplam {hist['total_fatura']:.0f} adet faturalamis "
            f"(son alim: {hist['last_month']}); duzenli talep gordugu bir urun."
        )
    if hist["months_bought"] >= 1:
        return (
            f"- **{product_code}**: {qty} adet — bu bayi bu urunden daha once sinirli miktarda "
            f"({hist['total_fatura']:.0f} adet, {hist['months_bought']} ay) almis; kategori genelinde "
            "talep gordugu icin oneriliyor."
        )
    return (
        f"- **{product_code}**: {qty} adet — bu bayi bu urunden son {HISTORY_MONTHS} ayda hic almamis; "
        "pazar genelinde talep gordugu icin tanitim/genisleme amacli oneriliyor."
    )


def _format_facts(
    dealer_name: str,
    remaining: float,
    recommendations: pd.Series,
    transactions: pd.DataFrame,
) -> tuple[str, list[str], int, float]:
    """Facts metnini ve gerekce cumlelerini en yuksek adetli ilk
    MAX_DISPLAYED_PRODUCTS urun icin uretir; kalanini (varsa) adet/urun
    sayisi olarak ayrica dondurur (uzun listelerin saha raporunu
    okunmaz hale getirmemesi icin)."""
    covered = float(recommendations.sum())
    lines = [
        f"Bayi: {dealer_name}",
        f"Kalan hedef: {remaining:.0f} adet",
        f"Onerilen toplam: {covered:.0f} adet",
    ]
    positive = recommendations[recommendations > 0].sort_values(ascending=False)
    shown = positive.iloc[:MAX_DISPLAYED_PRODUCTS]
    rest = positive.iloc[MAX_DISPLAYED_PRODUCTS:]

    reasons = []
    for product_code, qty in shown.items():
        hist = _dealer_product_history(transactions, dealer_name, product_code)
        reasons.append(_reason_sentence(product_code, int(qty), hist))
        lines.append(
            f"- {product_code}: onerilen {int(qty)} adet, son {HISTORY_MONTHS} ayda faturalanan "
            f"{hist['total_fatura']:.0f} adet, {hist['months_bought']} farkli ayda alim, "
            f"son alim ayi: {hist['last_month'] or 'yok'}"
        )
    return "\n".join(lines), reasons, len(rest), float(rest.sum())


def _template_narrative(
    dealer_name: str,
    remaining: float,
    shortfall: float,
    recommendations: pd.Series,
    transactions: pd.DataFrame,
) -> str:
    facts, reasons, rest_count, rest_qty = _format_facts(dealer_name, remaining, recommendations, transactions)
    covered = float(recommendations.sum())

    parts = [
        f"## {dealer_name} - Siparis Onerisi",
        "",
        f"**Kalan hedef:** {remaining:.0f} adet  |  **Onerilen toplam:** {covered:.0f} adet",
    ]
    if shortfall > 0.5:
        parts.append(
            f"**Not:** urun kotasi kisitli oldugu icin {shortfall:.0f} adet kalan hedef bu turda "
            "karsilanamadi; bir sonraki kota doneminde tekrar degerlendirilmeli."
        )
    parts.append("")
    parts.append("### Urun bazinda oneri ve gerekce")
    if reasons:
        parts.extend(reasons)
        if rest_count:
            parts.append(f"- ...ve {rest_count} urun daha, toplam {rest_qty:.0f} adet (tam liste icin Excel raporuna bakin).")
    else:
        parts.append("- Bu bayi icin kota/veri kisiti nedeniyle bu turda urun onerisi olusturulamadi.")
    parts.append("")
    parts.append("### Destekleyici veriler")
    parts.append("```")
    parts.append(facts)
    parts.append("```")
    return "\n".join(parts)


def _llm_narrative(
    dealer_name: str,
    remaining: float,
    shortfall: float,
    recommendations: pd.Series,
    transactions: pd.DataFrame,
    model: str,
) -> str:
    import anthropic

    client = anthropic.Anthropic()
    facts, _, rest_count, rest_qty = _format_facts(dealer_name, remaining, recommendations, transactions)
    if rest_count:
        facts += f"\n(+ {rest_count} urun daha, toplam {rest_qty:.0f} adet -- yukarida listelenmedi)"
    shortfall_note = (
        f"Kota kisitlamasi nedeniyle {shortfall:.0f} adetlik hedef bu turda karsilanamadi, bunu belirt."
        if shortfall > 0.5
        else "Kalan hedefin tamami onerilen urunlerle karsilaniyor."
    )
    prompt = (
        "Asagida bir bayiye, kalan siparis hedefini kapatmak icin onerilen urun ve adetler, "
        "ve bu bayinin bu urunlerdeki gecmis fatura verisi var. Saha satis temsilcisinin bayi "
        "ziyaretinde kullanacagi, Turkce, kisa ve somut veriye dayali bir siparis onerisi metni yaz. "
        "Uydurma veri veya rakam ekleme, sadece verilen bilgileri kullan. "
        f"{shortfall_note} "
        "Cikti Markdown formatinda olsun ve su bolumleri icersin: "
        "'## <Bayi Adi> - Siparis Onerisi', '### Urun bazinda oneri ve gerekce' (madde madde).\n\n"
        f"Veriler:\n{facts}"
    )
    response = client.messages.create(model=model, max_tokens=800, messages=[{"role": "user", "content": prompt}])
    return "".join(block.text for block in response.content if getattr(block, "type", None) == "text").strip()


def generate_order_insight(
    dealer_name: str,
    remaining: float,
    shortfall: float,
    recommendations: pd.Series,
    transactions: pd.DataFrame,
    model: str = DEFAULT_MODEL,
) -> DealerOrderInsight:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if api_key:
        try:
            narrative = _llm_narrative(dealer_name, remaining, shortfall, recommendations, transactions, model)
            if narrative:
                return DealerOrderInsight(dealer_name=dealer_name, narrative=narrative, source="llm")
        except Exception:
            pass  # API erisimi basarisizsa sablon fallback'e dusulur
    narrative = _template_narrative(dealer_name, remaining, shortfall, recommendations, transactions)
    return DealerOrderInsight(dealer_name=dealer_name, narrative=narrative, source="template")
