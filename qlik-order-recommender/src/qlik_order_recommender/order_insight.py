"""Bayi bazinda urun-adet onerisini, gecmis satin alma verisine + sezonsal/
pazar geneli baglama dayanan ikna edici bir metne donusturur.
ANTHROPIC_API_KEY tanimli degilse ya da API cagrisi basarisiz olursa kural
tabanli sablon fallback kullanilir."""
from __future__ import annotations

import os
from dataclasses import dataclass

import pandas as pd

from .order_features import FISCAL_QUARTER_MONTHS, SELL_THROUGH_HIGH, SELL_THROUGH_LOW, SeasonalityResult

DEFAULT_MODEL = "claude-sonnet-5"
HISTORY_MONTHS = 12
MAX_DISPLAYED_PRODUCTS = 15

# YoY buyume/dusus metinde ancak makul bir aralikta (%-95 ile %500 arasinda)
# gosterilir -- cok kucuk bir taban uzerinden hesaplanan (ör. 2 adetten 20
# adede) oranlar sayisal olarak dogru ama okuyucuyu yanlis yonlendirebilir.
YOY_DISPLAY_BOUNDS = (-0.95, 5.0)

QUARTER_MONTH_NAMES_TR = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
    7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık",
}


def quarter_label(fiscal_year: int, quarter: int) -> str:
    months = FISCAL_QUARTER_MONTHS[quarter]
    names = "-".join(QUARTER_MONTH_NAMES_TR[m] for m in months)
    return f"{names} {fiscal_year}"


@dataclass
class MarketContext:
    """order_pipeline'in bir kez hesaplayip tum bayiler icin paylastigi
    pazar-geneli/sezonsal baglam."""

    sell_through: pd.DataFrame  # index=(dealer_name, product_code), col: sell_through_ratio
    product_seasonality: SeasonalityResult  # index=product_code
    dealer_seasonality: SeasonalityResult  # index=dealer_name
    target_fiscal_year: int
    target_quarter: int

    @property
    def period_label(self) -> str:
        return quarter_label(self.target_fiscal_year, self.target_quarter)


@dataclass
class DealerOrderInsight:
    dealer_name: str
    narrative: str
    source: str  # "llm" veya "template"


def _dealer_product_history(
    transactions: pd.DataFrame, dealer_name: str, product_code: str, market: MarketContext
) -> dict:
    df = transactions[(transactions["dealer_name"] == dealer_name) & (transactions["product_code"] == product_code)]

    if df.empty:
        result = {"total_fatura": 0.0, "months_bought": 0, "last_month": None, "sellout_total": 0.0}
    else:
        df = df.sort_values("month").tail(HISTORY_MONTHS)
        bought = df[df["fatura"] > 0]
        result = {
            "total_fatura": float(df["fatura"].clip(lower=0).sum()),
            "months_bought": int((df["fatura"] > 0).sum()),
            "last_month": bought["month"].max() if not bought.empty else None,
            "sellout_total": float((df["bildirim"].clip(lower=0) + df["montaj"].clip(lower=0)).sum()),
        }

    sell_through_ratio = None
    if (dealer_name, product_code) in market.sell_through.index:
        val = market.sell_through.loc[(dealer_name, product_code), "sell_through_ratio"]
        sell_through_ratio = None if pd.isna(val) else float(val)
    result["sell_through_ratio"] = sell_through_ratio

    result["product_seasonal_index"] = float(market.product_seasonality.seasonal_index.get(product_code, 1.0))
    yoy = market.product_seasonality.yoy_growth.get(product_code)
    result["product_yoy_growth"] = None if (yoy is None or pd.isna(yoy)) else float(yoy)

    return result


def _reason_sentence(product_code: str, qty: int, hist: dict, market: MarketContext) -> str:
    if hist["months_bought"] >= 3:
        base = (
            f"- **{product_code}**: {qty} adet — bu bayi son {HISTORY_MONTHS} ayda {hist['months_bought']} "
            f"farkli ayda bu urunden toplam {hist['total_fatura']:.0f} adet faturalamis, "
            f"{hist['sellout_total']:.0f} adet son musteriye satmis/kurmus (son alim: {hist['last_month']})."
        )
    elif hist["months_bought"] >= 1:
        base = (
            f"- **{product_code}**: {qty} adet — bu bayi bu urunden daha once sinirli miktarda "
            f"({hist['total_fatura']:.0f} adet fatura, {hist['sellout_total']:.0f} adet satis/kurulum) almis."
        )
    else:
        base = (
            f"- **{product_code}**: {qty} adet — bu bayi bu urunden son {HISTORY_MONTHS} ayda hic almamis; "
            "pazar genelindeki talebe dayanarak tanitim/genisleme amacli oneriliyor."
        )

    extras = []
    ratio = hist["sell_through_ratio"]
    if ratio is not None:
        if ratio < SELL_THROUGH_LOW:
            extras.append(f"dikkat: sell-through orani dusuk ({ratio:.2f}) — mevcut stok muhtemelen hala eritilmedi")
        elif ratio > SELL_THROUGH_HIGH:
            extras.append(f"sell-through orani yuksek ({ratio:.2f}) — hizli satiyor, restok mantikli")

    seasonal = hist["product_seasonal_index"]
    if seasonal >= 1.3:
        extras.append(f"bu urun pazar genelinde {market.period_label} doneminde tipik olarak guclu (endeks {seasonal:.2f})")
    elif seasonal <= 0.7:
        extras.append(f"bu urun pazar genelinde {market.period_label} doneminde tipik olarak zayif (endeks {seasonal:.2f})")

    yoy = hist["product_yoy_growth"]
    if yoy is not None and YOY_DISPLAY_BOUNDS[0] <= yoy <= YOY_DISPLAY_BOUNDS[1]:
        direction = "artis" if yoy >= 0 else "azalis"
        extras.append(f"pazar genelinde gecen yila gore %{abs(yoy) * 100:.0f} {direction}")

    if extras:
        base += " (" + "; ".join(extras) + ")"
    return base


def _dealer_seasonal_intro(dealer_name: str, market: MarketContext) -> str:
    idx = float(market.dealer_seasonality.seasonal_index.get(dealer_name, 1.0))
    yoy = market.dealer_seasonality.yoy_growth.get(dealer_name)
    yoy_val = None if (yoy is None or pd.isna(yoy)) else float(yoy)

    parts = []
    if idx >= 1.15:
        parts.append(f"bu bayi tarihsel olarak {market.period_label} doneminde ortalamadan guclu satis yapiyor (mevsimsellik endeksi {idx:.2f})")
    elif idx <= 0.85:
        parts.append(f"bu bayi tarihsel olarak {market.period_label} doneminde ortalamadan zayif satis yapiyor (mevsimsellik endeksi {idx:.2f})")

    if yoy_val is not None and YOY_DISPLAY_BOUNDS[0] <= yoy_val <= YOY_DISPLAY_BOUNDS[1]:
        direction = "buyume" if yoy_val >= 0 else "gerileme"
        prior_period = quarter_label(market.target_fiscal_year - 1, market.target_quarter)
        parts.append(f"gecen yilin ayni donemine ({prior_period}) gore %{abs(yoy_val) * 100:.0f} {direction}")

    if not parts:
        return ""
    return "**Sezonsal baglam:** " + "; ".join(parts) + "."


def _format_facts(
    dealer_name: str,
    remaining: float,
    recommendations: pd.Series,
    transactions: pd.DataFrame,
    market: MarketContext,
) -> tuple[str, list[str], int, float]:
    """Facts metnini ve gerekce cumlelerini en yuksek adetli ilk
    MAX_DISPLAYED_PRODUCTS urun icin uretir; kalanini (varsa) adet/urun
    sayisi olarak ayrica dondurur (uzun listelerin saha raporunu
    okunmaz hale getirmemesi icin)."""
    covered = float(recommendations.sum())
    lines = [
        f"Bayi: {dealer_name}",
        f"Hedef donem: {market.period_label}",
        f"Kalan hedef: {remaining:.0f} adet",
        f"Onerilen toplam: {covered:.0f} adet",
    ]
    seasonal_intro = _dealer_seasonal_intro(dealer_name, market)
    if seasonal_intro:
        lines.append(seasonal_intro)

    positive = recommendations[recommendations > 0].sort_values(ascending=False)
    shown = positive.iloc[:MAX_DISPLAYED_PRODUCTS]
    rest = positive.iloc[MAX_DISPLAYED_PRODUCTS:]

    reasons = []
    for product_code, qty in shown.items():
        hist = _dealer_product_history(transactions, dealer_name, product_code, market)
        reasons.append(_reason_sentence(product_code, int(qty), hist, market))
        line = (
            f"- {product_code}: onerilen {int(qty)} adet, son {HISTORY_MONTHS} ayda faturalanan "
            f"{hist['total_fatura']:.0f} adet / son musteriye satilan-kurulan {hist['sellout_total']:.0f} adet"
        )
        if hist["sell_through_ratio"] is not None:
            line += f", sell-through: {hist['sell_through_ratio']:.2f}"
        lines.append(line)
    return "\n".join(lines), reasons, len(rest), float(rest.sum())


def _template_narrative(
    dealer_name: str,
    remaining: float,
    shortfall: float,
    recommendations: pd.Series,
    transactions: pd.DataFrame,
    market: MarketContext,
) -> str:
    facts, reasons, rest_count, rest_qty = _format_facts(dealer_name, remaining, recommendations, transactions, market)
    covered = float(recommendations.sum())

    parts = [
        f"## {dealer_name} - Siparis Onerisi ({market.period_label})",
        "",
        f"**Kalan hedef:** {remaining:.0f} adet  |  **Onerilen toplam:** {covered:.0f} adet",
    ]
    seasonal_intro = _dealer_seasonal_intro(dealer_name, market)
    if seasonal_intro:
        parts.append(seasonal_intro)
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
    market: MarketContext,
    model: str,
) -> str:
    import anthropic

    client = anthropic.Anthropic()
    facts, _, rest_count, rest_qty = _format_facts(dealer_name, remaining, recommendations, transactions, market)
    if rest_count:
        facts += f"\n(+ {rest_count} urun daha, toplam {rest_qty:.0f} adet -- yukarida listelenmedi)"
    shortfall_note = (
        f"Kota kisitlamasi nedeniyle {shortfall:.0f} adetlik hedef bu turda karsilanamadi, bunu belirt."
        if shortfall > 0.5
        else "Kalan hedefin tamami onerilen urunlerle karsilaniyor."
    )
    prompt = (
        "Asagida bir bayiye, kalan siparis hedefini kapatmak icin onerilen urun ve adetler, bu bayinin "
        "gecmis fatura/satis verisi, sell-through orani ve hem bayi hem pazar genelinde sezonsal/YoY baglam var. "
        "Saha satis temsilcisinin bayi ziyaretinde kullanacagi, Turkce, kisa ve somut veriye dayali bir "
        "siparis onerisi metni yaz. Uydurma veri veya rakam ekleme, sadece verilen bilgileri kullan; "
        "sezonsal/pazar geneli baglami sadece belirginse (notrden acikca farkliysa) vurgula. "
        f"{shortfall_note} "
        "Cikti Markdown formatinda olsun ve su bolumleri icersin: "
        "'## <Bayi Adi> - Siparis Onerisi', '### Urun bazinda oneri ve gerekce' (madde madde).\n\n"
        f"Veriler:\n{facts}"
    )
    response = client.messages.create(model=model, max_tokens=900, messages=[{"role": "user", "content": prompt}])
    return "".join(block.text for block in response.content if getattr(block, "type", None) == "text").strip()


def generate_order_insight(
    dealer_name: str,
    remaining: float,
    shortfall: float,
    recommendations: pd.Series,
    transactions: pd.DataFrame,
    market: MarketContext,
    model: str = DEFAULT_MODEL,
) -> DealerOrderInsight:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if api_key:
        try:
            narrative = _llm_narrative(dealer_name, remaining, shortfall, recommendations, transactions, market, model)
            if narrative:
                return DealerOrderInsight(dealer_name=dealer_name, narrative=narrative, source="llm")
        except Exception:
            pass  # API erisimi basarisizsa sablon fallback'e dusulur
    narrative = _template_narrative(dealer_name, remaining, shortfall, recommendations, transactions, market)
    return DealerOrderInsight(dealer_name=dealer_name, narrative=narrative, source="template")
