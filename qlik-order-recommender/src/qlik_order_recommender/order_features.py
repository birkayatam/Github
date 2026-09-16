"""Gercek talep sinyali, sezonsallik/YoY ve sell-through hesaplari.

Sutun anlamlari (kullanicidan netlestirildi):
 - Fatura        = bizim bayiye sevkiyatimiz (sell-in)
 - Diva Bildirim + Diva Ileri Teslim = bayinin son musteriye SATTIGI adet (sell-out)
 - Montaj        = kurulum sayisi -- en guclu "gercekten tuketildi" sinyali

Bu modul iki sey uretir:
 1. Agirlikli "gercek talep" sinyali (sell-out agirlikli, sell-in dusuk agirlikli)
    ve bundan turetilen sell-through orani (stokta bekleyen urunleri isaretlemek icin).
 2. Sirketin mali ceyrek tanimina (Q1=Aralik-Ocak-Subat, ..., Q4=Eylul-Ekim-Kasim)
    gore, hedef donemin gecmis yillardaki mevsimsellik endeksi ve YoY buyumesi --
    hem bayi bazinda hem pazar (tum bayiler) genelinde."""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

FISCAL_QUARTER_MONTHS: dict[int, list[int]] = {
    1: [12, 1, 2],
    2: [3, 4, 5],
    3: [6, 7, 8],
    4: [9, 10, 11],
}

# Gercek talep (sell-out) agirlikli sinyal: Bildirim/Ileri Teslim/Montaj = bayinin
# gercekten sattigi/kurdugu; Fatura = bizim sevkiyatimiz (sell-in), dusuk agirlikla
# -- hic sell-out verisi olmayan ama duzenli sipariş edilen bir urunu tamamen
# dislamamak icin kucuk bir taban olarak tutulur.
BILDIRIM_WEIGHT = 1.0
ILERI_TESLIM_WEIGHT = 1.0
MONTAJ_WEIGHT = 1.5
FATURA_WEIGHT = 0.3

SELL_THROUGH_TRAILING_MONTHS = 12
SELL_THROUGH_MIN_FATURA = 1.0  # bunun altinda fatura toplami varsa oran anlamsiz (NaN birakilir)
SELL_THROUGH_LOW = 0.3  # bu esigin altinda: muhtemelen hala stokta bekliyor
SELL_THROUGH_HIGH = 1.2  # bu esigin ustunde: hizli satiyor, restok mantikli
SELL_THROUGH_DAMPENING_FACTOR = 0.4  # dusuk sell-through'lu urunlerin tahsis agirligi bu katsayiyla azaltilir

# Cok az hacimli (dusuk toplam) gruplarda tek bir yuksek/dusuk ceyrek orantisiz
# buyuk bir endeks uretebilir (ör. 2 birimlik urun bir ceyrekte "hic yok" olunca
# endeks sifira, digerinde patlayabilir); bu yuzden hem minimum hacim sarti hem
# de mantikli bir ust/alt sinir uygulanir.
SEASONALITY_MIN_OVERALL_AVG = 1.0
SEASONALITY_INDEX_BOUNDS = (0.2, 5.0)


def demand_signal(transactions: pd.DataFrame) -> pd.Series:
    """Her satir icin agirlikli gercek talep (sell-out) sinyali."""
    return (
        BILDIRIM_WEIGHT * transactions["bildirim"].clip(lower=0)
        + ILERI_TESLIM_WEIGHT * transactions["ileri_teslim"].clip(lower=0)
        + MONTAJ_WEIGHT * transactions["montaj"].clip(lower=0)
        + FATURA_WEIGHT * transactions["fatura"].clip(lower=0)
    )


def quarter_periods(fiscal_year: int, quarter: int) -> list[str]:
    """Verilen mali yil + ceyrek icin "YYYY-MM" ay listesi. Q1'in Aralik
    ayi bir onceki takvim yilina aittir (ör. mali yil 2026 Q1 = 2025-12,
    2026-01, 2026-02)."""
    periods = []
    for m in FISCAL_QUARTER_MONTHS[quarter]:
        calendar_year = fiscal_year - 1 if (quarter == 1 and m == 12) else fiscal_year
        periods.append(f"{calendar_year:04d}-{m:02d}")
    return periods


def detect_target_quarter(transactions: pd.DataFrame) -> tuple[int, int]:
    """Verideki en guncel ayin ait oldugu mali ceyregi (ve mali yili) tespit
    eder -- bu, planlanan/devam eden donem olarak kullanilir."""
    latest = max(transactions["month"])
    year, month = int(latest[:4]), int(latest[5:7])
    quarter = next(q for q, months in FISCAL_QUARTER_MONTHS.items() if month in months)
    fiscal_year = year + 1 if (quarter == 1 and month == 12) else year
    return fiscal_year, quarter


def _complete_quarters(available_months: set[str]) -> list[tuple[int, int]]:
    """Verideki butun 3 ayi da bulunan (fiscal_year, quarter) ciftlerini bulur."""
    years = sorted({int(m[:4]) for m in available_months})
    if not years:
        return []
    complete = []
    for y in range(years[0], years[-1] + 2):
        for q in (1, 2, 3, 4):
            if all(m in available_months for m in quarter_periods(y, q)):
                complete.append((y, q))
    return sorted(set(complete))


def _quarter_group_totals(
    transactions: pd.DataFrame, group_col: str, signal: pd.Series, fiscal_year: int, quarter: int
) -> pd.Series:
    months = set(quarter_periods(fiscal_year, quarter))
    mask = transactions["month"].isin(months)
    return signal[mask].groupby(transactions.loc[mask, group_col]).sum()


@dataclass
class SeasonalityResult:
    seasonal_index: pd.Series  # group -> hedef ceyrek / tipik ceyrek orani (1.0 = notr)
    yoy_growth: pd.Series  # group -> hedef ceyregin son iki tamamlanmis tekrari arasi buyume (NaN = veri yok)
    history: dict[tuple[int, int], pd.Series] = field(default_factory=dict)  # (yil,ceyrek) -> group totalleri


def compute_seasonality(
    transactions: pd.DataFrame, group_col: str, target_fiscal_year: int, target_quarter: int
) -> SeasonalityResult:
    """group_col ('dealer_name' veya 'product_code') bazinda, hedef mali
    ceyregin gecmiste (tamamlanmis ceyrekler) tipik bir ceyreğe gore ne
    kadar guclu/zayif oldugunu (seasonal_index) ve son iki tekrari
    arasindaki buyumeyi (yoy_growth) hesaplar. Hedef donemin kendisi asla
    baseline'a dahil edilmez."""
    signal = demand_signal(transactions)
    groups = pd.Index(sorted(transactions[group_col].unique()))

    available_months = set(transactions["month"].unique())
    complete_quarters = [
        (y, q) for (y, q) in _complete_quarters(available_months) if (y, q) != (target_fiscal_year, target_quarter)
    ]

    if not complete_quarters:
        return SeasonalityResult(
            seasonal_index=pd.Series(1.0, index=groups),
            yoy_growth=pd.Series(float("nan"), index=groups),
        )

    history = {
        (y, q): _quarter_group_totals(transactions, group_col, signal, y, q) for (y, q) in complete_quarters
    }
    combined = pd.DataFrame(history).reindex(groups).fillna(0.0)
    overall_avg = combined.mean(axis=1)

    target_cols = [c for c in combined.columns if c[1] == target_quarter]
    if target_cols:
        target_avg = combined[target_cols].mean(axis=1)
        low_volume = overall_avg < SEASONALITY_MIN_OVERALL_AVG
        seasonal_index = (target_avg / overall_avg.replace(0, float("nan"))).fillna(1.0)
        seasonal_index = seasonal_index.clip(*SEASONALITY_INDEX_BOUNDS)
        seasonal_index[low_volume] = 1.0  # hacim cok dusukse guvenilir degil, notr birak
    else:
        seasonal_index = pd.Series(1.0, index=groups)

    yoy_growth = pd.Series(float("nan"), index=groups)
    target_years_sorted = sorted(y for (y, _q) in target_cols)
    if len(target_years_sorted) >= 2:
        recent_y, prior_y = target_years_sorted[-1], target_years_sorted[-2]
        recent = combined[(recent_y, target_quarter)]
        prior = combined[(prior_y, target_quarter)]
        yoy_growth = (recent - prior) / prior.replace(0, float("nan"))

    return SeasonalityResult(seasonal_index=seasonal_index, yoy_growth=yoy_growth, history=history)


def compute_sell_through(transactions: pd.DataFrame) -> pd.DataFrame:
    """dealer_name x product_code bazinda, son SELL_THROUGH_TRAILING_MONTHS
    ayin (Bildirim+Montaj) toplami / Fatura toplami oranini hesaplar.
    Dusuk oran = urun muhtemelen hala stokta bekliyor (fazla sipariş
    onerilmemeli); yuksek oran = hizli satiyor (restok mantikli)."""
    months_sorted = sorted(transactions["month"].unique())
    recent_months = set(months_sorted[-SELL_THROUGH_TRAILING_MONTHS:])
    df = transactions[transactions["month"].isin(recent_months)]

    grouped = df.groupby(["dealer_name", "product_code"]).agg(
        bildirim_total=("bildirim", lambda s: s.clip(lower=0).sum()),
        ileri_teslim_total=("ileri_teslim", lambda s: s.clip(lower=0).sum()),
        montaj_total=("montaj", lambda s: s.clip(lower=0).sum()),
        fatura_total=("fatura", lambda s: s.clip(lower=0).sum()),
    )
    has_meaningful_fatura = grouped["fatura_total"] >= SELL_THROUGH_MIN_FATURA
    grouped["sell_through_ratio"] = float("nan")
    grouped.loc[has_meaningful_fatura, "sell_through_ratio"] = (
        grouped.loc[has_meaningful_fatura, "bildirim_total"] + grouped.loc[has_meaningful_fatura, "montaj_total"]
    ) / grouped.loc[has_meaningful_fatura, "fatura_total"]
    return grouped
