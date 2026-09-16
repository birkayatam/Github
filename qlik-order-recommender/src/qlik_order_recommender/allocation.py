"""Bayi kalan hedefini, urun bazli sirket kotasini asmadan, bayinin gecmis
satin alma egilimine gore urunlere dagitan tahsis (allocation) motoru.

Yontem: yinelemeli oransal dagitim (iterative proportional fitting).
Her turda, hala ihtiyaci olan bayilerin kalan hedefi, hala kotasi kalan
urunlere agirliklarina (affinity) oranla dagitilir; bir urunun toplam
talebi kotasini asarsa o urundeki tum paylar orantili olarak kisilir ve
acikta kalan ihtiyac bir sonraki turda diger urunlere yeniden dagitilir.
Bu, hem "bayinin kalanini olabildigince kapat" hem de "urun kotasini asma"
kisitlarini birlikte, aciklanabilir bir sekilde saglar."""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from .order_features import SELL_THROUGH_DAMPENING_FACTOR, SELL_THROUGH_LOW
from .order_features import demand_signal as _demand_signal

MARKET_SMOOTHING = 0.15
MAX_ITERATIONS = 25
CONVERGENCE_EPS = 1e-6


def compute_affinity(
    transactions: pd.DataFrame,
    dealer_names: list[str],
    product_codes: list[str],
    product_seasonal_index: pd.Series | None = None,
    sell_through: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """dealer_name x product_code agirlik matrisi. Temel sinyal
    `order_features.demand_signal` -- yani gercek talep (Bildirim/Ileri
    Teslim/Montaj = sell-out) agirlikli, Fatura (sell-in) dusuk agirlikli.
    Bir bayinin hic gecmisi olmadigi urunler icin, o urunun pazar
    ortalamasindan kucuk bir pay (smoothing) eklenir -- boylece yeni/az
    satilmis urunler de tamamen dislanmaz.

    `product_seasonal_index` verilirse (order_features.compute_seasonality
    cikisi), hedef donemde tipik olarak guclu/zayif olan urunler buna gore
    agirliklandirilir. `sell_through` verilirse (order_features.compute_
    sell_through cikisi), belirgin sekilde stokta bekleyen (dusuk sell-
    through) bayi-urun ciftlerinin agirligi hafifletilir -- zaten satilamamis
    stoku olan bir urunu tekrar onermemek icin."""
    df = transactions[
        transactions["product_code"].isin(product_codes) & transactions["dealer_name"].isin(dealer_names)
    ].copy()
    df["signal"] = _demand_signal(df)

    pivot = df.pivot_table(index="dealer_name", columns="product_code", values="signal", aggfunc="sum", fill_value=0.0)
    pivot = pivot.reindex(index=dealer_names, columns=product_codes, fill_value=0.0)

    market_avg = pivot.mean(axis=0)
    weight = pivot.add(market_avg * MARKET_SMOOTHING, axis=1)

    if product_seasonal_index is not None:
        factor = product_seasonal_index.reindex(product_codes).fillna(1.0)
        weight = weight.mul(factor, axis=1)

    if sell_through is not None and "sell_through_ratio" in sell_through.columns:
        ratio = sell_through["sell_through_ratio"]
        # Varsayilan notr (1.0): fatura gecmisi yetersizse (NaN) oran zaten
        # yargilanamaz, ceza uygulanmaz. Sadece bilinen ve esigin altinda
        # kalan ciftler hafifletilir.
        damp = pd.Series(1.0, index=ratio.index)
        damp[ratio < SELL_THROUGH_LOW] = SELL_THROUGH_DAMPENING_FACTOR
        damp_matrix = damp.unstack("product_code").reindex(index=dealer_names, columns=product_codes, fill_value=1.0)
        weight = weight.mul(damp_matrix)

    return weight


@dataclass
class AllocationResult:
    allocations: pd.DataFrame  # dealer_name x product_code, tamsayi adet
    dealer_shortfall: pd.Series  # dealer_name -> kalan - tahsis edilen (0 = tam karsilandi)
    product_leftover: pd.Series  # product_code -> kota - tahsis edilen (0 = kota tukendi)
    product_oversubscribed: pd.Series  # product_code -> bool (ilk turdaki toplam talep kotayi asti mi)


def allocate(remaining: pd.Series, quota: pd.Series, affinity: pd.DataFrame) -> AllocationResult:
    dealers = list(remaining.index)
    products = list(quota.index)

    need = remaining.astype(float).copy()
    capacity = quota.astype(float).copy()
    alloc = pd.DataFrame(0.0, index=dealers, columns=products)
    weight = affinity.reindex(index=dealers, columns=products, fill_value=0.0)

    initial_total_demand = pd.Series(0.0, index=products)

    for iteration in range(MAX_ITERATIONS):
        active_products = capacity[capacity > CONVERGENCE_EPS].index
        active_dealers = need[need > CONVERGENCE_EPS].index
        if len(active_products) == 0 or len(active_dealers) == 0:
            break

        w = weight.loc[active_dealers, active_products]
        row_sums = w.sum(axis=1)
        distributable = row_sums[row_sums > CONVERGENCE_EPS].index
        if len(distributable) == 0:
            break

        desired = w.loc[distributable].div(row_sums.loc[distributable], axis=0).mul(need.loc[distributable], axis=0)

        if iteration == 0:
            initial_total_demand = initial_total_demand.add(desired.sum(axis=0), fill_value=0.0)

        demand_per_product = desired.sum(axis=0)
        scale = (capacity.loc[active_products] / demand_per_product.replace(0, np.nan)).clip(upper=1.0).fillna(1.0)
        actual = desired.mul(scale, axis=1)

        alloc.loc[distributable, active_products] = alloc.loc[distributable, active_products].add(actual, fill_value=0.0)
        need.loc[distributable] = (need.loc[distributable] - actual.sum(axis=1)).clip(lower=0)
        capacity.loc[active_products] = (capacity.loc[active_products] - actual.sum(axis=0)).clip(lower=0)

        if actual.to_numpy().sum() < CONVERGENCE_EPS:
            break

    alloc_int = _round_preserving_quota(alloc, quota)

    dealer_shortfall = (remaining - alloc_int.sum(axis=1)).clip(lower=0)
    product_leftover = (quota - alloc_int.sum(axis=0)).clip(lower=0)
    product_oversubscribed = (initial_total_demand > quota.reindex(initial_total_demand.index).fillna(0)).reindex(
        products, fill_value=False
    )

    return AllocationResult(
        allocations=alloc_int,
        dealer_shortfall=dealer_shortfall,
        product_leftover=product_leftover,
        product_oversubscribed=product_oversubscribed,
    )


def _round_preserving_quota(alloc: pd.DataFrame, quota: pd.Series) -> pd.DataFrame:
    """Sonuclari tamsayiya yuvarlar; yuvarlama nedeniyle bir urunde kota
    asilirsa, en yuksek paya sahip bayilerden sirayla birer adet dusurerek
    kotanin kesinlikle asilmamasini saglar."""
    alloc_int = alloc.round().clip(lower=0)
    for product in alloc_int.columns:
        cap = int(quota.get(product, 0))
        col = alloc_int[product]
        over = int(col.sum() - cap)
        if over <= 0:
            continue
        order = col[col > 0].sort_values(ascending=False).index.tolist()
        i = 0
        guard = 0
        while over > 0 and order and guard < len(order) * 10_000 + 10_000:
            dealer = order[i % len(order)]
            if alloc_int.loc[dealer, product] > 0:
                alloc_int.loc[dealer, product] -= 1
                over -= 1
            i += 1
            guard += 1
    return alloc_int.astype(int)
