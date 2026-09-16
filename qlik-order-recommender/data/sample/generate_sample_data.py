"""Generates synthetic sample CSVs (Turkish Qlik-export-style headers) for
end-to-end testing of the pipeline without access to real Qlik data.

Run: python data/sample/generate_sample_data.py
"""
from __future__ import annotations

import csv
import random
from datetime import date, timedelta

random.seed(42)

OUT_DIR = __file__.rsplit("/", 1)[0]

DEALERS = [
    {"dealer_id": "D001", "dealer_name": "Anadolu Ticaret", "region": "Ankara", "segment": "Buyuk", "base": 180000, "growth": 0.018, "volatility": 0.08},
    {"dealer_id": "D002", "dealer_name": "Ege Dis Ticaret", "region": "Izmir", "segment": "Orta", "base": 95000, "growth": 0.010, "volatility": 0.12},
    {"dealer_id": "D003", "dealer_name": "Marmara Bayi", "region": "Istanbul", "segment": "Buyuk", "base": 210000, "growth": -0.004, "volatility": 0.10},
    {"dealer_id": "D004", "dealer_name": "Akdeniz Toptan", "region": "Antalya", "segment": "Kucuk", "base": 45000, "growth": 0.025, "volatility": 0.18},
    {"dealer_id": "D005", "dealer_name": "Bursa Endustri", "region": "Bursa", "segment": "Orta", "base": 110000, "growth": 0.006, "volatility": 0.09},
]

CATEGORIES = ["Elektronik", "Beyaz Esya", "Aksesuar"]
CATEGORY_SHARE = {"Elektronik": 0.5, "Beyaz Esya": 0.35, "Aksesuar": 0.15}

# 32 months of history: 2024-01 .. 2026-08
MONTHS = []
y, m = 2024, 1
for _ in range(32):
    MONTHS.append(date(y, m, 1))
    m += 1
    if m > 12:
        m = 1
        y += 1

SEASONALITY = {1: 0.85, 2: 0.88, 3: 0.95, 4: 1.0, 5: 1.05, 6: 1.1, 7: 0.9, 8: 0.8, 9: 1.0, 10: 1.1, 11: 1.2, 12: 1.35}


def month_sales(dealer: dict, month_idx: int, month: date) -> float:
    trend = dealer["base"] * ((1 + dealer["growth"]) ** month_idx)
    seasonal = SEASONALITY[month.month]
    noise = 1 + random.uniform(-dealer["volatility"], dealer["volatility"])
    return round(trend * seasonal * noise, 2)


def write_dealers():
    path = f"{OUT_DIR}/dealers.csv"
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Bayi Kodu", "Bayi Adi", "Bolge", "Segment"])
        for d in DEALERS:
            w.writerow([d["dealer_id"], d["dealer_name"], d["region"], d["segment"]])


def write_sales_and_targets():
    sales_path = f"{OUT_DIR}/sales_history.csv"
    targets_path = f"{OUT_DIR}/targets_history.csv"
    with open(sales_path, "w", newline="", encoding="utf-8") as sf, \
            open(targets_path, "w", newline="", encoding="utf-8") as tf:
        sw = csv.writer(sf)
        tw = csv.writer(tf)
        sw.writerow(["Bayi Kodu", "Tarih", "Urun Kategorisi", "Satis Tutari", "Miktar"])
        tw.writerow(["Bayi Kodu", "Ay", "Hedef Tutari", "Gerceklesen Tutar"])

        for d in DEALERS:
            for idx, month in enumerate(MONTHS):
                total = month_sales(d, idx, month)
                for cat in CATEGORIES:
                    amount = round(total * CATEGORY_SHARE[cat], 2)
                    qty = max(1, round(amount / random.uniform(800, 1400)))
                    sw.writerow([d["dealer_id"], month.isoformat(), cat, amount, qty])

                # Target set ~2-6% above prior trailing average, achieved = actual total
                target = round(total * random.uniform(0.95, 1.05), 2)
                tw.writerow([d["dealer_id"], month.strftime("%Y-%m"), target, total])


def write_invoices():
    path = f"{OUT_DIR}/invoices.csv"
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Bayi Kodu", "Fatura Tarihi", "Fatura Tutari", "Vade Tarihi", "Odeme Tarihi", "Odeme Durumu"])
        for d in DEALERS:
            for idx, month in enumerate(MONTHS):
                total = month_sales(d, idx, month)
                n_invoices = random.randint(2, 5)
                remaining = total
                for i in range(n_invoices):
                    amount = round(remaining / (n_invoices - i), 2) if i < n_invoices - 1 else round(remaining, 2)
                    remaining -= amount
                    invoice_date = month + timedelta(days=random.randint(0, 27))
                    due_date = invoice_date + timedelta(days=30)
                    # Dealers with lower volatility pay on time more often; small chance of overdue
                    late_chance = 0.08 + (0.15 if d["segment"] == "Kucuk" else 0.0)
                    if random.random() < late_chance:
                        paid_date = due_date + timedelta(days=random.randint(5, 45))
                        status = "Gecikti"
                    elif month >= MONTHS[-2]:
                        # last couple months: some invoices still open
                        if random.random() < 0.3:
                            paid_date = ""
                            status = "Beklemede"
                        else:
                            paid_date = (due_date - timedelta(days=random.randint(0, 10))).isoformat()
                            status = "Odendi"
                    else:
                        paid_date = (due_date - timedelta(days=random.randint(0, 10))).isoformat()
                        status = "Odendi"
                    w.writerow([d["dealer_id"], invoice_date.isoformat(), amount, due_date.isoformat(), paid_date, status])


if __name__ == "__main__":
    write_dealers()
    write_sales_and_targets()
    write_invoices()
    print("Sample data generated in", OUT_DIR)
