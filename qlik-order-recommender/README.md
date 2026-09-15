# Qlik Sipariş / Hedef Öneri Sistemi

Qlik'ten dışa aktarılan satış ve fatura raporlarını kullanarak, her bayi için
ilgili ayda satıcıya verilecek **hedefi** öneren ve satıcının bayiyi ikna
ederken kullanabileceği **veriye dayalı argümanları** üreten yapay zeka
destekli bir araç.

## Neden bu yapı?

- **Fire vermeden hedefleme**: Tek bir hedef sayısı yerine bir güven aralığı
  (`min`/`max`) ve risk bayrakları üretilir. Amaç ne bayinin masada para
  bırakması (çok düşük hedef) ne de gerçekçi olmayan bir hedefle
  motivasyonun kırılması (çok yüksek hedef).
- **İkna edilebilirlik**: Hedefi destekleyen somut veri noktaları (büyüme
  trendi, mevsimsellik, ürün karması, ödeme/fatura sağlığı, geçmiş hedef
  isabet oranı) satıcıya doğal dilde sunulur.

## Kurulum

```bash
cd qlik-order-recommender
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
pip install -e .
```

LLM destekli ikna metni üretimi için (isteğe bağlı, yoksa şablon tabanlı
fallback kullanılır):

```bash
export ANTHROPIC_API_KEY=sk-...
```

## Gerçek Qlik verisini bağlama

Qlik raporlarından dışa aktardığınız CSV/Excel dosyalarının sütun adları
muhtemelen `config/column_mapping.example.yaml` içindeki iç şema alanlarından
farklı olacaktır (örn. "Bayi Adı", "Fatura Tutarı" gibi Türkçe başlıklar).
Bu dosyayı kopyalayıp (`config/column_mapping.yaml`) kendi rapor
sütunlarınızla eşleştirin — kod değişikliği gerekmez.

Gereken 4 girdi dosyası ve iç şeması:

| Dosya | Alanlar |
|---|---|
| `dealers` | `dealer_id, dealer_name, region, segment` |
| `sales_history` | `dealer_id, date, product_category, sales_amount, quantity` |
| `invoices` | `dealer_id, invoice_date, invoice_amount, due_date, paid_date, payment_status` |
| `targets_history` | `dealer_id, month, target_amount, achieved_amount` |

`data/sample/` altında sentetik örnek veriler bulunur, uçtan uca deneme için
kullanılabilir.

## Çalıştırma

```bash
python -m qlik_order_recommender \
  --dealers data/sample/dealers.csv \
  --sales data/sample/sales_history.csv \
  --invoices data/sample/invoices.csv \
  --targets data/sample/targets_history.csv \
  --mapping config/column_mapping.example.yaml \
  --month 2026-10 \
  --out reports/
```

Çıktılar:
- `reports/hedef_onerileri_2026-10.xlsx` — tüm bayiler için tek satırlık özet
  (önerilen hedef, aralık, güven, risk bayrakları, kısa gerekçe).
- `reports/<dealer_id>_2026-10.md` — her bayi için satıcının saha
  ziyaretinde kullanacağı detaylı ikna metni.

## Testler

```bash
pytest tests/
```

## Modül yapısı

```
src/qlik_order_recommender/
  ingestion.py         # CSV/Excel okuma + mapping ile normalize etme
  features.py           # trend, mevsimsellik, ödeme sağlığı, hedef isabet oranı
  target_model.py        # ML + heuristik fallback ile hedef aralığı önerisi
  insight_generator.py    # Claude API (veya şablon fallback) ile ikna metni
  report.py                # Excel + Markdown rapor üretimi
  pipeline.py               # uçtan uca orkestrasyon + CLI
```
