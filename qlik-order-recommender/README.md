# Qlik Sipariş / Hedef Öneri Sistemi

Qlik'ten dışa aktarılan bayi satış/fatura/kampanya verisini kullanarak,
şirketten gelen bayi ve ürün bazlı hedefleri en isabetli şekilde dağıtan ve
satıcının bayiyi ikna ederken kullanabileceği **veriye dayalı argümanları**
üreten yapay zeka destekli bir araç.

İki ayrı çalışma modu var:

1. **Sipariş dağıtım motoru** (`order_pipeline`) — gerçek iş sürecine karşılık
   gelen, gerçek verilerle doğrulanmış asıl araç. Hedefler (bayi bazlı
   dönemsel hedef/fiili ve ürün bazlı şirket kotası) dışarıdan geliyor;
   bu araç "hangi bayiye hangi üründen kaç adet önerilmeli" sorusunu, o
   bayinin geçmiş satın alma eğilimine göre, şirket kotasını aşmadan
   cevaplar.
2. **Hedef tahmin motoru** (`pipeline`) — dışarıdan verilmiş bir hedef
   olmadığı, hedefin geçmiş satış trendinden tahmin edilmesi gerektiği
   senaryolar için (bkz. [Hedef tahmin motoru](#hedef-tahmin-motoru-pipeline)).

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

## Sipariş dağıtım motoru (`order_pipeline`)

### Neden bu yapı?

- **Fire vermeden dağıtım**: Şirketin ürün başına verdiği kota kesin sınır
  olarak korunur (bir ürün için tüm bayilere önerilen toplam adet, o ürünün
  kotasını asla aşmaz); bir bayinin favori ürününde kota yetersiz kalırsa,
  kalan ihtiyacı otomatik olarak diğer uygun ürünlere yeniden dağıtılır.
- **İkna edilebilirlik**: Her ürün-adet önerisi, o bayinin bu üründeki
  gerçek geçmiş fatura verisiyle (kaç ay, kaç adet, en son ne zaman alım
  yapmış) gerekçelendirilir.
- **Yeni ürünler dışlanmaz**: Bir bayinin hiç almadığı ama pazarda talep
  gören bir ürün, küçük bir "pazar ortalaması" payıyla yine de önerilebilir.

### Girdi dosyaları

| Dosya | Varsayılan sütunlar (gerçek Qlik/Excel disa aktarımı) | İç şema |
|---|---|---|
| `transactions` (Qlik ham veri) | `Nihai Ünvan, LPH2, LPH3, UrunKodu, YılAy, Diva Bildirim, Diva Stok, Diva - İleri Teslim, Fatura, Montaj` | `dealer_name, lph2, lph3, product_code, month, bildirim, stok, ileri_teslim, fatura, montaj` |
| `dealer_targets` (bayi dönemsel hedefi) | `Ünvan, LPH-2, Hedef Q4, Fiili Q4` | `dealer_name, lph2, hedef, fiili` |
| `product_quota` (şirket ürün kotası) | `ÜRÜNLER, ADET` | `product_code, quota` |

Sütun adları farklıysa `--mapping` ile bir YAML dosyası verilebilir (bkz.
`config/column_mapping.example.yaml` deki gibi, tablo adı → `{ic_alan:
gercek_sutun_adi}`). Varsayılanlar gerçek Qlik/Diva dışa aktarımlarındaki
başlıklarla birebir eşleşecek şekilde ayarlıdır, çoğu durumda mapping
gerekmez.

Notlar:
- `transactions` içindeki geçersiz/eksik ay (`YılAy = "-"`) satırları
  otomatik atlanır ve konsola kaç satırın atlandığı yazılır.
  `Fatura`/`Stok` gibi alanlarda negatif değerler (iade/düzeltme) geçerlidir.
- `dealer_targets`'taki bayi adı Qlik geçmişiyle birebir eşleşmezse (ör.
  kısaltılmış ünvan), otomatik bir ön-ek eşleştirmesi denenir; hâlâ
  eşleşmezse konsola uyarı yazılır ve o bayi sadece pazar ortalamasına göre
  değerlendirilir.
- Bir bayinin `Kalan` (Hedef − Fiili) değeri 0 veya negatifse (hedefi zaten
  karşılamış/aşmış), o bayi için ürün önerisi üretilmez.

### Çalıştırma

```bash
python -m qlik_order_recommender.order_pipeline \
  --transactions qlik_export.xlsx \
  --dealer-targets dealer_targets.xlsx \
  --product-quota product_quota.xlsx \
  --out reports/
```

Çıktılar:
- `reports/siparis_onerileri.xlsx` — 3 sayfa:
  - **Bayi Özet**: her bayi için hedef/fiili/kalan/önerilen toplam/karşılanamayan.
  - **Bayi x Ürün Tahsis**: bayi × ürün × önerilen adet (uzun format, tam detay).
  - **Siparis Dosyasi**: girdi kota dosyanızın, `SİPARİŞ` sütunu her ürün için
    tahsis edilen toplam adetle, kalan kota ve "kota aşıldı mı" bilgisiyle
    doldurulmuş hali.
- `reports/bayi_raporlari/<bayi adı>.md` — her bayi için saha ziyaretinde
  kullanılacak, geçmiş veriye dayalı ürün/adet gerekçe metni (okunabilirlik
  için en yüksek adetli ilk 15 ürün listelenir, tam liste Excel'de).

### Masaüstü arayüzü (GUI) ve meslektaşlarla paylaşma

Komut satırı yerine, dosya seçip "Çalıştır" demeye dayanan basit bir pencere
için:

```bash
python -m qlik_order_recommender.gui
```

Bu, Python kurulu bir makinede doğrudan çalışır. Python/Git kurulumu olmayan
meslektaşlarınızla paylaşmak için tek bir `.exe` dosyası olarak
paketleyebilirsiniz — **bu derleme adımı Windows üzerinde yapılmalı**
(PyInstaller çapraz derleme yapmaz), yani kendi Windows bilgisayarınızda,
kurulu olan `.venv` içinde çalıştırın:

```bash
pip install -r requirements-build.txt
pyinstaller --onefile --windowed --name qlik-siparis-onerisi run_gui.py
```

(`run_gui.py` proje kök dizininde, `gui.py`'yi paket bağlamında çalıştıran
küçük bir giriş noktası — PyInstaller'ı doğrudan `gui.py`'ye yöneltmeyin,
o zaman `.order_pipeline` gibi göreli import'lar "attempted relative
import with no known parent package" hatasıyla çöker.)

Üretilen `dist/qlik-siparis-onerisi.exe` dosyası (~150-250 MB, pandas/
scikit-learn içerdiği için) tek başına çalışır — Python kurulumu gerektirmez;
e-posta, Teams veya paylaşımlı bir sürücü ile dağıtabilirsiniz.

**Bilinen risk:** Kurumsal güvenlik yazılımları (ör. Bosch AppLocker),
imzasız/tanınmayan `.exe` dosyalarını "potansiyel zararlı yazılım" diye
engelleyebilir — bu README'yi hazırlarken tam olarak bu sorunla karşılaşıldı
(bkz. Python'un kendisi Downloads klasöründen engellenmişti). Böyle bir
durumda IT ile görüşüp dosyayı güvenilir listeye ekletmeniz gerekebilir.

## Hedef tahmin motoru (`pipeline`)

Dışarıdan verilmiş bir bayi/ürün hedefi olmadığı, hedefin bizzat geçmiş
satış trendinden (büyüme, mevsimsellik, ödeme sağlığı, geçmiş hedef isabet
oranı) tahmin edilmesi gerektiği senaryolar için. Girdi olarak parasal
`sales_amount` bekler (yukarıdaki adet-bazlı Diva verisinden farklı bir
şema).

### Girdi dosyaları

| Dosya | Alanlar |
|---|---|
| `dealers` | `dealer_id, dealer_name, region, segment` |
| `sales_history` | `dealer_id, date, product_category, sales_amount, quantity` |
| `invoices` | `dealer_id, invoice_date, invoice_amount, due_date, paid_date, payment_status` |
| `targets_history` | `dealer_id, month, target_amount, achieved_amount` |

`data/sample/` altında bu şema için sentetik örnek veriler bulunur.

### Çalıştırma

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

Çıktılar: `reports/hedef_onerileri_2026-10.xlsx` (önerilen hedef, min/max
aralık, güven, risk bayrakları) ve bayi başına Markdown ikna metni.

## Testler

```bash
pytest tests/
```

## Modül yapısı

```
src/qlik_order_recommender/
  # Siparis dagitim motoru (gercek is sureci)
  order_ingestion.py    # transactions / dealer_targets / product_quota okuma + bayi adi eslendirme
  allocation.py           # yinelemeli oransal dagitim: bayi kalanini urun kotasini asmadan dagitir
  order_insight.py          # Claude API (veya sablon) ile urun-adet gerekcesi
  order_report.py             # Excel (bayi ozet + detay + doldurulmus kota dosyasi) + Markdown
  order_pipeline.py             # uctan uca orkestrasyon + CLI
  gui.py                          # tkinter masaustu arayuzu (.exe paketleme icin)

  # Hedef tahmin motoru (harici hedef olmadigi senaryo icin)
  ingestion.py           # CSV/Excel okuma + mapping ile normalize etme
  features.py              # trend, mevsimsellik, odeme sagligi, hedef isabet orani
  target_model.py            # ML + heuristik fallback ile hedef araligi onerisi
  insight_generator.py         # Claude API (veya sablon) ile ikna metni
  report.py                      # Excel + Markdown rapor uretimi
  pipeline.py                      # uctan uca orkestrasyon + CLI
```
