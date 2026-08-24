# Bayi DİVA Kampanya / Hakediş Kontrol Sistemi

BSH bayilerinin DİVA üzerinden yaptığı kampanya kodlu satışları, E2E portaldan alınan
nakliye/montaj sonuçlarını ve BSH'nin bayiye kestiği hakediş faturası detayını tek dosyada
eşleştirip hakedişin doğru hesaplanıp doğru ödenip ödenmediğini otomatik kontrol eden Excel
aracı.

## Dosyalar

- `Bayi_DIVA_Kampanya_Hakedis_Kontrol.xlsx` — asıl çalışma dosyası (bayiye/yöneticiye dağıtılacak).
- `build_workbook.py` — dosyayı sıfırdan üreten Python (openpyxl) betiği. Kampanya listesi,
  sütun yapısı veya formül mantığı değiştiğinde bu betik güncellenip yeniden çalıştırılır.

## Sekmeler ve sahiplik

| Sekme | Kim doldurur | Koruma |
|---|---|---|
| Talimatlar | — (salt okunur rehber) | — |
| Parametreler | Yönetici (bayi kodu, stok mutabakatı durumu, %75 eşiği) | Yapı korumalı |
| Kampanyalar | **Yönetici** — aylık kampanya duyuru PDF'i işlenip satır olarak eklenir | **Şifre korumalı (admin)** |
| DİVA Satışları | **Bayi** — DİVA'daki "Kampanya Kodlu Satış" raporunu (35 sütun) A3'ten yapıştırır | Yapı korumalı, veri alanı açık |
| E2E Nakliye Montaj | **Bayi** — E2E portal raporunu (38 sütun) A3'ten yapıştırır | Yapı korumalı, veri alanı açık |
| Hakedis Fatura Detay | **Bayi** — BSH hakediş faturası detayını yapıştırır (yapı geçici, örnek dosya ile netleşecek) | Yapı korumalı, veri alanı açık |
| Kontrol | Otomatik | Tamamen formül, korumalı |
| Yil Sonu Kontrolu | Otomatik (Kural A / stok mutabakatlı bayiler için %75 nakliye-montaj oranı) | Tamamen formül, korumalı |
| Ozet | Otomatik dashboard | Tamamen formül, korumalı |

Sekme koruma şifreleri dosyada saklanmaz; ilgili konuşmada ayrıca paylaşılmıştır.

## İş kuralı özeti

- **Kural A (Stok Mutabakatlı Bayi):** Hakediş DİVA'da hemen teslim faturalandığı anda ödenir.
  30.11 itibarıyla yıl içinde hakediş ödenen ürünlerin nakliye/montaj tamamlanma oranı en az
  %75 olmalı; altında kalırsa fark mahsuplaşılır (`Yil Sonu Kontrolu` sekmesi).
- **Kural B (Stok Mutabakatı Yapılmamış Bayi):** Hakediş, nakliye tamamlandıysa nakliyeyi
  takip eden ayda; nakliye verisi yoksa montajı takip eden ayda ödenir. Bu tamamlanma,
  kampanyanın "Hakediş Son Tarihi"nden önce gerçekleşmelidir; aksi halde satır risk olarak
  işaretlenir.
- Eşleştirme anahtarı tüm sekmelerde **DİVA Fatura No (26 ile başlayan) + Ürün Kodu**'dur.

## Bilinen sınırlamalar / açık konular

- `Hakedis Fatura Detay` sekmesinin sütun yapısı geçicidir; gerçek Temmuz 2026 hakediş detay
  dosyası incelendiğinde kesinleştirilecektir.
- E2E raporundaki tarih alanları metin olarak farklı biçimlerde geldiğinden, montajın gerçek
  tamamlanma tarihi yerine "Montaj İstenilen Tarih" alanı vekil (proxy) olarak kullanılmıştır.
- Bundle kampanyalarında bayiye ödenecek sabit hakediş tutarları `Kampanyalar` sekmesinde bazı
  satırlarda "TEYİT EDİLMELİ" notuyla işaretlenmiştir; yönetici tarafından kampanya yazısına
  göre kesinleştirilmelidir.
