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
| Hakedis Fatura Detay | **Bayi** — BSH'nin aylık hakediş detay raporunu ("... DETAY" sekmesi, 32 sütun) A3'ten yapıştırır | Yapı korumalı, veri alanı açık |
| Kontrol | Otomatik | Tamamen formül, korumalı |
| Sorun Raporu | Otomatik — Kontrol'de "Uygun" olmayan (eşleşmeyen, riskli, tutar uyuşmazlığı olan) tüm satırların filtrelenmiş listesi | Tamamen formül, korumalı |
| Yil Sonu Kontrolu | Otomatik (Kural A / stok mutabakatlı bayiler için %75 nakliye-montaj oranı) | Tamamen formül, korumalı |
| Ozet | Otomatik dashboard | Tamamen formül, korumalı |

Sekme koruma şifreleri dosyada saklanmaz; ilgili konuşmada ayrıca paylaşılmıştır.

## İş kuralı özeti

- **Kural A (Stok Mutabakatlı Bayi / "Mutabakat Var"):** Hakediş DİVA'da hemen teslim
  faturalandığı anda ödenir. 30.11 itibarıyla yıl içinde hakediş ödenen ürünlerin nakliye/montaj
  tamamlanma oranı en az %75 olmalı; altında kalırsa fark mahsuplaşılır (`Yil Sonu Kontrolu`
  sekmesi).
- **Kural B (Stok Mutabakatı Yapılmamış Bayi / "Mutabakat Yoksa"):** Hakediş, nakliye
  tamamlandıysa nakliyeyi takip eden ayda; nakliye verisi yoksa montajı takip eden ayda ödenir.
  Bu tamamlanma, kampanyanın "Hakediş Son Tarihi"nden önce gerçekleşmelidir; aksi halde satır
  risk olarak işaretlenir.
- Hangi kuralın uygulandığının **asıl kaynağı**, `Hakedis Fatura Detay` sekmesindeki "Ödeme
  Yöntemi" ve "Yöntem 1" sütunlarıdır (BSH'nin fiilen hesapladığı yöntem); `Parametreler`
  sekmesindeki bayi bazlı işaret, hakediş dosyası gelmeden önceki bir **ön beklenti**dir. Kontrol
  sekmesi ikisini yan yana gösterir (J: ön beklenti, M: hakediş dosyasına göre gerçek).
- Bir ürün aynı anda birden fazla kampanyaya konu olabildiğinden (ör. tekil + bundle),
  eşleştirme **DİVA Fatura No + Ürün Kodu + Kampanya Adı** üçlüsüyle yapılır (yalnızca E2E
  nakliye/montaj eşleştirmesi, kampanyadan bağımsız olduğu için Fatura No + Ürün Kodu ikilisini
  kullanır).

## Bilinen sınırlamalar / açık konular

- E2E raporundaki tarih alanları metin olarak farklı biçimlerde geldiğinden, montajın gerçek
  tamamlanma tarihi yerine "Montaj İstenilen Tarih" alanı vekil (proxy) olarak kullanılmıştır.
  Hakedis Fatura Detay raporu geldikten sonra oradaki gerçek Montaj/Nakliye Tarihi esas alınır.
- Bundle kampanyalarında bayiye ödenecek sabit hakediş tutarları `Kampanyalar` sekmesinde bazı
  satırlarda "TEYİT EDİLMELİ" notuyla işaretlenmiştir; yönetici tarafından kampanya yazısına
  göre kesinleştirilmelidir.
- `Yil Sonu Kontrolu` sekmesindeki %75 oranı, mevcut haliyle Kural A bayileri için E2E raporundan
  hesaplanır; hakediş dosyası biriktikçe bu hesaplama `Hakedis Fatura Detay`'daki gerçek
  Nakliye/Montaj Tarihi alanlarına taşınabilir.
