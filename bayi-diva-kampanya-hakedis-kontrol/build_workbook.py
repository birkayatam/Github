import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

FONT = "Arial"
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FILL = PatternFill("solid", fgColor="2E75B6")
RAW_HEADER_FILL = PatternFill("solid", fgColor="808080")
CALC_HEADER_FILL = PatternFill("solid", fgColor="2E75B6")
INPUT_FILL = PatternFill("solid", fgColor="FFFDE7")
EXAMPLE_FILL = PatternFill("solid", fgColor="F2F2F2")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
ADMIN_FILL = PatternFill("solid", fgColor="FFF2CC")

HEADER_FONT = Font(name=FONT, size=9, bold=True, color="FFFFFF")
RAW_HEADER_FONT = Font(name=FONT, size=9, bold=True, color="FFFFFF")
LABEL_FONT = Font(name=FONT, size=10, bold=True, color="1F4E78")
INPUT_FONT = Font(name=FONT, size=10, color="0000FF")
FORMULA_FONT = Font(name=FONT, size=9, color="000000")
NOTE_FONT = Font(name=FONT, size=9, italic=True, color="7F7F7F")
EXAMPLE_FONT = Font(name=FONT, size=9, italic=True, color="808080")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CUR = '#,##0.00 "TL"'
DATE_FMT = 'dd.mm.yyyy'
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_AL = Alignment(horizontal="left", vertical="center", wrap_text=True)

wb = openpyxl.Workbook()
wb.remove(wb.active)


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def header_row(ws, row, headers, fill, font, height=34, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = font
        c.fill = fill
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[row].height = height


# ============================================================
# 1) TALIMATLAR
# ============================================================
ws = wb.create_sheet("Talimatlar")
ws.sheet_view.showGridLines = False
set_widths(ws, [3, 105])
ws["B2"] = "BSH BAYİ - DİVA SATIŞ / KAMPANYA / HAKEDİŞ KONTROL SİSTEMİ"
ws["B2"].font = Font(name=FONT, size=15, bold=True, color="1F4E78")

blocks = [
    ("AMAÇ",
     "Bu dosya; bayinin DİVA üzerinden yaptığı kampanya kodlu satışları, E2E portaldan alınan nakliye/montaj "
     "sonuçlarını ve BSH'nin bayiye kestiği hakediş faturası detayını tek noktada eşleştirerek, hakedişin doğru "
     "hesaplanıp doğru ödenip ödenmediğini otomatik olarak kontrol eder."),
    ("İKİ FARKLI HAKEDİŞ KURALI (kaynak: aylık 'Perakende Uygulama Yazısı')",
     "Kural A - Stok Mutabakatlı Bayi: Hakediş, DİVA'da hemen teslim faturalandırıldığı anda ödenir. Buna karşılık "
     "30.11 itibarıyla yıl içinde hakediş ödenen ürünlerin nakliye VEYA montaj tamamlanma ORANININ toplam hakedişe "
     "oranı en az %75 olmalıdır; altında kalırsa aradaki fark mahsuplaşılır. (Bkz. 'Yil Sonu Kontrolu' sekmesi)\n"
     "Kural B - Stok Mutabakatı Yapılmamış Bayi: Hakediş, DİVA hemen teslim faturası olsa dahi hemen ödenmez; "
     "nakliye verisi varsa nakliyenin tamamlanmasını takip eden ayda, nakliye verisi yoksa montajın tamamlanmasını "
     "takip eden ayda ödenir. Bu tamamlanma, kampanyanın 'Hakediş Son Tarihi'nden önce gerçekleşmiş olmalıdır; "
     "aksi halde hakediş RİSK olarak işaretlenir."),
    ("HANGİ BAYİ HANGİ KURALA TABİ?",
     "'Parametreler' sekmesinde tek bir alanla belirlenir: 'Stok Mutabakatı Yapıldı mı?'. Bu bilgiyi BSH saha/finans "
     "ekibinizden teyit edin. Bu alan tüm 'DİVA Satışları' sekmesindeki her satır için otomatik kural ataması yapar."),
    ("VERİ AKIŞI VE SEKME SAHİPLİĞİ",
     "'Kampanyalar' sekmesi YÖNETİCİ (BSH) tarafından doldurulur ve şifre ile korumalıdır; bayi bu sekmeyi görebilir "
     "ama değiştiremez. Aylık kampanya duyuru yazısı (PDF) BSH tarafından işlenip bu sekmeye satır olarak eklenir.\n"
     "'DİVA Satışları' sekmesine bayi, DİVA'dan alınan 'Kampanya Kodlu Satış' raporunu (Şirket Kodu...Kategori, "
     "35 sütun) olduğu gibi kopyalayıp A3 hücresinden başlayarak yapıştırır.\n"
     "'E2E Nakliye Montaj' sekmesine bayi, E2E portaldan alınan raporu (MerkezKodu...PlatformName, 38 sütun) A3 "
     "hücresinden başlayarak yapıştırır.\n"
     "'Hakedis Fatura Detay' sekmesine bayi, BSH'nin kestiği hakediş faturasının satır detayını yapıştırır "
     "(bu sekmenin sütun yapısı, ilk gerçek örnek dosya incelendikten sonra kesinleştirilecektir; şimdilik "
     "eşleştirme için gereken asgari alanlar tanımlanmıştır)."),
    ("EŞLEŞTİRME ANAHTARI",
     "Tüm kaynaklar 'DİVA Fatura No (26 ile başlayan) + Ürün Kodu' ikilisi ile eşleştirilir. 'Kontrol' sekmesi bu "
     "anahtarla otomatik birleştirme yapar ve nihai durumu üretir."),
    ("BİLİNEN VERİ SINIRLAMALARI",
     "E2E raporundaki tarih alanları (SiparisTarihi, TeslimatIstenilenTarih, MontajIstenilenilenTarih, "
     "IrsaliyeTarihi) DİVA/E2E tarafından METİN olarak farklı biçimlerde gelmektedir; 'E2E Nakliye Montaj' "
     "sekmesinde bu alanlar otomatik olarak gerçek tarihe çevrilir (sütun AN-AQ). İrsaliye tarihi, nakliyenin "
     "tamamlanma tarihi için; montaj istenilen tarihi ise gerçek tamamlanma tarihi bilgisi rapor dışı olduğundan "
     "montajın tamamlanma tarihi için en yakın vekil (proxy) olarak kullanılmıştır. Gerçek 'montaj tamamlanma "
     "tarihi' alanı E2E raporuna eklenirse formüller kolayca güncellenebilir.\n"
     "'Bayi Aracı' ile yapılan teslimatlarda nakliye süreci BSH tarafından izlenemediğinden (bkz. kampanya yazısı), "
     "bu satışlarda sistem otomatik olarak montaj verisini esas alır."),
    ("RENK KODLARI",
     "Sarı zemin + mavi yazı = elle doldurulacak/yapıştırılacak alan. Gri başlık = kaynak rapordan birebir gelen "
     "ham sütun. Mavi başlık = otomatik hesaplanan sütun, dokunmayın. Turuncu zemin = risk/uyuşmazlık uyarısı."),
]
r = 4
for title, body in blocks:
    c = ws.cell(row=r, column=2, value=title)
    c.font = Font(name=FONT, size=11, bold=True, color="1F4E78")
    r += 1
    c = ws.cell(row=r, column=2, value=body)
    c.font = Font(name=FONT, size=10)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 18 * (body.count("\n") + 1 + len(body) // 100)
    r += 2

# ============================================================
# 2) PARAMETRELER
# ============================================================
ws = wb.create_sheet("Parametreler")
ws.sheet_view.showGridLines = False
set_widths(ws, [3, 55, 22, 3])
ws["B2"] = "PARAMETRELER"
ws["B2"].font = Font(name=FONT, size=14, bold=True, color="1F4E78")

r = 4
ws.cell(row=r, column=2, value="Bayi Kodu").font = LABEL_FONT
c_bayikodu = ws.cell(row=r, column=3, value="7000129546")
c_bayikodu.fill = INPUT_FILL; c_bayikodu.font = INPUT_FONT; c_bayikodu.border = BORDER
r += 1
ws.cell(row=r, column=2, value="Bayi Adı").font = LABEL_FONT
c_bayiadi = ws.cell(row=r, column=3, value="MUSTAFA DEMİREL")
c_bayiadi.fill = INPUT_FILL; c_bayiadi.font = INPUT_FONT; c_bayiadi.border = BORDER
r += 2

ws.cell(row=r, column=2,
        value="Bu bayi için STOK MUTABAKATI yapıldı mı? (Evet = Kural A: hakediş DİVA çıkışında hemen ödenir. "
              "Hayır = Kural B: hakediş nakliye/montaj tamamlandıktan sonraki ay ödenir.)").font = LABEL_FONT
ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[r].height = 48
c_mutabakat = ws.cell(row=r, column=3, value="Evet")
c_mutabakat.fill = INPUT_FILL; c_mutabakat.font = INPUT_FONT; c_mutabakat.border = BORDER
c_mutabakat.alignment = Alignment(horizontal="center")
MUTABAKAT_ROW = r
r += 2

ws.cell(row=r, column=2,
        value="Yıl Sonu Nakliye/Montaj Oranı Kontrol Tarihi (Kural A bayileri için)").font = LABEL_FONT
c_yilsonu = ws.cell(row=r, column=3, value="30.11.2026")
c_yilsonu.fill = INPUT_FILL; c_yilsonu.font = INPUT_FONT; c_yilsonu.number_format = DATE_FMT
c_yilsonu.border = BORDER; c_yilsonu.alignment = Alignment(horizontal="center")
YILSONU_ROW = r
r += 1
ws.cell(row=r, column=2, value="Minimum Nakliye/Montaj Oranı Eşiği (Kural A)").font = LABEL_FONT
c_esik = ws.cell(row=r, column=3, value=0.75)
c_esik.fill = INPUT_FILL; c_esik.font = INPUT_FONT; c_esik.number_format = "0.0%"
c_esik.border = BORDER; c_esik.alignment = Alignment(horizontal="center")
ESIK_ROW = r
r += 2

note = ws.cell(row=r, column=2,
    value="Not: 'Stok Mutabakatı' durumunu ve oranları BSH saha/finans ekibinizden teyit edin. Değer "
          "değiştiğinde tüm 'DİVA Satışları' ve 'Kontrol' sekmeleri otomatik olarak yeniden hesaplanır.")
note.font = NOTE_FONT
note.alignment = Alignment(wrap_text=True)
ws.row_dimensions[r].height = 30

dv_yn = DataValidation(type="list", formula1='"Evet,Hayır"', allow_blank=False)
ws.add_data_validation(dv_yn)
dv_yn.add(c_mutabakat)

P_MUTABAKAT = f"Parametreler!$C${MUTABAKAT_ROW}"
P_YILSONU = f"Parametreler!$C${YILSONU_ROW}"
P_ESIK = f"Parametreler!$C${ESIK_ROW}"

print("Part 1 done")

# ============================================================
# 3) KAMPANYALAR (ADMIN ONLY)
# ============================================================
ws = wb.create_sheet("Kampanyalar")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"
kamp_headers = [
    "Kampanya Adı\n(DİVA 'Paket Adı' ile\nBİREBİR aynı olmalı)", "Dönem\n(AAYYYY)", "Kampanya Tipi",
    "Ürün Grubu / Kapsam", "Kampanya\nBaşlangıç", "Kampanya\nBitiş", "Hakediş\nSon Tarihi",
    "Bayiye Ödenecek\nSabit Hakediş Tutarı\n(TL) (boşsa DİVA\nindirimi esas alınır)",
    "Adet Limiti", "Diğer Kampanyalarla\nBirleşebilir mi", "Kaynak / Not",
]
set_widths(ws, [30, 10, 14, 24, 12, 12, 12, 16, 10, 16, 34])
header_row(ws, 1, kamp_headers, ADMIN_FILL, LABEL_FONT, height=48)
for c in range(1, len(kamp_headers) + 1):
    ws.cell(row=1, column=c).font = Font(name=FONT, size=9, bold=True, color="1F4E78")

kamp_rows = [
    ("Aralık Ankastre Set Alana 6000 TL İndirim-202512", "202512", "Bundle-Set", "Ankastre Set (Fırın+Ocak+Davlumbaz)", "04.12.2025", "31.12.2025", "31.05.2026", 4500, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Aralık 2025 Perakende Uygulama Yazısı. PDF notu: Tüm set ürünleri aynı Diva faturasında olmalı"),
    ("Aralık Ankastre Set Alana 8000 TL İndirim-202512", "202512", "Bundle-Set", "Ankastre Set (Fırın+Ocak+Davlumbaz)", "04.12.2025", "31.12.2025", "31.05.2026", 6200, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Aralık 2025 Perakende Uygulama Yazısı. PDF notu: Tüm set ürünleri aynı Diva faturasında olmalı"),
    ("Aralık Ayı Fırsat Ürünleri Kampanyası-202512", "202512", "Diğer", "Çeşitli (Buzdolabı/Derin Dondurucu/Bulaşık Mak./Çamaşır Mak./Ankastre Fırın - hareketi düşük/2 yıl önce satışa kapanan modeller)", "04.12.2025", "31.12.2025", "30.01.2026", None, None, "",
     "Kaynak: Aralık 2025 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; hakediş son tarihi diğer kampanyalardan farklı (30.01.2026); Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; hakediş son tarihi diğer kampanyalardan farklı (31.01.2026)"),
    ("Aralık Kurutma+Bulaşık Alımına 10000 TL İndirim-202512", "202512", "Bundle-Set", "Kurutma Makinesi + Seçili Beyaz Eşya (Buzdolabı/Çamaşır/Bulaşık/Derin Dondurucu)", "04.12.2025", "31.12.2025", "31.05.2026", 8000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Aralık 2025 Perakende Uygulama Yazısı. PDF notu: Kampanya birleşme kısıtı: sadece tekil ile birleşir"),
    ("Aralık Soğutucu + Bulaşık Alımına 5000 TL İndirim-202512", "202512", "Bundle-Set", "Beyaz Eşya 2'li Set (Buzdolabı/Çamaşır Mak./Bulaşık Mak./Derin Dondurucu)", "04.12.2025", "31.12.2025", "31.05.2026", 3900, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Aralık 2025 Perakende Uygulama Yazısı. PDF notu: Kampanya birleşme kısıtı: sadece tekil ile birleşir; Kampanya birleşme kısıtı: sadece tekil kampanyalar ile birleşir, diğer bundle kampanyalarla birleşmez"),
    ("Aralık Tekil Bulaşık Makinesi Kampanyası-202512", "202512", "Tekil", "Bulaşık Makinesi", "04.12.2025", "31.12.2025", "31.05.2026", None, None, "Evet",
     "Kaynak: Aralık 2025 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş"),
    ("Aralık Tekil Kurutma Makinesi Kampanyası-202512", "202512", "Tekil", "Kurutma Makinesi", "04.12.2025", "31.12.2025", "31.05.2026", None, None, "Evet",
     "Kaynak: Aralık 2025 Perakende Uygulama Yazısı. PDF notu: 30 gün koşulsuz iade uygulaması + fiyat farkı yöntemi Montaj/Nakliye Adedi, sabit TL yok"),
    ("Aralık Tekil Soğutucu Kampanyası-202512", "202512", "Tekil", "Soğutucu", "04.12.2025", "31.12.2025", "31.05.2026", None, None, "Evet",
     "Kaynak: Aralık 2025 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş; Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş; model listesi ile sınırlı"),
    ("Aralık Tekil Çamaşır Kampanyası-202512", "202512", "Tekil", "Çamaşır Makinesi", "04.12.2025", "31.12.2025", "31.05.2026", None, None, "Evet",
     "Kaynak: Aralık 2025 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş"),
    ("Aralık-Aile Bakanlığı Kampanyası – Bosch-202512", "202512", "Diğer", "Çeşitli (Buzdolabı/Derin Dondurucu/Çamaşır/Kurutma/Bulaşık/Fırın/Mikrodalga/Ocak/Davlumbaz/Küçük Ev Aletleri)", "04.12.2025", "31.12.2025", "31.05.2026", None, None, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: TRPF üzerinden ekstra %5 indirim (sabit TL değil); KEA modellerinde hakediş son tarihi 31.12.2025, diğer ürünlerde 31.05.2026; Tüketiciye TRPF üzerinden ekstra %5 indirim (sabit TL değil), Aile Bakanlığı evlilik kredisi indirim kodu ile; KEA (Küçük Ev Aletleri) modellerinde hakediş son tarihi 15.12."),
    ("Aralık-Ankastre Set Fırsat 3: Seçili Ankastre Set ile Birlikte Ankastre Bulaşık Makinesi Alan Tüketicilerimize 12.000 TL İndirim-202512", "202512", "Bundle-Set", "Ankastre Set (Fırın+Ocak+Davlumbaz) + Ankastre Bulaşık Makinesi", "04.12.2025", "31.12.2025", "31.05.2026", 12000, 300, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Tüm set ürünleri (fırın, ocak, davlumbaz, bulaşık makinesi) aynı Diva faturasında olmalı; Tüm set ürünleri aynı Diva faturasında olmalı"),
    ("Aralık-Bosch Premium Seçili Büyük Beyaz Eşya Seti - Seçili modelleden (Buzdolabı, Çamaşır Makinesi, Kurutma Makinesi, Bulaşık Makinesi ve Ankastre Fırın) Üçlü Beyaz Eşya Seti Alana BCS1041WAC Unlimited Süpürge Hediye-202512", "202512", "Küçük Ev Aletleri", "Büyük Beyaz Eşya 3'lü Set (Buzdolabı/Çamaşır/Kurutma/Bulaşık/Ankastre Fırın)", "04.12.2025", "31.12.2025", "31.05.2026", 25000, 100, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Hediye ürün BCS1041WAC alımı zorunlu; Hediye ürün BCS1041WAC alımı zorunlu, alınmazsa kampanya kapsamı dışı kalır; tüketiciye TL indirimi değil hediye ürün verilir"),
    ("Aralık-Derin Dondurucu Kampanyası-202512", "202512", "Bundle-Set", "Derin Dondurucu", "04.12.2025", "31.12.2025", "31.05.2026", None, None, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş"),
    ("Aralık-Süpürge Değişim Kampanyası (Eskiyi Getir, Yeniyi Götür)-202512", "202512", "Küçük Ev Aletleri", "Süpürge", "04.12.2025", "31.12.2025", "31.12.2025", None, 1250, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Değişim/trade-in kampanyası; hakediş tutarı belirtilmemiş, sadece yeni TRPF listesi verilmiş; eski süpürgenin kargo ile BSH İade Depo'ya gönderilmesi şart; hakediş için sabit tutar yok, hemen teslim satışı zorunlu (ileri teslim kapsam dışı); Değişim/trade-in kampanyası; hakediş tutarı belirtilmemiş,"),
    ("Aralık-Tekil Ankastre Kampanyası-202512", "202512", "Tekil", "Ankastre (Fırın hariç Ocak/Davlumbaz)", "04.12.2025", "31.12.2025", "31.05.2026", None, None, "Evet",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş"),
    ("Ocak-Ankastre Set Alana 6000 TL İndirim-202601", "202601", "Bundle-Set", "Ankastre Set (Fırın+Ocak+Davlumbaz)", "07.01.2026", "31.01.2026", "31.05.2026", 4500, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Ocak 2026 Perakende Uygulama Yazısı. PDF notu: Tüm set ürünleri aynı Diva faturasında olmalı"),
    ("Ocak-Soğutucu+Kurutma Alımına 10000 TL İndirim-202601", "202601", "Bundle-Set", "Kurutma Makinesi + Seçili Beyaz Eşya (Buzdolabı/Çamaşır/Bulaşık/Derin Dondurucu)", "07.01.2026", "31.01.2026", "31.05.2026", 8000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Ocak 2026 Perakende Uygulama Yazısı. PDF notu: Kampanya birleşme kısıtı: sadece tekil ile birleşir"),
    ("Ocak-Tekil Bulaşık Makinesi Kampanyası-202601", "202601", "Tekil", "Bulaşık Makinesi", "07.01.2026", "31.01.2026", "31.05.2026", None, None, "Evet",
     "Kaynak: Ocak 2026 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş; Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; stok mutabakatı %75 kuralı (30.11.2026) geçerli"),
    ("Ocak-Tekil Kurutma Makinesi Kampanyası-202601", "202601", "Tekil", "Kurutma Makinesi", "07.01.2026", "31.01.2026", "31.05.2026", None, None, "Evet",
     "Kaynak: Ocak 2026 Perakende Uygulama Yazısı. PDF notu: 30 gün koşulsuz iade + fiyat farkı yöntemi Montaj/Nakliye Adedi, sabit TL yok; stok mutabakatı %75 kuralı geçerli; 30 gün koşulsuz iade uygulaması + fiyat farkı yöntemi Montaj/Nakliye Adedi, sabit TL yok"),
    ("Ocak-Tekil Soğutucu Kampanyası-202601", "202601", "Tekil", "Soğutucu", "07.01.2026", "31.01.2026", "31.05.2026", None, None, "Evet",
     "Kaynak: Ocak 2026 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş; Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok. Bu aydan itibaren stok mutabakatı yapılmış bayilerde 30.11.2026 itibariyle nakliye/montaj oranı toplam hakedişin min. %75'i olmalı, altındaysa fark mahsuplaşılır"),
    ("Ocak-Tekil Çamaşır Kampanyası-202601", "202601", "Tekil", "Çamaşır Makinesi", "07.01.2026", "31.01.2026", "31.05.2026", None, None, "Evet",
     "Kaynak: Ocak 2026 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş; Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; stok mutabakatı %75 kuralı (30.11.2026) geçerli"),
    ("Ocak-Aile Bakanlığı Kampanyası – Bosch-202601", "202601", "Diğer", "Çeşitli (Buzdolabı/Derin Dondurucu/Çamaşır/Kurutma/Bulaşık/Fırın/Mikrodalga/Ocak/Davlumbaz/Küçük Ev Aletleri)", "07.01.2026", "31.01.2026", "31.05.2026", None, None, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: TRPF üzerinden ekstra %5 indirim (sabit TL değil); KEA modellerinde hakediş son tarihi 15.01.2026, diğer ürünlerde 31.05.2026; TRPF üzerinden ekstra %5 indirim (sabit TL değil); KEA modellerinde hakediş son tarihi 31.01.2026, diğer ürünlerde 31.05.2026"),
    ("Ocak-Ankastre Set Fırsat 2: Seçili Ankastre Set Alan Tüketicilerimize 8.000 TL İndirim-202601", "202601", "Bundle-Set", "Ankastre Set (Fırın+Ocak+Davlumbaz)", "07.01.2026", "31.01.2026", "31.05.2026", 6200, 2500, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Tüm set ürünleri aynı Diva faturasında olmalı"),
    ("Ocak-Ankastre Set Fırsat 3: Seçili Ankastre Set ile Birlikte Ankastre Bulaşık Makinesi Alan Tüketicilerimize 12.000 TL İndirim-202601", "202601", "Bundle-Set", "Ankastre Set (Fırın+Ocak+Davlumbaz) + Ankastre Bulaşık Makinesi", "07.01.2026", "31.01.2026", "31.05.2026", 12000, 300, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Tüm set ürünleri aynı Diva faturasında olmalı"),
    ("Ocak-Bosch Premium Seçili Büyük Beyaz Eşya Seti - Seçili modelleden (Buzdolabı, Çamaşır Makinesi, Kurutma Makinesi, Bulaşık Makinesi) 4'lü Beyaz Eşya Seti Alana BCS1041WAC Unlimited Süpürge Hediye-202601", "202601", "Küçük Ev Aletleri", "Büyük Beyaz Eşya 4'lü Set (Buzdolabı/Çamaşır/Kurutma/Bulaşık)", "07.01.2026", "31.01.2026", "31.05.2026", 27000, 100, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Hediye alınmazsa 34.000 TL indirim uygulanır; Hediye alınmazsa 34.000 TL indirim uygulanır (Ankastre Fırın bu ay set kapsamı dışı, önceki aydan farklı)"),
    ("Ocak-Derin Dondurucu Kampanyası-202601", "202601", "Bundle-Set", "Derin Dondurucu", "20.01.2026", "31.01.2026", "31.05.2026", None, None, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; model tablosu tekrar mevcut (önceki PDF'te yoktu)"),
    ("Ocak-Ocak Ayı Bulaşık Makineli Çeyiz Seti - 3 Büyük Beyaz Eşya (Buzdolabı, Çamaşır Makinesi, Kurutma Makinesi, Derin Dondurucu) ile Birlikte SMS26DW00T Bulaşık Makinesi Alımına Hediye Edilecektir-202601", "202601", "Bundle-Set", "Büyük Beyaz Eşya 3'lü Set + Bulaşık Makinesi", "07.01.2026", "31.01.2026", "31.05.2026", 17900, 2000, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: SMS26DW00T hediye edilir; farklı bulaşık makinesi tercih edilirse 22.000 TL indirim uygulanır; tekil ile birleşir, bundle ile birleşmez"),
    ("Ocak-Ocak Ayı Fırsat Ürünleri Kampanyası-202601", "202601", "Diğer", "Çeşitli (Buzdolabı/Derin Dondurucu/Bulaşık Mak./Çamaşır Mak./Ankastre Fırın - hareketi düşük modeller)", "07.01.2026", "31.01.2026", "28.02.2026", None, None, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; hakediş son tarihi diğer kampanyalardan farklı (28.02.2026)"),
    ("Ocak-Tekil Ankastre Kampanyası-202601", "202601", "Tekil", "Ankastre (Fırın hariç Ocak/Davlumbaz)", "07.01.2026", "31.01.2026", "31.05.2026", None, None, "Evet",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş"),
    ("Şubat Ayı Fırsat Ürünleri Kampanyası-202602", "202602", "Diğer", "Çeşitli (Buzdolabı/Derin Dondurucu/Bulaşık Mak./Çamaşır Mak./Ankastre Fırın - hareketi düşük modeller)", "04.02.2026", "28.02.2026", "31.03.2026", None, None, "",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; hakediş son tarihi diğer kampanyalardan farklı (31.03.2026); Hareket hızı düşük/2 yıl önce satışa kapanan modeller; hakediş son tarihi diğerlerinden farklı olarak 31.03.2026"),
    ("Şubat XL/XXL Soğutucu-RDW1276 Su Sebili Kampanyası-202602", "202602", "Bundle-Set", "Buzdolabı (XL/XXL)", "04.02.2026", "28.02.2026", "30.11.2026", 7200, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: RDW1276 su sebili hediye; alınmazsa 9.700 TL indirim; Yeni kampanya (önceki 4 PDF'te yoktu); hediye RDW1276 su sebili alınmazsa 9.700 TL indirim uygulanır; tekil ile birleşir, bundle ile birleşmez"),
    ("Şubat-Ankastre Set Alana 6000 TL İndirim-202602", "202602", "Bundle-Set", "Ankastre Set (Fırın+Ocak+Davlumbaz)", "04.02.2026", "28.02.2026", "30.11.2026", 4500, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: Tüm set ürünleri aynı Diva faturasında olmalı; ankastre fırın model listesi daralmış (HUA717 serisi bu kampanyadan çıkmış)"),
    ("Şubat-Bulaşık Makineli Çeyiz Seti 2-202602", "202602", "Bundle-Set", "3'lü Büyük Beyaz Eşya + Bulaşık Makinesi", "17.02.2026", "28.02.2026", "30.11.2026", 8000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. DİKKAT: Ay içindeki iki uygulama yazısı arasında tutar farkı görüldü (tüketici indirimleri: [10000, 23090], bayi hakedişleri: [8000, 17900]); en güncel (son yarı) değer alındı, admin teyit etmeli. PDF notu: En az 3 büyük beyaz eşya + SMS26DW00T bulaşık makinesi alımına hediye/indirim; Metinde 'Şubat Ayı Kurutmalı 2'li Beyaz Eşya Fırsatı' olarak da anılıyor"),
    ("Şubat-Bulaşık Makineli Çeyiz Seti 4-202602", "202602", "Bundle-Set", "Büyük Beyaz Eşya 4'lü Set (Buzdolabı/Çamaşır/Kurutma/Bulaşık)", "04.02.2026", "15.02.2026", "30.11.2026", 27000, 100, "Tekil Evet / Bundle Hayır",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: Hediye alınmazsa 35.450 TL indirim uygulanır (önceki ayda 34.000 TL idi)"),
    ("Şubat-Tekil Bulaşık Makinesi Kampanyası-202602", "202602", "Tekil", "Bulaşık Makinesi", "04.02.2026", "28.02.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; stok mutabakatı %75 kuralı (30.11.2026) geçerli"),
    ("Şubat-Tekil Derin Dondurucu Kampanyası-202602", "202602", "Tekil", "Derin Dondurucu", "04.02.2026", "28.02.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok"),
    ("Şubat-Tekil Kurutma Makinesi Kampanyası-202602", "202602", "Tekil", "Kurutma Makinesi", "04.02.2026", "28.02.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: 30 gün koşulsuz iade + fiyat farkı yöntemi Montaj/Nakliye Adedi, sabit TL yok; stok mutabakatı %75 kuralı geçerli; 30 gün koşulsuz iade uygulaması da mevcut"),
    ("Şubat-Tekil Soğutucu Kampanyası-202602", "202602", "Tekil", "Soğutucu", "04.02.2026", "28.02.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı Montaj/Nakliye Adedi yöntemiyle hesaplanır, sabit TL belirtilmemiş; Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; hakediş invoicing/montaj penceresi bu ay itibariyle 30.11.2026'ya uzatıldı; stok mutabakatı %75 kuralı (30.11.2026) geçerli"),
    ("Şubat-Tekil Çamaşır Kampanyası-202602", "202602", "Tekil", "Çamaşır Makinesi", "04.02.2026", "28.02.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL yok; stok mutabakatı %75 kuralı (30.11.2026) geçerli"),
    ("Şubat-Çamaşır+Kurutma Alımına 10000 TL İndirim-202602", "202602", "Bundle-Set", "Kurutma Makinesi + Seçili Beyaz Eşya (Buzdolabı/Çamaşır/Bulaşık/Derin Dondurucu)", "04.02.2026", "15.02.2026", "30.11.2026", 8000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Şubat 2026 Perakende Uygulama Yazısı. PDF notu: Kampanya birleşme kısıtı: sadece tekil ile birleşir"),
    ("Şubat-Aile Bakanlığı Kampanyası – Bosch-202602", "202602", "Diğer", "Çeşitli (Buzdolabı/Derin Dondurucu/Çamaşır/Kurutma/Bulaşık/Fırın/Mikrodalga/Ocak/Davlumbaz/Küçük Ev Aletleri)", "04.02.2026", "28.02.2026", "30.11.2026", None, None, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Evlenecek gençlere indirim kodu ile ekstra %5 indirim; KEA modeller için hakediş son tarihi 28.02.2026; TRPF üzerinden ekstra %5 indirim (sabit TL değil); KEA modellerinde hakediş son tarihi 15.02.2026, diğer ürünlerde 30.11.2026"),
    ("Şubat-Ankastre Set Fırsat 2: Seçili Ankastre Set Alan Tüketicilerimize 8.000 TL İndirim-202602", "202602", "Bundle-Set", "Ankastre Set (Fırın+Ocak+Davlumbaz)", "04.02.2026", "28.02.2026", "30.11.2026", 6200, 2500, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Tüm set ürünleri aynı Diva faturasında olmalı"),
    ("Şubat-Ankastre Set Fırsat 3: Seçili Ankastre Set ile Birlikte Ankastre Bulaşık Makinesi Alan Tüketicilerimize 12.000 TL İndirim-202602", "202602", "Bundle-Set", "Ankastre Set (Fırın+Ocak+Davlumbaz) + Ankastre Bulaşık Makinesi", "04.02.2026", "28.02.2026", "30.11.2026", 12000, 300, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Tüm set ürünleri aynı Diva faturasında olmalı; kapsama CDG714XB1 ve CMG7241B1 ankastre fırın kodları eklendi"),
    ("Şubat-Bosch Premium Seçili Büyük Beyaz Eşya Seti-202602", "202602", "Bundle-Set", "4'lü Beyaz Eşya Seti (Buzdolabı+Çamaşır+Kurutma+Bulaşık)", "17.02.2026", "28.02.2026", "30.11.2026", 27000, 100, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: BCS1041WAC Unlimited süpürge hediye; alınmazsa 35.450 TL indirim"),
    ("Şubat-Tekil Ankastre Kampanyası-202602", "202602", "Tekil", "Ankastre (Fırın hariç Ocak/Davlumbaz)", "04.02.2026", "28.02.2026", "30.11.2026", None, None, "Evet",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Fiyat farkı yöntemi: Montaj/Nakliye Adedi, sabit TL tutarı belirtilmemiş; model listesi güncellendi"),
    ("Şubat-Şubat Ayı Bulaşık Makineli Çeyiz Seti - 3 Büyük Beyaz Eşya (Buzdolabı, Çamaşır Makinesi, Kurutma Makinesi, Derin Dondurucu) ile Birlikte SMS26DW00T Bulaşık Makinesi Alımına Hediye Edilecektir-202602", "202602", "Bundle-Set", "Büyük Beyaz Eşya 3'lü Set + Bulaşık Makinesi", "04.02.2026", "15.02.2026", "30.11.2026", 17900, 2000, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: SMS26DW00T hediye edilir; farklı bulaşık makinesi tercih edilirse 23.090 TL indirim uygulanır (önceki aylarda 22.000 TL idi); tekil ile birleşir, bundle ile birleşmez"),
    ("Şubat-Şubat Ayı Küçük Ev Aletleri Fırsatları (Unlimited Kampanyası)-202602", "202602", "Küçük Ev Aletleri", "Küçük Ev Aletleri (Unlimited süpürge/hijyenik süpürge)", "17.02.2026", "28.02.2026", "28.02.2026", None, 3000, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Diva Hemen Teslim Fatura yöntemi, sabit TL yok; 'Şubat Ayı Unlimited Kampanyası-202602' adıyla anılıyor"),
    ("Mart Ayı Fırsat Ürünleri Kampanyası-202603", "202603", "Diğer", "Karma (Buzdolabı, Derin Dondurucu, Bulaşık Mak., Çamaşır Mak., Ankastre Fırın)", "05.03.2026", "31.03.2026", "30.04.2026", None, None, "",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı. PDF notu: Hareket hızı düşük/eski modeller"),
    ("Mart XL/XXL Soğutucu-RDW1276 Su Sebili Kampanyası-202603", "202603", "Bundle-Set", "Buzdolabı (XL/XXL)", "05.03.2026", "31.03.2026", "30.11.2026", 7200, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı."),
    ("Mart-4'lü Çeyiz Seti 2- 202603", "202603", "Bundle-Set", "Beyaz Eşya 2'li Set", "17.03.2026", "31.03.2026", "30.11.2026", 4000, 4000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı."),
    ("Mart-4'lü Çeyiz Seti 4 -202603", "202603", "Bundle-Set", "4'lü Büyük Beyaz Eşya Seti (Buzdolabı+Çamaşır+Kurutma+Derin Dondurucu)", "05.03.2026", "31.03.2026", "30.11.2026", 27000, 100, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı. DİKKAT: Ay içindeki iki uygulama yazısı arasında tutar farkı görüldü (tüketici indirimleri: [15000, 36510], bayi hakedişleri: [12200, 27000]); en güncel (son yarı) değer alındı, admin teyit etmeli. PDF notu: Bu ay ürün grubuna Bulaşık Makinesi de eklendi (5 kategoriden herhangi 4 adet alımı yeterli); Hediye alınmazsa 36.510 TL indirim"),
    ("Mart-Ankastre Set Alana 6000 TL İndirim-202603", "202603", "Bundle-Set", "Ankastre Set", "05.03.2026", "31.03.2026", "30.11.2026", 4500, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı."),
    ("Mart-Ankastre Set Alana 8000 TL İndirim-202603", "202603", "Bundle-Set", "Ankastre Set", "05.03.2026", "31.03.2026", "30.11.2026", 6200, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı."),
    ("Mart-Tekil Bulaşık Makinesi Kampanyası-202603", "202603", "Tekil", "Bulaşık Makinesi", "05.03.2026", "31.03.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı."),
    ("Mart-Tekil Derin Dondurucu Kampanyası-202603", "202603", "Tekil", "Derin Dondurucu", "05.03.2026", "31.03.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı."),
    ("Mart-Tekil Kurutma Makinesi Kampanyası-202603", "202603", "Tekil", "Kurutma Makinesi", "05.03.2026", "31.03.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı. PDF notu: 30 gün koşulsuz iade uygulaması da mevcut"),
    ("Mart-Tekil Soğutucu Kampanyası-202603", "202603", "Tekil", "Soğutucu", "05.03.2026", "31.03.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı."),
    ("Mart-Tekil Çamaşır Kampanyası-202603", "202603", "Tekil", "Çamaşır Makinesi", "05.03.2026", "31.03.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı."),
    ("Mart-Çamaşır+Kurutma Alımına 10000 TL İndirim-202603", "202603", "Bundle-Set", "Kurutma Makinesi + Seçili Beyaz Eşya", "05.03.2026", "31.03.2026", "30.11.2026", 8000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mart 2026 Perakende Uygulama Yazısı."),
    ("Mart-Aile Bakanlığı Kampanyası – Bosch-202603", "202603", "Diğer", "Karma", "05.03.2026", "31.03.2026", "30.11.2026", None, None, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: %5 ekstra indirim; KEA modeller için hakediş son tarihi 15.03.2026; %5 ekstra indirim; KEA modeller için hakediş son tarihi 31.03.2026"),
    ("Mart-Ankastre Set Fırsat 3: Seçili Ankastre Set ile Birlikte Ankastre Bulaşık Makinesi veya Ankastre Mikrodalga Alan Tüketicilerimize 12.000 TL İndirim-202603", "202603", "Bundle-Set", "Ankastre Set + Ankastre Bulaşık Makinesi/Mikrodalga", "05.03.2026", "31.03.2026", "30.11.2026", 12000, 300, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Mart-Mart Ayı 2'li Beyaz Eşya Fırsatı-202603", "202603", "Bundle-Set", "Beyaz Eşya 2'li Set (Buzdolabı/Çamaşır/Bulaşık/Derin Dondurucu)", "05.03.2026", "15.03.2026", "30.11.2026", 4000, 4000, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Mart-Mart Ayı Küçük Ev Aletleri Fırsatları (Seçili Mutfak Makinesi Kampanyası)-202603", "202603", "Küçük Ev Aletleri", "Küçük Ev Aletleri (Mutfak Makineleri)", "05.03.2026", "31.03.2026", "31.03.2026", None, 5000, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Diva Hemen Teslim Fatura yöntemi, sabit TL belirtilmemiş"),
    ("Mart-Stanley Mug Hediyeli Espresso Makineleri-202603", "202603", "Küçük Ev Aletleri", "Küçük Ev Aletleri (Tam Otomatik Espresso Makinesi)", "05.03.2026", "31.03.2026", "31.03.2026", None, 150, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Stanley Mug hediye; fiyat farkı belirtilmemiş; montaj/tanıtım 31.03.2026'ya kadar; Stanley Mug hediye; fiyat farkı tutarı sonradan bildirilecek; montaj/tanıtım 31.03.2026'ya kadar"),
    ("Mart-Tekil Ankastre Kampanyası-202603", "202603", "Tekil", "Ankastre (Fırın/Ocak/Davlumbaz)", "05.03.2026", "31.03.2026", "30.11.2026", None, None, "Evet",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Nisan Ayı Fırsat Ürünleri Kampanyası-202604", "202604", "Diğer", "Karma (Buzdolabı, Derin Dondurucu, Bulaşık Mak., Çamaşır Mak., Ankastre Fırın)", "04.04.2026", "30.04.2026", "31.05.2026", None, None, "",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı. PDF notu: Hareket hızı düşük/eski modeller"),
    ("Nisan XL/XXL Soğutucu-RDW1276 Su Sebili Kampanyası-202604", "202604", "Bundle-Set", "Buzdolabı (XL/XXL)", "04.04.2026", "30.04.2026", "30.11.2026", 7900, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı. PDF notu: İndirim/hakediş tutarı önceki aylara göre artırıldı (Şubat-Mart'ta 9.700/7.200 TL idi)"),
    ("Nisan Çamaşır+Bulaşık Alımına 5000 TL İndirim-202604", "202604", "Bundle-Set", "Beyaz Eşya 2'li Set", "04.04.2026", "30.04.2026", "30.11.2026", 4000, 4000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı."),
    ("Nisan-Ankastre Set Alana 11.400 TL İndirim-202604", "202604", "Bundle-Set", "Ankastre Set + Davlumbaz", "04.04.2026", "30.04.2026", "30.11.2026", 8850, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı. PDF notu: DWK63PJ20T/60T/61T davlumbaz hediye; alınmazsa 11.400 TL indirim"),
    ("Nisan-Ankastre Set Alana 6000 TL İndirim-202604", "202604", "Bundle-Set", "Ankastre Set", "04.04.2026", "30.04.2026", "30.11.2026", 4500, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı."),
    ("Nisan-Tekil Bulaşık Makinesi Kampanyası-202604", "202604", "Tekil", "Bulaşık Makinesi", "04.04.2026", "30.04.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı."),
    ("Nisan-Tekil Kurutma Makinesi Kampanyası-202604", "202604", "Tekil", "Kurutma Makinesi", "04.04.2026", "30.04.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı. PDF notu: 30 gün koşulsuz iade uygulaması da mevcut"),
    ("Nisan-Tekil Soğutucu Kampanyası-202604", "202604", "Tekil", "Soğutucu", "04.04.2026", "30.04.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı."),
    ("Nisan-Tekil Çamaşır Kampanyası-202604", "202604", "Tekil", "Çamaşır Makinesi", "04.04.2026", "30.04.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı."),
    ("Nisan-Çamaşır+Kurutma Alımına 10000 TL İndirim-202604", "202604", "Bundle-Set", "Kurutma Makinesi + Seçili Beyaz Eşya", "04.04.2026", "30.04.2026", "30.11.2026", 8000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Nisan 2026 Perakende Uygulama Yazısı."),
    ("Nisan-4'lü Beyaz Eşya Seti Alana BCS1041WAC Unlimited Süpürge Hediye-202604", "202604", "Küçük Ev Aletleri", "4'lü Beyaz Eşya Seti (Buzdolabı+Çamaşır+Kurutma+Bulaşık)", "04.04.2026", "30.04.2026", "30.11.2026", 27000, 100, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Nisan-Aile Bakanlığı Kampanyası – Bosch-202604", "202604", "Diğer", "Karma", "04.04.2026", "30.04.2026", "30.11.2026", None, None, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: %5 ekstra indirim; KEA modeller için hakediş son tarihi 15.04.2026; %5 ekstra indirim; KEA modeller için hakediş son tarihi 30.04.2026"),
    ("Nisan-Ankastre Set Fırsat 3: Seçili Ankastre Set ile Birlikte Ankastre Bulaşık Makinesi veya Ankastre Mikrodalga Alan Tüketicilerimize 12.000 TL İndirim-202604", "202604", "Bundle-Set", "Ankastre Set + Ankastre Bulaşık Makinesi/Mikrodalga", "04.04.2026", "30.04.2026", "30.11.2026", 12000, 300, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Nisan-Derin Dondurucu Kampanyası-202604", "202604", "Bundle-Set", "Derin Dondurucu", "04.04.2026", "30.04.2026", "30.11.2026", None, None, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Nisan-Nisan Ayı 4'lü Çeyiz Seti: 4 Büyük Beyaz Eşya Alan Tüketicilere 15.000 TL İndirim-202604", "202604", "Bundle-Set", "4'lü Büyük Beyaz Eşya Seti (5 kategoriden 4 adet)", "04.04.2026", "30.04.2026", "30.11.2026", 12200, 2000, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Nisan-Stanley Mug Hediyeli Espresso Makineleri-202604", "202604", "Küçük Ev Aletleri", "Küçük Ev Aletleri (Espresso Makinesi)", "04.04.2026", "30.04.2026", "31.05.2026", None, 150, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Fiyat farkı tutarı sonradan bildirilecek; montaj/tanıtım 30.04.2026'ya kadar; Montaj/tanıtım 31.05.2026'ya kadar"),
    ("Nisan-Süpürge Eskiyi Getir Yeniyi Götür Kampanyası-202604", "202604", "Küçük Ev Aletleri", "Süpürge (Trade-in)", "04.04.2026", "30.04.2026", "30.11.2026", None, 900, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Eski süpürge iadesi karşılığı model bazlı fiyat farkı: BGC41PET 17.920 TL, BGL41HYG2H 10.770 TL, BGL41PET2H 13.630 TL"),
    ("Nisan-Tekil Ankastre Kampanyası-202604", "202604", "Tekil", "Ankastre (Fırın/Ocak/Davlumbaz)", "04.04.2026", "30.04.2026", "30.11.2026", None, None, "Evet",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Anneler Günü Motorlu Ürünler Kampanyası - 202605", "202605", "Küçük Ev Aletleri", "Küçük Ev Aletleri (içecek hazırlama, motorlu ürünler, mutfak makinesi, süpürge, tost makinesi)", "06.05.2026", "15.05.2026", None, None, None, "",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: Diva Hemen Teslim Bildirim yöntemi; ek dosyada fiyat farkları var. Alt kategori adet limitleri: İçecek hazırlama 4000, Motorlu ürünler 2800, Mutfak Makinesi 1000, Süpürge 2300, Tost makinesi 1000."),
    ("Mayıs Ayı Fırsat Ürünleri Kampanyası-202605", "202605", "Diğer", "Buzdolabı/Derin Dondurucu/Bulaşık/Çamaşır/Ankastre Fırın (eski model envanteri)", "06.05.2026", "31.05.2026", "30.06.2026", None, None, "",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: İki yıl önce satışa kapanan/hareket hızı düşük seçili modeller; TL tutarı belirtilmemiş."),
    ("Mayıs Ayı Mug Hediyeli Tam Otomatik Espresso Makinesi Kampanyası-202605", "202605", "Küçük Ev Aletleri", "Tam Otomatik Espresso Makinesi", "06.05.2026", "31.05.2026", "30.06.2026", None, 100, "",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: Hediye Mug kargolanacak; fiyat farkı tutarı sonradan bildirilecek."),
    ("Mayıs Soğutucu + Derin Dondurucu Alımına 7500 TL İndirim-202605", "202605", "Bundle-Set", "Derin Dondurucu + Beyaz Eşya (2'li Set)", "06.05.2026", "31.05.2026", "30.11.2026", 6000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: PDF'te adet limiti cümlesi başka kampanyaya atıfta bulunuyor olabilir (belirsiz), sonraki dönem PDF'i 3000 diyor."),
    ("Mayıs XL/XXL Soğutucu-RDW1276 Su Sebili Kampanyası-202605", "202605", "Bundle-Set", "Buzdolabı (XL/XXL)", "06.05.2026", "31.05.2026", "30.11.2026", 7900, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: RDW1276 su sebili hediye edilir; alınmazsa 10.720 TL indirim."),
    ("Mayıs Çamaşır+Bulaşık Alımına 5000 TL İndirim-202605", "202605", "Bundle-Set", "Buzdolabı/Çamaşır/Bulaşık/Kurutma (2'li Set)", "06.05.2026", "31.05.2026", "30.11.2026", 4000, 4000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı."),
    ("Mayıs-Ankastre Set Alana 6000 TL İndirim-202605", "202605", "Bundle-Set", "Ankastre Set", "06.05.2026", "31.05.2026", "30.11.2026", 4500, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı."),
    ("Mayıs-Tekil Bulaşık Makinesi Kampanyası-202605", "202605", "Tekil", "Bulaşık Makinesi", "06.05.2026", "31.05.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: Montaj/Nakliye Adedi yöntemi, TL tutarı belirtilmemiş."),
    ("Mayıs-Tekil Derin Dondurucu Kampanyası-202605", "202605", "Tekil", "Derin Dondurucu", "06.05.2026", "31.05.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: Montaj/Nakliye Adedi yöntemi, TL tutarı belirtilmemiş."),
    ("Mayıs-Tekil Klima Kampanyası-202605", "202605", "Tekil", "Klima", "06.05.2026", "31.05.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: Montaj/Nakliye Adedi yöntemi, TL tutarı belirtilmemiş."),
    ("Mayıs-Tekil Kurutma Makinesi Kampanyası-202605", "202605", "Tekil", "Kurutma Makinesi", "06.05.2026", "31.05.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: 30 gün koşulsuz iade ayrıca belirtiliyor.; 30 gün koşulsuz iade uygulaması ayrıca belirtiliyor; TL tutarı yok."),
    ("Mayıs-Tekil Soğutucu Kampanyası-202605", "202605", "Tekil", "Soğutucu", "06.05.2026", "31.05.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: Fiyat farkı yöntemi Montaj/Nakliye Adedi; güncellenen TRPF ile devam, ayrı TL tutarı belirtilmemiş.; Montaj/Nakliye Adedi yöntemi, TL tutarı belirtilmemiş."),
    ("Mayıs-Tekil Çamaşır Kampanyası-202605", "202605", "Tekil", "Çamaşır Makinesi", "06.05.2026", "31.05.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: Montaj/Nakliye Adedi yöntemi, TL tutarı belirtilmemiş."),
    ("Mayıs-Çamaşır+Kurutma Alımına 10000 TL İndirim-202605", "202605", "Bundle-Set", "Çamaşır+Kurutma Set", "06.05.2026", "31.05.2026", "30.11.2026", 8000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Mayıs 2026 Perakende Uygulama Yazısı. PDF notu: Bundle; tekil kampanyalarla birleşebilir, diğer bundle'larla birleşemez."),
    ("Mayıs-4'lü Beyaz Eşya Seti Alana BCS1041WAC Unlimited Süpürge Hediye-202605", "202605", "Küçük Ev Aletleri", "Buzdolabı+Çamaşır+Kurutma+Bulaşık (4'lü Set)", "06.05.2026", "31.05.2026", "30.11.2026", 27000, 100, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Hediye alınmazsa 36.510 TL indirim."),
    ("Mayıs-Aile Bakanlığı Kampanyası – Bosch-202605", "202605", "Diğer", "Seçili Modeller (Buzdolabı/Derin Dondurucu/Çamaşır/Kurutma/Bulaşık/Fırın/Mikrodalga/Ocak/Davlumbaz/Süpürge/Çay-Kahve Makinesi vb.)", "06.05.2026", "15.05.2026", "30.11.2026", None, None, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: %5 indirim. KEA modeller için hakediş son tarihi 15.05.2026."),
    ("Mayıs-Ankastre Set Fırsat 2: Seçili Ankastre Set Alan Tüketicilerimize Davlumbaz Hediye-202605", "202605", "Bundle-Set", "Ankastre Set", "06.05.2026", "31.05.2026", "30.11.2026", 8850, 2500, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. DİKKAT: Ay içindeki iki uygulama yazısı arasında tutar farkı görüldü (tüketici indirimleri: [11400, 11920], bayi hakedişleri: [8850]); en güncel (son yarı) değer alındı, admin teyit etmeli. PDF notu: Davlumbaz hediye edilir; alınmazsa 11.920 TL indirim (DWK63PJ20T/60T/61T).; İndirim tutarı bu dönemde 11.400 TL (önceki dönemde 11.920 TL idi)."),
    ("Mayıs-Ankastre Set Fırsat 3: Seçili Ankastre Set ile Birlikte Ankastre Bulaşık Makinesi veya Ankastre Mikrodalga Alan Tüketicilerimize 13.000 TL İndirim-202605", "202605", "Bundle-Set", "Ankastre Set + Ankastre Bulaşık/Mikrodalga", "06.05.2026", "31.05.2026", "30.11.2026", 13000, 300, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Mayıs-Tekil Ankastre Kampanyası-202605", "202605", "Tekil", "Ankastre (Fırın/Ocak/Davlumbaz)", "06.05.2026", "31.05.2026", "30.11.2026", None, None, "Evet",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Ürün kodu tabloları var, TL tutarı belirtilmemiş."),
    ("Haziran Ayı Fırsat Ürünleri Kampanyası-202606", "202606", "Diğer", "Buzdolabı/Derin Dondurucu/Bulaşık/Çamaşır/Ankastre Fırın (eski model envanteri)", "03.06.2026", "30.06.2026", "31.07.2026", None, None, "",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı."),
    ("Haziran Çamaşır+Bulaşık Alımına 5000 TL İndirim-202606", "202606", "Bundle-Set", "Buzdolabı/Çamaşır/Bulaşık/Kurutma (2'li Set)", "17.06.2026", "30.06.2026", "30.11.2026", 4000, 4000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı."),
    ("Haziran-Ankastre Set Alana 5000 TL İndirim-202606", "202606", "Bundle-Set", "Ankastre Set", "03.06.2026", "30.06.2026", "30.11.2026", 3700, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı. PDF notu: Tutar 6.000 TL'den 5.000 TL'ye düştü."),
    ("Haziran-Tekil Bulaşık Makinesi Kampanyası-202606", "202606", "Tekil", "Bulaşık Makinesi", "03.06.2026", "30.06.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı."),
    ("Haziran-Tekil Derin Dondurucu Kampanyası-202606", "202606", "Tekil", "Derin Dondurucu", "03.06.2026", "30.06.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı."),
    ("Haziran-Tekil Kurutma Makinesi Kampanyası-202606", "202606", "Tekil", "Kurutma Makinesi", "03.06.2026", "30.06.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı. PDF notu: 30 gün koşulsuz iade ayrıca belirtiliyor."),
    ("Haziran-Tekil Soğutucu Kampanyası-202606", "202606", "Tekil", "Soğutucu", "03.06.2026", "30.06.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı."),
    ("Haziran-Tekil Çamaşır Kampanyası-202606", "202606", "Tekil", "Çamaşır Makinesi", "03.06.2026", "30.06.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı. PDF notu: Kapsam bu dönemde daralmış (WGK242ZXTR, WGK244Z0TR, WGA25202TR, WGA25203TR)."),
    ("Haziran-XL/XXL Soğutucu-RDW1276 Su Sebili Kampanyası-202606", "202606", "Bundle-Set", "Buzdolabı (XL/XXL)", "03.06.2026", "30.06.2026", "30.11.2026", 7900, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı. PDF notu: Bu dönemde '2'li Beyaz Eşya Fırsatı' ve 'XL/XXL 2'li Paketleri' bundle kampanyaları PDF'te yok."),
    ("Haziran-Çamaşır+Dondurucu Alımına 7500 TL İndirim-202606", "202606", "Bundle-Set", "Derin Dondurucu + Beyaz Eşya (2'li Set)", "03.06.2026", "30.06.2026", "30.11.2026", 6000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı."),
    ("Haziran-Çamaşır+Kurutma Alımına 10000 TL İndirim-202606", "202606", "Bundle-Set", "Çamaşır+Kurutma Set", "03.06.2026", "30.06.2026", "30.11.2026", 8000, 3000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Haziran 2026 Perakende Uygulama Yazısı."),
    ("Haziran-4'lü Beyaz Eşya Seti Alana BCS1041WAC Unlimited Süpürge Hediye-202606", "202606", "Küçük Ev Aletleri", "Buzdolabı+Çamaşır+Kurutma+Bulaşık (4'lü Set)", "03.06.2026", "30.06.2026", "30.11.2026", 27000, 100, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: İndirim tutarı yükselmiş (önceki dönemlerde 36.510 TL idi)."),
    ("Haziran-Ankastre Set Fırsat 2: Seçili Ankastre Set Alan Tüketicilerimize 9.000 TL İndirim-202606", "202606", "Bundle-Set", "Ankastre Set", "03.06.2026", "30.06.2026", "30.11.2026", 6900, 2500, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Bu dönemde davlumbaz hediye yapısı yerine doğrudan 9.000 TL indirim uygulanıyor."),
    ("Haziran-Ankastre Set Fırsat 3: Seçili Ankastre Set ile Birlikte Ankastre Bulaşık Makinesi veya Ankastre Mikrodalga Alan Tüketicilerimize 13.000 TL İndirim-202606", "202606", "Bundle-Set", "Ankastre Set + Ankastre Bulaşık/Mikrodalga", "03.06.2026", "30.06.2026", "30.11.2026", 13000, 300, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Haziran-Klima Kampanyası-202606", "202606", "Bundle-Set", "Klima", "03.06.2026", "30.06.2026", "30.11.2026", None, None, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
    ("Haziran-Stanley Mug Hediyeli Espresso Makineleri-202606", "202606", "Küçük Ev Aletleri", "Tam Otomatik Espresso Makinesi", "03.06.2026", "30.06.2026", "31.07.2026", None, 50, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Adet limiti 50'ye düşürülmüş (önceki aylarda 100 idi)."),
    ("Haziran-Tekil Ankastre Kampanyası-202606", "202606", "Tekil", "Ankastre (genel)", "03.06.2026", "30.06.2026", "30.11.2026", None, None, "Evet",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Bu dönemde ürün tablosu/Montaj-Nakliye mekanizması yok, sadece 'avantajlı TRPF' ibaresi var."),
    ("Temmuz Ayı Fırsat Ürünleri Kampanyası-202607", "202607", "Diğer", "Unlimited (Dikey Süpürge)", "03.07.2026", "31.07.2026", "31.08.2026", None, 2200, "",
     "Kaynak: Temmuz 2026 Perakende Uygulama Yazısı. PDF notu: Diva Hemen Teslim Fatura yöntemi, sabit TL belirtilmemiş; Hareket hızı düşük/2 yıl önce satışa kapanan seçili modeller için ekstra indirim, TL/adet limiti belirtilmemiş; PDF'de 'Hakkediş son tarihi' olarak yazılmış (yazım hatası); TL/adet limiti belirtilmemiş"),
    ("Temmuz- Bulaşık+Dondurucu Alımına 5.000 TL İndirim-202607", "202607", "Bundle-Set", "Buzdolabı+Çamaşır+Kurutma+Bulaşık+Derin Dondurucu (2'li Set)", "03.07.2026", "31.07.2026", "30.11.2026", 4000, 4000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Temmuz 2026 Perakende Uygulama Yazısı. PDF notu: Bu dönemde Buzdolabı grubu bundle'a dahil ama XL/XXL modeller hariç (önceki dönemde tüm buzdolabı dahildi, bu dönem farklı); hakediş son tarihi bölümde belirtilmemiş, genel tarih kullanıldı; Hakediş son tarihi bölümde ayrıca belirtilmemiş, genel tarih kullanıldı; tekil kampanyalarla birleşebilir, di"),
    ("Temmuz-Ankastre Set Alana 5000 TL İndirim-202607", "202607", "Bundle-Set", "Ankastre Set", "03.07.2026", "31.07.2026", "30.11.2026", 3700, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Temmuz 2026 Perakende Uygulama Yazısı. PDF notu: Hakediş son tarihi bölümde belirtilmemiş, genel tarih kullanıldı"),
    ("Temmuz-Ankastre Set Alana 9.000 TL İndirim-202607", "202607", "Bundle-Set", "Ankastre Set", "03.07.2026", "31.07.2026", "30.11.2026", 6900, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Temmuz 2026 Perakende Uygulama Yazısı. PDF notu: Hakediş son tarihi bölümde belirtilmemiş, genel tarih kullanıldı"),
    ("Temmuz-Tekil Bulaşık Makinesi Kampanyası-202607", "202607", "Tekil", "Bulaşık Makinesi", "03.07.2026", "31.07.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Temmuz 2026 Perakende Uygulama Yazısı. PDF notu: Diva Hemen Teslim Fatura yöntemi, sabit TL belirtilmemiş"),
    ("Temmuz-Tekil Kurutma Makinesi Kampanyası-202607", "202607", "Tekil", "Kurutma Makinesi", "03.07.2026", "31.07.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Temmuz 2026 Perakende Uygulama Yazısı. PDF notu: 30 gün koşulsuz iade uygulaması devam ediyor"),
    ("Temmuz-Tekil Soğutucu Kampanyası-202607", "202607", "Tekil", "Soğutucu", "03.07.2026", "31.07.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Temmuz 2026 Perakende Uygulama Yazısı. PDF notu: Diva Hemen Teslim Fatura yöntemi, sabit TL belirtilmemiş"),
    ("Temmuz-Tekil Çamaşır Kampanyası-202607", "202607", "Tekil", "Çamaşır Makinesi", "03.07.2026", "31.07.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Temmuz 2026 Perakende Uygulama Yazısı. PDF notu: Tek model: WGA25202TR"),
    ("Temmuz-XL/XXL Soğutucu-RDW1276 Su Sebili Kampanyası-202607", "202607", "Bundle-Set", "Buzdolabı (XL/XXL)", "03.07.2026", "15.07.2026", "30.11.2026", 8500, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Temmuz 2026 Perakende Uygulama Yazısı. PDF notu: 11.900 TL indirim yalnızca RDW1276 su sebili hediyesi alınmazsa uygulanır; hakediş son tarihi bölümde belirtilmemiş, genel tarih kullanıldı"),
    ("Temmuz-Derin Dondurucu Kampanyası-202607", "202607", "Bundle-Set", "Derin Dondurucu", "03.07.2026", "31.07.2026", "30.11.2026", None, None, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Tek model: GSN33VWE0N"),
    ("Temmuz-Klima Kampanyası-202607", "202607", "Bundle-Set", "Klima", "03.07.2026", "20.07.2026", "30.11.2026", None, None, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: DİKKAT: Bu kampanyanın geçerlilik aralığı PDF döneminin geri kalanından farklı ve kısa (18-20 Temmuz, PDF genelinde 18-31 Temmuz), PDF'de kırmızı ile vurgulanmış"),
    ("Temmuz-Stanley Mug Hediyeli Espresso Makineleri-202607", "202607", "Küçük Ev Aletleri", "Tam Otomatik Espresso Makinesi", "03.07.2026", "31.07.2026", "31.08.2026", None, 50, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: TL indirim yok; hediye Stanley Mug montaj sonrası kargolanır"),
    ("Temmuz-Temmuz Ayı Tam Otomatik Espresso Makinesi Kampanyası-202607", "202607", "Küçük Ev Aletleri", "Tam Otomatik Espresso Makinesi", "03.07.2026", "31.07.2026", "30.11.2026", None, 300, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Diva Hemen Teslim Fatura yöntemi, sabit TL belirtilmemiş; Diva Hemen Teslim Fatura yöntemi, sabit TL belirtilmemiş; bu dönemde ayrı 'Unlimited Kampanyası' yok"),
    ("Ağustos Ayı Fırsat Ürünleri Kampanyası-202608", "202608", "Diğer", "Fırsat Ürünleri (Buzdolabı+Bulaşık Makinesi+Çamaşır Makinesi+Ankastre Fırın)", "05.08.2026", "31.08.2026", "30.09.2026", None, None, "",
     "Kaynak: Ağustos 2026 Perakende Uygulama Yazısı. PDF notu: Hareket hızı düşük/2 yıl önce satışa kapanan modeller; PDF'de 'Hakkediş son tarihi' olarak yazılmış (yazım hatası); TL/adet limiti belirtilmemiş"),
    ("Ağustos-Ankastre Set Alana 5000 TL İndirim-202608", "202608", "Bundle-Set", "Ankastre Set", "05.08.2026", "31.08.2026", "30.11.2026", 3700, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Ağustos 2026 Perakende Uygulama Yazısı. PDF notu: Hakediş son tarihi bölümde belirtilmemiş, genel tarih kullanıldı; Tekil ile birleşir, bundle ile birleşmez"),
    ("Ağustos-Kurutma+Bulaşık Alımına 5000 TL İndirim-202608", "202608", "Bundle-Set", "Çamaşır+Kurutma+Bulaşık+Derin Dondurucu+Seçili Buzdolabı (2'li Set)", "05.08.2026", "31.08.2026", "30.11.2026", 4000, 4000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Ağustos 2026 Perakende Uygulama Yazısı. PDF notu: Seçili buzdolapları XL ve XXL hariç; hakediş son tarihi bölümde belirtilmemiş, genel tarih kullanıldı; Tekil ile birleşir, bundle ile birleşmez"),
    ("Ağustos-Tekil Bulaşık Makinesi Kampanyası-202608", "202608", "Tekil", "Bulaşık Makinesi", "05.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Ağustos 2026 Perakende Uygulama Yazısı. PDF notu: Diva Hemen Teslim Fatura yöntemi, sabit TL belirtilmemiş"),
    ("Ağustos-Tekil Klima Kampanyası-202608", "202608", "Tekil", "Klima", "05.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Ağustos 2026 Perakende Uygulama Yazısı."),
    ("Ağustos-Tekil Kurutma Makinesi Kampanyası-202608", "202608", "Tekil", "Kurutma Makinesi", "05.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Ağustos 2026 Perakende Uygulama Yazısı. PDF notu: 30 gün koşulsuz iade uygulaması devam; 30 gün koşulsuz iade uygulaması devam ediyor"),
    ("Ağustos-Tekil Soğutucu Kampanyası-202608", "202608", "Tekil", "Soğutucu", "05.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Ağustos 2026 Perakende Uygulama Yazısı. PDF notu: Diva Hemen Teslim Fatura yöntemi, sabit TL belirtilmemiş"),
    ("Ağustos-Tekil Çamaşır Kampanyası-202608", "202608", "Tekil", "Çamaşır Makinesi", "05.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Ağustos 2026 Perakende Uygulama Yazısı. PDF notu: Modeller: WGA242X3TR, WGA25202TR, WGA25203TR"),
    ("Ağustos-Ankastre Set Fırsat 2: Seçili Ankastre Set Alan Tüketicilerimize 12.100 TL İndirim-202608", "202608", "Bundle-Set", "Ankastre Set (Davlumbazlı)", "05.08.2026", "31.08.2026", "30.11.2026", 9400, 2500, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: 12.100 TL indirim yalnızca hediye davlumbaz (DWK63PJ20T/60T/61T) tercih edilmezse uygulanır; bu ay tutar önceki aylardaki 9.000 TL'lik Fırsat 2 kampanyasından farklı ve daha yüksek; hakediş son tarihi bölümde belirtilmemiş, genel tarih kullanıldı; Davlumbaz hediye (DWK63PJ20T/60T/61T) alınmazsa 12.1"),
    ("Ağustos-Ağustos Ayı Tam Otomatik Espresso Makinesi Kampanyası-202608", "202608", "Küçük Ev Aletleri", "Tam Otomatik Espresso Makinesi", "05.08.2026", "31.08.2026", "30.11.2026", None, 300, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Diva Hemen Teslim Fatura yöntemi; Diva Hemen Teslim Fatura yöntemi, sabit TL belirtilmemiş"),
    ("Ağustos-Derin Dondurucu Kampanyası-202608", "202608", "Bundle-Set", "Derin Dondurucu", "05.08.2026", "31.08.2026", "30.11.2026", None, None, "Tekil Evet / Bundle Hayır",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Tek model: GSN33VWE0N"),
    ("Ağustos-Stanley Mug Hediyeli Espresso Makineleri-202608", "202608", "Küçük Ev Aletleri", "Tam Otomatik Espresso Makinesi", "05.08.2026", "31.08.2026", "30.09.2026", None, 50, "",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI. PDF notu: Hediye Mug montaj sonrası kargolanır; TL indirim yok; hediye Stanley Mug montaj sonrası kargolanır; hakediş son tarihi bu ay 30.09.2026 (önceki aylarda 31.08.2026'ydı)"),
    ("Ağustos-Tekil Ankastre Kampanyası-202608", "202608", "Tekil", "Ankastre", "18.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "DİVA Paket Adı ORNEK SATIS VERISINDE TEYIT EDILEMEDI (o ay bu kampanyadan satis olmamis olabilir) - Kampanya Adi sutunu DIVA'daki gercek Paket Adi ile ADMIN tarafindan TEYIT EDILMELI."),
]
r = 2
for row in kamp_rows:
    for i, val in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=val)
        c.font = INPUT_FONT
        c.fill = ADMIN_FILL
        c.border = BORDER
        c.alignment = LEFT_AL if i in (1, 4, 11) else CENTER
        if i in (5, 6, 7):
            c.number_format = DATE_FMT
        if i == 8 and val is not None:
            c.number_format = CUR
    ws.row_dimensions[r].height = 30
    r += 1

KAMP_FIRST, KAMP_LAST = 2, 400
for row in range(r, KAMP_LAST + 1):
    for i in range(1, len(kamp_headers) + 1):
        c = ws.cell(row=row, column=i)
        c.fill = ADMIN_FILL
        c.font = INPUT_FONT
        c.border = BORDER
        c.alignment = LEFT_AL if i in (1, 4, 11) else CENTER
        if i in (5, 6, 7):
            c.number_format = DATE_FMT
        if i == 8:
            c.number_format = CUR

note = ws.cell(row=KAMP_LAST + 2, column=1,
    value="YÖNETİCİ NOTU: Her ay yeni kampanya duyuru PDF'i geldiğinde bu tabloya yeni satırlar eklenir. Kampanya "
          "Adı sütunu mutlaka DİVA'daki 'Paket Adı' ile birebir aynı yazılmalıdır, aksi halde 'DİVA Satışları' ve "
          "'Kontrol' sekmelerindeki eşleştirme formülleri kampanyayı bulamaz.")
note.font = NOTE_FONT
note.alignment = Alignment(wrap_text=True)
ws.merge_cells(start_row=KAMP_LAST + 2, start_column=1, end_row=KAMP_LAST + 2, end_column=11)
ws.row_dimensions[KAMP_LAST + 2].height = 40

KAMP_NAME_RNG = f"Kampanyalar!$A${KAMP_FIRST}:$A${KAMP_LAST}"
KAMP_START_RNG = f"Kampanyalar!$E${KAMP_FIRST}:$E${KAMP_LAST}"
KAMP_END_RNG = f"Kampanyalar!$F${KAMP_FIRST}:$F${KAMP_LAST}"
KAMP_DUE_RNG = f"Kampanyalar!$G${KAMP_FIRST}:$G${KAMP_LAST}"
KAMP_AMT_RNG = f"Kampanyalar!$H${KAMP_FIRST}:$H${KAMP_LAST}"

print("Part 2 (Kampanyalar) done")

# ============================================================
# 4) DIVA SATISLARI  (35 raw columns, exact DİVA export layout)
# ============================================================
DIVA_HEADERS = [
    "Şirket Kodu", "Şirket Adı", "Şube Kodu", "Şube Adı", "Bayi Kodu", "Aile", "Firma Adı",
    "Fatura Tarihi", "Fatura No", "Sipariş No", "Müşteri Kart No", "Hesap Kodu", "Ad Soyad",
    "Kampanya Tipi", "Paket Adı", "Satır İndirim Tipi", "Dip İndirim Tipi", "Ürün Kodu", "Ürün Adı",
    "Barkod", "Miktar", "Kazanılan Kod", "Kullanılan Kod", "Seri No", "Brüt", "Satır İndirimi",
    "İndirim (%)", "Dip İndirim", "İndirim (%)1", "Toplam İndirim", "İndirim (%)2", "Fatura Tutarı",
    "KDV Tutarı", "Özel Kod 2", "Kategori",
]
assert len(DIVA_HEADERS) == 35
DIVA_CALC_HEADERS = [
    "Eşleştirme Anahtarı\n(Ürün Düzeyi -\nFaturaNo+Ürün)",
    "Eşleştirme Anahtarı\n(Kampanya Düzeyi -\nFaturaNo+Ürün+Kampanya)",
    "Kampanya\nBulundu mu", "Kampanya\nBaşlangıç", "Kampanya\nBitiş",
    "Hakediş\nSon Tarihi", "Fatura Tarihi\nKampanya\nAralığında mı", "Sabit Hakediş\nTutarı (varsa)",
    "Beklenen\nHakediş Tutarı", "Uygulanacak\nKural (Ön Beklenti)",
]

ws = wb.create_sheet("DIVA Satislari")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"
set_widths(ws, [11] * 35 + [18, 20, 12, 11, 11, 11, 12, 13, 13, 20])
header_row(ws, 1, DIVA_HEADERS, RAW_HEADER_FILL, RAW_HEADER_FONT, height=44)
header_row(ws, 1, DIVA_CALC_HEADERS, CALC_HEADER_FILL, HEADER_FONT, height=44, start_col=36)

DS_FIRST, DS_LAST = 3, 1002
FATURANO_C, URUNKODU_C, FATURATARIHI_C, PAKETADI_C, TOPLAMINDIRIM_C = 9, 18, 8, 15, 30
DSKEY2_C = 36  # urun duzeyi (E2E eslestirmesi icin: bir urun birden fazla kampanyaya konu olabilir)
KEY_C = 37     # kampanya duzeyi (Hakedis Fatura Detay eslestirmesi icin)
KAMP_BULUNDU_C, KBAS_C, KBIT_C, KSON_C, ARALIK_C, SABIT_C, BEKLENEN_C, KURAL_C = 38, 39, 40, 41, 42, 43, 44, 45

example_diva = ["BSHMDEMIR", "MUSTAFA DEMİREL", "BSHMDEMIR", "MUSTAFA DEMİREL", "7000129546", "Bosch-LDA",
    "", "24.08.2026", "2600000650", "", "", "X000002267", "ÖRNEK MÜŞTERİ", "Fatura altı indirim",
    "Ağustos-Tekil Bulaşık Makinesi Kampanyası-202608", "", "", "SMS4IKW62T", "Solo bulaşık makinesi",
    "4242005342716", 1, "", "", "", 37000, 5761.9, 15.57, 0, 0, 5761.9, 15.6, 31238.1, 5206.35, "", ""]
for i, val in enumerate(example_diva, start=1):
    c = ws.cell(row=2, column=i, value=val)
    c.font = EXAMPLE_FONT
    c.fill = EXAMPLE_FILL
    c.border = BORDER
    c.alignment = CENTER
    if i == FATURATARIHI_C:
        c.number_format = DATE_FMT
    if i in (25, 26, 28, 30, 32, 33):
        c.number_format = CUR

for row in range(DS_FIRST, DS_LAST + 1):
    for col in range(1, 36):
        c = ws.cell(row=row, column=col)
        c.fill = INPUT_FILL
        c.font = INPUT_FONT
        c.border = BORDER
        c.alignment = CENTER
        if col == FATURATARIHI_C:
            c.number_format = DATE_FMT
        if col in (25, 26, 28, 30, 32, 33):
            c.number_format = CUR

    L = get_column_letter
    fno = f"{L(FATURANO_C)}{row}"
    urun = f"{L(URUNKODU_C)}{row}"
    ftar = f"{L(FATURATARIHI_C)}{row}"
    paket = f"{L(PAKETADI_C)}{row}"
    toplamind = f"{L(TOPLAMINDIRIM_C)}{row}"
    kbas = f"{L(KBAS_C)}{row}"
    kbit = f"{L(KBIT_C)}{row}"
    sabit = f"{L(SABIT_C)}{row}"

    formulas = {
        DSKEY2_C: f'=IF({fno}="","",{fno}&"|"&{urun})',
        KEY_C: f'=IF({fno}="","",{fno}&"|"&{urun}&"|"&{paket})',
        KAMP_BULUNDU_C: f'=IF({fno}="","",IF(ISNA(MATCH({paket},{KAMP_NAME_RNG},0)),"Hayır - Kampanyalar\'da Yok","Evet"))',
        KBAS_C: f'=IF({fno}="","",IFERROR(INDEX({KAMP_START_RNG},MATCH({paket},{KAMP_NAME_RNG},0)),""))',
        KBIT_C: f'=IF({fno}="","",IFERROR(INDEX({KAMP_END_RNG},MATCH({paket},{KAMP_NAME_RNG},0)),""))',
        KSON_C: f'=IF({fno}="","",IFERROR(INDEX({KAMP_DUE_RNG},MATCH({paket},{KAMP_NAME_RNG},0)),""))',
        ARALIK_C: (f'=IF(OR({fno}="",{kbas}="",{kbit}=""),"",'
                   f'IF(AND({ftar}>={kbas},{ftar}<={kbit}),"Evet","HAYIR - Tarih Dışı"))'),
        SABIT_C: f'=IF({fno}="","",IFERROR(INDEX({KAMP_AMT_RNG},MATCH({paket},{KAMP_NAME_RNG},0)),""))',
        BEKLENEN_C: f'=IF({fno}="","",IF(ISNUMBER({sabit}),{sabit},{toplamind}))',
        KURAL_C: (f'=IF({fno}="","",IF({P_MUTABAKAT}="Evet","Kural A - Hemen Ödeme (Stok Mutabakatlı)",'
                  f'"Kural B - Nakliye/Montaj Sonrası Ödeme"))'),
    }
    for col, formula in formulas.items():
        c = ws.cell(row=row, column=col, value=formula)
        c.font = FORMULA_FONT
        c.border = BORDER
        c.alignment = CENTER
        if col in (KBAS_C, KBIT_C, KSON_C):
            c.number_format = DATE_FMT
        if col in (SABIT_C, BEKLENEN_C):
            c.number_format = CUR

DS_KEY_RNG = f"'DIVA Satislari'!${get_column_letter(KEY_C)}${DS_FIRST}:${get_column_letter(KEY_C)}${DS_LAST}"
DS_KEY2_RNG = f"'DIVA Satislari'!${get_column_letter(DSKEY2_C)}${DS_FIRST}:${get_column_letter(DSKEY2_C)}${DS_LAST}"

print("Part 3 (DIVA Satislari) done")

# ============================================================
# 5) E2E NAKLIYE MONTAJ  (38 raw columns, exact E2E export layout)
# ============================================================
E2E_HEADERS = [
    "MerkezKodu", "BayiKodu", "UrunCikisDepoKodu", "BayiUnvani", "SiparisNo", "DivaSiparisFaturaNo",
    "DivaFaturaNo", "E2ESiparisNo", "SiparisTarihi", "SiparisTuru", "SiparisKanali", "TuketiciAdi",
    "TuketiciSoyadi", "TuketiciTelefon", "TeslimEdilecekKisiEmail", "UrunKodu", "UrunAdi", "UrunTipi",
    "TeslimatFisNo", "TeslimatDurum", "IadeDurum", "TeslimatAdresBilgisi", "TeslimatTuketiciIl",
    "TeslimatTuketiciIlce", "TeslimatIstenilenTarih", "TeslimatRandevuSaati", "MontajFisNo", "MontajDurum",
    "MontajIstenilenilenTarih", "MontajRandevuSaati", "SatisTipi", "TeslimatTipi", "Tasiyici", "DepoYeri",
    "IrsaliyeTarihi", "IrsaliyeNo", "Platform", "PlatformName",
]
assert len(E2E_HEADERS) == 38
E2E_CALC_HEADERS = [
    "Eşleştirme\nAnahtarı", "Sipariş Tarihi\n(Tarih)", "Teslimat İstenilen\nTarih (Tarih)",
    "Montaj İstenilen\nTarih (Tarih)", "İrsaliye Tarih\nMetni (yardımcı)", "İrsaliye Tarihi\n(Tarih)",
    "Nakliye\nİzlenebilir mi", "Nakliye\nTamamlandı mı", "Montaj\nTamamlandı mı",
    "Hakediş Tetikleyici\nTarih",
]

ws = wb.create_sheet("E2E Nakliye Montaj")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"
set_widths(ws, [11] * 38 + [16, 12, 12, 12, 12, 12, 12, 11, 11, 13])
header_row(ws, 1, E2E_HEADERS, RAW_HEADER_FILL, RAW_HEADER_FONT, height=44)
header_row(ws, 1, E2E_CALC_HEADERS, CALC_HEADER_FILL, HEADER_FONT, height=44, start_col=39)

E2E_FIRST, E2E_LAST = 3, 1002
DIVAFNO_C, URUNKODU_C2, SIPTAR_C, TESDURUM_C, TESISTTAR_C, MONTFISNO_C, MONTDURUM_C = 7, 16, 9, 20, 25, 27, 28
MONTISTTAR_C, TASIYICI_C, IRSTAR_C = 29, 33, 35
KEY2_C, SIPTARP_C, TESISTP_C, MONTISTP_C, IRSTXT_C, IRSTARP_C = 39, 40, 41, 42, 43, 44
NAKIZLE_C, NAKTAM_C, MONTAM_C, TETIK_C = 45, 46, 47, 48

example_e2e = ["7000129546", "7000129546", "7000129546", "MUSTAFA DEMIREL", "TR01-20764369", "2600000650",
    "2600000650", "5371889", "24-08-2026", "Offline", "DIVA", "ÖRNEK", "MÜŞTERİ", "5xxxxxxxxx", "",
    "SMS4IKW62T", "Solo bulaşık makinesi", "LDA", "", "Siparişiniz Teslim Edildi", "-", "Örnek adres",
    "ANTALYA", "ALANYA", "24-08-2026", "", "158360746", "Siparişinizin Montajı Yapıldı", "29-08-2026",
    "Tüm Gün", "Hemen Teslim Satış", "Mağaza Depodan Tüketiciye", "Servis", "Magaza Depo",
    "8/24/2026 6:24:05 PM", "MDI2026000000099", "1354", "Antalya Platform"]
for i, val in enumerate(example_e2e, start=1):
    c = ws.cell(row=2, column=i, value=val)
    c.font = EXAMPLE_FONT
    c.fill = EXAMPLE_FILL
    c.border = BORDER
    c.alignment = CENTER

for row in range(E2E_FIRST, E2E_LAST + 1):
    for col in range(1, 39):
        c = ws.cell(row=row, column=col)
        c.fill = INPUT_FILL
        c.font = INPUT_FONT
        c.border = BORDER
        c.alignment = CENTER

    L = get_column_letter
    divf = f"{L(DIVAFNO_C)}{row}"
    urun = f"{L(URUNKODU_C2)}{row}"
    siptar = f"{L(SIPTAR_C)}{row}"
    tesdurum = f"{L(TESDURUM_C)}{row}"
    tesisttar = f"{L(TESISTTAR_C)}{row}"
    montdurum = f"{L(MONTDURUM_C)}{row}"
    montisttar = f"{L(MONTISTTAR_C)}{row}"
    tasiyici = f"{L(TASIYICI_C)}{row}"
    irstar = f"{L(IRSTAR_C)}{row}"
    irstxt = f"{L(IRSTXT_C)}{row}"
    irstarp = f"{L(IRSTARP_C)}{row}"
    nakizle = f"{L(NAKIZLE_C)}{row}"
    naktam = f"{L(NAKTAM_C)}{row}"
    montam = f"{L(MONTAM_C)}{row}"
    montistp = f"{L(MONTISTP_C)}{row}"

    def ddmmyyyy(src):
        return f'=IF({src}="","",IFERROR(DATE(VALUE(RIGHT({src},4)),VALUE(MID({src},4,2)),VALUE(LEFT({src},2))),""))'

    formulas = {
        KEY2_C: f'=IF({divf}="","",{divf}&"|"&{urun})',
        SIPTARP_C: ddmmyyyy(siptar),
        TESISTP_C: ddmmyyyy(tesisttar),
        MONTISTP_C: ddmmyyyy(montisttar),
        IRSTXT_C: f'=IF({irstar}="","",TRIM(LEFT({irstar},FIND(" ",{irstar}&" ")-1)))',
        IRSTARP_C: (f'=IF({irstxt}="","",IFERROR(DATE('
                    f'VALUE(MID({irstxt},FIND("/",{irstxt},FIND("/",{irstxt})+1)+1,4)),'
                    f'VALUE(LEFT({irstxt},FIND("/",{irstxt})-1)),'
                    f'VALUE(MID({irstxt},FIND("/",{irstxt})+1,FIND("/",{irstxt},FIND("/",{irstxt})+1)-FIND("/",{irstxt})-1))'
                    f'),""))'),
        NAKIZLE_C: f'=IF({divf}="","",IF(AND({tasiyici}<>"",{tasiyici}<>"Bayi Aracı"),"Evet","Hayır"))',
        NAKTAM_C: f'=IF({divf}="","",IF(AND({nakizle}="Evet",{tesdurum}="Siparişiniz Teslim Edildi"),"Evet","Hayır"))',
        MONTAM_C: f'=IF({divf}="","",IF({montdurum}="Siparişinizin Montajı Yapıldı","Evet","Hayır"))',
        TETIK_C: (f'=IF({divf}="","",IF({naktam}="Evet",{irstarp},IF({montam}="Evet",{montistp},"")))'),
    }
    for col, formula in formulas.items():
        c = ws.cell(row=row, column=col, value=formula)
        c.font = FORMULA_FONT
        c.border = BORDER
        c.alignment = CENTER
        if col in (SIPTARP_C, TESISTP_C, MONTISTP_C, IRSTARP_C, TETIK_C):
            c.number_format = DATE_FMT

E2E_KEY_RNG = f"'E2E Nakliye Montaj'!${get_column_letter(KEY2_C)}${E2E_FIRST}:${get_column_letter(KEY2_C)}${E2E_LAST}"
E2E_NAKIZLE_RNG = f"'E2E Nakliye Montaj'!${get_column_letter(NAKIZLE_C)}${E2E_FIRST}:${get_column_letter(NAKIZLE_C)}${E2E_LAST}"
E2E_NAKTAM_RNG = f"'E2E Nakliye Montaj'!${get_column_letter(NAKTAM_C)}${E2E_FIRST}:${get_column_letter(NAKTAM_C)}${E2E_LAST}"
E2E_MONTAM_RNG = f"'E2E Nakliye Montaj'!${get_column_letter(MONTAM_C)}${E2E_FIRST}:${get_column_letter(MONTAM_C)}${E2E_LAST}"
E2E_TETIK_RNG = f"'E2E Nakliye Montaj'!${get_column_letter(TETIK_C)}${E2E_FIRST}:${get_column_letter(TETIK_C)}${E2E_LAST}"

print("Part 4 (E2E) done")

# ============================================================
# 6) HAKEDIS FATURA DETAY  (gercek Temmuz 2026 "TEMMUZ DETAY" export yapisi)
# ============================================================
HFD_HEADERS = [
    "Satış id", "Bölge", "Alt Bölge", "Rut", "Bayi Kodu", "Nihai Kod", "Bayi Ünvanı", "Fatura No",
    "Değişen Fatura No", "Ürün Kodu", "LPH 2", "LPH 3", "İade Durumu", "Satış Tipi", "Fatura Tipi",
    "eLogo", "Kime Satış", "Ödeme Yöntemi", "Diva Yaratma Tarihi", "Fatura Tarihi", "Montaj Tarihi",
    "Montaj Fiş No", "Nakliye Tarihi", "Nakliye Kaynağı", "Nakliye Fiş No", "Kampanya Adı 1",
    "Kampanya Türü 1", "Yöntem 1", "Toplam Tutar 1", "Prim Fonundan 1", "Karlılıktan 1",
    "Toplam Tutar H 1",
]
assert len(HFD_HEADERS) == 32
HFD_CALC_HEADERS = [
    "Eşleştirme Anahtarı\n(Ürün Düzeyi)", "Eşleştirme Anahtarı\n(Kampanya Düzeyi)",
    "Kural (Ödeme\nYöntemine Göre)", "Nakliye Tarihi\n(Ayrıştırılmış)", "Hakediş\nTetikleyici Tarih",
]

ws = wb.create_sheet("Hakedis Fatura Detay")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"
set_widths(ws, [11] * 32 + [18, 20, 20, 16, 14])
header_row(ws, 1, HFD_HEADERS, RAW_HEADER_FILL, RAW_HEADER_FONT, height=44)
header_row(ws, 1, HFD_CALC_HEADERS, CALC_HEADER_FILL, HEADER_FONT, height=44, start_col=33)

HFD_FIRST, HFD_LAST = 3, 1500
(HFD_BAYIKODU_C, HFD_FATNO_C, HFD_URUN_C, HFD_ODEME_C, HFD_FATTAR_C, HFD_MONTTAR_C, HFD_MONTFIS_C,
 HFD_NAKTAR_C, HFD_KAMPANYAADI_C, HFD_YONTEM1_C, HFD_TUTARH_C) = 5, 8, 10, 18, 20, 21, 22, 23, 26, 28, 32
HFD_KEY2_C, HFD_KEY_C, HFD_KURAL_C, HFD_NAKTARP_C, HFD_TETIK_C = 33, 34, 35, 36, 37

note = ws.cell(row=1, column=39,
    value="Kaynak: BSH'nin bayiye gönderdiği aylık hakediş detay raporu (\"... DETAY\" sekmesi). Bir ürün "
          "aynı anda birden fazla kampanyaya konu olabildiğinden (ör. tekil + bundle), aynı Fatura No + Ürün "
          "Kodu için birden fazla satır olabilir; her satır 'Kampanya Adı 1' ile ayrı bir kampanya uygulamasını "
          "temsil eder. Bu yüzden eşleştirme Fatura No + Ürün Kodu + Kampanya Adı üçlüsüyle yapılır.")
note.font = NOTE_FONT
note.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[1].height = 44

example_hfd = ["7000129546|2600007552", "TRBB", "TB3", "000TRRBB32", 7000129546, 7000129546,
    "ÖRNEK BAYİ ÜNVANI LTD.ŞTİ.", 2600007552, "-", "SMS4IKW62T", "LDA", "Bosch-LDA-BULAŞIK MAKİNESİ",
    "Geçerli", "Hemen Teslim", "e-Arşiv", "1", "Tüketici", "Mutabakat Yoksa",
    "30.07.2026", "30.07.2026", "-", "-", "2026-07-31", "e2e", "TR01-20674473",
    "Temmuz-Tekil Bulaşık Makinesi Kampanyası-202607", "Muhtelif Ürünler Fiyat Farkı",
    "Montaj veya Nakliye", 3777.92, 0, 3777.92, 3777.92]
for i, val in enumerate(example_hfd, start=1):
    c = ws.cell(row=2, column=i, value=val)
    c.font = EXAMPLE_FONT
    c.fill = EXAMPLE_FILL
    c.border = BORDER
    c.alignment = CENTER
    if i in (HFD_FATTAR_C, HFD_MONTTAR_C):
        c.number_format = DATE_FMT
    if i in (29, 30, 31, HFD_TUTARH_C):
        c.number_format = CUR

dv_odeme = DataValidation(type="list", formula1='"Mutabakat Var,Mutabakat Yoksa"', allow_blank=True)
ws.add_data_validation(dv_odeme)
dv_odeme.add(f"{get_column_letter(HFD_ODEME_C)}{HFD_FIRST}:{get_column_letter(HFD_ODEME_C)}{HFD_LAST}")
dv_yontem = DataValidation(type="list", formula1='"Diva Fatura,Montaj veya Nakliye"', allow_blank=True)
ws.add_data_validation(dv_yontem)
dv_yontem.add(f"{get_column_letter(HFD_YONTEM1_C)}{HFD_FIRST}:{get_column_letter(HFD_YONTEM1_C)}{HFD_LAST}")

for row in range(HFD_FIRST, HFD_LAST + 1):
    for col in range(1, 33):
        c = ws.cell(row=row, column=col)
        c.fill = INPUT_FILL
        c.font = INPUT_FONT
        c.border = BORDER
        c.alignment = CENTER
        if col in (HFD_FATTAR_C, HFD_MONTTAR_C):
            c.number_format = DATE_FMT
        if col in (29, 30, 31, HFD_TUTARH_C):
            c.number_format = CUR

    L = get_column_letter
    fno = f"{L(HFD_FATNO_C)}{row}"
    urun = f"{L(HFD_URUN_C)}{row}"
    odeme = f"{L(HFD_ODEME_C)}{row}"
    fattar = f"{L(HFD_FATTAR_C)}{row}"
    monttar = f"{L(HFD_MONTTAR_C)}{row}"
    naktar = f"{L(HFD_NAKTAR_C)}{row}"
    kampanya = f"{L(HFD_KAMPANYAADI_C)}{row}"
    yontem1 = f"{L(HFD_YONTEM1_C)}{row}"
    naktarp = f"{L(HFD_NAKTARP_C)}{row}"

    formulas = {
        HFD_KEY2_C: f'=IF({fno}="","",TEXT({fno},"0")&"|"&{urun})',
        HFD_KEY_C: f'=IF({fno}="","",TEXT({fno},"0")&"|"&{urun}&"|"&{kampanya})',
        HFD_KURAL_C: (f'=IF({fno}="","",IF({odeme}="Mutabakat Yoksa",'
                      f'"Kural B - Nakliye/Montaj Sonrası Ödeme","Kural A - Hemen Ödeme (Stok Mutabakatlı)"))'),
        HFD_NAKTARP_C: (f'=IF(OR({naktar}="",{naktar}="-"),"",IFERROR(DATE(VALUE(LEFT({naktar},4)),'
                        f'VALUE(MID({naktar},6,2)),VALUE(RIGHT({naktar},2))),""))'),
        HFD_TETIK_C: (f'=IF({fno}="","",IF({yontem1}="Diva Fatura",{fattar},'
                      f'IF({naktarp}<>"",{naktarp},IF(ISNUMBER({monttar}),{monttar},""))))'),
    }
    for col, formula in formulas.items():
        c = ws.cell(row=row, column=col, value=formula)
        c.font = FORMULA_FONT
        c.border = BORDER
        c.alignment = CENTER
        if col in (HFD_NAKTARP_C, HFD_TETIK_C):
            c.number_format = DATE_FMT

HFD_KEY_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_KEY_C)}${HFD_FIRST}:${get_column_letter(HFD_KEY_C)}${HFD_LAST}"
HFD_KEY2_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_KEY2_C)}${HFD_FIRST}:${get_column_letter(HFD_KEY2_C)}${HFD_LAST}"
HFD_ODEME_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_ODEME_C)}${HFD_FIRST}:${get_column_letter(HFD_ODEME_C)}${HFD_LAST}"
HFD_KURAL_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_KURAL_C)}${HFD_FIRST}:${get_column_letter(HFD_KURAL_C)}${HFD_LAST}"
HFD_YONTEM1_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_YONTEM1_C)}${HFD_FIRST}:${get_column_letter(HFD_YONTEM1_C)}${HFD_LAST}"
HFD_TETIK_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_TETIK_C)}${HFD_FIRST}:${get_column_letter(HFD_TETIK_C)}${HFD_LAST}"
HFD_TUTARH_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_TUTARH_C)}${HFD_FIRST}:${get_column_letter(HFD_TUTARH_C)}${HFD_LAST}"

print("Part 5 (Hakedis Fatura Detay) done")

# ============================================================
# 7) KONTROL  (otomatik - tum kaynaklari birlestirir)
# ============================================================
ws = wb.create_sheet("Kontrol")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"

KONTROL_COLS = [
    "Fatura No", "Ürün Kodu", "Ad Soyad", "Fatura\nTarihi", "Paket Adı\n(Kampanya)",
    "Kampanya\nBulundu mu", "Kampanya\nAralığında mı", "Hakediş\nSon Tarihi",
    "Beklenen\nHakediş Tutarı", "Uygulanacak Kural\n(Ön Beklenti)", "Hakediş\nFaturasında\nBulundu mu",
    "HFD Ödeme\nYöntemi", "HFD'ye Göre\nKural", "HFD\nYöntem 1", "HFD Hakediş\nTetikleyici Tarih",
    "Süre İçinde\nTamamlandı mı", "E2E: Nakliye\nTamamlandı mı", "E2E: Montaj\nTamamlandı mı",
    "E2E-HFD\nTutarlılığı", "Faturada\nBulunan Tutar", "Tutar\nFarkı", "Tutar\nDurumu", "GENEL DURUM",
    "Sorun Raporu\nSıra No\n(yardımcı)",
]
set_widths(ws, [13, 12, 16, 11, 26, 12, 12, 11, 13, 18, 13, 13, 16, 13, 14, 12, 12, 12, 14, 13, 11, 15, 24, 12])
header_row(ws, 1, KONTROL_COLS, CALC_HEADER_FILL, HEADER_FONT, height=48)

K_FIRST, K_LAST = DS_FIRST, DS_LAST  # 1:1 aligned with DIVA Satislari rows
(A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X) = range(1, 25)

L_ = get_column_letter
for row in range(K_FIRST, K_LAST + 1):
    ds = lambda col: f"'DIVA Satislari'!{L_(col)}{row}"
    ds_fno = ds(FATURANO_C)
    ds_urun = ds(URUNKODU_C)
    ds_ad = ds(13)
    ds_tar = ds(FATURATARIHI_C)
    ds_paket = ds(PAKETADI_C)
    ds_bulundu = ds(KAMP_BULUNDU_C)
    ds_aralik = ds(ARALIK_C)
    ds_son = ds(KSON_C)
    ds_beklenen = ds(BEKLENEN_C)
    ds_kural = ds(KURAL_C)
    ds_key = ds(KEY_C)     # kampanya duzeyi -> Hakedis Fatura Detay eslestirmesi
    ds_key2 = ds(DSKEY2_C)   # urun duzeyi -> E2E eslestirmesi

    a = f"{L_(A)}{row}"; f = f"{L_(F)}{row}"; g = f"{L_(G)}{row}"; h = f"{L_(H)}{row}"
    i_ = f"{L_(I)}{row}"; k = f"{L_(K)}{row}"; n = f"{L_(N)}{row}"; o = f"{L_(O)}{row}"
    p = f"{L_(P)}{row}"; q = f"{L_(Q)}{row}"; r_ = f"{L_(R)}{row}"; s_ = f"{L_(S)}{row}"
    t = f"{L_(T)}{row}"; u = f"{L_(U)}{row}"; v = f"{L_(V)}{row}"; w = f"{L_(W)}{row}"
    w_grow = f"${L_(W)}$3:{L_(W)}{row}"

    match_e2e = f'MATCH({ds_key2},{E2E_KEY_RNG},0)'
    match_hfd = f'MATCH({ds_key},{HFD_KEY_RNG},0)'
    cells = {
        A: (f'=IF({ds_fno}="","",{ds_fno})', None),
        B: (f'=IF({ds_fno}="","",{ds_urun})', None),
        C: (f'=IF({ds_fno}="","",{ds_ad})', None),
        D: (f'=IF({ds_fno}="","",{ds_tar})', DATE_FMT),
        E: (f'=IF({ds_fno}="","",{ds_paket})', None),
        F: (f'=IF({ds_fno}="","",{ds_bulundu})', None),
        G: (f'=IF({ds_fno}="","",{ds_aralik})', None),
        H: (f'=IF({ds_fno}="","",{ds_son})', DATE_FMT),
        I: (f'=IF({ds_fno}="","",{ds_beklenen})', CUR),
        J: (f'=IF({ds_fno}="","",{ds_kural})', None),
        K: (f'=IF({ds_fno}="","",IF(ISNA({match_hfd}),"Hayır - Henüz Faturalanmamış","Evet"))', None),
        L: (f'=IF({k}<>"Evet","",IFERROR(INDEX({HFD_ODEME_RNG},{match_hfd}),""))', None),
        M: (f'=IF({k}<>"Evet","",IFERROR(INDEX({HFD_KURAL_RNG},{match_hfd}),""))', None),
        N: (f'=IF({k}<>"Evet","",IFERROR(INDEX({HFD_YONTEM1_RNG},{match_hfd}),""))', None),
        O: (f'=IF({k}<>"Evet","",IFERROR(INDEX({HFD_TETIK_RNG},{match_hfd}),""))', DATE_FMT),
        P: (f'=IF(OR({k}<>"Evet",{o}="",{h}=""),"",IF({o}<={h},"Evet","HAYIR - RİSK"))', None),
        Q: (f'=IF({a}="","",IFERROR(INDEX({E2E_NAKTAM_RNG},{match_e2e}),"Hayır"))', None),
        R: (f'=IF({a}="","",IFERROR(INDEX({E2E_MONTAM_RNG},{match_e2e}),"Hayır"))', None),
        S: (f'=IF(OR({k}<>"Evet",{n}=""),"",IF({n}="Montaj veya Nakliye",'
            f'IF(OR({q}="Evet",{r_}="Evet"),"Tutarlı","UYUŞMUYOR - E2E\'de Tamamlanma Yok"),'
            f'"N/A - Hemen Ödeme Yöntemi"))', None),
        T: (f'=IF({a}="","",SUMIF({HFD_KEY_RNG},{ds_key},{HFD_TUTARH_RNG}))', CUR),
        U: (f'=IF({a}="","",{i_}-{t})', CUR),
        V: (f'=IF({a}="","",IF({k}<>"Evet","Henüz Faturalanmadı",'
            f'IF({u}=0,"Doğru",IF({u}>0,"EKSİK ÖDENMİŞ","FAZLA ÖDENMİŞ"))))', None),
        W: (f'=IF({a}="","",IF({f}<>"Evet","İNCELEME GEREKLİ - Kampanya Tanımsız",'
            f'IF({g}<>"Evet","İNCELEME GEREKLİ - Tarih Dışı Satış",'
            f'IF({k}<>"Evet","Bekleniyor - Henüz Faturalanmadı",'
            f'IF({p}="HAYIR - RİSK","İNCELEME GEREKLİ - Hakediş Süresi Aşıldı",'
            f'IF(LEFT({s_},9)="UYUŞMUYOR","İNCELEME GEREKLİ - Nakliye/Montaj Tutarsızlığı",'
            f'IF(OR({v}="EKSİK ÖDENMİŞ",{v}="FAZLA ÖDENMİŞ"),"İNCELEME GEREKLİ - Tutar Uyuşmazlığı",'
            f'"Uygun")))))))', None),
        X: (f'=IF(OR({w}="",{w}="Uygun"),"",SUMPRODUCT(({w_grow}<>"Uygun")*({w_grow}<>"")*1))', None),
    }
    for col, (formula, fmt) in cells.items():
        c = ws.cell(row=row, column=col, value=formula)
        c.font = FORMULA_FONT
        c.border = BORDER
        c.alignment = CENTER
        if fmt:
            c.number_format = fmt
        if col == W:
            c.font = Font(name=FONT, size=9, bold=True)

K_A_RNG = f"Kontrol!${L_(A)}${K_FIRST}:${L_(A)}${K_LAST}"
K_B_RNG = f"Kontrol!${L_(B)}${K_FIRST}:${L_(B)}${K_LAST}"
K_C_RNG = f"Kontrol!${L_(C)}${K_FIRST}:${L_(C)}${K_LAST}"
K_D_RNG = f"Kontrol!${L_(D)}${K_FIRST}:${L_(D)}${K_LAST}"
K_E_RNG = f"Kontrol!${L_(E)}${K_FIRST}:${L_(E)}${K_LAST}"
K_K_RNG = f"Kontrol!${L_(K)}${K_FIRST}:${L_(K)}${K_LAST}"
K_F_RNG = f"Kontrol!${L_(F)}${K_FIRST}:${L_(F)}${K_LAST}"
K_I_RNG = f"Kontrol!${L_(I)}${K_FIRST}:${L_(I)}${K_LAST}"
K_J_RNG = f"Kontrol!${L_(J)}${K_FIRST}:${L_(J)}${K_LAST}"
K_Q_RNG = f"Kontrol!${L_(Q)}${K_FIRST}:${L_(Q)}${K_LAST}"
K_R_RNG = f"Kontrol!${L_(R)}${K_FIRST}:${L_(R)}${K_LAST}"
K_T_RNG = f"Kontrol!${L_(T)}${K_FIRST}:${L_(T)}${K_LAST}"
K_U_RNG = f"Kontrol!${L_(U)}${K_FIRST}:${L_(U)}${K_LAST}"
K_V_RNG = f"Kontrol!${L_(V)}${K_FIRST}:${L_(V)}${K_LAST}"
K_W_RNG = f"Kontrol!${L_(W)}${K_FIRST}:${L_(W)}${K_LAST}"
K_X_RNG = f"Kontrol!${L_(X)}${K_FIRST}:${L_(X)}${K_LAST}"

print("Part 6 (Kontrol) done")

# ============================================================
# 7b) SORUN RAPORU  (Kontrol'deki "Uygun" olmayan tum satirlarin ozetlenmis listesi)
# ============================================================
ws = wb.create_sheet("Sorun Raporu")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"

SR_COLS = [
    "Sıra", "Fatura No", "Ürün Kodu", "Ad Soyad", "Fatura\nTarihi", "Paket Adı\n(Kampanya)",
    "Beklenen\nHakediş Tutarı", "Hakediş\nFaturasında\nBulundu mu", "Faturada\nBulunan Tutar",
    "Tutar\nFarkı", "SORUN / DURUM",
]
set_widths(ws, [7, 13, 12, 16, 11, 26, 13, 13, 13, 11, 30])
header_row(ws, 1, SR_COLS, WARN_FILL, Font(name=FONT, size=9, bold=True, color="C00000"), height=44)

note = ws.cell(row=1, column=13,
    value="Bu liste, 'Kontrol' sekmesinde GENEL DURUM sütunu 'Uygun' OLMAYAN (henüz faturalanmamış, "
          "kampanya/tarih dışı, süresi aşmış, tutar uyuşmazlığı veya nakliye-montaj tutarsızlığı olan) "
          "tüm satırları otomatik listeler. Kaynak veriler güncellendikçe bu sayfa da otomatik güncellenir.")
note.font = NOTE_FONT
note.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[1].height = 44

SR_FIRST, SR_LAST = 3, K_LAST - K_FIRST + 3  # ayni kapasite (Kontrol satir sayisi kadar)
(SA, SB, SC, SD, SE, SF, SG, SH, SI, SJ, SK) = range(1, 12)
for row in range(SR_FIRST, SR_LAST + 1):
    rank = row - (SR_FIRST - 1)
    ca = ws.cell(row=row, column=SA, value=rank)
    ca.font = FORMULA_FONT
    ca.border = BORDER
    ca.alignment = CENTER
    sa = f"{L_(SA)}{row}"
    match_rank = f'MATCH({sa},{K_X_RNG},0)'
    cells = {
        SB: (f'=IFERROR(INDEX({K_A_RNG},{match_rank}),"")', None),
        SC: (f'=IFERROR(INDEX({K_B_RNG},{match_rank}),"")', None),
        SD: (f'=IFERROR(INDEX({K_C_RNG},{match_rank}),"")', None),
        SE: (f'=IFERROR(INDEX({K_D_RNG},{match_rank}),"")', DATE_FMT),
        SF: (f'=IFERROR(INDEX({K_E_RNG},{match_rank}),"")', None),
        SG: (f'=IFERROR(INDEX({K_I_RNG},{match_rank}),"")', CUR),
        SH: (f'=IFERROR(INDEX({K_K_RNG},{match_rank}),"")', None),
        SI: (f'=IFERROR(INDEX({K_T_RNG},{match_rank}),"")', CUR),
        SJ: (f'=IFERROR(INDEX({K_U_RNG},{match_rank}),"")', CUR),
        SK: (f'=IFERROR(INDEX({K_W_RNG},{match_rank}),"")', None),
    }
    for col, (formula, fmt) in cells.items():
        c = ws.cell(row=row, column=col, value=formula)
        c.font = FORMULA_FONT
        c.border = BORDER
        c.alignment = CENTER
        if fmt:
            c.number_format = fmt
        if col == SK:
            c.font = Font(name=FONT, size=9, bold=True, color="C00000")
        if col == SA:
            c.font = FORMULA_FONT

print("Part 6b (Sorun Raporu) done")

# ============================================================
# 8) YIL SONU KONTROLU  (Kural A bayileri icin %75 nakliye/montaj orani)
# ============================================================
ws = wb.create_sheet("Yil Sonu Kontrolu")
ws.sheet_view.showGridLines = False
set_widths(ws, [3, 52, 20, 3])
ws["B2"] = "YIL SONU %75 NAKLİYE/MONTAJ ORANI KONTROLÜ (Kural A - Stok Mutabakatlı Bayi)"
ws["B2"].font = Font(name=FONT, size=13, bold=True, color="1F4E78")
ws.merge_cells("B2:C2")

info = ws.cell(row=3, column=2,
    value="Bu kontrol yalnızca Parametreler sekmesinde 'Stok Mutabakatı Yapıldı mı?' = Evet olan bayiler için "
          "anlamlıdır. Kampanya yazısına göre: 30.11 itibarıyla yıl içinde hakediş ödenen ürünlerin nakliye VEYA "
          "montaj tamamlanma oranı en az %75 olmalıdır; altında kalırsa fark mahsuplaşılır.")
info.font = NOTE_FONT
info.alignment = Alignment(wrap_text=True)
ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=3)
ws.row_dimensions[3].height = 45

rows = [
    ("Kural A Kapsamında Hakediş Ödenen Satış Sayısı",
     f'=SUMPRODUCT((LEFT({K_J_RNG},8)="Kural A ")*({K_A_RNG}<>""))', None),
    ("Bunlardan Nakliye VEYA Montaj Tamamlanan Sayısı",
     f'=SUMPRODUCT((LEFT({K_J_RNG},8)="Kural A ")*((({K_Q_RNG}="Evet")+({K_R_RNG}="Evet"))>0))', None),
    ("Nakliye/Montaj Tamamlanma Oranı", None, "0.0%"),
    ("Minimum Eşik", f'={P_ESIK}', "0.0%"),
    ("Eşik Karşılanıyor mu", None, None),
    ("Kural A Kapsamında Toplam Beklenen Hakediş Tutarı",
     f'=SUMPRODUCT((LEFT({K_J_RNG},8)="Kural A ")*({K_I_RNG}))', CUR),
    ("Eşik Altı Kalan Oran Farkı", None, "0.0%"),
    ("Tahmini Mahsuplaşma Tutarı (oran farkı x toplam hakediş - yaklaşık tahmindir)", None, CUR),
]
r = 5
row_refs = {}
for label, formula, fmt in rows:
    ws.cell(row=r, column=2, value=label).font = LABEL_FONT
    ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 30
    c = ws.cell(row=r, column=3, value=formula)
    c.font = Font(name=FONT, size=11, bold=True)
    c.border = BORDER
    c.alignment = Alignment(horizontal="center")
    if fmt:
        c.number_format = fmt
    row_refs[label] = r
    r += 1

R_TOPLAM = row_refs["Kural A Kapsamında Hakediş Ödenen Satış Sayısı"]
R_TAMAM = row_refs["Bunlardan Nakliye VEYA Montaj Tamamlanan Sayısı"]
R_ORAN = row_refs["Nakliye/Montaj Tamamlanma Oranı"]
R_ESIK = row_refs["Minimum Eşik"]
R_KARSIL = row_refs["Eşik Karşılanıyor mu"]
R_TUTAR = row_refs["Kural A Kapsamında Toplam Beklenen Hakediş Tutarı"]
R_FARK = row_refs["Eşik Altı Kalan Oran Farkı"]
R_MAHSUP = row_refs["Tahmini Mahsuplaşma Tutarı (oran farkı x toplam hakediş - yaklaşık tahmindir)"]

ws.cell(row=R_ORAN, column=3, value=f'=IF(C{R_TOPLAM}=0,"",C{R_TAMAM}/C{R_TOPLAM})')
c = ws.cell(row=R_KARSIL, column=3, value=f'=IF(C{R_TOPLAM}=0,"",IF(C{R_ORAN}>=C{R_ESIK},"EVET - Uygun","HAYIR - Mahsuplaşma Riski"))')
c.number_format = "General"
ws.cell(row=R_FARK, column=3, value=f'=IF(C{R_TOPLAM}=0,0,IF(C{R_ORAN}<C{R_ESIK},C{R_ESIK}-C{R_ORAN},0))')
ws.cell(row=R_MAHSUP, column=3, value=f'=C{R_FARK}*C{R_TUTAR}')

for lbl in ("Eşik Karşılanıyor mu",):
    ws.cell(row=row_refs[lbl], column=3).fill = WARN_FILL
    ws.cell(row=row_refs[lbl], column=2).fill = WARN_FILL

note2 = ws.cell(row=r + 1, column=2,
    value="Not: Mahsuplaşma tutarı, kampanya yazısındaki hesaplama yöntemi netleşinceye kadar basit bir "
          "yaklaşıklıkla (oran açığı x toplam beklenen hakediş) hesaplanmıştır; gerçek mahsuplaşma yöntemi "
          "BSH finans ekibi ile teyit edilip formül buna göre güncellenmelidir.")
note2.font = NOTE_FONT
note2.alignment = Alignment(wrap_text=True)
ws.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=3)
ws.row_dimensions[r + 1].height = 40

print("Part 7 (Yil Sonu Kontrolu) done")

# ============================================================
# 9) OZET
# ============================================================
ws = wb.create_sheet("Ozet")
ws.sheet_view.showGridLines = False
set_widths(ws, [3, 48, 18, 3])
ws["B2"] = "ÖZET"
ws["B2"].font = Font(name=FONT, size=14, bold=True, color="1F4E78")

summary = [
    ("Toplam Satış Sayısı", f'=COUNTIF({K_A_RNG},"<>")', None, False),
    ("Kural A (Hemen Ödeme) Satış Sayısı", f'=COUNTIF({K_J_RNG},"Kural A*")', None, False),
    ("Kural B (Nakliye/Montaj Sonrası) Satış Sayısı", f'=COUNTIF({K_J_RNG},"Kural B*")', None, False),
    ("Toplam Beklenen Hakediş Tutarı", f'=SUMIF({K_A_RNG},"<>",{K_I_RNG})', CUR, False),
    ("Faturada Bulunan Toplam Tutar", f'=SUMIF({K_A_RNG},"<>",{K_T_RNG})', CUR, False),
    ("Toplam Tutar Farkı", f'=SUMIF({K_A_RNG},"<>",{K_U_RNG})', CUR, False),
    ("Eksik Ödenmiş Satış Sayısı", f'=COUNTIF({K_V_RNG},"EKSİK ÖDENMİŞ")', None, True),
    ("Fazla Ödenmiş Satış Sayısı", f'=COUNTIF({K_V_RNG},"FAZLA ÖDENMİŞ")', None, True),
    ("Kampanya Tanımsız (Kampanyalar Sekmesinde Yok)", f'=COUNTIF({K_W_RNG},"İNCELEME GEREKLİ - Kampanya Tanımsız")', None, True),
    ("Tarih Dışı Satış Sayısı", f'=COUNTIF({K_W_RNG},"İNCELEME GEREKLİ - Tarih Dışı Satış")', None, True),
    ("Hakediş Süresi Aşılan (RİSK) Satış Sayısı", f'=COUNTIF({K_W_RNG},"İNCELEME GEREKLİ - Hakediş Süresi Aşıldı")', None, True),
    ("Nakliye/Montaj Tutarsızlığı Olan Satış Sayısı", f'=COUNTIF({K_W_RNG},"İNCELEME GEREKLİ - Nakliye/Montaj Tutarsızlığı")', None, True),
    ("Henüz Faturalanmamış (Bekleyen) Satış Sayısı", f'=COUNTIF({K_W_RNG},"Bekleniyor*")', None, False),
    ("İncelenmesi Gereken Toplam Satış Sayısı", f'=SUMPRODUCT(--(LEFT({K_W_RNG},16)="İNCELEME GEREKLİ"))', None, True),
    ("Uygun / Sorunsuz Satış Sayısı", f'=COUNTIF({K_W_RNG},"Uygun")', None, False),
]
r = 4
for label, formula, fmt, warn in summary:
    lc = ws.cell(row=r, column=2, value=label)
    lc.font = LABEL_FONT
    lc.alignment = Alignment(wrap_text=True, vertical="center")
    c = ws.cell(row=r, column=3, value=formula)
    c.font = Font(name=FONT, size=11, bold=True)
    c.border = BORDER
    c.alignment = Alignment(horizontal="center")
    if fmt:
        c.number_format = fmt
    if warn:
        lc.fill = WARN_FILL
        c.fill = WARN_FILL
    ws.row_dimensions[r].height = 26
    r += 1

r += 1
note = ws.cell(row=r, column=2,
    value="Not: Bu tablo 'Kontrol' sekmesindeki tüm satırları otomatik özetler; 'Yıl Sonu Kontrolü' sekmesi ise "
          "Kural A bayileri için ayrıca %75 nakliye/montaj oranı kontrolünü gösterir. Kaynak sekmeleri "
          "güncelledikçe bu sayfa otomatik yenilenir.")
note.font = NOTE_FONT
note.alignment = Alignment(wrap_text=True)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
ws.row_dimensions[r].height = 40

print("Part 8 (Ozet) done")

# ============================================================
# 10) SEKME KORUMASI VE DUZEN
# ============================================================
ADMIN_PASSWORD = "BSHKampanya2026!"
STRUCT_PASSWORD = "BSHYapi2026!"

kampanyalar_ws = wb["Kampanyalar"]
for row in kampanyalar_ws.iter_rows(min_row=1, max_row=KAMP_LAST + 3, max_col=11):
    for cell in row:
        cell.protection = openpyxl.styles.Protection(locked=True)
kampanyalar_ws.protection.sheet = True
kampanyalar_ws.protection.password = ADMIN_PASSWORD
kampanyalar_ws.protection.formatCells = True
kampanyalar_ws.protection.formatColumns = True
kampanyalar_ws.protection.formatRows = True
kampanyalar_ws.protection.insertRows = True
kampanyalar_ws.protection.insertColumns = False
kampanyalar_ws.protection.deleteRows = True
kampanyalar_ws.protection.sort = True
kampanyalar_ws.protection.autoFilter = True

def protect_formula_sheet(ws_obj, max_row, max_col, unlock_ranges=()):
    unlocked = openpyxl.styles.Protection(locked=False)
    locked = openpyxl.styles.Protection(locked=True)
    for row in ws_obj.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            is_unlocked = any(
                lo <= cell.row <= hi and c1 <= cell.column <= c2
                for (lo, hi, c1, c2) in unlock_ranges
            )
            cell.protection = unlocked if is_unlocked else locked
    ws_obj.protection.sheet = True
    ws_obj.protection.password = STRUCT_PASSWORD
    ws_obj.protection.insertRows = False
    ws_obj.protection.deleteRows = False

# DIVA Satislari: kolon 1-35 (ham veri girisi) acik, 36-45 (formul) kilitli
protect_formula_sheet(wb["DIVA Satislari"], DS_LAST, KURAL_C, unlock_ranges=[(2, DS_LAST, 1, 35)])
# E2E: kolon 1-38 acik, 39-48 kilitli
protect_formula_sheet(wb["E2E Nakliye Montaj"], E2E_LAST, 48, unlock_ranges=[(2, E2E_LAST, 1, 38)])
# Hakedis Fatura Detay: kolon 1-32 (ham veri girisi) acik, 33-37 (formul) kilitli
protect_formula_sheet(wb["Hakedis Fatura Detay"], HFD_LAST, HFD_TETIK_C, unlock_ranges=[(2, HFD_LAST, 1, 32)])
# Kontrol, Sorun Raporu, Yil Sonu Kontrolu, Ozet: tamamen formul, hicbir yer acik degil
protect_formula_sheet(wb["Kontrol"], K_LAST, X)
protect_formula_sheet(wb["Sorun Raporu"], SR_LAST, SK)
protect_formula_sheet(wb["Yil Sonu Kontrolu"], 40, 4)
protect_formula_sheet(wb["Ozet"], 40, 4)
# Parametreler: sadece C sutunundaki giris hucreleri acik
protect_formula_sheet(wb["Parametreler"], 20, 4, unlock_ranges=[(4, 20, 3, 3)])

wb.security = openpyxl.workbook.protection.WorkbookProtection(
    workbookPassword=ADMIN_PASSWORD, lockStructure=True, lockWindows=False
)

ORDER = ["Talimatlar", "Parametreler", "Kampanyalar", "DIVA Satislari", "E2E Nakliye Montaj",
         "Hakedis Fatura Detay", "Kontrol", "Sorun Raporu", "Yil Sonu Kontrolu", "Ozet"]
wb._sheets = [wb[name] for name in ORDER]
wb.active = 0

TAB_COLORS = {
    "Talimatlar": "1F4E78", "Parametreler": "1F4E78", "Kampanyalar": "BF8F00",
    "DIVA Satislari": "2E75B6", "E2E Nakliye Montaj": "2E75B6", "Hakedis Fatura Detay": "2E75B6",
    "Kontrol": "548235", "Sorun Raporu": "C00000", "Yil Sonu Kontrolu": "C00000", "Ozet": "548235",
}
for name, color in TAB_COLORS.items():
    wb[name].sheet_properties.tabColor = color

OUT = os.path.join(SCRIPT_DIR, "Bayi_DIVA_Kampanya_Hakedis_Kontrol.xlsx")
wb.save(OUT)
print("SAVED", OUT)
print("ADMIN_PASSWORD:", ADMIN_PASSWORD)
print("STRUCT_PASSWORD:", STRUCT_PASSWORD)
