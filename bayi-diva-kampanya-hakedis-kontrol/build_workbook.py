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
    ("Ağustos-Tekil Soğutucu Kampanyası-202608", "082026", "Tekil", "Soğutucu/Buzdolabı-Dondurucu",
     "01.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: 18-31 Ağustos 2026 Perakende Uygulama Yazısı - Tekil Soğutucu Kampanyası. Hakediş tutarı DİVA'daki "
     "'Toplam İndirim' esas alınarak hesaplanır."),
    ("Ağustos-Tekil Bulaşık Makinesi Kampanyası-202608", "082026", "Tekil", "Bulaşık Makinesi",
     "01.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Ağustos 2026 Uygulama Yazısı - Tekil Bulaşık Makinesi Kampanyası."),
    ("Ağustos-Tekil Çamaşır Kampanyası-202608", "082026", "Tekil", "Çamaşır Makinesi",
     "01.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Ağustos 2026 Uygulama Yazısı - Tekil Çamaşır Makinesi Kampanyası."),
    ("Ağustos-Tekil Kurutma Makinesi Kampanyası-202608", "082026", "Tekil", "Kurutma Makinesi",
     "01.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Ağustos 2026 Uygulama Yazısı - Kurutma Makinesi Kampanyası."),
    ("Ağustos-Tekil Klima Kampanyası-202608", "082026", "Tekil", "Klima",
     "01.08.2026", "31.08.2026", "30.11.2026", None, None, "Evet",
     "Kaynak: Ağustos 2026 Uygulama Yazısı - Klima Kampanyası."),
    ("Ağustos-Ankastre Set Alana 5000 TL İndirim-202608", "082026", "Bundle-Set", "Ankastre Fırın+Ocak/Davlumbaz Set",
     "01.08.2026", "31.08.2026", "30.11.2026", 3700, 2500, "Tekil Evet / Bundle Hayır",
     "Kaynak: Ankastre Set Fırsat 1 (5.000 TL tüketici indirimi / 3.700 TL bayi hakediş). Tüketici indirimi ile "
     "bayi hakedişi FARKLIDIR - kontrol formülü sabit tutarı esas alır."),
    ("Ağustos-Kurutma+Bulaşık Alımına 5000 TL İndirim-202608", "082026", "Bundle-Set", "Kurutma+Bulaşık/Set",
     "01.08.2026", "31.08.2026", "30.11.2026", 4000, 4000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Ağustos Ayı 2'li Beyaz Eşya Fırsatı ailesi (5.000 TL tüketici indirimi / 4.000 TL bayi hakediş, "
     "4.000 adet ile TÜM 2'li set kampanyaları ortak sınırlıdır). TEYİT EDİLMELİ."),
    ("Ağustos Seçili Soğutucu + Çamaşır Alımına 5000 TL İndirim-202608", "082026", "Bundle-Set", "Soğutucu+Çamaşır/Set",
     "01.08.2026", "31.08.2026", "30.11.2026", 4000, 4000, "Tekil Evet / Bundle Hayır",
     "Kaynak: Ağustos Ayı 2'li Beyaz Eşya Fırsatı ailesi. TEYİT EDİLMELİ - tutar/adet limiti admin tarafından "
     "kesinleştirilecektir."),
    ("Ağustos Ayı Fırsat Ürünleri Kampanyası-202608", "082026", "Tekil", "Fırsat Ürünleri (Hareket Hızı Düşük)",
     "01.08.2026", "31.08.2026", "30.09.2026", None, None, "Evet",
     "Kaynak: Ağustos Ayı Fırsat Ürünleri Kampanyası. DİKKAT: Hakediş son tarihi diğer kampanyalardan farklı "
     "olarak 30.09.2026'dır."),
    ("Stanley Mug Hediyeli Espresso Makineleri Kampanyası-202608", "082026", "Küçük Ev Aletleri (Hediye)",
     "Tam Otomatik Espresso Makinesi", "18.08.2026", "31.08.2026", "30.09.2026", 0, 50, "Evet",
     "Kaynak: Stanley Mug Hediyeli Espresso Makineleri. Fiyat farkı hakedişi YOK; kampanya şartı sağlanan montajı "
     "takip eden hafta BSH tarafından hediye kargolanır. Bu satır sadece takip amaçlıdır, tutar 0 girilmiştir."),
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

KAMP_FIRST, KAMP_LAST = 2, 150
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
    "Eşleştirme\nAnahtarı", "Kampanya\nBulundu mu", "Kampanya\nBaşlangıç", "Kampanya\nBitiş",
    "Hakediş\nSon Tarihi", "Fatura Tarihi\nKampanya\nAralığında mı", "Sabit Hakediş\nTutarı (varsa)",
    "Beklenen\nHakediş Tutarı", "Uygulanacak\nKural",
]

ws = wb.create_sheet("DIVA Satislari")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"
set_widths(ws, [11] * 35 + [16, 12, 11, 11, 11, 12, 13, 13, 20])
header_row(ws, 1, DIVA_HEADERS, RAW_HEADER_FILL, RAW_HEADER_FONT, height=44)
header_row(ws, 1, DIVA_CALC_HEADERS, CALC_HEADER_FILL, HEADER_FONT, height=44, start_col=36)

DS_FIRST, DS_LAST = 3, 1002
FATURANO_C, URUNKODU_C, FATURATARIHI_C, PAKETADI_C, TOPLAMINDIRIM_C = 9, 18, 8, 15, 30
KEY_C = 36
KAMP_BULUNDU_C, KBAS_C, KBIT_C, KSON_C, ARALIK_C, SABIT_C, BEKLENEN_C, KURAL_C = 37, 38, 39, 40, 41, 42, 43, 44

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
        KEY_C: f'=IF({fno}="","",{fno}&"|"&{urun})',
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
# 6) HAKEDIS FATURA DETAY  (placeholder yapı - gerçek örnek gelince kesinleşecek)
# ============================================================
ws = wb.create_sheet("Hakedis Fatura Detay")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"
hfd_headers = ["Hakediş Fatura No", "Fatura Tarihi", "DİVA Fatura No\n(26...)", "Ürün Kodu",
               "Kampanya / Paket Adı\n(bilgi amaçlı)", "Kalem Türü", "Tutar"]
hfd_calc = ["Eşleştirme\nAnahtarı"]
set_widths(ws, [16, 13, 14, 12, 26, 14, 13, 16])
header_row(ws, 1, hfd_headers, RAW_HEADER_FILL, RAW_HEADER_FONT, height=36)
header_row(ws, 1, hfd_calc, CALC_HEADER_FILL, HEADER_FONT, height=36, start_col=len(hfd_headers) + 1)

note = ws.cell(row=1, column=len(hfd_headers) + 3,
    value="GEÇİCİ YAPI: Bu sekmenin sütunları, Temmuz 2026 hakediş detay excel örneği incelendikten sonra "
          "gerçek yapıya göre güncellenecektir. Şimdilik eşleştirme için gereken asgari alanlar (DİVA Fatura "
          "No + Ürün Kodu + Kalem Türü + Tutar) kullanılmıştır.")
note.font = NOTE_FONT
note.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[1].height = 36

HFD_FIRST, HFD_LAST = 3, 1200
HFD_FATNO_C, HFD_URUN_C, HFD_TUR_C, HFD_TUTAR_C, HFD_KEY_C = 3, 4, 6, 7, 8

example_hfd = ["HFT-2026-000123", "05.09.2026", "2600000453", "SMS26DW00T",
               "Mayıs-Tekil Bulaşık Makinesi Kampanyası-202605", "Fiyat Farkı", 3777.92]
for i, val in enumerate(example_hfd, start=1):
    c = ws.cell(row=2, column=i, value=val)
    c.font = EXAMPLE_FONT
    c.fill = EXAMPLE_FILL
    c.border = BORDER
    c.alignment = CENTER
    if i == 2:
        c.number_format = DATE_FMT
    if i == 7:
        c.number_format = CUR

dv_tur = DataValidation(type="list", formula1='"Fiyat Farkı,Bundle,Nakliye,Montaj,Diğer"', allow_blank=True)
ws.add_data_validation(dv_tur)
dv_tur.add(f"{get_column_letter(HFD_TUR_C)}{HFD_FIRST}:{get_column_letter(HFD_TUR_C)}{HFD_LAST}")

for row in range(HFD_FIRST, HFD_LAST + 1):
    for col in range(1, 8):
        c = ws.cell(row=row, column=col)
        c.fill = INPUT_FILL
        c.font = INPUT_FONT
        c.border = BORDER
        c.alignment = CENTER
        if col == 2:
            c.number_format = DATE_FMT
        if col == 7:
            c.number_format = CUR
    fatno = f"{get_column_letter(HFD_FATNO_C)}{row}"
    urun = f"{get_column_letter(HFD_URUN_C)}{row}"
    c = ws.cell(row=row, column=HFD_KEY_C, value=f'=IF({fatno}="","",{fatno}&"|"&{urun})')
    c.font = FORMULA_FONT
    c.border = BORDER
    c.alignment = CENTER

HFD_KEY_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_KEY_C)}${HFD_FIRST}:${get_column_letter(HFD_KEY_C)}${HFD_LAST}"
HFD_TUR_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_TUR_C)}${HFD_FIRST}:${get_column_letter(HFD_TUR_C)}${HFD_LAST}"
HFD_TUTAR_RNG = f"'Hakedis Fatura Detay'!${get_column_letter(HFD_TUTAR_C)}${HFD_FIRST}:${get_column_letter(HFD_TUTAR_C)}${HFD_LAST}"

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
    "Beklenen\nHakediş Tutarı", "Uygulanacak\nKural", "Nakliye\nİzlenebilir mi",
    "Nakliye\nTamamlandı mı", "Montaj\nTamamlandı mı", "Hakediş\nTetikleyici Tarih",
    "Süre İçinde\nTamamlandı mı", "Ödeme Zamanlaması\nDurumu", "Faturada\nBulunan Tutar",
    "Tutar\nFarkı", "Tutar\nDurumu", "GENEL DURUM",
]
set_widths(ws, [13, 12, 16, 11, 26, 12, 12, 11, 13, 20, 11, 11, 11, 13, 12, 20, 13, 11, 15, 24])
header_row(ws, 1, KONTROL_COLS, CALC_HEADER_FILL, HEADER_FONT, height=44)

K_FIRST, K_LAST = DS_FIRST, DS_LAST  # 1:1 aligned with DIVA Satislari rows
(A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T) = range(1, 21)

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
    ds_key = ds(KEY_C)

    a = f"{L_(A)}{row}"; e = f"{L_(E)}{row}"; f = f"{L_(F)}{row}"; g = f"{L_(G)}{row}"
    h = f"{L_(H)}{row}"; j = f"{L_(J)}{row}"
    k = f"{L_(K)}{row}"; l = f"{L_(L)}{row}"; m = f"{L_(M)}{row}"; n = f"{L_(N)}{row}"; o = f"{L_(O)}{row}"
    p = f"{L_(P)}{row}"; i_ = f"{L_(I)}{row}"; q = f"{L_(Q)}{row}"; r_ = f"{L_(R)}{row}"; s_ = f"{L_(S)}{row}"

    match_e2e = f'MATCH({ds_key},{E2E_KEY_RNG},0)'
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
        K: (f'=IF({ds_fno}="","",IFERROR(INDEX({E2E_NAKIZLE_RNG},{match_e2e}),"Eşleşme Yok"))', None),
        L: (f'=IF({ds_fno}="","",IFERROR(INDEX({E2E_NAKTAM_RNG},{match_e2e}),"Hayır"))', None),
        M: (f'=IF({ds_fno}="","",IFERROR(INDEX({E2E_MONTAM_RNG},{match_e2e}),"Hayır"))', None),
        N: (f'=IF({ds_fno}="","",IFERROR(INDEX({E2E_TETIK_RNG},{match_e2e}),""))', DATE_FMT),
        O: (f'=IF(OR({a}="",{n}=""),"",IF({n}<={h},"Evet","HAYIR - RİSK"))', None),
        P: (f'=IF({a}="","",IF(LEFT({j},8)="Kural A ","Uygun - Hemen Ödenir (Yıl Sonu %75 Kontrolüne Tabi)",'
            f'IF(OR({l}="Evet",{m}="Evet"),IF({o}="Evet","Uygun","RİSK - Hakediş Süresi Aşıldı"),'
            f'"Bekleniyor - Nakliye/Montaj Tamamlanmadı")))', None),
        Q: (f'=IF({a}="","",SUMIFS({HFD_TUTAR_RNG},{HFD_KEY_RNG},{ds_key},{HFD_TUR_RNG},"Fiyat Farkı")'
            f'+SUMIFS({HFD_TUTAR_RNG},{HFD_KEY_RNG},{ds_key},{HFD_TUR_RNG},"Bundle"))', CUR),
        R: (f'=IF({a}="","",{i_}-{q})', CUR),
        S: (f'=IF({a}="","",IF(LEFT({p},7)="Bekleni","N/A - Henüz Ödenmemeli",'
            f'IF(LEFT({p},4)="RİSK","İNCELEME GEREKLİ - Süre Aşımı",'
            f'IF({r_}=0,"Doğru",IF({r_}>0,"EKSİK ÖDENMİŞ","FAZLA ÖDENMİŞ")))))', None),
        T: (f'=IF({a}="","",IF({f}<>"Evet","İNCELEME GEREKLİ - Kampanya Tanımsız",'
            f'IF({g}<>"Evet","İNCELEME GEREKLİ - Tarih Dışı Satış",'
            f'IF(LEFT({p},4)="RİSK","İNCELEME GEREKLİ - Hakediş Süresi Aşıldı",'
            f'IF(OR({s_}="EKSİK ÖDENMİŞ",{s_}="FAZLA ÖDENMİŞ"),"İNCELEME GEREKLİ - Tutar Uyuşmazlığı",'
            f'IF(LEFT({p},7)="Bekleni","Bekleniyor","Uygun"))))))', None),
    }
    for col, (formula, fmt) in cells.items():
        c = ws.cell(row=row, column=col, value=formula)
        c.font = FORMULA_FONT
        c.border = BORDER
        c.alignment = CENTER
        if fmt:
            c.number_format = fmt
        if col == T:
            c.font = Font(name=FONT, size=9, bold=True)

K_A_RNG = f"Kontrol!${L_(A)}${K_FIRST}:${L_(A)}${K_LAST}"
K_F_RNG = f"Kontrol!${L_(F)}${K_FIRST}:${L_(F)}${K_LAST}"
K_I_RNG = f"Kontrol!${L_(I)}${K_FIRST}:${L_(I)}${K_LAST}"
K_J_RNG = f"Kontrol!${L_(J)}${K_FIRST}:${L_(J)}${K_LAST}"
K_L_RNG = f"Kontrol!${L_(L)}${K_FIRST}:${L_(L)}${K_LAST}"
K_M_RNG = f"Kontrol!${L_(M)}${K_FIRST}:${L_(M)}${K_LAST}"
K_Q_RNG = f"Kontrol!${L_(Q)}${K_FIRST}:${L_(Q)}${K_LAST}"
K_R_RNG = f"Kontrol!${L_(R)}${K_FIRST}:${L_(R)}${K_LAST}"
K_S_RNG = f"Kontrol!${L_(S)}${K_FIRST}:${L_(S)}${K_LAST}"
K_T_RNG = f"Kontrol!${L_(T)}${K_FIRST}:${L_(T)}${K_LAST}"

print("Part 6 (Kontrol) done")

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
     f'=SUMPRODUCT((LEFT({K_J_RNG},8)="Kural A ")*((({K_L_RNG}="Evet")+({K_M_RNG}="Evet"))>0))', None),
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
    ("Faturada Bulunan Toplam Tutar", f'=SUMIF({K_A_RNG},"<>",{K_Q_RNG})', CUR, False),
    ("Toplam Tutar Farkı", f'=SUMIF({K_A_RNG},"<>",{K_R_RNG})', CUR, False),
    ("Eksik Ödenmiş Satış Sayısı", f'=COUNTIF({K_S_RNG},"EKSİK ÖDENMİŞ")', None, True),
    ("Fazla Ödenmiş Satış Sayısı", f'=COUNTIF({K_S_RNG},"FAZLA ÖDENMİŞ")', None, True),
    ("Kampanya Tanımsız (Kampanyalar Sekmesinde Yok)", f'=COUNTIF({K_T_RNG},"İNCELEME GEREKLİ - Kampanya Tanımsız")', None, True),
    ("Tarih Dışı Satış Sayısı", f'=COUNTIF({K_T_RNG},"İNCELEME GEREKLİ - Tarih Dışı Satış")', None, True),
    ("Hakediş Süresi Aşılan (RİSK) Satış Sayısı", f'=COUNTIF({K_T_RNG},"İNCELEME GEREKLİ - Hakediş Süresi Aşıldı")', None, True),
    ("Henüz Ödenmemesi Gereken (Bekleyen) Satış Sayısı", f'=COUNTIF({K_T_RNG},"Bekleniyor")', None, False),
    ("İncelenmesi Gereken Toplam Satış Sayısı", f'=SUMPRODUCT(--(LEFT({K_T_RNG},16)="İNCELEME GEREKLİ"))', None, True),
    ("Uygun / Sorunsuz Satış Sayısı", f'=COUNTIF({K_T_RNG},"Uygun")', None, False),
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

# DIVA Satislari: kolon 1-35 (ham veri girisi) acik, 36-44 (formul) kilitli
protect_formula_sheet(wb["DIVA Satislari"], DS_LAST, 44, unlock_ranges=[(2, DS_LAST, 1, 35)])
# E2E: kolon 1-38 acik, 39-48 kilitli
protect_formula_sheet(wb["E2E Nakliye Montaj"], E2E_LAST, 48, unlock_ranges=[(2, E2E_LAST, 1, 38)])
# Hakedis Fatura Detay: kolon 1-7 acik, 8 (anahtar) kilitli
protect_formula_sheet(wb["Hakedis Fatura Detay"], HFD_LAST, 8, unlock_ranges=[(2, HFD_LAST, 1, 7)])
# Kontrol, Yil Sonu Kontrolu, Ozet: tamamen formul, hicbir yer acik degil
protect_formula_sheet(wb["Kontrol"], K_LAST, 20)
protect_formula_sheet(wb["Yil Sonu Kontrolu"], 40, 4)
protect_formula_sheet(wb["Ozet"], 40, 4)
# Parametreler: sadece C sutunundaki giris hucreleri acik
protect_formula_sheet(wb["Parametreler"], 20, 4, unlock_ranges=[(4, 20, 3, 3)])

wb.security = openpyxl.workbook.protection.WorkbookProtection(
    workbookPassword=ADMIN_PASSWORD, lockStructure=True, lockWindows=False
)

ORDER = ["Talimatlar", "Parametreler", "Kampanyalar", "DIVA Satislari", "E2E Nakliye Montaj",
         "Hakedis Fatura Detay", "Kontrol", "Yil Sonu Kontrolu", "Ozet"]
wb._sheets = [wb[name] for name in ORDER]
wb.active = 0

TAB_COLORS = {
    "Talimatlar": "1F4E78", "Parametreler": "1F4E78", "Kampanyalar": "BF8F00",
    "DIVA Satislari": "2E75B6", "E2E Nakliye Montaj": "2E75B6", "Hakedis Fatura Detay": "2E75B6",
    "Kontrol": "548235", "Yil Sonu Kontrolu": "C00000", "Ozet": "548235",
}
for name, color in TAB_COLORS.items():
    wb[name].sheet_properties.tabColor = color

OUT = os.path.join(SCRIPT_DIR, "Bayi_DIVA_Kampanya_Hakedis_Kontrol.xlsx")
wb.save(OUT)
print("SAVED", OUT)
print("ADMIN_PASSWORD:", ADMIN_PASSWORD)
print("STRUCT_PASSWORD:", STRUCT_PASSWORD)
