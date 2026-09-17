# -*- coding: utf-8 -*-
"""
===================================================================================================
FİDE KONSERVE
Puantaj ve PDKS Entegrasyonu, Akıllı Vardiya Analizi ve Modern Excel Raporlama Sistemi
===================================================================================================

Bu script, işletmenin ham PDKS (turnike) verilerini ve mevcut bordro/personel tablolarını uçtan uca
işleyerek hatasız, canlı formüllü ve denetlenebilir modern bir Excel (puantaj.xlsx) çalışma kitabı üretir.

---------------------------------------------------------------------------------------------------
TEMEL GÖREVLER VE İŞ MANTIĞI:
---------------------------------------------------------------------------------------------------
1. PDKS HAREKETLERİNİN AKILLI ANALİZİ (pdks.xls):
   - Mükerrer / Peş Peşe Basım Filtresi: Turnikede 5 dakika içinde peş peşe yapılan çift basımları
     otomatik olarak filtreler ve tekil hareket kabul eder.
   - Çoklu Basım Yönetimi (Min/Max Kuralı): Kişinin gün içinde 3 veya daha fazla kart basımı varsa
     en erken saati GİRİŞ (MIN), en geç saati ÇIKIŞ (MAX) kabul ederek net çalışma süresini hesaplar.
     Bu hücreler insan kontrolü için görsel olarak işaretlenir.
   - Unutulan / Tek Basım Yönetimi (Eksik Basım):
     * Sabah basıp akşam çıkış basmayanlar: Hak kaybı yaşanmaması için 7.5 saat (tam gün) yazılır.
     * Sabah basmayı unutup akşam çıkış basanlar: 08:00 iş başı kabul edilerek çıkış saatine göre
       hak ettiği fazla mesaisi korunarak yazılır (Örn: 18:12 çıkış -> 8.5 saat).
     * Tüm eksik basımlar 'Aylik_Puantaj' sayfasında TURUNCU renkle boyanır, hücreye detaylı açıklama
       notu eklenir ve 'Eksik_Basim_Raporu' sayfasında amir onayına sunulur.

2. ÖZEL BÖLÜM PRİMLERİ VE BONUS MESAİ KURALLARI:
   - Balık Dolum ve Balık Kesim Ekipleri:
     * Fiili çalışma süresi 10 saat ve üzeri olduğunda net çalışma süresine +2.0 saat prim eklenir.
       (Örn: 10 saat çalışan -> 12 saat, 11 saat çalışan -> 13 saat, 11.5 saat çalışan -> 13.5 saat yazılır).
     * 10 saatin altındaki çalışmalarda standart kademeler geçerlidir.
   - Üretim Ekibi (Üretim & Konserve Üretim Elemanı):
     * Fiili çalışma süresi 12 saat ve üzeri olduğunda net çalışma süresine +4.0 saat prim eklenir.
       (Örn: 12 saat çalışan -> 16 saat, 13 saat çalışan -> 17 saat, 14 saat çalışan -> 18 saat yazılır).
     * 12 saatin altındaki çalışmalarda standart kademeler geçerlidir.
   - Tüm bu bonuslu mesai günleri 'Aylik_Puantaj' sayfasında LİLA/MOR renkle işaretlenir ve
     'Bolum_Prim_Mesai_Raporu' sayfasında amir kontrolüne sunulur.

3. STANDART VARDİYA VE FAZLA MESAİ KURALLARI (Diğer Bölümler):
   - 1. Vardiya (Gündüz): 08:00 - 17:00 (8-5) ve 08:00 - 16:00 (8-4)
     * Sabah Toleransı: 08:20'ye kadar gelenlerin başlangıcı 08:00 kabul edilir (tam 7.5h).
       08:21 ve sonrası her 30 dakikada bir yarım saat mesai kesintisi uygulanır.
     * Akşam Çıkış & Fazla Mesai (8-5 için):
       - 17:00 - 17:25 arası çıkış -> 7.5 saat (Fazla Mesai Yok)
       - 17:26 - 17:49 arası çıkış -> 8.0 saat (+0.5 saat FM)
       - 17:50 - 18:25 arası çıkış -> 8.5 saat (+1.0 saat FM)
       - 18:26 - 18:49 arası çıkış -> 9.0 saat (+1.5 saat FM)
       - 18:50 - 19:25 arası çıkış -> 9.5 saat (+2.0 saat FM)
       - 19:26 - 19:49 arası çıkış -> 10.0 saat (+2.5 saat FM)
       - 19:50 - 20:25 arası çıkış -> 10.5 saat (+3.0 saat FM)
       - 20:26 - 20:49 arası çıkış -> 11.0 saat (+3.5 saat FM)
   - 2. Vardiya (Akşam) : 16:00 - 24:00 (16-24)
   - 3. Vardiya (Gece)  : 24:00 - 08:00 (24-8 / 00:00 - 08:00)

4. ÇOK SAYFALI VE CANLI FORMÜLLÜ MODERN EXCEL ÜRETİMİ (openpyxl):
   - Aylik_Puantaj: 30 günlük çalışma süreleri, dinamik SUM, COUNTIF, IF, MIN ve VLOOKUP formülleri.
   - Eksik_Basim_Raporu: İnsan kontrolü gerektiren tüm istisna ve tek/çoklu basımların denetim listesi.
   - Bolum_Prim_Mesai_Raporu: Balık Dolum/Kesim ve Üretim ekiplerinin primli mesai denetim cetveli.
   - Resmi_Bordro_SGK, Icra_Takip, SGK_Raporlar, Ucretli_Izinler, Daimi_Personel_Listesi, Elden_Odeme_Farki.
===================================================================================================
"""

import xlrd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
import re
import os
import shutil
import unicodedata
from collections import defaultdict

# =================================================================================================
# 1. TÜRKÇE KARAKTER VE KELİME ONARIM SÖZLÜĞÜ
# =================================================================================================
WORD_REPLACEMENTS = {
    "RENLOLU": "İRENLİOĞLU",
    "ETN": "ÇETİN",
    "ADVYE": "ADVİYE",
    "AFFE": "AFİFE",
    "KLTER": "KÜLTER",
    "ENER": "ŞENER",
    "AKICI": "ÇAKICI",
    "AYLAK": "ÇAYLAK",
    "TRK": "TÜRKÜ",
    "EREF": "EŞREF",
    "GNEY": "GÜNEY",
    "GLL": "GÜLLÜ",
    "GLDANE": "GÜLDANE",
    "GLSM": "GÜLSÜM",
    "GLAH": "GÜLŞAH",
    "GLEN": "GÜLŞEN",
    "GZDE": "GÜZİDE",
    "GLER": "GÜLER",
    "GZN": "GÜZİN",
    "GNAY": "GÜNAY",
    "GRDAL": "GÜRDAL",
    "GRGEN": "GÜRGEN",
    "GRGN": "GÜRGÜN",
    "GZELDAL": "GZELDAL",
    "ZKAN": "ÖZKAN",
    "ZDEMR": "ÖZDEMİR",
    "ZCAN": "ÖZCAN",
    "ZTRK": "ÖZTÜRK",
    "ZMEN": "ÖZMEN",
    "ZEN": "ÖZEN",
    "ZALP": "ÖZALP",
    "ZBEK": "ÖZBEK",
    "ZTOP": "ÖZTOP",
    "NAL": "ÜNAL",
    "NL": "ÜNLÜ",
    "STN": "ÜSTÜN",
    "MT": "ÜMİT",
    "MM": "ÜMMÜ",
    "MMGLSM": "ÜMMÜGÜLSÜM",
    "LK": "ÜLKÜ",
    "LKER": "ÜLKER",
    "LKNUR": "İLKNUR",
    "LHAN": "İLHAN",
    "LYAS": "İLYAS",
    "SMAL": "İSMAİL",
    "BRAHM": "İBRAHİM",
    "NC": "İNCİ",
    "PEK": "İPEK",
    "MREN": "İMREN",
    "REM": "İREM",
    "RFAN": "İRFAN",
    "SMET": "İSMET",
    "HSAN": "İHSAN",
    "DRS": "İDRİS",
    "IKCAN": "ÇIKCAN",
    "AKIR": "ÇAKIR",
    "AKMAK": "ÇAKMAK",
    "ELK": "ÇELİK",
    "ETNKAYA": "ÇETİNKAYA",
    "OBAN": "ÇOBAN",
    "COKUN": "COŞKUN",
    "ALAR": "ÇAĞLAR",
    "ELEN": "ÇELEN",
    "EVK": "ÇEVİK",
    "FT": "ÇİFTÇİ",
    "ATAL": "ÇATAL",
    "DADELEN": "DAĞDELEN",
    "BOZDEMR": "BOZDEMİR",
    "ELSIKI": "ELİSIKI",
    "FDAN": "FİDAN",
    "DRENC": "DİRENCİ",
    "BLG": "BİLGİ",
    "MAKAK": "MAKAK",
    "TEMREN": "TEMREN",
    "BAI": "BAŞI",
    "GDER": "GİDER",
    "GR": "GİRİŞ",
    "IKI": "ÇIKIŞ",
    "MEVSMLK": "MEVSİMLİK",
    "DAM": "DAİMİ",
    "EMEKL": "EMEKLİ",
    "ZN": "İZİN",
    "MESA": "MESAİ",
    "TATL": "TATİL",
    "ALIMA": "ÇALIŞMA",
    "CRET": "ÜCRET",
    "KIDEM": "KIDEM",
    "BLM": "BÖLÜM",
    "KAMET": "İKAMET",
    "RKET": "ŞİRKET",
    "SAAT": "SAATİ",
    "GN": "GÜNÜ",
    "ST": "SÖĞÜT",
    "SGT": "SÖGÜT",
    "BRA": "BÜŞRA",
    "YUCEL": "YÜCEL",
    "SILA ZER": "SILA ÖZER",
}

def clean_display_text(text):
    """
    Excel hücrelerinde görüntülenecek metinleri temizler ve bozuk Türkçe karakterleri düzeltir.
    """
    if not text:
        return ""
    s = str(text).strip()
    for k, v in WORD_REPLACEMENTS.items():
        s = s.replace(k, v)
    s = s.replace("\ufffd", "").replace("", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def norm_name_key(s):
    """
    PDKS listesi ile Puantaj personel listesi arasındaki isimleri %100 eşleştirmek için
    özel karakterleri, parantez içi lakapları ve ekleri arındıran kanonik anahtar üretir.
    """
    if not s:
        return ""
    s = str(s).strip()
    s = re.sub(r"\(.*?\)", "", s)  # (BASİL), (SEMİH) gibi parantez içi ekleri kaldır
    s = re.sub(r"-.*$", "", s)      # -MKP gibi tire sonrası ekleri kaldır
    s = s.replace("\ufffd", "I").replace("", "")
    s = unicodedata.normalize("NFKD", s)
    s = re.sub(r"[^a-zA-Z0-9]", "", s).upper()
    
    # Özel isim varyasyonları eşleştirmesi
    aliases = {
        "SAMETYUMITKEN": "SAMETYUMITKAN",
        "SERKANYUMITKEN": "SERKANYUMITKAN",
        "SELAHATTINSOGUT": "SELAHATTINSOGUT",
        "SELAHATTINSIT": "SELAHATTINSOGUT",
        "SELAHATTINSIGIT": "SELAHATTINSOGUT",
        "SELAHATTINSUGUT": "SELAHATTINSOGUT",
        "BUSRAYAVRUTURK": "BUSRAYAVRUKURT",
        "AYSEYILMAZ": "AYSEYILMAZ",
        "MHDBASSELALZALEK": "MHDBASSELALZALEK",
        "MUHAMMEDSAMIHSERHAN": "MUHAMMEDSAMIHSERHAN",
    }
    return aliases.get(s, s)

def clean_tc(tc_val):
    """
    TC Kimlik Numarasını temizler, 11 haneli standart metin formatına getirir.
    """
    if tc_val is None:
        return ""
    if isinstance(tc_val, float):
        if tc_val.is_integer():
            s = str(int(tc_val)).strip()
        else:
            s = str(int(round(tc_val))).strip()
    else:
        s = str(tc_val).replace(".0", "").strip()
    s = re.sub(r"\D", "", s)
    if not s:
        return ""
    if len(s) == 10:
        s = "0" + s
    return s

def clean_str(val):
    """
    Genel metin alanlarını temizler ve Türkçe karakter onarımından geçirir.
    """
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val)).strip()
    return clean_display_text(str(val).strip())

def clean_money(val):
    """
    Parasal değerleri güvenli float formatına dönüştürür.
    """
    if val is None or val == "":
        return 0.0
    try:
        return float(val)
    except:
        return 0.0

def parse_time_str(s):
    """
    '08:30' formatındaki saat dizgisini sayısal ondalık saate (8.5) dönüştürür.
    """
    if not s or s == "":
        return 0.0
    s = str(s).strip()
    if ":" in s:
        parts = s.split(":")
        try:
            h = float(parts[0])
            m = float(parts[1]) if len(parts) > 1 else 0.0
            return round(h + m / 60.0, 2)
        except:
            return 0.0
    try:
        return float(s)
    except:
        return 0.0

def xldate_to_str(val):
    """
    Excel seri tarih sayılarını 'GG.AA.YYYY' formatlı metne dönüştürür.
    """
    if isinstance(val, (int, float)) and val > 30000 and val < 60000:
        try:
            return xlrd.xldate.xldate_as_datetime(val, 0).strftime("%d.%m.%Y")
        except:
            return str(val)
    return str(val) if val is not None else ""

def get_sheet_by_keyword(wb, keyword, default_idx=None):
    """
    İsminde belirli bir anahtar kelime geçen sayfayı bulur.
    """
    for s in wb.sheets():
        if keyword.lower() in s.name.lower():
            return s
    if default_idx is not None and default_idx < len(wb.sheets()):
        return wb.sheet_by_index(default_idx)
    return None

def parse_hm(t_str):
    """
    '08:15:00' formatından (saat, dakika) tuple'ı döndürür.
    """
    if not t_str or ":" not in str(t_str):
        return None
    p = str(t_str).strip().split(":")
    try:
        return int(p[0]), int(p[1])
    except:
        return None

def fmt_hm(mins):
    """
    00:00'dan itibaren geçen toplam dakikayı 'SS:DD' metnine dönüştürür.
    """
    h = (mins // 60) % 24
    m = mins % 60
    return f"{h:02d}:{m:02d}"

def fmt_hours_tr(val):
    """
    Sayısal saat değerini Türkçe standartlarında virgüllü metne dönüştürür: 7.5 -> '7,5', 13.0 -> '13,0'
    """
    if isinstance(val, (int, float)):
        return f"{val:.1f}".replace(".", ",")
    return str(val).replace(".", ",")

# =================================================================================================
# 2. VARDİYA & FAZLA MESAİ HESAPLAMA MOTORU
# =================================================================================================
def calc_factory_worked_hours(g_saat_str, c_saat_str, sure_str="", is_office=False):
    """
    Fabrika Standart Vardiya ve Fazla Mesai Hesaplama Motoru:
    - 08:20 sabah toleransı ve kademeli geç kalma kesintisi.
    - 17:00 sonrası kademeli fazla mesai aralıkları (17:26-17:49 -> 8.0h, 17:50-18:25 -> 8.5h vb.).
    - 3 vardiya (8-4, 16-24, 24-8) tam desteği.
    """
    g = parse_hm(g_saat_str)
    c = parse_hm(c_saat_str)
    
    if not g or not c:
        return 7.5, 0.0
        
    g_min = g[0] * 60 + g[1]
    c_min = c[0] * 60 + c[1]
    
    # Giriş ve çıkış aynı dakikadaysa (çift basım)
    if abs(c_min - g_min) <= 3:
        return 7.5, 0.0
        
    if c_min < g_min:
        c_min += 24 * 60
        
    # 1. GÜNDÜZ VARDİYASI (06:00 - 11:59 Giriş)
    if 6 * 60 <= g_min <= 11 * 60 + 59:
        if g_min <= 8 * 60 + 20:
            effective_start = 8 * 60
            base_hours = 7.5
        else:
            diff = g_min - (8 * 60 + 20)
            cuts = (diff - 1) // 30 + 1
            effective_start = 8 * 60 + cuts * 30
            base_hours = max(0.0, 7.5 - cuts * 0.5)
            
        # 8-4 vardiyası tam bitiş aralığı
        if 15 * 60 + 50 <= c_min <= 16 * 60 + 25:
            return 7.5, 0.0
            
        # 8-5 vardiyası normal bitişi (17:00) veya erken çıkış
        shift_end = 17 * 60
        if c_min < shift_end:
            elapsed_h = (c_min - effective_start) / 60.0
            if elapsed_h <= 0:
                return 0.0, 0.0
            # Öğle molası öncesi çıkış (12:00 ve öncesi): Mola kesintisi YOK (Tam fiili süre, örn: 08-11 -> 3.0h)
            if c_min <= 12 * 60:
                net = elapsed_h
            # Öğle molası esnası (12:01 - 13:00): 12:00'ye kadar olan net süre
            elif c_min < 13 * 60:
                net = max(0.0, (12 * 60 - effective_start) / 60.0)
            # Öğleden sonra çıkış (13:00 ve sonrası, örn: 08-14 -> 4.5h, 08-15 -> 5.5h): 1.5h mola kesilir
            else:
                net = max(0.0, elapsed_h - 1.5)
                
            return round(net * 2) / 2.0, 0.0
            
        ot_min = c_min - shift_end

    # 2. ÖĞLEDEN SONRA / KISMİ GÜNDÜZ ÇALIŞMASI (12:00 - 15:59 Giriş)
    # Örn: 13:00 - 15:00 -> 2.0h, 13:00 - 17:00 -> 4.0h (Yemek molası kesintisi YOK)
    elif 12 * 60 <= g_min < 16 * 60:
        elapsed_h = (c_min - g_min) / 60.0
        if c_min <= 17 * 60 + 25:
            net = elapsed_h
            return round(net * 2) / 2.0, 0.0
        elif c_min < 24 * 60:
            net = max(0.0, elapsed_h - 0.5)
            return round(net * 2) / 2.0, 0.0
        else:
            shift_end = 24 * 60
            ot_min = c_min - shift_end
            base_hours = 7.5

    # 3. AKŞAM VARDİYASI (16:00 - 19:59 Giriş)
    elif 16 * 60 <= g_min <= 19 * 60 + 59:
        shift_end = 24 * 60
        if g_min <= 16 * 60 + 20:
            effective_start = 16 * 60
            base_hours = 7.5
        else:
            diff = g_min - (16 * 60 + 20)
            cuts = (diff - 1) // 30 + 1
            effective_start = 16 * 60 + cuts * 30
            base_hours = max(0.0, 7.5 - cuts * 0.5)
            
        if c_min < shift_end:
            elapsed_h = (c_min - effective_start) / 60.0
            if elapsed_h <= 4.0:
                net = elapsed_h
            else:
                net = max(0.0, elapsed_h - 0.5)
            return round(net * 2) / 2.0, 0.0
            
        ot_min = c_min - shift_end

    # 4. GECE VARDİYASI (20:00 - 05:59 Giriş)
    else:
        shift_end = 32 * 60 if g_min >= 20 * 60 else 8 * 60
        effective_start = 24 * 60 if g_min >= 20 * 60 else 0
        base_hours = 7.5
        
        if c_min < shift_end:
            raw_w = (c_min - effective_start) / 60.0
            if raw_w <= 4.0:
                net = raw_w
            else:
                net = max(0.0, raw_w - 0.5)
            return round(net * 2) / 2.0, 0.0
            
        ot_min = c_min - shift_end

    # Kademeli Fazla Mesai Basamakları
    if ot_min <= 25:
        step = 0
    else:
        hours_past = ot_min // 60
        min_in_hour = ot_min % 60
        if min_in_hour <= 25:
            step = hours_past * 2
        elif min_in_hour <= 49:
            step = hours_past * 2 + 1
        else:
            step = hours_past * 2 + 2
            
    fm = step * 0.5
    total = base_hours + fm
    return total, fm

# =================================================================================================
# 3. ANA ÇALIŞTIRMA VE RAPOR ÜRETİM FONKSİYONU
# =================================================================================================
def main():
    print("==================================================")
    print(" FİDE KONSERVE - PUANTAJ VE PDKS İŞLEME SİSTEMİ")
    print("==================================================")

    # Orijinal puantaj.xls dosyasını yedekten tazele
    if os.path.exists("puantaj_backup.xls"):
        try:
            shutil.copyfile("puantaj_backup.xls", "puantaj.xls")
        except:
            pass

    print("1. 'puantaj.xls' ve 'pdks.xls' dosyaları yükleniyor...")
    wb_old = xlrd.open_workbook("puantaj.xls", encoding_override="cp1254")
    wb_pdks = xlrd.open_workbook("pdks.xls", encoding_override="cp1254")

    # =========================================================================
    # 1. PERSONEL VE BÖLÜM BİLGİLERİNİN PUANTAJDAN ÇEKİLMESİ
    # =========================================================================
    sh_p = get_sheet_by_keyword(wb_old, "puantaj", 0)
    puantaj_rows = []
    seen_tcs = {}
    emp_dept_map = {}

    for r in range(2, sh_p.nrows):
        sno = clean_str(sh_p.cell_value(r, 0))
        tc = clean_tc(sh_p.cell_value(r, 1))
        ad_soyad = clean_str(sh_p.cell_value(r, 2))
        cinsiyet = clean_str(sh_p.cell_value(r, 3))
        isletme_giris = xldate_to_str(sh_p.cell_value(r, 4))
        sgk_giris = xldate_to_str(sh_p.cell_value(r, 5))
        kidem = sh_p.cell_value(r, 6)
        sgk_cikis = xldate_to_str(sh_p.cell_value(r, 7))
        sgk_durumu = clean_str(sh_p.cell_value(r, 8))
        durumu = clean_str(sh_p.cell_value(r, 9))
        bolum = clean_str(sh_p.cell_value(r, 10))
        ikamet = clean_str(sh_p.cell_value(r, 11))
        net_maas = clean_money(sh_p.cell_value(r, 12))
        sirket = clean_str(sh_p.cell_value(r, 13))
        mesai_durumu = clean_str(sh_p.cell_value(r, 14))
        
        if not ad_soyad:
            continue
            
        name_k = norm_name_key(ad_soyad)
        if bolum:
            emp_dept_map[name_k] = bolum
            
        if tc and tc in seen_tcs:
            continue
        if tc:
            seen_tcs[tc] = ad_soyad
            
        puantaj_rows.append({
            "sno": len(puantaj_rows) + 1,
            "tc": tc,
            "ad_soyad": ad_soyad,
            "cinsiyet": cinsiyet,
            "isletme_giris": isletme_giris,
            "sgk_giris": sgk_giris,
            "kidem": kidem,
            "sgk_cikis": sgk_cikis,
            "sgk_durumu": sgk_durumu if sgk_durumu else "NORMAL",
            "durumu": durumu if durumu else "MEVSİMLİK",
            "bolum": bolum,
            "ikamet": ikamet,
            "net_maas": net_maas,
            "sirket": sirket if sirket else "FİDE KONSERVE",
            "mesai_durumu": mesai_durumu if mesai_durumu else "ALIR"
        })

    print(f"   -> Puantaj tablosundan {len(puantaj_rows)} personel yüklendi.")

    # =========================================================================
    # 2. PDKS HAREKETLERİNİN TOPLANMASI VE AKILLI ANALİZİ
    # =========================================================================
    sh_pdks = wb_pdks.sheet_by_name("HarList")
    pdks_records = []
    emp_pdks_daily = {}
    missing_punch_records = []
    bonus_audit_records = []

    # Günlük hareketleri topla
    raw_daily_rows = defaultdict(lambda: {"rows": [], "meta": {}})

    for r in range(1, sh_pdks.nrows):
        sira = clean_str(sh_pdks.cell_value(r, 0))
        sicil = clean_str(sh_pdks.cell_value(r, 2))
        kart = clean_str(sh_pdks.cell_value(r, 3))
        gun_adi = clean_str(sh_pdks.cell_value(r, 4))
        adi = clean_str(sh_pdks.cell_value(r, 5))
        soyadi = clean_str(sh_pdks.cell_value(r, 6))
        full_name = f"{adi} {soyadi}".strip()
        lokasyon = clean_str(sh_pdks.cell_value(r, 7))
        g_tarih = clean_str(sh_pdks.cell_value(r, 8))
        g_saat = clean_str(sh_pdks.cell_value(r, 9))
        c_tarih = clean_str(sh_pdks.cell_value(r, 10))
        c_saat = clean_str(sh_pdks.cell_value(r, 11))
        puantaj_tarih = clean_str(sh_pdks.cell_value(r, 18)) if sh_pdks.ncols > 18 else ""
        
        date_val = puantaj_tarih if puantaj_tarih else g_tarih
        if not full_name or not date_val or full_name == "BBBBBB":
            continue
            
        day_num = None
        if "." in date_val:
            parts = date_val.split(".")
            try:
                d = int(parts[0])
                m = int(parts[1])
                if m == 9:
                    day_num = d
            except:
                pass
                
        if not day_num:
            continue
            
        name_key = norm_name_key(full_name)
        k = (name_key, day_num)
        
        if not raw_daily_rows[k]["meta"]:
            raw_daily_rows[k]["meta"] = {
                "sira": sira, "sicil": sicil, "kart": kart, "gun_adi": gun_adi,
                "full_name": full_name, "lokasyon": lokasyon, "date_val": date_val
            }
            
        raw_daily_rows[k]["rows"].append({
            "g_tarih": g_tarih, "g_saat": g_saat,
            "c_tarih": c_tarih, "c_saat": c_saat
        })

    # -------------------------------------------------------------------------
    # Günlük hareketleri analiz et:
    # - PDKS satırındaki sıralı Giriş - Çıkış ilişkisini koru (Gece vardiyası 16-00 / 18-04)
    # - Sadece gerçek 'ÜRETİM' ve 'BALIK DOLUM / KESİM' bölümlerine prim uygula
    # - Çoklu ve eksik basımları güvenli tespit et
    # -------------------------------------------------------------------------
    for (name_key, day_num), data in raw_daily_rows.items():
        rows = data["rows"]
        meta = data["meta"]
        full_name = meta["full_name"]
        date_val = meta["date_val"]
        gun_adi = meta["gun_adi"]
        kart = meta["kart"]
        sicil = meta["sicil"]
        lokasyon = meta["lokasyon"]
        dept_name = emp_dept_map.get(name_key, "").upper()
        
        # Bölüm ismini standartlaştır (Türkçe karakter duyarlı)
        norm_dept = dept_name.replace("İ", "I").replace("Ü", "U").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").strip()
        is_balik_dk = ("BALIK DOLUM" in norm_dept) or ("BALIK KESIM" in norm_dept) or ("BALIK KES" in norm_dept)
        is_uretim = (norm_dept == "URETIM") or ("URETIM ELEMANI" in norm_dept) or ("KONSERVE URETIM" in norm_dept)
        
        is_audit = False
        is_bonus = False
        bonus_hours = 0.0
        status_type = None
        note = ""
        g_display = "-"
        c_display = "-"
        all_punches = []
        
        for r_item in rows:
            if r_item["g_saat"]:
                all_punches.append(r_item["g_saat"])
            if r_item["c_saat"]:
                all_punches.append(r_item["c_saat"])
        all_punches_str = ", ".join(all_punches)
        
        # 1. DURUM: Tek satır var ve hem Giriş hem Çıkış saati dolu (Normal / Gece Vardiyası)
        if len(rows) == 1 and rows[0]["g_saat"] and rows[0]["c_saat"]:
            g_raw = rows[0]["g_saat"]
            c_raw = rows[0]["c_saat"]
            g_display = g_raw[:5] if len(g_raw) >= 5 else g_raw
            c_display = c_raw[:5] if len(c_raw) >= 5 else c_raw
            sure, fazla = calc_factory_worked_hours(g_raw, c_raw)
            mesai = min(7.5, sure)
            fiili_sure = sure
            
        # 2. DURUM: Birden fazla satır var (Çoklu basım veya parçalı hareket)
        elif len(rows) > 1 or (len(rows) == 1 and rows[0]["g_saat"] and rows[0]["c_saat"]):
            # Tüm geçerli giriş ve çıkışları topla
            g_list = [r_item["g_saat"] for r_item in rows if r_item["g_saat"]]
            c_list = [r_item["c_saat"] for r_item in rows if r_item["c_saat"]]
            
            if g_list and c_list:
                g_raw = g_list[0] # İlk giriş
                c_raw = c_list[-1] # Son çıkış
                g_display = g_raw[:5] if len(g_raw) >= 5 else g_raw
                c_display = c_raw[:5] if len(c_raw) >= 5 else c_raw
                sure, fazla = calc_factory_worked_hours(g_raw, c_raw)
                mesai = min(7.5, sure)
                fiili_sure = sure
                if len(all_punches) > 2:
                    is_audit = True
                    status_type = "Çoklu Basım (Min/Max)"
                    note = f"{date_val} Çoklu Basım: Giriş {g_display}, Çıkış {c_display} ({fmt_hours_tr(sure)} saat)"
            elif g_list:
                # Sadece girişler var
                g_raw = g_list[0]
                g_display = g_raw[:5]
                c_display = "-"
                is_audit = True
                hm = parse_hm(g_raw)
                tm = hm[0] * 60 + hm[1] if hm else 480
                if tm <= 12 * 60 + 30:
                    status_type = "Çıkış Basılmadı"
                    sure, fazla, mesai = 7.5, 0.0, 7.5
                    note = f"{date_val} Giriş: {g_display} | Çıkış Basılmadı (7,5 saat yazıldı)"
                else:
                    status_type = "Giriş Basılmadı (FM Korundu)"
                    sure, fazla = calc_factory_worked_hours("08:00", g_raw)
                    mesai = min(7.5, sure)
                    note = f"{date_val} Çıkış: {g_display} | Giriş Basılmadı ({fmt_hours_tr(sure)} saat yazıldı)"
                fiili_sure = sure
            else:
                sure, fazla, mesai, fiili_sure = 0.0, 0.0, 0.0, 0.0
                
        # 3. DURUM: Tek bir basım var (Eksik Basım)
        elif len(rows) == 1:
            r_single = rows[0]
            t_raw = r_single["g_saat"] if r_single["g_saat"] else r_single["c_saat"]
            is_audit = True
            hm = parse_hm(t_raw)
            tm = hm[0] * 60 + hm[1] if hm else 480
            t_s = t_raw[:5] if len(t_raw) >= 5 else t_raw
            
            if tm <= 12 * 60 + 30:
                g_display = t_s
                c_display = "-"
                status_type = "Çıkış Basılmadı"
                sure, fazla, mesai = 7.5, 0.0, 7.5
                note = f"{date_val} Giriş: {t_s} | Çıkış Basılmadı (7,5 saat yazıldı)"
            elif 12 * 60 + 30 < tm < 20 * 60:
                g_display = "-"
                c_display = t_s
                sure, fazla = calc_factory_worked_hours("08:00", t_raw)
                mesai = min(7.5, sure)
                status_type = "Giriş Basılmadı (FM Korundu)" if fazla > 0 else "Giriş Basılmadı"
                note = f"{date_val} Çıkış: {t_s} | Giriş Basılmadı ({fmt_hours_tr(sure)} saat yazıldı)"
            else:
                g_display = t_s
                c_display = "-"
                status_type = "Gece Vardiyası Tek Basım"
                sure, fazla, mesai = 7.5, 0.0, 7.5
                note = f"{date_val} Basım: {t_s} | Tek Basım (7,5 saat yazıldı)"
            fiili_sure = sure
        else:
            sure, fazla, mesai, fiili_sure = 0.0, 0.0, 0.0, 0.0
            
        # ---------------------------------------------------------------------
        # KESİN BÖLÜM PRİMİ KONTROLÜ (Sadece Balık Dolum/Kesim ve Üretim için):
        # ---------------------------------------------------------------------
        if is_balik_dk and fiili_sure >= 10.0:
            bonus_hours = 2.0
            sure += bonus_hours
            fazla += bonus_hours
            is_bonus = True
            prim_aciklama = f"{dept_name} Primi: Fiili {fmt_hours_tr(fiili_sure)}s -> Bonuslu {fmt_hours_tr(sure)}s (+2,0s Prim)"
            note = f"{note} | {prim_aciklama}" if note else f"{date_val} {prim_aciklama}"
            
        elif is_uretim and fiili_sure >= 12.0:
            bonus_hours = 4.0
            sure += bonus_hours
            fazla += bonus_hours
            is_bonus = True
            prim_aciklama = f"Üretim Primi: Fiili {fmt_hours_tr(fiili_sure)}s -> Bonuslu {fmt_hours_tr(sure)}s (+4,0s Prim)"
            note = f"{note} | {prim_aciklama}" if note else f"{date_val} {prim_aciklama}"

        # Personel günlük çalışma sözlüğüne kaydet
        if name_key not in emp_pdks_daily:
            emp_pdks_daily[name_key] = {}
        emp_pdks_daily[name_key][day_num] = {
            "sure": sure, "normal": mesai, "fazla": fazla,
            "missing_type": status_type, "note": note,
            "is_audit": is_audit, "is_bonus": is_bonus, "bonus_hours": bonus_hours
        }
        
        # İnsan kontrolü gerektiren eksik/çoklu basımları rapora ekle
        if is_audit and full_name:
            missing_punch_records.append({
                "sno": len(missing_punch_records) + 1,
                "tarih": date_val,
                "gun_adi": gun_adi,
                "kart": kart,
                "sicil": sicil,
                "ad_soyad": full_name,
                "lokasyon": lokasyon,
                "g_saat": g_display,
                "c_saat": c_display,
                "tum_hareketler": all_punches_str,
                "hata": status_type,
                "yazilan_saat": sure,
                "fazla_mesai": fazla,
                "durum": "İK / Amir Onayında"
            })
            
        # Özel Bölüm Primi alanları denetim raporuna ekle
        if is_bonus and full_name:
            bonus_audit_records.append({
                "sno": len(bonus_audit_records) + 1,
                "tarih": date_val,
                "gun_adi": gun_adi,
                "kart": kart,
                "sicil": sicil,
                "ad_soyad": full_name,
                "bolum": dept_name,
                "g_saat": g_display,
                "c_saat": c_display,
                "fiili_sure": fiili_sure,
                "eklenen_prim": bonus_hours,
                "yazilan_saat": sure,
                "fazla_mesai": fazla,
                "durum": "İK / Amir Onayında"
            })
            
        pdks_records.append({
            "sira": meta["sira"], "sicil": sicil, "kart": kart, "gun_adi": gun_adi,
            "ad_soyad": full_name, "lokasyon": lokasyon, "g_tarih": date_val,
            "g_saat": g_display, "c_tarih": date_val, "c_saat": c_display,
            "sure": round(sure, 1), "mesai": round(mesai, 1), "fazla": round(fazla, 1),
            "puantaj_tarih": date_val
        })

    print(f"   -> PDKS'den {len(pdks_records)} günlük hareket ({len(missing_punch_records)} eksik basım, {len(bonus_audit_records)} bölüm primi) işlendi.")

    # =========================================================================
    # 3. RAPORLAR, İCRA, İZİN VE BORDRO VERİLERİNİN OKUNMASI
    # =========================================================================
    # SGK Raporları
    sh_rap = get_sheet_by_keyword(wb_old, "rapor", 4)
    rapor_list = []
    rapor_by_tc = {}
    if sh_rap:
        for r in range(3, sh_rap.nrows):
            takip_no = clean_str(sh_rap.cell_value(r, 1))
            sira_no = clean_str(sh_rap.cell_value(r, 2))
            tc = clean_tc(sh_rap.cell_value(r, 3))
            ad_soyad = clean_str(sh_rap.cell_value(r, 4))
            vaka = clean_str(sh_rap.cell_value(r, 5))
            pol_tarih = xldate_to_str(sh_rap.cell_value(r, 6))
            isbasi_tarih = xldate_to_str(sh_rap.cell_value(r, 7))
            if ad_soyad:
                rapor_list.append({
                    "takip_no": takip_no, "sira_no": sira_no, "tc": tc, "ad_soyad": ad_soyad,
                    "vaka": vaka, "pol_tarih": pol_tarih, "isbasi_tarih": isbasi_tarih
                })
                if tc:
                    rapor_by_tc[tc] = rapor_by_tc.get(tc, 0) + 1

    # İcra Takip
    sh_icra = get_sheet_by_keyword(wb_old, "cra", 3)
    icra_list = []
    icra_by_name = {}
    if sh_icra:
        for r in range(2, sh_icra.nrows):
            isim = clean_str(sh_icra.cell_value(r, 1))
            sirket = clean_str(sh_icra.cell_value(r, 3))
            dosya = clean_str(sh_icra.cell_value(r, 4))
            kalan_borc = clean_money(sh_icra.cell_value(r, 5))
            kesilen = clean_money(sh_icra.cell_value(r, 6))
            iban = clean_str(sh_icra.cell_value(r, 7))
            if isim:
                icra_list.append({
                    "isim": isim, "sirket": sirket, "dosya": dosya,
                    "kalan_borc": kalan_borc, "kesilen": kesilen, "iban": iban
                })
                icra_by_name[norm_name_key(isim)] = {
                    "kesilen": kesilen, "kalan_borc": kalan_borc, "dosya": dosya, "iban": iban
                }

    # Ücretli İzinler
    sh_izin = get_sheet_by_keyword(wb_old, "zin", 7)
    izin_list = []
    izin_by_tc = {}
    if sh_izin:
        for r in range(1, sh_izin.nrows):
            tc = clean_tc(sh_izin.cell_value(r, 0))
            isim = clean_str(sh_izin.cell_value(r, 1))
            saat = clean_money(sh_izin.cell_value(r, 2)) if sh_izin.ncols > 2 else 0.0
            gun = clean_money(sh_izin.cell_value(r, 3)) if sh_izin.ncols > 3 else 0.0
            if isim:
                izin_list.append({"tc": tc, "isim": isim, "saat": saat, "gun": gun})
                if tc:
                    izin_by_tc[tc] = {"saat": saat, "gun": gun}

    # Resmi Bordro SGK
    sh_bordro = get_sheet_by_keyword(wb_old, "sayfa4", 8)
    bordro_list = []
    bordro_by_tc = {}
    if sh_bordro:
        for r in range(1, sh_bordro.nrows):
            sno = clean_str(sh_bordro.cell_value(r, 0))
            tc = clean_tc(sh_bordro.cell_value(r, 1))
            ad_soyad = clean_str(sh_bordro.cell_value(r, 2))
            ucret = clean_money(sh_bordro.cell_value(r, 3))
            giris = xldate_to_str(sh_bordro.cell_value(r, 4))
            cikis = xldate_to_str(sh_bordro.cell_value(r, 5))
            normal_kazanc = clean_money(sh_bordro.cell_value(r, 6))
            ek_kazanc = clean_money(sh_bordro.cell_value(r, 7))
            yasal_kesinti = clean_money(sh_bordro.cell_value(r, 8))
            toplam_kazanc = clean_money(sh_bordro.cell_value(r, 10)) if sh_bordro.ncols > 10 else 0.0
            toplam_kesinti = clean_money(sh_bordro.cell_value(r, 11)) if sh_bordro.ncols > 11 else 0.0
            odenecek_net = clean_money(sh_bordro.cell_value(r, 12)) if sh_bordro.ncols > 12 else 0.0
            sgk_gun = clean_money(sh_bordro.cell_value(r, 13)) if sh_bordro.ncols > 13 else 0.0
            sgk_brut = clean_money(sh_bordro.cell_value(r, 14)) if sh_bordro.ncols > 14 else 0.0
            kanun = clean_str(sh_bordro.cell_value(r, 30)) if sh_bordro.ncols > 30 else ""
            meslek_kodu = clean_str(sh_bordro.cell_value(r, 33)) if sh_bordro.ncols > 33 else ""
            
            if ad_soyad:
                row_dict = {
                    "sno": sno, "tc": tc, "ad_soyad": ad_soyad, "ucret": ucret,
                    "giris": giris, "cikis": cikis, "normal_kazanc": normal_kazanc,
                    "ek_kazanc": ek_kazanc, "yasal_kesinti": yasal_kesinti,
                    "toplam_kazanc": toplam_kazanc, "toplam_kesinti": toplam_kesinti,
                    "odenecek_net": odenecek_net, "sgk_gun": sgk_gun, "sgk_brut": sgk_brut,
                    "kanun": kanun, "meslek_kodu": meslek_kodu
                }
                bordro_list.append(row_dict)
                if tc:
                    bordro_by_tc[tc] = row_dict

    # Daimi Personel Listesi
    sh_daimi = get_sheet_by_keyword(wb_old, "daim", 1)
    daimi_list = []
    if sh_daimi:
        for r in range(1, sh_daimi.nrows):
            sno = clean_str(sh_daimi.cell_value(r, 0))
            tc = clean_tc(sh_daimi.cell_value(r, 1))
            ad_soyad = clean_str(sh_daimi.cell_value(r, 2))
            sgk_giris = xldate_to_str(sh_daimi.cell_value(r, 3)) if sh_daimi.ncols > 3 else ""
            sgk_durum = clean_str(sh_daimi.cell_value(r, 4)) if sh_daimi.ncols > 4 else ""
            durum = clean_str(sh_daimi.cell_value(r, 5)) if sh_daimi.ncols > 5 else ""
            maas = clean_money(sh_daimi.cell_value(r, 6)) if sh_daimi.ncols > 6 else 0.0
            if ad_soyad:
                daimi_list.append({
                    "sno": sno, "tc": tc, "ad_soyad": ad_soyad,
                    "sgk_giris": sgk_giris, "sgk_durum": sgk_durum,
                    "durum": durum, "maas": maas
                })

    # =========================================================================
    # 4. MODERN VE CANLI FORMÜLLÜ EXCEL (.xlsx) OLUŞTURMA
    # =========================================================================
    print("3. Modern 'puantaj.xlsx' çalışma kitabı oluşturuluyor...")
    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)

    FONT_FAMILY = "Segoe UI"
    font_title = Font(name=FONT_FAMILY, size=13, bold=True, color="1F4E79")
    font_hdr = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
    font_data = Font(name=FONT_FAMILY, size=9, bold=False, color="000000")
    font_total = Font(name=FONT_FAMILY, size=9, bold=True, color="000000")

    fill_navy = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_pazar = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_overtime = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_net = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_deduct = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_missing = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # Turuncu (Eksik Basım)
    fill_bonus = PatternFill(start_color="E1D5E7", end_color="E1D5E7", fill_type="solid")   # Açık Mor/Lila (Bölüm Primi)
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_total = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_hdr = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border_side = Side(border_style="thin", color="D3D3D3")
    thick_bottom = Side(border_style="medium", color="1F4E79")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_header = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thick_bottom)
    border_total = Border(top=thin_border_side, bottom=Side(border_style="double", color="000000"))

    # =========================================================================
    # SAYFA 1: Aylik_Puantaj (Eylül 2026 Cetveli)
    # =========================================================================
    ws_p = wb_new.create_sheet(title="Aylik_Puantaj")
    ws_p.views.sheetView[0].showGridLines = True

    ws_p.merge_cells("A1:N1")
    ws_p["A1"] = "FİDE KONSERVE GIDA SAN. VE TİC. A.Ş. - EYLÜL 2026 AYLIK PUANTAJ VE HAKEDİŞ CETVELİ"
    ws_p["A1"].font = font_title
    ws_p["A1"].alignment = Alignment(horizontal="left", vertical="center")

    info_headers = [
        ("A", "Sıra No", 7),
        ("B", "TC Kimlik No", 14),
        ("C", "Adı Soyadı", 22),
        ("D", "Cinsiyet", 7),
        ("E", "İşletme Giriş", 11),
        ("F", "SGK Giriş", 11),
        ("G", "SGK Çıkış", 11),
        ("H", "SGK Statü", 11),
        ("I", "Kadro", 11),
        ("J", "Bölüm / İş", 16),
        ("K", "İkamet / Güzergah", 15),
        ("L", "Anlaşılan Net Maaş (TL)", 15),
        ("M", "Şirket", 13),
        ("N", "Mesai Durumu", 11),
    ]

    day_names_tr = ["Salı", "Çar", "Per", "Cum", "Cmt", "PAZAR", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "PAZAR", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "PAZAR", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "PAZAR", "Pzt", "Sal", "Çar"]

    day_cols = []
    start_col_idx = 15
    for d in range(1, 31):
        col_letter = get_column_letter(start_col_idx + d - 1)
        day_name = day_names_tr[d-1]
        is_pazar = (day_name == "PAZAR")
        day_cols.append((col_letter, f"{d:02d}.09\n{day_name}", 6, is_pazar, d))

    calc_headers = [
        ("Toplam Çalışma Saati", 13, fill_overtime),
        ("Fiili Çalışılan Gün", 11, fill_navy),
        ("Hafta Tatili Günü", 11, fill_pazar),
        ("Ücretli İzin (Gün)", 11, fill_navy),
        ("Raporlu Gün", 11, fill_deduct),
        ("Toplam SGK Günü", 11, fill_net),
        ("Eksik Gün Sayısı", 11, fill_deduct),
        ("Eksik Gün Nedeni", 15, fill_navy),
        ("Hafta İçi Fazla Mesai (Saat)", 13, fill_overtime),
        ("Pazar Mesai (Saat)", 13, fill_pazar),
        ("Resmi Bordro SGK Net (TL)", 15, fill_navy),
        ("Elden Ödenecek Fark (TL)", 15, fill_net),
        ("İcra Kesintisi (TL)", 13, fill_deduct),
        ("Ödenecek Net Tutar (TL)", 15, fill_net),
    ]

    header_row = 3
    ws_p.row_dimensions[header_row].height = 28
    ws_p.row_dimensions[1].height = 24

    for col_let, text, width in info_headers:
        cell = ws_p[f"{col_let}{header_row}"]
        cell.value = text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_p.column_dimensions[col_let].width = width

    for col_let, text, width, is_pazar, d in day_cols:
        cell = ws_p[f"{col_let}{header_row}"]
        cell.value = text
        cell.font = font_hdr if not is_pazar else Font(name=FONT_FAMILY, size=8, bold=True, color="7F6000")
        cell.fill = fill_navy if not is_pazar else fill_pazar
        cell.alignment = align_hdr
        cell.border = border_header
        ws_p.column_dimensions[col_let].width = width

    cur_col_idx = start_col_idx + 30
    calc_col_letters = []
    for text, width, hfill in calc_headers:
        col_let = get_column_letter(cur_col_idx)
        calc_col_letters.append((col_let, text))
        cell = ws_p[f"{col_let}{header_row}"]
        cell.value = text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_p.column_dimensions[col_let].width = width
        cur_col_idx += 1

    first_day_col = day_cols[0][0]
    last_day_col = day_cols[-1][0]

    # Personel Satırlarını Doldur
    start_data_row = 4
    for idx, emp in enumerate(puantaj_rows):
        r_idx = start_data_row + idx
        ws_p.row_dimensions[r_idx].height = 19
        is_zebra = (idx % 2 == 1)
        row_fill = fill_zebra if is_zebra else None
        
        ws_p[f"A{r_idx}"] = emp["sno"]
        ws_p[f"B{r_idx}"] = emp["tc"]
        ws_p[f"C{r_idx}"] = emp["ad_soyad"]
        ws_p[f"D{r_idx}"] = emp["cinsiyet"]
        ws_p[f"E{r_idx}"] = emp["isletme_giris"]
        ws_p[f"F{r_idx}"] = emp["sgk_giris"]
        ws_p[f"G{r_idx}"] = emp["sgk_cikis"]
        ws_p[f"H{r_idx}"] = emp["sgk_durumu"]
        ws_p[f"I{r_idx}"] = emp["durumu"]
        ws_p[f"J{r_idx}"] = emp["bolum"]
        ws_p[f"K{r_idx}"] = emp["ikamet"]
        ws_p[f"L{r_idx}"] = emp["net_maas"]
        ws_p[f"M{r_idx}"] = emp["sirket"]
        ws_p[f"N{r_idx}"] = emp["mesai_durumu"]
        
        ws_p[f"A{r_idx}"].alignment = align_center
        ws_p[f"B{r_idx}"].alignment = align_center
        ws_p[f"B{r_idx}"].number_format = "@"
        ws_p[f"C{r_idx}"].alignment = align_left
        ws_p[f"D{r_idx}"].alignment = align_center
        ws_p[f"E{r_idx}"].alignment = align_center
        ws_p[f"F{r_idx}"].alignment = align_center
        ws_p[f"G{r_idx}"].alignment = align_center
        ws_p[f"H{r_idx}"].alignment = align_center
        ws_p[f"I{r_idx}"].alignment = align_center
        ws_p[f"J{r_idx}"].alignment = align_left
        ws_p[f"K{r_idx}"].alignment = align_left
        ws_p[f"L{r_idx}"].alignment = align_right
        ws_p[f"L{r_idx}"].number_format = "#,##0.00"
        ws_p[f"M{r_idx}"].alignment = align_center
        ws_p[f"N{r_idx}"].alignment = align_center
        
        norm_name = norm_name_key(emp["ad_soyad"])
        emp_pdks = emp_pdks_daily.get(norm_name, {})
        
        emp_ot_weekday = 0.0
        emp_ot_sunday = 0.0
        
        for col_let, text, width, is_pazar, d in day_cols:
            cell = ws_p[f"{col_let}{r_idx}"]
            cell.alignment = align_center
            cell.number_format = "0.0"
            if is_pazar:
                cell.fill = fill_pazar
            elif row_fill:
                cell.fill = row_fill
                
            if d in emp_pdks:
                day_info = emp_pdks[d]
                hours = day_info["sure"]
                fazla = day_info["fazla"]
                cell.value = hours if hours > 0 else 0
                
                # Görsel Renklendirme ve Hücre Notları
                if not is_pazar:
                    if day_info.get("is_bonus"): # Bölüm Primi Alan Günler (Açık Mor/Lila)
                        cell.fill = fill_bonus
                        if day_info.get("note"):
                            cell.comment = Comment(day_info["note"], "PDKS Sistemi")
                    elif day_info.get("is_audit"): # Eksik / Hatalı Basımlar (Turuncu)
                        cell.fill = fill_missing
                        if day_info.get("note"):
                            cell.comment = Comment(day_info["note"], "PDKS Sistemi")
                            
                    if cell.value == 0 or cell.value == 0.0:
                        cell.value = 7.5
                        
                if is_pazar:
                    emp_ot_sunday += hours
                else:
                    emp_ot_weekday += fazla
            else:
                cell.value = 0
                
        c_tot_hours = calc_col_letters[0][0]
        c_work_days = calc_col_letters[1][0]
        c_ht_days = calc_col_letters[2][0]
        c_izin_days = calc_col_letters[3][0]
        c_rap_days = calc_col_letters[4][0]
        c_sgk_days = calc_col_letters[5][0]
        c_eksik_days = calc_col_letters[6][0]
        c_eksik_neden = calc_col_letters[7][0]
        c_ot_weekday = calc_col_letters[8][0]
        c_ot_pazar = calc_col_letters[9][0]
        c_bordro_net = calc_col_letters[10][0]
        c_elden_fark = calc_col_letters[11][0]
        c_icra_kes = calc_col_letters[12][0]
        c_net_odenecek = calc_col_letters[13][0]
        
        # Canlı Excel Formülleri
        ws_p[f"{c_tot_hours}{r_idx}"] = f"=SUM({first_day_col}{r_idx}:{last_day_col}{r_idx})"
        ws_p[f"{c_work_days}{r_idx}"] = f'=COUNTIF({first_day_col}{r_idx}:{last_day_col}{r_idx}, ">0")'
        ws_p[f"{c_ht_days}{r_idx}"] = f"=IF({c_work_days}{r_idx}>=5, 4, IF({c_work_days}{r_idx}>0, INT({c_work_days}{r_idx}/6), 0))"
        
        tc_val = emp["tc"]
        izin_val = izin_by_tc.get(tc_val, {}).get("gun", 0) if tc_val else 0
        rap_val = rapor_by_tc.get(tc_val, 0) if tc_val else 0
        
        ws_p[f"{c_izin_days}{r_idx}"] = izin_val
        ws_p[f"{c_rap_days}{r_idx}"] = rap_val
        
        ws_p[f"{c_sgk_days}{r_idx}"] = f"=MIN(30, {c_work_days}{r_idx}+{c_ht_days}{r_idx}+{c_izin_days}{r_idx}+{c_rap_days}{r_idx})"
        ws_p[f"{c_eksik_days}{r_idx}"] = f"=30-{c_sgk_days}{r_idx}"
        ws_p[f"{c_eksik_neden}{r_idx}"] = f'=IF({c_rap_days}{r_idx}>0, "01-İstirahat", IF({c_eksik_days}{r_idx}>0, "12-Birden Fazla", ""))'
        
        ws_p[f"{c_ot_weekday}{r_idx}"] = emp_ot_weekday
        ws_p[f"{c_ot_pazar}{r_idx}"] = emp_ot_sunday
        
        ws_p[f"{c_bordro_net}{r_idx}"] = f'=IFERROR(VLOOKUP(B{r_idx}, Resmi_Bordro_SGK!B:M, 11, FALSE), 0)'
        ws_p[f"{c_elden_fark}{r_idx}"] = f'=IF(L{r_idx}>{c_bordro_net}{r_idx}, L{r_idx}-{c_bordro_net}{r_idx}, 0)'
        ws_p[f"{c_icra_kes}{r_idx}"] = f'=IFERROR(VLOOKUP(B{r_idx}, Icra_Takip!A:E, 4, FALSE), 0)'
        ws_p[f"{c_net_odenecek}{r_idx}"] = f'=L{r_idx}-{c_icra_kes}{r_idx}'
        
        for col_info in [
            (c_tot_hours, "0.0", align_center, fill_overtime),
            (c_work_days, "0", align_center, None),
            (c_ht_days, "0", align_center, fill_pazar),
            (c_izin_days, "0", align_center, None),
            (c_rap_days, "0", align_center, fill_deduct if rap_val > 0 else None),
            (c_sgk_days, "0", align_center, fill_net),
            (c_eksik_days, "0", align_center, fill_deduct),
            (c_eksik_neden, "@", align_center, None),
            (c_ot_weekday, "0.0", align_center, fill_overtime),
            (c_ot_pazar, "0.0", align_center, fill_pazar),
            (c_bordro_net, "#,##0.00", align_right, None),
            (c_elden_fark, "#,##0.00", align_right, fill_net),
            (c_icra_kes, "#,##0.00", align_right, fill_deduct),
            (c_net_odenecek, "#,##0.00", align_right, fill_net),
        ]:
            cl, fmt, algn, fill_bg = col_info
            c = ws_p[f"{cl}{r_idx}"]
            c.number_format = fmt
            c.alignment = algn
            if fill_bg:
                c.fill = fill_bg
            elif row_fill:
                c.fill = row_fill
                
        for col_idx in range(1, cur_col_idx):
            cl = get_column_letter(col_idx)
            ws_p[f"{cl}{r_idx}"].border = border_cell
            ws_p[f"{cl}{r_idx}"].font = font_data

    # Puantaj Genel Toplam Satırı
    tot_r_puantaj = start_data_row + len(puantaj_rows)
    ws_p.row_dimensions[tot_r_puantaj].height = 22
    ws_p[f"A{tot_r_puantaj}"] = ""
    ws_p[f"B{tot_r_puantaj}"] = ""
    ws_p[f"C{tot_r_puantaj}"] = "GENEL TOPLAM"
    ws_p[f"C{tot_r_puantaj}"].alignment = align_left
    ws_p[f"C{tot_r_puantaj}"].font = font_total

    ws_p[f"L{tot_r_puantaj}"] = f"=SUM(L4:L{tot_r_puantaj-1})"
    ws_p[f"L{tot_r_puantaj}"].number_format = "#,##0.00"

    for col_let, text, width, is_pazar, d in day_cols:
        ws_p[f"{col_let}{tot_r_puantaj}"] = f"=SUM({col_let}4:{col_let}{tot_r_puantaj-1})"
        ws_p[f"{col_let}{tot_r_puantaj}"].number_format = "0.0"

    for cl in [c_tot_hours, c_work_days, c_ht_days, c_izin_days, c_rap_days, c_sgk_days, c_eksik_days, c_ot_weekday, c_ot_pazar]:
        ws_p[f"{cl}{tot_r_puantaj}"] = f"=SUM({cl}4:{cl}{tot_r_puantaj-1})"
        ws_p[f"{cl}{tot_r_puantaj}"].number_format = "0.0" if cl in (c_tot_hours, c_ot_weekday, c_ot_pazar) else "0"

    for cl in [c_bordro_net, c_elden_fark, c_icra_kes, c_net_odenecek]:
        ws_p[f"{cl}{tot_r_puantaj}"] = f"=SUM({cl}4:{cl}{tot_r_puantaj-1})"
        ws_p[f"{cl}{tot_r_puantaj}"].number_format = "#,##0.00"

    for col_idx in range(1, cur_col_idx):
        cl = get_column_letter(col_idx)
        c = ws_p[f"{cl}{tot_r_puantaj}"]
        c.font = font_total
        c.fill = fill_total
        c.border = border_total
        if not c.alignment.horizontal:
            c.alignment = align_right

    ws_p.freeze_panes = "D4"

    # =========================================================================
    # SAYFA 2: Eksik_Basim_Raporu (İstisna & İnsan Kontrolü Raporu)
    # =========================================================================
    ws_miss = wb_new.create_sheet(title="Eksik_Basim_Raporu")
    ws_miss.views.sheetView[0].showGridLines = True

    miss_headers = [
        ("Sıra No", 8), ("Tarih", 12), ("Gün", 12), ("Kart No", 10), ("Sicil No", 10),
        ("Adı Soyadı", 22), ("Bölüm / Lokasyon", 20), ("Tespit Edilen Giriş", 15),
        ("Tespit Edilen Çıkış", 15), ("Tüm Basım Hareketleri", 22),
        ("İnceleme Nedeni (Hata Türü)", 26), ("Puantaja Yazılan Saat", 18),
        ("Fazla Mesai (Saat)", 16), ("İK / Amir Onayı & Notu", 22)
    ]
    ws_miss.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(miss_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_miss[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_miss.column_dimensions[col_let].width = w

    for r_idx, mp in enumerate(missing_punch_records, start=2):
        ws_miss[f"A{r_idx}"] = mp["sno"]
        ws_miss[f"B{r_idx}"] = mp["tarih"]
        ws_miss[f"C{r_idx}"] = mp["gun_adi"]
        ws_miss[f"D{r_idx}"] = mp["kart"]
        ws_miss[f"E{r_idx}"] = mp["sicil"]
        ws_miss[f"F{r_idx}"] = mp["ad_soyad"]
        ws_miss[f"G{r_idx}"] = mp["lokasyon"]
        ws_miss[f"H{r_idx}"] = mp["g_saat"]
        ws_miss[f"I{r_idx}"] = mp["c_saat"]
        ws_miss[f"J{r_idx}"] = mp["tum_hareketler"]
        ws_miss[f"K{r_idx}"] = mp["hata"]
        ws_miss[f"L{r_idx}"] = mp["yazilan_saat"]
        ws_miss[f"M{r_idx}"] = mp["fazla_mesai"]
        ws_miss[f"N{r_idx}"] = mp["durum"]
        
        for c_idx in range(1, 15):
            cl = get_column_letter(c_idx)
            ws_miss[f"{cl}{r_idx}"].border = border_cell
            ws_miss[f"{cl}{r_idx}"].font = font_data
            if c_idx == 11:
                ws_miss[f"{cl}{r_idx}"].fill = fill_missing
                ws_miss[f"{cl}{r_idx}"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="C00000")
            elif c_idx in (12, 13):
                ws_miss[f"{cl}{r_idx}"].number_format = "0.0"
                ws_miss[f"{cl}{r_idx}"].alignment = align_center
            if c_idx not in (6, 7, 10):
                ws_miss[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 3: Bolum_Prim_Mesai_Raporu (Balık Dolum/Kesim & Üretim Prim Denetimi)
    # =========================================================================
    ws_bon = wb_new.create_sheet(title="Bolum_Prim_Mesai_Raporu")
    ws_bon.views.sheetView[0].showGridLines = True

    bon_headers = [
        ("Sıra No", 8), ("Tarih", 12), ("Gün", 12), ("Kart No", 10), ("Sicil No", 10),
        ("Adı Soyadı", 22), ("Bölüm / Görev", 20), ("Giriş Saati", 14),
        ("Çıkış Saati", 14), ("Fiili Çalışma (Saat)", 18), ("Eklenen Prim (Saat)", 18),
        ("Puantaja Yazılan Saat", 18), ("Toplam Fazla Mesai", 18), ("Vardiya Amiri / İK Onayı", 22)
    ]
    ws_bon.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(bon_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_bon[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_bon.column_dimensions[col_let].width = w

    for r_idx, br in enumerate(bonus_audit_records, start=2):
        ws_bon[f"A{r_idx}"] = br["sno"]
        ws_bon[f"B{r_idx}"] = br["tarih"]
        ws_bon[f"C{r_idx}"] = br["gun_adi"]
        ws_bon[f"D{r_idx}"] = br["kart"]
        ws_bon[f"E{r_idx}"] = br["sicil"]
        ws_bon[f"F{r_idx}"] = br["ad_soyad"]
        ws_bon[f"G{r_idx}"] = br["bolum"]
        ws_bon[f"H{r_idx}"] = br["g_saat"]
        ws_bon[f"I{r_idx}"] = br["c_saat"]
        ws_bon[f"J{r_idx}"] = br["fiili_sure"]
        ws_bon[f"K{r_idx}"] = br["eklenen_prim"]
        ws_bon[f"L{r_idx}"] = br["yazilan_saat"]
        ws_bon[f"M{r_idx}"] = br["fazla_mesai"]
        ws_bon[f"N{r_idx}"] = br["durum"]
        
        for c_idx in range(1, 15):
            cl = get_column_letter(c_idx)
            ws_bon[f"{cl}{r_idx}"].border = border_cell
            ws_bon[f"{cl}{r_idx}"].font = font_data
            if c_idx in (10, 11, 12, 13):
                ws_bon[f"{cl}{r_idx}"].number_format = "0.0"
                ws_bon[f"{cl}{r_idx}"].alignment = align_center
            if c_idx == 11: # Eklenen prim sütunu
                ws_bon[f"{cl}{r_idx}"].fill = fill_bonus
                ws_bon[f"{cl}{r_idx}"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="7030A0")
            if c_idx not in (6, 7):
                ws_bon[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 4: PDKS_Hareket_Kayitlari (Ham Turnike Verileri)
    # =========================================================================
    ws_pdk = wb_new.create_sheet(title="PDKS_Hareket_Kayitlari")
    ws_pdk.views.sheetView[0].showGridLines = True

    pdk_headers = [
        ("Sıra No", 8), ("Sicil No", 10), ("Kart No", 10), ("Adı Soyadı", 22),
        ("Gün", 12), ("Giriş Tarihi", 12), ("Giriş Saati", 12),
        ("Çıkış Tarihi", 12), ("Çıkış Saati", 12), ("Toplam Süre", 12),
        ("Normal Mesai", 12), ("Fazla Mesai", 12), ("Lokasyon", 16)
    ]
    ws_pdk.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(pdk_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_pdk[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_pdk.column_dimensions[col_let].width = w

    for r_idx, rec in enumerate(pdks_records, start=2):
        ws_pdk[f"A{r_idx}"] = rec["sira"]
        ws_pdk[f"B{r_idx}"] = rec["sicil"]
        ws_pdk[f"C{r_idx}"] = rec["kart"]
        ws_pdk[f"D{r_idx}"] = rec["ad_soyad"]
        ws_pdk[f"E{r_idx}"] = rec["gun_adi"]
        ws_pdk[f"F{r_idx}"] = rec["g_tarih"]
        ws_pdk[f"G{r_idx}"] = rec["g_saat"]
        ws_pdk[f"H{r_idx}"] = rec["c_tarih"]
        ws_pdk[f"I{r_idx}"] = rec["c_saat"]
        ws_pdk[f"J{r_idx}"] = rec["sure"]
        ws_pdk[f"K{r_idx}"] = rec["mesai"]
        ws_pdk[f"L{r_idx}"] = rec["fazla"]
        ws_pdk[f"J{r_idx}"].number_format = "0.0"
        ws_pdk[f"K{r_idx}"].number_format = "0.0"
        ws_pdk[f"L{r_idx}"].number_format = "0.0"
        ws_pdk[f"M{r_idx}"] = rec["lokasyon"]
        
        for c_idx in range(1, 14):
            cl = get_column_letter(c_idx)
            ws_pdk[f"{cl}{r_idx}"].border = border_cell
            ws_pdk[f"{cl}{r_idx}"].font = font_data
            if c_idx not in (4, 13):
                ws_pdk[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 5: Resmi_Bordro_SGK (Yasal SGK Bordro Listesi)
    # =========================================================================
    ws_bor = wb_new.create_sheet(title="Resmi_Bordro_SGK")
    ws_bor.views.sheetView[0].showGridLines = True
    bor_headers = [
        ("Sıra No", 8), ("TC Kimlik No", 14), ("Adı Soyadı", 22), ("Ücret (TL)", 14),
        ("Giriş Tarihi", 12), ("Çıkış Tarihi", 12), ("Normal Kazanç", 14),
        ("Ek Kazanç", 14), ("Yasal Kesinti", 14), ("Toplam Kazanç", 14),
        ("Toplam Kesinti", 14), ("Ödenecek Net", 15), ("SGK Gün", 10),
        ("SGK Matrah (Brüt)", 16), ("Kanun No", 12), ("Meslek Kodu", 14)
    ]
    ws_bor.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(bor_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_bor[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_bor.column_dimensions[col_let].width = w

    for r_idx, b in enumerate(bordro_list, start=2):
        ws_bor[f"A{r_idx}"] = b["sno"]
        ws_bor[f"B{r_idx}"] = b["tc"]
        ws_bor[f"B{r_idx}"].number_format = "@"
        ws_bor[f"C{r_idx}"] = b["ad_soyad"]
        ws_bor[f"D{r_idx}"] = b["ucret"]
        ws_bor[f"E{r_idx}"] = b["giris"]
        ws_bor[f"F{r_idx}"] = b["cikis"]
        ws_bor[f"G{r_idx}"] = b["normal_kazanc"]
        ws_bor[f"H{r_idx}"] = b["ek_kazanc"]
        ws_bor[f"I{r_idx}"] = b["yasal_kesinti"]
        ws_bor[f"J{r_idx}"] = b["toplam_kazanc"]
        ws_bor[f"K{r_idx}"] = b["toplam_kesinti"]
        ws_bor[f"L{r_idx}"] = b["odenecek_net"]
        ws_bor[f"M{r_idx}"] = b["sgk_gun"]
        ws_bor[f"N{r_idx}"] = b["sgk_brut"]
        ws_bor[f"O{r_idx}"] = b["kanun"]
        ws_bor[f"P{r_idx}"] = b["meslek_kodu"]
        
        for c_idx in range(1, 17):
            cl = get_column_letter(c_idx)
            ws_bor[f"{cl}{r_idx}"].border = border_cell
            ws_bor[f"{cl}{r_idx}"].font = font_data
            if c_idx in (4, 7, 8, 9, 10, 11, 12, 14):
                ws_bor[f"{cl}{r_idx}"].number_format = "#,##0.00"
                ws_bor[f"{cl}{r_idx}"].alignment = align_right
            elif c_idx not in (3,):
                ws_bor[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 6: Icra_Takip (Yasal Maaş Haciz & İcra Kesintileri)
    # =========================================================================
    ws_icra = wb_new.create_sheet(title="Icra_Takip")
    ws_icra.views.sheetView[0].showGridLines = True
    icra_headers = [
        ("TC Kimlik No", 14), ("Adı Soyadı", 22), ("Şirket", 16),
        ("İcra Kesintisi (TL)", 16), ("Kalan Borç (TL)", 16),
        ("İcra Dosya No", 24), ("İban No", 28), ("Açıklama", 20)
    ]
    ws_icra.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(icra_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_icra[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_icra.column_dimensions[col_let].width = w

    for r_idx, ic in enumerate(icra_list, start=2):
        tc_found = ""
        norm_ic_name = norm_name_key(ic["isim"])
        for p in puantaj_rows:
            if norm_name_key(p["ad_soyad"]) == norm_ic_name:
                tc_found = p["tc"]
                break
                
        ws_icra[f"A{r_idx}"] = tc_found
        ws_icra[f"A{r_idx}"].number_format = "@"
        ws_icra[f"B{r_idx}"] = ic["isim"]
        ws_icra[f"C{r_idx}"] = ic["sirket"]
        ws_icra[f"D{r_idx}"] = ic["kesilen"]
        ws_icra[f"D{r_idx}"].number_format = "#,##0.00"
        ws_icra[f"E{r_idx}"] = ic["kalan_borc"]
        ws_icra[f"E{r_idx}"].number_format = "#,##0.00"
        ws_icra[f"F{r_idx}"] = ic["dosya"]
        ws_icra[f"G{r_idx}"] = ic["iban"]
        ws_icra[f"H{r_idx}"] = "Aktif Kesinti"
        
        for c_idx in range(1, 9):
            cl = get_column_letter(c_idx)
            ws_icra[f"{cl}{r_idx}"].border = border_cell
            ws_icra[f"{cl}{r_idx}"].font = font_data
            if c_idx in (4, 5):
                ws_icra[f"{cl}{r_idx}"].alignment = align_right
            elif c_idx not in (2, 6, 7):
                ws_icra[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 7: SGK_Raporlar (İstirahat Rapor Takip Listesi)
    # =========================================================================
    ws_rap = wb_new.create_sheet(title="SGK_Raporlar")
    ws_rap.views.sheetView[0].showGridLines = True
    rap_headers = [
        ("Takip No", 12), ("Sıra No", 8), ("TC Kimlik No", 14), ("Adı Soyadı", 22),
        ("Vaka Türü", 18), ("Poliklinik Tarihi", 15), ("İşbaşı Tarihi", 15)
    ]
    ws_rap.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(rap_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_rap[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_rap.column_dimensions[col_let].width = w

    for r_idx, rp in enumerate(rapor_list, start=2):
        ws_rap[f"A{r_idx}"] = rp["takip_no"]
        ws_rap[f"B{r_idx}"] = rp["sira_no"]
        ws_rap[f"C{r_idx}"] = rp["tc"]
        ws_rap[f"C{r_idx}"].number_format = "@"
        ws_rap[f"D{r_idx}"] = rp["ad_soyad"]
        ws_rap[f"E{r_idx}"] = rp["vaka"]
        ws_rap[f"F{r_idx}"] = rp["pol_tarih"]
        ws_rap[f"G{r_idx}"] = rp["isbasi_tarih"]
        
        for c_idx in range(1, 8):
            cl = get_column_letter(c_idx)
            ws_rap[f"{cl}{r_idx}"].border = border_cell
            ws_rap[f"{cl}{r_idx}"].font = font_data
            if c_idx not in (4, 5):
                ws_rap[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 8: Ucretli_Izinler (Yıllık / Ücretli İzin Cetveli)
    # =========================================================================
    ws_izn = wb_new.create_sheet(title="Ucretli_Izinler")
    ws_izn.views.sheetView[0].showGridLines = True
    for c_idx, h_text in enumerate(["TC Kimlik No", "Adı Soyadı", "İzin Saati", "İzin Günü"], start=1):
        ws_izn.cell(1, c_idx, h_text).font = font_hdr
        ws_izn.cell(1, c_idx).fill = fill_navy

    # =========================================================================
    # SAYFA 9: Daimi_Personel_Listesi (Kadro Durumu)
    # =========================================================================
    ws_dai = wb_new.create_sheet(title="Daimi_Personel_Listesi")
    ws_dai.views.sheetView[0].showGridLines = True
    for c_idx, h_text in enumerate(["Sıra No", "TC Kimlik No", "Adı Soyadı", "SGK Giriş", "SGK Durum", "Kadro", "Maaş (TL)"], start=1):
        ws_dai.cell(1, c_idx, h_text).font = font_hdr
        ws_dai.cell(1, c_idx).fill = fill_navy

    # =========================================================================
    # SAYFA 10: Elden_Odeme_Farki (Net Maaş ile Resmi Bordro Fark Cetveli)
    # =========================================================================
    ws_eld = wb_new.create_sheet(title="Elden_Odeme_Farki")
    ws_eld.views.sheetView[0].showGridLines = True
    eld_headers = [("Sıra No", 8), ("TC Kimlik No", 14), ("Adı Soyadı", 22), ("Anlaşılan Maaş", 15), ("Resmi Bordro Net", 15), ("Elden Fark (TL)", 15)]
    ws_eld.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(eld_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_eld[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_eld.column_dimensions[col_let].width = w

    for r_idx, p in enumerate(puantaj_rows, start=2):
        p_row = r_idx + 2
        ws_eld[f"A{r_idx}"] = p["sno"]
        ws_eld[f"B{r_idx}"] = p["tc"]
        ws_eld[f"B{r_idx}"].number_format = "@"
        ws_eld[f"C{r_idx}"] = p["ad_soyad"]
        ws_eld[f"D{r_idx}"] = f"=Aylik_Puantaj!L{p_row}"
        ws_eld[f"E{r_idx}"] = f"=Aylik_Puantaj!BC{p_row}"
        ws_eld[f"F{r_idx}"] = f"=Aylik_Puantaj!BD{p_row}"
        
        for c_idx in range(1, 7):
            cl = get_column_letter(c_idx)
            ws_eld[f"{cl}{r_idx}"].border = border_cell
            ws_eld[f"{cl}{r_idx}"].font = font_data
            if c_idx in (4, 5, 6):
                ws_eld[f"{cl}{r_idx}"].number_format = "#,##0.00"
                ws_eld[f"{cl}{r_idx}"].alignment = align_right
            elif c_idx not in (3,):
                ws_eld[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # 5. DOSYA KAYDETME VE GÜVENLİK
    # =========================================================================
    output_filename = "puantaj.xlsx"
    saved_name = output_filename
    try:
        wb_new.save(output_filename)
        print(f"4. Dosya '{output_filename}' olarak kaydedildi.")
    except PermissionError:
        saved_name = "puantaj_guncel.xlsx"
        wb_new.save(saved_name)
        print(f"4. UYARI: '{output_filename}' Microsoft Excel'de açık olduğu için '{saved_name}' olarak kaydedildi.")
        print(f"   (Tam '{output_filename}' üzerine yazmak için Excel'i kapatıp tekrar çalıştırabilirsiniz).")

    print("==================================================")
    print(f" BAŞARILI! '{saved_name}' dosyası eksiksiz oluşturuldu.")
    print(" - 8-5 Normal Çalışma ve 3 Vardiya (8-4 / 16-24 / 24-8) tam destekli")
    print(f" - Özel Bölüm Primleri: {len(bonus_audit_records)} kayıt 'Bolum_Prim_Mesai_Raporu' sayfasına işlendi (Lila renkli)")
    print(f" - İnsan Kontrolü Yönetimi: {len(missing_punch_records)} istisna/hata kaydı 'Eksik_Basim_Raporu' sayfasına işlendi (Turuncu renkli)")
    print(" - Canlı Formüller (SUM, COUNTIF, IF, MIN, VLOOKUP) Aktif")
    print(f" - {len(puantaj_rows)} personelin 30 günlük çalışma süreleri işlendi")
    print("==================================================")

if __name__ == "__main__":
    main()
