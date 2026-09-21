# -*- coding: utf-8 -*-
"""
===================================================================================================
PDKS ve Puantaj Çekirdek Hesaplama Motoru (pdks_engine.py)
===================================================================================================
Bu modül, ham PDKS (turnike) ve personel verilerini alıp İş Kanunu ve fabrika kurallarına göre
gün gün çalışma saati, fazla mesai, mola kesintileri ve bölüm primlerini en yüksek doğrulukla hesaplar.
"""

import xlrd
import re
import unicodedata
from collections import defaultdict
from datetime import datetime

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
    if not text:
        return ""
    s = str(text).strip()
    words = s.split()
    fixed_words = []
    for w in words:
        if w in WORD_REPLACEMENTS:
            fixed_words.append(WORD_REPLACEMENTS[w])
        else:
            w_fixed = w
            for k, v in WORD_REPLACEMENTS.items():
                if len(k) > 2:
                    w_fixed = w_fixed.replace(k, v)
                else:
                    w_fixed = re.sub(r'\b' + re.escape(k) + r'\b', v, w_fixed)
            fixed_words.append(w_fixed)
    s = " ".join(fixed_words)
    s = s.replace("\ufffd", "").replace("", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def get_day_name(date_str):
    """Tarih metninden (GG.AA.YYYY) Türkçe gün adını döndürür."""
    if not date_str:
        return ""
    try:
        parts = str(date_str).split(".")
        if len(parts) >= 3:
            d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
            dt = datetime(y, m, d)
            days = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
            return days[dt.weekday()]
    except:
        pass
    return ""

def norm_hdr(h):
    """Sütun başlıklarını eşleştirmek için normalize eder."""
    s = str(h).strip().lower()
    s = s.replace("ı", "i").replace("İ", "i").replace("ş", "s").replace("Ş", "s")
    s = s.replace("ç", "c").replace("Ç", "c").replace("ğ", "g").replace("Ğ", "g")
    s = s.replace("ü", "u").replace("Ü", "u").replace("ö", "o").replace("Ö", "o")
    return re.sub(r"[^a-z0-9]", "", s)

def resolve_pdks_columns(sh):
    """
    PDKS Excel sayfasındaki başlıkları dinamik olarak analiz eder ve sütun indekslerini döndürür.
    Böylece 8 sütunlu, 12 sütunlu, 29 sütunlu veya sütun sırası değişen tüm formatları hatasız destekler.
    """
    hdr_map = {}
    if sh.nrows > 0:
        for c in range(sh.ncols):
            hdr_map[norm_hdr(sh.cell_value(0, c))] = c

    def find_col(candidates, default=None):
        for cand in candidates:
            if cand in hdr_map:
                return hdr_map[cand]
        return default

    is_legacy_29 = sh.ncols >= 20

    return {
        "sira": find_col(["sira", "sirano", "sno"], 0 if is_legacy_29 else None),
        "sicil": find_col(["sicil", "sicilno", "tc", "tckimlik", "tckimlikno", "personelno"], 2 if is_legacy_29 else None),
        "kart": find_col(["kart", "kartno"], 3 if is_legacy_29 else None),
        "gun_adi": find_col(["gun", "gunadi", "gunler"], 4 if is_legacy_29 else None),
        "adi": find_col(["ad", "adi", "personeladi", "isim"], 5 if is_legacy_29 else 0),
        "soyadi": find_col(["soyad", "soyadi", "personelsoyadi"], 6 if is_legacy_29 else 1),
        "ad_soyad": find_col(["adsoyad", "personel", "adisoyadi"]),
        "lokasyon": find_col(["lokasyon", "lokasyonkodu", "kapi", "yer", "bolge"], 7 if is_legacy_29 else 2),
        "g_tarih": find_col(["giristarihi", "gtarih", "giristarih", "bastarih", "tarih"], 8 if is_legacy_29 else 3),
        "g_saat": find_col(["girissaati", "gsaat", "girissaat", "bassaat"], 9 if is_legacy_29 else 4),
        "c_tarih": find_col(["cikistarihi", "ctarih", "cikistarih", "bittarih"], 10 if is_legacy_29 else 5),
        "c_saat": find_col(["cikissaati", "csaat", "cikissaat", "bitsaat"], 11 if is_legacy_29 else 6),
        "puantaj_tarih": find_col(["puantajtarihi", "ptarih", "puantajtarih"], 18 if sh.ncols > 18 else None),
    }

def norm_name_key(s):
    if not s:
        return ""
    s = str(s).strip()
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"-.*$", "", s)
    s = s.replace("\ufffd", "I").replace("", "")
    s = unicodedata.normalize("NFKD", s)
    s = re.sub(r"[^a-zA-Z0-9]", "", s).upper()
    
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
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val)).strip()
    return clean_display_text(str(val).strip())

def parse_hm(t_str):
    if not t_str or ":" not in str(t_str):
        return None
    p = str(t_str).strip().split(":")
    try:
        return int(p[0]), int(p[1])
    except:
        return None

def parse_d_hm(d_str, t_str):
    if not d_str or not t_str or ":" not in str(t_str):
        return None
    try:
        p_d = str(d_str).strip().split(".")
        d, m, y = int(p_d[0]), int(p_d[1]), int(p_d[2])
        p_t = str(t_str).strip().split(":")
        h, mn = int(p_t[0]), int(p_t[1])
        return datetime(y, m, d, h, mn)
    except:
        return None

def fmt_hm(mins):
    h = (mins // 60) % 24
    m = mins % 60
    return f"{h:02d}:{m:02d}"

def fmt_hours_tr(val):
    if isinstance(val, (int, float)):
        return f"{val:.1f}".replace(".", ",")
    return str(val).replace(".", ",")

# =================================================================================================
# 2. VARDİYA & FAZLA MESAİ HESAPLAMA MOTORU (1.5 SAAT MOLA POLİTİKASI)
# =================================================================================================
def calculate_shift_hours(g_saat_str, c_saat_str):
    """
    Giriş ve çıkış saatlerine göre:
    - 08:20 toleransı
    - 1.5 saat standart yemek ve dinlenme molası
    - Kademeli fazla mesai (17:26 sonrası)
    - 3 vardiya desteğini hesaplar.
    
    Döndürdüğü Değerler (tuple):
    (total_hours, base_hours, ot_hours, break_hours, effective_g, effective_c, gross_hours)
    """
    g = parse_hm(g_saat_str)
    c = parse_hm(c_saat_str)
    
    if not g or not c:
        return 7.5, 7.5, 0.0, 0.0, "08:00", "17:00", 7.5
        
    g_min = g[0] * 60 + g[1]
    c_min = c[0] * 60 + c[1]
    
    # Çift basım / Anlık giriş-çıkış
    if abs(c_min - g_min) <= 3:
        return 7.5, 7.5, 0.0, 0.0, g_saat_str[:5], c_saat_str[:5], 0.05
        
    if c_min < g_min:
        c_min += 24 * 60
        
    gross_hours = round((c_min - g_min) / 60.0, 2)
    effective_c = fmt_hm(c_min)
    break_hours = 0.0
    
    # -------------------------------------------------------------------------
    # 1. GÜNDÜZ VARDİYASI (06:00 - 11:59 Giriş)
    # -------------------------------------------------------------------------
    if 6 * 60 <= g_min <= 11 * 60 + 59:
        if g_min <= 8 * 60 + 20:
            effective_start = 8 * 60
            base_hours = 7.5
        else:
            diff = g_min - (8 * 60 + 20)
            cuts = (diff - 1) // 30 + 1
            effective_start = 8 * 60 + cuts * 30
            base_hours = max(0.0, 7.5 - cuts * 0.5)
            
        effective_g = fmt_hm(effective_start)
        
        # 8-4 vardiyası tam bitiş aralığı (15:50 - 16:25)
        if 15 * 60 + 50 <= c_min <= 16 * 60 + 25:
            return 7.5, 7.5, 0.0, 0.5, effective_g, effective_c, gross_hours
            
        # 8-5 vardiyası bitişi (17:00) veya erken çıkış
        shift_end = 17 * 60
        if c_min < shift_end:
            elapsed_h = (c_min - effective_start) / 60.0
            if elapsed_h <= 0:
                return 0.0, 0.0, 0.0, 0.0, effective_g, effective_c, gross_hours
            # Öğle molası öncesi çıkış (<= 12:00): Mola kesintisi YOK (Tam fiili süre)
            if c_min <= 12 * 60:
                net = elapsed_h
                break_hours = 0.0
            # Öğle molası esnasında çıkış (12:01 - 13:00): 12:00'ye kadar olan net süre
            elif c_min < 13 * 60:
                net = max(0.0, (12 * 60 - effective_start) / 60.0)
                break_hours = round((c_min - 12 * 60) / 60.0, 2)
            # Öğleden sonra çıkış (13:00 ve sonrası): 1.5 saat mola kesilir
            else:
                net = max(0.0, elapsed_h - 1.5)
                break_hours = 1.5
                
            final_h = round(net * 2) / 2.0
            return final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours
            
        ot_min = c_min - shift_end
        break_hours = 1.5

    # -------------------------------------------------------------------------
    # 2. ÖĞLEDEN SONRA / KISMİ GÜNDÜZ ÇALIŞMASI (12:00 - 15:59 Giriş)
    # -------------------------------------------------------------------------
    elif 12 * 60 <= g_min < 16 * 60:
        effective_start = g_min
        effective_g = fmt_hm(g_min)
        elapsed_h = (c_min - g_min) / 60.0
        
        if c_min <= 17 * 60 + 25:
            net = elapsed_h
            break_hours = 0.0
            final_h = round(net * 2) / 2.0
            return final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours
        elif c_min < 24 * 60:
            net = max(0.0, elapsed_h - 0.5)
            break_hours = 0.5
            final_h = round(net * 2) / 2.0
            return final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours
        else:
            shift_end = 24 * 60
            ot_min = c_min - shift_end
            base_hours = 7.5
            break_hours = 0.5

    # -------------------------------------------------------------------------
    # 3. AKŞAM VARDİYASI (16:00 - 19:59 Giriş)
    # -------------------------------------------------------------------------
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
            
        effective_g = fmt_hm(effective_start)
        
        if c_min < shift_end:
            elapsed_h = (c_min - effective_start) / 60.0
            if elapsed_h <= 4.0:
                net = elapsed_h
                break_hours = 0.0
            else:
                net = max(0.0, elapsed_h - 0.5)
                break_hours = 0.5
            final_h = round(net * 2) / 2.0
            return final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours
            
        ot_min = c_min - shift_end
        break_hours = 0.5

    # -------------------------------------------------------------------------
    # 4. GECE VARDİYASI (20:00 - 05:59 Giriş)
    # -------------------------------------------------------------------------
    else:
        shift_end = 32 * 60 if g_min >= 20 * 60 else 8 * 60
        effective_start = 24 * 60 if g_min >= 20 * 60 else 0
        effective_g = fmt_hm(effective_start)
        base_hours = 7.5
        
        if c_min < shift_end:
            raw_w = (c_min - effective_start) / 60.0
            if raw_w <= 4.0:
                net = raw_w
                break_hours = 0.0
            else:
                net = max(0.0, raw_w - 0.5)
                break_hours = 0.5
            final_h = round(net * 2) / 2.0
            return final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours
            
        ot_min = c_min - shift_end
        break_hours = 0.5

    # -------------------------------------------------------------------------
    # Kademeli Fazla Mesai Dilimleri:
    # 0 - 25 dk   -> 0 saat FM
    # 26 - 49 dk  -> +0.5 saat FM
    # 50 - 85 dk  -> +1.0 saat FM
    # 86 - 109 dk -> +1.5 saat FM ...
    # -------------------------------------------------------------------------
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
            
    ot_hours = step * 0.5
    total_hours = base_hours + ot_hours
    return total_hours, base_hours, ot_hours, break_hours, effective_g, effective_c, gross_hours

# =================================================================================================
# 3. ANA MOTOR SINIFI (PDKSEngine)
# =================================================================================================
class PDKSEngine:
    def __init__(self, pdks_path="pdks.xls", puantaj_path="puantaj.xls", target_month=9, target_year=2026):
        self.pdks_path = pdks_path
        self.puantaj_path = puantaj_path
        self.target_month = target_month
        self.target_year = target_year
        
        self.personnel_list = []
        self.personnel_by_key = {}
        self.personnel_by_tc = {}
        self.dept_by_key = {}
        
        self.raw_daily_punches = defaultdict(lambda: {"rows": [], "meta": {}})
        self.daily_results = {}
        self.audit_records = []
        self.exception_records = []

    def load_personnel(self):
        """Puantaj dosyasından personel ana bilgilerini ve bölümlerini yükler."""
        wb = xlrd.open_workbook(self.puantaj_path, encoding_override="cp1254")
        sheet = None
        for s in wb.sheets():
            if "puantaj" in s.name.lower():
                sheet = s
                break
        if not sheet:
            sheet = wb.sheet_by_index(0)

        seen_tcs = set()
        for r in range(2, sheet.nrows):
            tc = clean_tc(sheet.cell_value(r, 1))
            ad_soyad = clean_str(sheet.cell_value(r, 2))
            bolum = clean_str(sheet.cell_value(r, 10))
            durumu = clean_str(sheet.cell_value(r, 9)) or "MEVSİMLİK"
            net_maas = sheet.cell_value(r, 12) or 0.0
            
            if not ad_soyad:
                continue
                
            name_k = norm_name_key(ad_soyad)
            if bolum:
                self.dept_by_key[name_k] = bolum
                
            if tc and tc in seen_tcs:
                continue
            if tc:
                seen_tcs.add(tc)
                
            p_data = {
                "sno": len(self.personnel_list) + 1,
                "tc": tc,
                "ad_soyad": ad_soyad,
                "name_key": name_k,
                "bolum": bolum,
                "durumu": durumu,
                "net_maas": net_maas
            }
            self.personnel_list.append(p_data)
            self.personnel_by_key[name_k] = p_data
            if tc:
                self.personnel_by_tc[tc] = p_data

        return len(self.personnel_list)

    def load_and_process_pdks(self):
        """PDKS hareketlerini okur, mükerrer basımları filtreler ve günlük hareketleri toplar."""
        wb_pdks = xlrd.open_workbook(self.pdks_path, encoding_override="cp1254")
        sh_pdks = None
        for s_name in wb_pdks.sheet_names():
            if "har" in s_name.lower() or "list" in s_name.lower():
                sh_pdks = wb_pdks.sheet_by_name(s_name)
                break
        if not sh_pdks:
            sh_pdks = wb_pdks.sheet_by_index(0)

        cols = resolve_pdks_columns(sh_pdks)

        def get_val(r_idx, col_key):
            c_idx = cols.get(col_key)
            if c_idx is not None and 0 <= c_idx < sh_pdks.ncols and c_idx < len(sh_pdks.row_values(r_idx)):
                return clean_str(sh_pdks.cell_value(r_idx, c_idx))
            return ""

        for r in range(1, sh_pdks.nrows):
            ad_soyad_single = get_val(r, "ad_soyad")
            if ad_soyad_single:
                full_name = ad_soyad_single
            else:
                adi = get_val(r, "adi")
                soyadi = get_val(r, "soyadi")
                full_name = f"{adi} {soyadi}".strip()
                
            if not full_name or full_name in ("BBBBBB", "DDDDD"):
                continue
            if re.match(r"^(.)\1+$", full_name.replace(" ", "")):
                continue
                
            name_k = norm_name_key(full_name)
            sira = get_val(r, "sira") or str(r)
            sicil = get_val(r, "sicil")
            kart = get_val(r, "kart")
            lokasyon = get_val(r, "lokasyon")
            g_tarih = get_val(r, "g_tarih")
            g_saat = get_val(r, "g_saat")
            c_tarih = get_val(r, "c_tarih")
            c_saat = get_val(r, "c_saat")
            p_tarih = get_val(r, "puantaj_tarih")
            target_date = p_tarih if p_tarih else (g_tarih if g_tarih else c_tarih)
            gun_adi = get_val(r, "gun_adi") or (get_day_name(target_date) if target_date else "")

            # Personel listesinde yoksa ekle
            if name_k not in self.personnel_by_key:
                p_data = {
                    "sno": len(self.personnel_list) + 1,
                    "tc": sicil if sicil else "",
                    "ad_soyad": full_name,
                    "name_key": name_k,
                    "bolum": "DİĞER",
                    "durumu": "MEVSİMLİK",
                    "net_maas": 0.0
                }
                self.personnel_list.append(p_data)
                self.personnel_by_key[name_k] = p_data

            # 18+ saatlik hatalı turnike birleşimlerini 2 ayrı güne ayrıştır
            if g_tarih and g_saat and c_tarih and c_saat:
                dt_g = parse_d_hm(g_tarih, g_saat)
                dt_c = parse_d_hm(c_tarih, c_saat)
                if dt_g and dt_c and (dt_c - dt_g).total_seconds() > 18 * 3600:
                    self._add_punch_to_day(g_tarih, sira, sicil, kart, gun_adi, full_name, name_k, lokasyon, g_saat, "")
                    c_gun_adi = get_day_name(c_tarih) if c_tarih else gun_adi
                    self._add_punch_to_day(c_tarih, sira, sicil, kart, c_gun_adi, full_name, name_k, lokasyon, "", c_saat)
                    continue

            # Normal geçerli hareket
            self._add_punch_to_day(target_date, sira, sicil, kart, gun_adi, full_name, name_k, lokasyon, g_saat, c_saat)

        # Günlük hareketleri hesapla
        self._calculate_daily_results()

    def _add_punch_to_day(self, date_str, sira, sicil, kart, gun_adi, full_name, name_k, lokasyon, g_saat, c_saat):
        if not date_str or "." not in str(date_str):
            return
        parts = str(date_str).split(".")
        try:
            d = int(parts[0])
            m = int(parts[1])
            if m != self.target_month:
                return
        except:
            return
            
        k = (name_k, d)
        if not self.raw_daily_punches[k]["meta"]:
            self.raw_daily_punches[k]["meta"] = {
                "sira": sira, "sicil": sicil, "kart": kart, "gun_adi": gun_adi,
                "full_name": full_name, "lokasyon": lokasyon, "date_str": date_str, "day": d
            }
        self.raw_daily_punches[k]["rows"].append({"g_saat": g_saat, "c_saat": c_saat})

    def _calculate_daily_results(self):
        """Toplanan hareketleri kural motoruna sokar ve tüm denetim izlerini üretir."""
        for (name_k, day_num), data in self.raw_daily_punches.items():
            rows = data["rows"]
            meta = data["meta"]
            full_name = meta["full_name"]
            date_str = meta["date_str"]
            gun_adi = meta["gun_adi"]
            kart = meta["kart"]
            sicil = meta["sicil"]
            lokasyon = meta["lokasyon"]
            
            p_info = self.personnel_by_key.get(name_k, {})
            dept_name = p_info.get("bolum", "").upper()
            
            norm_dept = dept_name.replace("İ", "I").replace("Ü", "U").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").strip()
            is_balik_dk = ("BALIK DOLUM" in norm_dept) or ("BALIK KESIM" in norm_dept) or ("BALIK KES" in norm_dept)
            is_uretim = (norm_dept == "URETIM") or ("URETIM ELEMANI" in norm_dept) or ("KONSERVE URETIM" in norm_dept)
            
            all_punches = []
            for r in rows:
                if r["g_saat"]:
                    all_punches.append(r["g_saat"][:5])
                if r["c_saat"]:
                    all_punches.append(r["c_saat"][:5])
            all_punches_str = ", ".join(all_punches)
            
            status_type = "NORMAL"
            is_exception = False
            audit_note = ""
            bonus_hours = 0.0
            
            # 1. DURUM: Çift Basım (Tek satır, hem giriş hem çıkış dolu)
            if len(rows) == 1 and rows[0]["g_saat"] and rows[0]["c_saat"]:
                g_raw = rows[0]["g_saat"]
                c_raw = rows[0]["c_saat"]
                tot_h, base_h, ot_h, brk_h, eff_g, eff_c, gross_h = calculate_shift_hours(g_raw, c_raw)
                fiili_sure = tot_h
                raw_g_str = g_raw[:5]
                raw_c_str = c_raw[:5]
                
            # 2. DURUM: Çoklu Basım (Birden fazla satır)
            elif len(rows) > 1:
                g_list = [r["g_saat"] for r in rows if r["g_saat"]]
                c_list = [r["c_saat"] for r in rows if r["c_saat"]]
                
                if g_list and c_list:
                    g_raw = g_list[0]
                    c_raw = c_list[-1]
                    raw_g_str = g_raw[:5]
                    raw_c_str = c_raw[:5]
                    tot_h, base_h, ot_h, brk_h, eff_g, eff_c, gross_h = calculate_shift_hours(g_raw, c_raw)
                    fiili_sure = tot_h
                    if len(all_punches) > 2:
                        status_type = "ÇOKLU_BASIM_MIN_MAX"
                        is_exception = True
                        audit_note = f"Çoklu Basım: İlk Giriş {raw_g_str}, Son Çıkış {raw_c_str} ({fmt_hours_tr(tot_h)}s)"
                elif g_list:
                    # Sadece girişler var
                    g_raw = g_list[0]
                    raw_g_str = g_raw[:5]
                    raw_c_str = "-"
                    hm = parse_hm(g_raw)
                    tm = hm[0] * 60 + hm[1] if hm else 480
                    is_exception = True
                    
                    if tm <= 12 * 60 + 30:
                        status_type = "ÇIKIŞ_YOK_SABAH"
                        tot_h, base_h, ot_h, brk_h = 7.5, 7.5, 0.0, 1.5
                        eff_g, eff_c, gross_h = raw_g_str, "17:00", 7.5
                        audit_note = f"Giriş: {raw_g_str} | Çıkış Basılmadı (Tam Gün 7,5s yazıldı)"
                    elif 12 * 60 + 30 < tm < 20 * 60:
                        status_type = "GİRİŞ_YOK_AKŞAM"
                        tot_h, base_h, ot_h, brk_h, eff_g, eff_c, gross_h = calculate_shift_hours("08:00", g_raw)
                        audit_note = f"Çıkış: {raw_g_str} | Giriş Basılmadı (08:00 başı ile {fmt_hours_tr(tot_h)}s korundu)"
                    else:
                        status_type = "GECE_TEK_BASIM"
                        tot_h, base_h, ot_h, brk_h = 7.5, 7.5, 0.0, 0.5
                        eff_g, eff_c, gross_h = raw_g_str, "08:00", 7.5
                        audit_note = f"Gece Girişi: {raw_g_str} | Çıkış Basılmadı (7,5s yazıldı)"
                    fiili_sure = tot_h
                else:
                    tot_h, base_h, ot_h, brk_h, eff_g, eff_c, gross_h, fiili_sure = 0.0, 0.0, 0.0, 0.0, "-", "-", 0.0, 0.0
                    raw_g_str, raw_c_str = "-", "-"
                    
            # 3. DURUM: Tek Satır, Tek Basım (Eksik Basım)
            elif len(rows) == 1:
                r_s = rows[0]
                t_raw = r_s["g_saat"] if r_s["g_saat"] else r_s["c_saat"]
                hm = parse_hm(t_raw)
                tm = hm[0] * 60 + hm[1] if hm else 480
                t_s = t_raw[:5] if len(t_raw) >= 5 else t_raw
                is_exception = True
                
                if tm <= 12 * 60 + 30:
                    raw_g_str = t_s
                    raw_c_str = "-"
                    status_type = "ÇIKIŞ_YOK_SABAH"
                    tot_h, base_h, ot_h, brk_h = 7.5, 7.5, 0.0, 1.5
                    eff_g, eff_c, gross_h = t_s, "17:00", 7.5
                    audit_note = f"Giriş: {t_s} | Çıkış Basılmadı (Tam Gün 7,5s yazıldı)"
                elif 12 * 60 + 30 < tm < 20 * 60:
                    raw_g_str = "-"
                    raw_c_str = t_s
                    status_type = "GİRİŞ_YOK_AKŞAM"
                    tot_h, base_h, ot_h, brk_h, eff_g, eff_c, gross_h = calculate_shift_hours("08:00", t_raw)
                    audit_note = f"Çıkış: {t_s} | Giriş Basılmadı (08:00 başı ile {fmt_hours_tr(tot_h)}s korundu)"
                else:
                    raw_g_str = t_s
                    raw_c_str = "-"
                    status_type = "GECE_TEK_BASIM"
                    tot_h, base_h, ot_h, brk_h = 7.5, 7.5, 0.0, 0.5
                    eff_g, eff_c, gross_h = t_s, "08:00", 7.5
                    audit_note = f"Gece Basımı: {t_s} | Tek Basım (7,5s yazıldı)"
                fiili_sure = tot_h
            else:
                tot_h, base_h, ot_h, brk_h, eff_g, eff_c, gross_h, fiili_sure = 0.0, 0.0, 0.0, 0.0, "-", "-", 0.0, 0.0
                raw_g_str, raw_c_str = "-", "-"

            # -----------------------------------------------------------------
            # ÖZEL BÖLÜM PRİMLERİ (Balık Dolum/Kesim ve Üretim):
            # -----------------------------------------------------------------
            if is_balik_dk and fiili_sure >= 10.0:
                bonus_hours = 2.0
                tot_h += bonus_hours
                ot_h += bonus_hours
                status_type = "BÖLÜM_PRİMLİ"
                prim_msg = f"{dept_name} Primi: Fiili {fmt_hours_tr(fiili_sure)}s -> Bonuslu {fmt_hours_tr(tot_h)}s (+2,0s Prim)"
                audit_note = f"{audit_note} | {prim_msg}" if audit_note else prim_msg
            elif is_uretim and fiili_sure >= 12.0:
                bonus_hours = 4.0
                tot_h += bonus_hours
                ot_h += bonus_hours
                status_type = "BÖLÜM_PRİMLİ"
                prim_msg = f"Üretim Primi: Fiili {fmt_hours_tr(fiili_sure)}s -> Bonuslu {fmt_hours_tr(tot_h)}s (+4,0s Prim)"
                audit_note = f"{audit_note} | {prim_msg}" if audit_note else prim_msg

            rec = {
                "tarih": date_str,
                "gun": day_num,
                "gun_adi": gun_adi,
                "sicil": sicil,
                "kart": kart,
                "ad_soyad": full_name,
                "tc": p_info.get("tc", sicil),
                "bolum": dept_name,
                "lokasyon": lokasyon,
                "raw_g_saat": raw_g_str,
                "raw_c_saat": raw_c_str,
                "all_punches": all_punches_str,
                "effective_g": eff_g,
                "effective_c": eff_c,
                "gross_hours": gross_h,
                "break_hours": brk_h,
                "fiili_sure": fiili_sure,
                "base_hours": base_h,
                "overtime_hours": ot_h,
                "bonus_hours": bonus_hours,
                "final_hours": tot_h,
                "status_type": status_type,
                "is_exception": is_exception,
                "audit_note": audit_note
            }
            
            self.daily_results[(name_k, day_num)] = rec
            self.audit_records.append(rec)
            if is_exception:
                self.exception_records.append(rec)

    def get_summary_matrix(self):
        """Personel bazında 1..30 günlük saat matrisini ve toplamlarını döndürür."""
        matrix = []
        for p in self.personnel_list:
            name_k = p["name_key"]
            row = {
                "sno": p["sno"],
                "tc": p["tc"],
                "ad_soyad": p["ad_soyad"],
                "name_key": name_k,
                "bolum": p["bolum"],
                "durumu": p["durumu"],
                "daily_hours": {},
                "total_work_days": 0,
                "total_base_hours": 0.0,
                "total_ot_hours": 0.0,
                "total_bonus_hours": 0.0,
                "total_hours": 0.0
            }
            
            for d in range(1, 31):
                rec = self.daily_results.get((name_k, d))
                if rec and rec["final_hours"] > 0:
                    h = rec["final_hours"]
                    row["daily_hours"][d] = h
                    row["total_work_days"] += 1
                    row["total_base_hours"] += rec["base_hours"]
                    row["total_ot_hours"] += rec["overtime_hours"]
                    row["total_bonus_hours"] += rec["bonus_hours"]
                    row["total_hours"] += h
                else:
                    row["daily_hours"][d] = 0.0
                    
            matrix.append(row)
        return matrix
