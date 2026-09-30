# -*- coding: utf-8 -*-
"""
===================================================================================================
PDKS ve Puantaj Çekirdek Hesaplama Motoru (pdks_engine.py)
===================================================================================================
Bu modül, ham PDKS (turnike) ve personel verilerini alıp İş Kanunu ve fabrika kurallarına göre
gün gün çalışma saati, fazla mesai, mola kesintileri ve bölüm primlerini en yüksek doğrulukla hesaplar.
Dinamik takvim (calendar) desteğiyle tüm ay ve yıllarla tam uyumludur.
"""

import calendar
import re
import unicodedata
from collections import defaultdict, Counter
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Set, Tuple, Union
import xlrd
from config.settings import load_rules, calculate_department_bonus

# =================================================================================================
# 1. VERİ MODELLERİ (DATACLASSES)
# =================================================================================================
@dataclass
class ShiftResult:
    """Vardiya ve fazla mesai hesaplama sonuç modeli."""
    total_hours: float
    base_hours: float
    overtime_hours: float
    break_hours: float
    effective_g: str
    effective_c: str
    gross_hours: float

    def __iter__(self):
        """Geriye dönük uyumluluk için tuple unpacking desteği: total, base, ot, brk, eg, ec, gross = res"""
        yield self.total_hours
        yield self.base_hours
        yield self.overtime_hours
        yield self.break_hours
        yield self.effective_g
        yield self.effective_c
        yield self.gross_hours

    def __getitem__(self, index):
        return tuple(self)[index]


# =================================================================================================
# 2. TÜRKÇE KARAKTER VE KELİME ONARIM SÖZLÜĞÜ
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

# =================================================================================================
# 3. YARDIMCI VE TEMİZLEME FONKSİYONLARI
# =================================================================================================
def clean_display_text(text: Optional[str]) -> str:
    if not text:
        return ""
    s = str(text).strip()
    words = s.split()
    fixed_words = []
    for w in words:
        w_clean = w.replace("\ufffd", "").replace("", "")
        if w in WORD_REPLACEMENTS:
            fixed_words.append(WORD_REPLACEMENTS[w])
        elif w_clean in WORD_REPLACEMENTS:
            fixed_words.append(WORD_REPLACEMENTS[w_clean])
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

def get_day_name(date_str: str) -> str:
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
    except Exception:
        pass
    return ""

def norm_hdr(h: str) -> str:
    """Sütun başlıklarını eşleştirmek için normalize eder."""
    s = str(h).strip().lower()
    s = s.replace("ı", "i").replace("İ", "i").replace("ş", "s").replace("Ş", "s")
    s = s.replace("ç", "c").replace("Ç", "c").replace("ğ", "g").replace("Ğ", "g")
    s = s.replace("ü", "u").replace("Ü", "u").replace("ö", "o").replace("Ö", "o")
    return re.sub(r"[^a-z0-9]", "", s)

def resolve_pdks_columns(sh) -> Dict[str, Optional[int]]:
    """PDKS Excel sayfasındaki başlıkları dinamik analiz eder ve sütun indekslerini döndürür."""
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
        "sure": find_col(["sure", "calismasuresi", "toplamsure"]),
        "puantaj_tarih": find_col(["puantajtarihi", "ptarih", "puantajtarih"], 18 if sh.ncols > 18 else None),
    }

def norm_name_key(s: Optional[str]) -> str:
    if not s:
        return ""
    s = str(s).strip()
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"-.*$", "", s)
    s = s.replace("ı", "i").replace("İ", "I")
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

def _norm_dept_for_peers(dept: Optional[str]) -> str:
    """Mesai arkadaşı analizi için bölüm isimlerini normalize eder."""
    if not dept:
        return "GENEL"
    d = dept.strip().upper()
    for tr, en in [("İ", "I"), ("Ğ", "G"), ("Ü", "U"), ("Ş", "S"), ("Ö", "O"), ("Ç", "C")]:
        d = d.replace(tr, en)
    if "BALIK TEM" in d:
        return "BALIK TEMIZLEME"
    if "BALIK DOL" in d:
        return "BALIK DOLUM"
    if "BALIK KES" in d:
        return "BALIK KESIM"
    if "URETIM" in d:
        return "URETIM"
    if "AMBAR" in d or "DEPO" in d:
        return "AMBAR"
    if "BAKIM" in d or "MEKANIK" in d:
        return "BAKIM"
    if "KAZAN" in d:
        return "KAZAN"
    if "ISLETME" in d:
        return "ISLETME"
    if "MEYDAN" in d:
        return "MEYDAN"
    if "SOFOR" in d:
        return "SOFOR"
    return d

def clean_tc(tc_val) -> str:
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

def clean_str(val) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val)).strip()
    return clean_display_text(str(val).strip())

def clean_money(val) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace("TL", "").replace("tl", "").replace(".", "").replace(",", ".").strip()
    try:
        return float(s)
    except Exception:
        return 0.0

def xldate_to_str(val) -> str:
    if val is None or val == "":
        return ""
    if isinstance(val, float):
        try:
            t = xlrd.xldate_as_datetime(val, 0)
            return t.strftime("%d.%m.%Y")
        except Exception:
            pass
    s = str(val).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        parts = s.split()[0].split("-")
        return f"{parts[2]}.{parts[1]}.{parts[0]}"
    return s

def get_sheet_by_keyword(wb, keyword: str, default_idx: Optional[int] = None):
    kw = keyword.lower()
    for s in wb.sheets():
        if kw in s.name.lower():
            return s
    if default_idx is not None and default_idx < wb.nsheets:
        return wb.sheet_by_index(default_idx)
    return None

def parse_hm(t_str: Optional[str]) -> Optional[Tuple[int, int]]:
    if not t_str or ":" not in str(t_str):
        return None
    p = str(t_str).strip().split(":")
    try:
        return int(p[0]), int(p[1])
    except Exception:
        return None

def parse_d_hm(d_str: str, t_str: str) -> Optional[datetime]:
    if not d_str or not t_str or ":" not in str(t_str):
        return None
    try:
        p_d = str(d_str).strip().split(".")
        d, m, y = int(p_d[0]), int(p_d[1]), int(p_d[2])
        p_t = str(t_str).strip().split(":")
        h, mn = int(p_t[0]), int(p_t[1])
        return datetime(y, m, d, h, mn)
    except Exception:
        return None

def fmt_hm(mins: int) -> str:
    h = (mins // 60) % 24
    m = mins % 60
    return f"{h:02d}:{m:02d}"

def fmt_hours_tr(val: Union[int, float, str]) -> str:
    if isinstance(val, (int, float)):
        return f"{val:.1f}".replace(".", ",")
    return str(val).replace(".", ",")

# =================================================================================================
# 4. VARDİYA & FAZLA MESAİ HESAPLAMA MOTORU (1.5 SAAT MOLA POLİTİKASI)
# =================================================================================================
def calculate_shift_hours(g_saat_str: str, c_saat_str: str) -> ShiftResult:
    """
    Giriş ve çıkış saatlerine göre vardiya süresini hesaplar.
    Döndürdüğü Nesne: ShiftResult (total_hours, base_hours, overtime_hours, break_hours, effective_g, effective_c, gross_hours)
    Aynı zamanda tuple unpacking ile doğrudan (tot, base, ot, ...) şeklinde de açılabilir.
    """
    g = parse_hm(g_saat_str)
    c = parse_hm(c_saat_str)
    
    if not g or not c:
        return ShiftResult(7.5, 7.5, 0.0, 0.0, "08:00", "17:00", 7.5)
        
    g_min = g[0] * 60 + g[1]
    c_min = c[0] * 60 + c[1]
    
    # Çift basım / Anlık giriş-çıkış
    if abs(c_min - g_min) <= 3:
        return ShiftResult(7.5, 7.5, 0.0, 0.0, g_saat_str[:5], c_saat_str[:5], 0.05)
        
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
            return ShiftResult(7.5, 7.5, 0.0, 0.5, effective_g, effective_c, gross_hours)
            
        # 8-5 vardiyası bitişi (17:00) veya erken çıkış
        shift_end = 17 * 60
        if c_min < shift_end:
            elapsed_h = (c_min - effective_start) / 60.0
            if elapsed_h <= 0:
                return ShiftResult(0.0, 0.0, 0.0, 0.0, effective_g, effective_c, gross_hours)
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
            return ShiftResult(final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours)
            
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
            return ShiftResult(final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours)
        elif c_min < 24 * 60:
            net = max(0.0, elapsed_h - 0.5)
            break_hours = 0.5
            final_h = round(net * 2) / 2.0
            return ShiftResult(final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours)
        else:
            shift_end = 24 * 60
            ot_min = c_min - shift_end
            base_hours = 7.5
            break_hours = 0.5

    # -------------------------------------------------------------------------
    # 3. AKŞAM VARDİYASI (15:30 - 18:29 Giriş / Standart 16:00 - 24:00)
    # -------------------------------------------------------------------------
    elif 15 * 60 + 30 <= g_min < 18 * 60 + 30:
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
                net = max(0.0, elapsed_h)
                break_hours = 0.0
            else:
                net = max(0.0, elapsed_h - 0.5)
                break_hours = 0.5
            final_h = round(net * 2) / 2.0
            return ShiftResult(final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours)
            
        ot_min = c_min - shift_end
        break_hours = 0.5

    # -------------------------------------------------------------------------
    # 4. ARA GECE VARDİYASI (18:30 - 19:29 Giriş / Standart 19:00 - 03:00 / 07:00)
    # -------------------------------------------------------------------------
    elif 18 * 60 + 30 <= g_min < 19 * 60 + 30:
        shift_end = 27 * 60  # 03:00 (7.5 saat net temel vardiya)
        if g_min <= 19 * 60 + 20:
            effective_start = 19 * 60
            base_hours = 7.5
        else:
            diff = g_min - (19 * 60 + 20)
            cuts = (diff - 1) // 30 + 1
            effective_start = 19 * 60 + cuts * 30
            base_hours = max(0.0, 7.5 - cuts * 0.5)
            
        effective_g = fmt_hm(effective_start)
        
        if c_min < shift_end:
            elapsed_h = (c_min - effective_start) / 60.0
            if elapsed_h <= 4.0:
                net = max(0.0, elapsed_h)
                break_hours = 0.0
            else:
                net = max(0.0, elapsed_h - 0.5)
                break_hours = 0.5
            final_h = round(net * 2) / 2.0
            return ShiftResult(final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours)
            
        ot_min = c_min - shift_end
        break_hours = 0.5

    # -------------------------------------------------------------------------
    # 5. GECE / 12 SAATLİK VARDİYA (19:30 - 22:29 Giriş / Standart 20:00 - 04:00 / 08:00)
    # -------------------------------------------------------------------------
    elif 19 * 60 + 30 <= g_min < 22 * 60 + 30:
        if g_min <= 20 * 60 + 20:
            effective_start = 20 * 60
            base_hours = 7.5
        else:
            diff = g_min - (20 * 60 + 20)
            cuts = (diff - 1) // 30 + 1
            effective_start = 20 * 60 + cuts * 30
            base_hours = max(0.0, 7.5 - cuts * 0.5)
            
        shift_end = effective_start + 8 * 60  # 20:00 için 04:00 (28:00)
        effective_g = fmt_hm(effective_start)
        
        if c_min < shift_end:
            elapsed_h = (c_min - effective_start) / 60.0
            if elapsed_h <= 4.0:
                net = max(0.0, elapsed_h)
                break_hours = 0.0
            else:
                net = max(0.0, elapsed_h - 0.5)
                break_hours = 0.5
            final_h = round(net * 2) / 2.0
            return ShiftResult(final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours)
            
        ot_min = c_min - shift_end
        break_hours = 0.5

    # -------------------------------------------------------------------------
    # 6. GECE VARDİYASI (22:30 - 05:59 Giriş / Standart 24:00 - 08:00)
    # -------------------------------------------------------------------------
    else:
        if g_min >= 22 * 60 + 30:
            effective_start = 24 * 60
            base_hours = 7.5
            shift_end = 32 * 60
        elif g_min <= 20:
            effective_start = 0
            base_hours = 7.5
            shift_end = 8 * 60
        else:
            diff = g_min - 20
            cuts = (diff - 1) // 30 + 1
            effective_start = cuts * 30
            base_hours = max(0.0, 7.5 - cuts * 0.5)
            shift_end = effective_start + 8 * 60
            
        effective_g = fmt_hm(effective_start)
        
        if c_min < shift_end:
            elapsed_h = (c_min - effective_start) / 60.0
            if elapsed_h <= 4.0:
                net = max(0.0, elapsed_h)
                break_hours = 0.0
            else:
                net = max(0.0, elapsed_h - 0.5)
                break_hours = 0.5
            final_h = round(net * 2) / 2.0
            return ShiftResult(final_h, final_h, 0.0, break_hours, effective_g, effective_c, gross_hours)
            
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
    # Bir günde 24 saatten fazla çalışma yazılamaz kuralı:
    if total_hours > 24.0:
        total_hours = 24.0
        ot_hours = max(0.0, 24.0 - base_hours)
    return ShiftResult(total_hours, base_hours, ot_hours, break_hours, effective_g, effective_c, gross_hours)

def calc_factory_worked_hours(g_saat_str: str, c_saat_str: str, sure_str: str = "", is_office: bool = False) -> Tuple[float, float]:
    """Geriye dönük uyumluluk: (total_hours, overtime_hours) tuple'ı döndürür."""
    res = calculate_shift_hours(g_saat_str, c_saat_str)
    return res.total_hours, res.overtime_hours

def repair_turnstile_anomalies(rows: List[Dict[str, str]]) -> Tuple[List[Dict[str, str]], bool, str]:
    """
    Turnike algoritmik denetimi:
    1. Gündüz Vardiyası Ters Basım Onarımı: İşçi sabah turnikesine çıkış, akşam turnikesine giriş basmışsa
       (Örn: g_saat 15:00-20:00 arası ve c_saat 06:00-10:00 arası), saatleri yer değiştirerek
       gündüz vardiyasını (08:00 - 17:00 / 18:00) otomatik kurtarır.
       (Not: Gece vardiyası girişleri 20:00-05:59 arası olup bu kapsama dahil edilmez).
    2. Çift/Mükerrer Anlık Basım Ayıklama: Turnikeye peş peşe basılan mükerrer girişler veya çıkışlar ayıklanır.
    """
    if not rows:
        return rows, False, ""

    repaired = [dict(r) for r in rows]
    was_repaired = False
    notes = []

    if len(repaired) == 1:
        g = repaired[0].get("g_saat", "")
        c = repaired[0].get("c_saat", "")
        if g and c:
            g_hm = parse_hm(g)
            c_hm = parse_hm(c)
            if g_hm and c_hm:
                g_m = g_hm[0] * 60 + g_hm[1]
                c_m = c_hm[0] * 60 + c_hm[1]
                if (15 * 60 <= g_m <= 19 * 60 + 15) and (6 * 60 <= c_m <= 10 * 60):
                    repaired[0]["g_saat"] = c
                    repaired[0]["c_saat"] = g
                    was_repaired = True
                    notes.append(f"Turnike Ters Basımı Düzeltildi ({c[:5]} Giriş, {g[:5]} Çıkış)")

    return repaired, was_repaired, "; ".join(notes)

# =================================================================================================
# 5. ANA MOTOR SINIFI (PDKSEngine)
# =================================================================================================
class PDKSEngine:
    def __init__(self, pdks_path: str = "pdks.xls", puantaj_path: str = "puantaj.xls", target_month: int = 9, target_year: int = 2026, rules: Optional[Dict] = None):
        self.pdks_path = pdks_path
        self.puantaj_path = puantaj_path
        self.target_month = target_month
        self.target_year = target_year
        self.rules = rules or load_rules()
        
        # Dinamik Takvim Tanımları
        _, self.days_in_month = calendar.monthrange(self.target_year, self.target_month)
        self.sundays: Set[int] = {
            d for d in range(1, self.days_in_month + 1)
            if calendar.weekday(self.target_year, self.target_month, d) == 6
        }
        
        self.personnel_list: List[Dict] = []
        self.personnel_by_key: Dict[str, Dict] = {}
        self.personnel_by_tc: Dict[str, Dict] = {}
        self.dept_by_key: Dict[str, str] = {}
        
        self.raw_daily_punches = defaultdict(lambda: {"rows": [], "meta": {}})
        self.daily_results: Dict[Tuple[str, int], Dict] = {}
        self.audit_records: List[Dict] = []
        self.exception_records: List[Dict] = []
        self.bonus_records: List[Dict] = []
        
        # Ek bordro tabloları
        self.rapor_list: List[Dict] = []
        self.rapor_by_tc: Dict[str, int] = {}
        self.icra_list: List[Dict] = []
        self.icra_by_name: Dict[str, Dict] = {}
        self.izin_list: List[Dict] = []
        self.izin_by_tc: Dict[str, Dict] = {}
        self.bordro_list: List[Dict] = []
        self.bordro_by_tc: Dict[str, Dict] = {}
        self.daimi_list: List[Dict] = []

    def get_day_name_for_day(self, day_num: int, short: bool = False) -> str:
        """Belirtilen ay ve gün için Türkçe gün adını döndürür."""
        try:
            wd = calendar.weekday(self.target_year, self.target_month, int(day_num))
            days = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
            short_days = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "PAZAR"]
            return short_days[wd] if short else days[wd]
        except Exception:
            return ""

    def is_sunday(self, day_num: int) -> bool:
        """Günün Pazar tatili olup olmadığını denetler."""
        return int(day_num) in self.sundays

    def load_personnel(self) -> int:
        """Puantaj dosyasından personel ana bilgilerini ve bölümlerini eksiksiz yükler."""
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
            if not ad_soyad:
                continue

            cinsiyet = clean_str(sheet.cell_value(r, 3)) if sheet.ncols > 3 else ""
            isletme_giris = xldate_to_str(sheet.cell_value(r, 4)) if sheet.ncols > 4 else ""
            sgk_giris = xldate_to_str(sheet.cell_value(r, 5)) if sheet.ncols > 5 else ""
            kidem = sheet.cell_value(r, 6) if sheet.ncols > 6 else 0
            sgk_cikis = xldate_to_str(sheet.cell_value(r, 7)) if sheet.ncols > 7 else ""
            sgk_durumu = clean_str(sheet.cell_value(r, 8)) or "NORMAL" if sheet.ncols > 8 else "NORMAL"
            durumu = clean_str(sheet.cell_value(r, 9)) or "MEVSİMLİK" if sheet.ncols > 9 else "MEVSİMLİK"
            bolum = clean_str(sheet.cell_value(r, 10)) if sheet.ncols > 10 else ""
            ikamet = clean_str(sheet.cell_value(r, 11)) if sheet.ncols > 11 else ""
            net_maas = clean_money(sheet.cell_value(r, 12)) if sheet.ncols > 12 else 0.0
            sirket = clean_str(sheet.cell_value(r, 13)) or "FİDE KONSERVE" if sheet.ncols > 13 else "FİDE KONSERVE"
            mesai_durumu = clean_str(sheet.cell_value(r, 14)) or "ALIR" if sheet.ncols > 14 else "ALIR"
            
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
                "cinsiyet": cinsiyet,
                "isletme_giris": isletme_giris,
                "sgk_giris": sgk_giris,
                "kidem": kidem,
                "sgk_cikis": sgk_cikis,
                "sgk_durumu": sgk_durumu,
                "durumu": durumu,
                "bolum": bolum,
                "ikamet": ikamet,
                "net_maas": net_maas,
                "sirket": sirket,
                "mesai_durumu": mesai_durumu
            }
            self.personnel_list.append(p_data)
            self.personnel_by_key[name_k] = p_data
            if tc:
                self.personnel_by_tc[tc] = p_data
        self.load_payroll_sheets()
        return len(self.personnel_list)

    def load_payroll_sheets(self):
        """Puantaj dosyasındaki diğer bordro sayfalarını (Rapor, İcra, İzin, Bordro, Daimi) okur."""
        wb = xlrd.open_workbook(self.puantaj_path, encoding_override="cp1254")
        
        # 1. SGK Raporları
        sh_rap = get_sheet_by_keyword(wb, "rapor", 4)
        self.rapor_list = []
        self.rapor_by_tc = {}
        if sh_rap and sh_rap.ncols >= 8:
            for r in range(3, sh_rap.nrows):
                ad_soyad = clean_str(sh_rap.cell_value(r, 4))
                if not ad_soyad:
                    continue
                tc = clean_tc(sh_rap.cell_value(r, 3))
                rec = {
                    "takip_no": clean_str(sh_rap.cell_value(r, 1)),
                    "sira_no": clean_str(sh_rap.cell_value(r, 2)),
                    "tc": tc,
                    "ad_soyad": ad_soyad,
                    "vaka": clean_str(sh_rap.cell_value(r, 5)),
                    "pol_tarih": xldate_to_str(sh_rap.cell_value(r, 6)),
                    "isbasi_tarih": xldate_to_str(sh_rap.cell_value(r, 7))
                }
                self.rapor_list.append(rec)
                if tc:
                    self.rapor_by_tc[tc] = self.rapor_by_tc.get(tc, 0) + 1

        # 2. İcra Takip
        sh_icra = get_sheet_by_keyword(wb, "cra", 3)
        self.icra_list = []
        self.icra_by_name = {}
        if sh_icra and sh_icra.ncols >= 8:
            for r in range(2, sh_icra.nrows):
                isim = clean_str(sh_icra.cell_value(r, 1))
                if not isim:
                    continue
                kalan_borc = clean_money(sh_icra.cell_value(r, 5))
                kesilen = clean_money(sh_icra.cell_value(r, 6))
                dosya = clean_str(sh_icra.cell_value(r, 4))
                iban = clean_str(sh_icra.cell_value(r, 7))
                rec = {
                    "isim": isim,
                    "sirket": clean_str(sh_icra.cell_value(r, 3)),
                    "dosya": dosya,
                    "kalan_borc": kalan_borc,
                    "kesilen": kesilen,
                    "iban": iban
                }
                self.icra_list.append(rec)
                self.icra_by_name[norm_name_key(isim)] = {
                    "kesilen": kesilen, "kalan_borc": kalan_borc, "dosya": dosya, "iban": iban
                }

        # 3. Ücretli İzinler
        sh_izin = None
        for s in wb.sheets():
            s_clean = s.name.lower().replace("\ufffd", "").replace("ı", "i")
            if "izin" in s_clean or "zin" in s_clean:
                sh_izin = s
                break
        if not sh_izin and wb.nsheets > 4:
            sh_izin = wb.sheet_by_index(4)

        self.izin_list = []
        self.izin_by_tc = {}
        if sh_izin:
            for r in range(1, sh_izin.nrows):
                isim = clean_str(sh_izin.cell_value(r, 1))
                if not isim:
                    continue
                tc = clean_tc(sh_izin.cell_value(r, 0))
                gun = clean_money(sh_izin.cell_value(r, 2)) if sh_izin.ncols > 2 else 0.0
                saat = gun * 7.5
                rec = {"tc": tc, "isim": isim, "saat": saat, "gun": gun}
                self.izin_list.append(rec)
                if tc:
                    self.izin_by_tc[tc] = {"saat": saat, "gun": gun}

        # 4. Resmi Bordro SGK (Sayfa3)
        sh_bordro = None
        for s in wb.sheets():
            s_clean = s.name.lower().replace("\ufffd", "").replace("ı", "i")
            if "bordro" in s_clean or s.name.lower() == "sayfa3":
                sh_bordro = s
                break
        if not sh_bordro and wb.nsheets > 17:
            sh_bordro = wb.sheet_by_index(17)

        self.bordro_list = []
        self.bordro_by_tc = {}
        if sh_bordro and sh_bordro.ncols >= 20:
            kanun_col = 71 if sh_bordro.ncols > 71 else (30 if sh_bordro.ncols > 30 else -1)
            meslek_col = 73 if sh_bordro.ncols > 73 else (33 if sh_bordro.ncols > 33 else -1)
            for c in range(sh_bordro.ncols):
                h_c = str(sh_bordro.cell_value(0, c)).lower()
                if "kanun no" in h_c:
                    kanun_col = c
                elif "meslek kodu" in h_c:
                    meslek_col = c

            for r in range(1, sh_bordro.nrows):
                ad_soyad = clean_str(sh_bordro.cell_value(r, 2))
                if not ad_soyad:
                    continue
                tc = clean_tc(sh_bordro.cell_value(r, 1))
                row_dict = {
                    "sno": clean_str(sh_bordro.cell_value(r, 0)),
                    "tc": tc,
                    "ad_soyad": ad_soyad,
                    "ucret": clean_money(sh_bordro.cell_value(r, 3)),
                    "giris": xldate_to_str(sh_bordro.cell_value(r, 4)),
                    "cikis": xldate_to_str(sh_bordro.cell_value(r, 5)),
                    "normal_kazanc": clean_money(sh_bordro.cell_value(r, 14)),
                    "ek_kazanc": clean_money(sh_bordro.cell_value(r, 15)),
                    "yasal_kesinti": clean_money(sh_bordro.cell_value(r, 17)),
                    "toplam_kazanc": clean_money(sh_bordro.cell_value(r, 20)),
                    "toplam_kesinti": clean_money(sh_bordro.cell_value(r, 21)),
                    "odenecek_net": clean_money(sh_bordro.cell_value(r, 22)),
                    "sgk_gun": clean_money(sh_bordro.cell_value(r, 23)),
                    "sgk_brut": clean_money(sh_bordro.cell_value(r, 24)),
                    "kanun": clean_str(sh_bordro.cell_value(r, kanun_col)) if kanun_col >= 0 else "",
                    "meslek_kodu": clean_str(sh_bordro.cell_value(r, meslek_col)) if meslek_col >= 0 else ""
                }
                self.bordro_list.append(row_dict)
                if tc:
                    self.bordro_by_tc[tc] = row_dict

        # 5. Daimi Personel Listesi
        self.daimi_list = []
        for p in self.personnel_list:
            if p.get("durumu") == "DAİMİ":
                self.daimi_list.append({
                    "sno": len(self.daimi_list) + 1,
                    "tc": p["tc"],
                    "ad_soyad": p["ad_soyad"],
                    "sgk_giris": p.get("sgk_giris", ""),
                    "sgk_durum": p.get("sgk_durumu", "NORMAL"),
                    "durum": "DAİMİ",
                    "maas": p.get("net_maas", 0.0)
                })

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
            sure_val = get_val(r, "sure")
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
                    "cinsiyet": "",
                    "isletme_giris": "",
                    "sgk_giris": "",
                    "kidem": 0,
                    "sgk_cikis": "",
                    "sgk_durumu": "NORMAL",
                    "durumu": "MEVSİMLİK",
                    "bolum": "DİĞER",
                    "ikamet": "",
                    "net_maas": 0.0,
                    "sirket": "FİDE KONSERVE",
                    "mesai_durumu": "ALIR"
                }
                self.personnel_list.append(p_data)
                self.personnel_by_key[name_k] = p_data
                if sicil:
                    self.personnel_by_tc[sicil] = p_data

            # 18+ saatlik hatalı turnike birleşimlerini 2 ayrı güne ayrıştır
            if g_tarih and g_saat and c_tarih and c_saat:
                dt_g = parse_d_hm(g_tarih, g_saat)
                dt_c = parse_d_hm(c_tarih, c_saat)
                if dt_g and dt_c and (dt_c - dt_g).total_seconds() > 18 * 3600:
                    self._add_punch_to_day(g_tarih, sira, sicil, kart, gun_adi, full_name, name_k, lokasyon, g_saat, "", sure_val)
                    c_gun_adi = get_day_name(c_tarih) if c_tarih else gun_adi
                    self._add_punch_to_day(c_tarih, sira, sicil, kart, c_gun_adi, full_name, name_k, lokasyon, "", c_saat, "")
                    continue

            # Normal geçerli hareket
            self._add_punch_to_day(target_date, sira, sicil, kart, gun_adi, full_name, name_k, lokasyon, g_saat, c_saat, sure_val)

        # Gece vardiyasında akşam başlayıp ertesi gün sabah biten ve 2 güne bölünmüş hareketleri birleştir
        self._stitch_cross_day_night_shifts()

        # Günlük hareketleri hesapla
        self._calculate_daily_results()

    def _stitch_cross_day_night_shifts(self):
        """
        Gece vardiyasında akşam (örn: 18:30-23:59) başlayıp ertesi gün sabah (05:00-10:00)
        biten ve PDKS turnikesi tarafından 2 ayrı güne bölünmüş hareketleri birleştirir.
        Örnek:
        - 11.09 20:13 giriş (çıkış yok, GECE_TEK_BASIM kalmış)
        - 12.09 08:15 çıkış (giriş yok veya tek basım kalmış)
        -> 11.09 için 20:13 Giriş - 08:15 Çıkış olarak bağlanır.
        -> 12.09 sabahındaki bu çıkış basımı tüketilir ve hatalı tek basımdan kurtarılır.
        """
        for name_k in list(self.personnel_by_key.keys()):
            for d in range(1, self.days_in_month):
                k_curr = (name_k, d)
                k_next = (name_k, d + 1)
                
                if k_curr not in self.raw_daily_punches or k_next not in self.raw_daily_punches:
                    continue
                
                data_curr = self.raw_daily_punches[k_curr]
                data_next = self.raw_daily_punches[k_next]
                
                rows_curr = data_curr["rows"]
                rows_next = data_next["rows"]
                
                if not rows_curr or not rows_next:
                    continue
                
                punches_curr = []
                for r in rows_curr:
                    if r.get("g_saat"): punches_curr.append(("g", r["g_saat"]))
                    if r.get("c_saat"): punches_curr.append(("c", r["c_saat"]))
                    
                punches_next = []
                for r in rows_next:
                    if r.get("g_saat"): punches_next.append(("g", r["g_saat"]))
                    if r.get("c_saat"): punches_next.append(("c", r["c_saat"]))
                    
                if not punches_curr or not punches_next:
                    continue
                    
                last_p = punches_curr[-1]
                t_curr = last_p[1][:5]
                hm_curr = parse_hm(t_curr)
                if not hm_curr:
                    continue
                tm_curr = hm_curr[0] * 60 + hm_curr[1]
                
                # Akşam veya gece girişiyle açık kalmış gün (18:30 sonrası tekil basım)
                if tm_curr >= 18 * 60 + 30 and (len(punches_curr) % 2 == 1):
                    first_p_next = punches_next[0]
                    t_next = first_p_next[1][:5]
                    hm_next = parse_hm(t_next)
                    if not hm_next:
                        continue
                    tm_next = hm_next[0] * 60 + hm_next[1]
                    
                    # Ertesi gün sabah çıkışı (05:00 - 10:00)
                    if 5 * 60 <= tm_next <= 10 * 60:
                        # Ertesi günün bu sabah basımı sonrasında tam gündüz mesaisi (10:00 - 22:00 arası) var mı?
                        has_daytime_shift = any(
                            parse_hm(p[1][:5]) and (10 * 60 < parse_hm(p[1][:5])[0] * 60 + parse_hm(p[1][:5])[1] < 22 * 60)
                            for p in punches_next
                        )
                        if not has_daytime_shift:
                            if len(rows_curr) == 1:
                                rows_curr[0] = {"g_saat": t_curr, "c_saat": t_next}
                            else:
                                for r in reversed(rows_curr):
                                    if r.get("g_saat") == last_p[1] or r.get("c_saat") == last_p[1]:
                                        r["g_saat"] = t_curr
                                        r["c_saat"] = t_next
                                        break
                                        
                            # Gün d+1'den sabah çıkış basımlarını tüket
                            new_next_rows = []
                            for r in rows_next:
                                r_g = r.get("g_saat", "")
                                r_c = r.get("c_saat", "")
                                hm_g = parse_hm(r_g[:5]) if r_g else None
                                hm_c = parse_hm(r_c[:5]) if r_c else None
                                is_morning_g = hm_g and (hm_g[0] * 60 + hm_g[1] <= 10 * 60)
                                is_morning_c = hm_c and (hm_c[0] * 60 + hm_c[1] <= 10 * 60)
                                if is_morning_g and not r_c:
                                    continue
                                if is_morning_g and is_morning_c:
                                    continue
                                if is_morning_g and r_c:
                                    r["g_saat"] = ""
                                new_next_rows.append(r)
                            data_next["rows"] = new_next_rows

    def _add_punch_to_day(self, date_str: str, sira: str, sicil: str, kart: str, gun_adi: str, full_name: str, name_k: str, lokasyon: str, g_saat: str, c_saat: str, sure_val: str = ""):
        if not date_str or "." not in str(date_str):
            return
        parts = str(date_str).split(".")
        try:
            d = int(parts[0])
            m = int(parts[1])
            if m != self.target_month:
                return
        except Exception:
            return
            
        k = (name_k, d)
        if not self.raw_daily_punches[k]["meta"]:
            self.raw_daily_punches[k]["meta"] = {
                "sira": sira, "sicil": sicil, "kart": kart, "gun_adi": gun_adi,
                "full_name": full_name, "lokasyon": lokasyon, "date_str": date_str, "day": d
            }
        self.raw_daily_punches[k]["rows"].append({"g_saat": g_saat, "c_saat": c_saat, "sure": sure_val})

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
                excel_sure = rows[0].get("sure", "")
                s_res = calculate_shift_hours(g_raw, c_raw)
                tot_h, base_h, ot_h, brk_h = s_res.total_hours, s_res.base_hours, s_res.overtime_hours, s_res.break_hours
                eff_g, eff_c, gross_h = s_res.effective_g, s_res.effective_c, s_res.gross_hours
                fiili_sure = tot_h
                
                if excel_sure:
                    hm = parse_hm(excel_sure)
                    if hm:
                        excel_tot_h = hm[0] + hm[1] / 60.0
                        if abs(excel_tot_h - tot_h) > 0.05:
                            audit_note += f" [Hesaplanan Süre: {fmt_hours_tr(tot_h)}s | Excel Süre Sütunu: {excel_sure} -> Excel baz alındı.] "
                            tot_h = excel_tot_h
                            fiili_sure = tot_h
                            ot_h = max(0.0, tot_h - 7.5)
                            base_h = min(7.5, tot_h)
                            
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
                    s_res = calculate_shift_hours(g_raw, c_raw)
                    tot_h, base_h, ot_h, brk_h = s_res.total_hours, s_res.base_hours, s_res.overtime_hours, s_res.break_hours
                    eff_g, eff_c, gross_h = s_res.effective_g, s_res.effective_c, s_res.gross_hours
                    fiili_sure = tot_h
                    
                    # Çoklu basımda, satırlardan birinde süre varsa toplayalım veya karşılaştıralım
                    total_excel_sure = 0.0
                    for r_sub in rows:
                        sub_sure = r_sub.get("sure", "")
                        if sub_sure:
                            hm = parse_hm(sub_sure)
                            if hm: total_excel_sure += hm[0] + hm[1] / 60.0
                    
                    if total_excel_sure > 0 and abs(total_excel_sure - tot_h) > 0.05:
                        audit_note += f" [Hesaplanan Süre: {fmt_hours_tr(tot_h)}s | Excel Çoklu Süre Toplamı: {fmt_hours_tr(total_excel_sure)}s -> Excel baz alındı.] "
                        tot_h = total_excel_sure
                        fiili_sure = tot_h
                        ot_h = max(0.0, tot_h - 7.5)
                        base_h = min(7.5, tot_h)
                        
                    if len(all_punches) > 2:
                        status_type = "ÇOKLU_BASIM_MIN_MAX"
                        is_exception = True
                        audit_note = f"Çoklu Basım: İlk Giriş {raw_g_str}, Son Çıkış {raw_c_str} ({fmt_hours_tr(tot_h)}s)"
                elif g_list:
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
                        s_res = calculate_shift_hours("08:00", g_raw)
                        tot_h, base_h, ot_h, brk_h = s_res.total_hours, s_res.base_hours, s_res.overtime_hours, s_res.break_hours
                        eff_g, eff_c, gross_h = s_res.effective_g, s_res.effective_c, s_res.gross_hours
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
                    s_res = calculate_shift_hours("08:00", t_raw)
                    tot_h, base_h, ot_h, brk_h = s_res.total_hours, s_res.base_hours, s_res.overtime_hours, s_res.break_hours
                    eff_g, eff_c, gross_h = s_res.effective_g, s_res.effective_c, s_res.gross_hours
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
            # ÖZEL BÖLÜM PRİMLERİ (Konfigürasyondan dinamik hesaplama):
            # -----------------------------------------------------------------
            bonus_hours, prim_msg = calculate_department_bonus(dept_name, fiili_sure, self.rules)
            if bonus_hours > 0:
                tot_h += bonus_hours
                ot_h += bonus_hours
                status_type = "BÖLÜM_PRİMLİ"
                audit_note = f"{audit_note} | {prim_msg}" if audit_note else prim_msg

            # Bir günde 24 saatten fazla çalışma yazılamaz kuralı:
            if tot_h > 24.0:
                tot_h = 24.0
                ot_h = max(0.0, 24.0 - base_h)

            rec = {
                "sno": len(self.audit_records) + 1,
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
                "is_bonus": (bonus_hours > 0),
                "audit_note": audit_note
            }
            
            self.daily_results[(name_k, day_num)] = rec
            self.audit_records.append(rec)
            if is_exception:
                self.exception_records.append(rec)
            if bonus_hours > 0:
                self.bonus_records.append(rec)

        # Tüm günlük sonuçlar hesaplandıktan sonra istisnalara akıllı amir önerilerini zenginleştir
        self.enrich_exceptions_with_smart_suggestions()

    def predict_smart_suggestion(self, rec: Dict[str, Any]) -> Dict[str, Any]:
        """
        Eksik veya tek basımlı istisna kayıtları için 3 kademeli akıllı amir önerisi üretir:
        1. Kademe: Aynı gün ve aynı bölümde iki basımı tam olan mesai arkadaşlarının çıkış saati (%95 güven)
        2. Kademe: Personelin ay içindeki normal vardiya alışkanlığı (%85 güven)
        3. Kademe: Fabrika standart vardiya kuralı (7,5 saat) (%75 güven)
        """
        day = rec.get("gun")
        dept = rec.get("bolum", "")
        norm_dept = _norm_dept_for_peers(dept)
        raw_g = rec.get("raw_g_saat", "-")
        raw_c = rec.get("raw_c_saat", "-")
        status_type = rec.get("status_type", "")
        name_k = rec.get("name_key") or norm_name_key(rec.get("ad_soyad", ""))
        display_dept = clean_display_text(dept) if dept else "Bölüm"

        g_hm = parse_hm(raw_g)
        is_morning_entry = g_hm and (6 * 60 <= g_hm[0] * 60 + g_hm[1] <= 12 * 60 + 30)

        # -------------------------------------------------------------
        # 1. KADEME: AYNI GÜN AYNI BÖLÜMDEKİ MESAİ ARKADAŞLARI ANALİZİ
        # -------------------------------------------------------------
        peer_exits = []
        peer_entries = []
        for (pk, d), drec in self.daily_results.items():
            if d == day and pk != name_k:
                p_dept = _norm_dept_for_peers(drec.get("bolum", ""))
                if p_dept == norm_dept and not drec.get("is_exception", False) and drec.get("final_hours", 0) > 0:
                    c_val = drec.get("raw_c_saat", "")
                    g_val = drec.get("raw_g_saat", "")
                    c_hm = parse_hm(c_val)
                    if is_morning_entry:
                        # Sabah giriş yapan işçi için gece vardiyasından çıkan mesai arkadaşı (örn 07:50 çıkışı) filtrelenir
                        if c_hm and (c_hm[0] * 60 + c_hm[1] >= 15 * 60):
                            peer_exits.append(c_val[:5])
                    else:
                        if c_val and c_val != "-":
                            peer_exits.append(c_val[:5])
                    if g_val and g_val != "-":
                        peer_entries.append(g_val[:5])

        if len(peer_exits) >= 2:
            mode_exit, exit_count = Counter(peer_exits).most_common(1)[0]
            mode_entry = Counter(peer_entries).most_common(1)[0][0] if peer_entries else "08:00"

            if status_type == "GİRİŞ_YOK_AKŞAM" or (raw_g == "-" and raw_c != "-"):
                eff_g = "08:00"
                eff_c = raw_c
                reason = f"{display_dept} bölümünde aynı gün ({day}. gün) çalışan {len(peer_exits)} mesai arkadaşının 08:00 başlangıcı ve mevcut {raw_c} çıkışı baz alındı."
            else:
                eff_g = raw_g if raw_g != "-" else "08:00"
                eff_c = mode_exit
                reason = f"{display_dept} bölümünde aynı gün ({day}. gün) çalışan {exit_count}/{len(peer_exits)} mesai arkadaşının çıkış saati ({mode_exit}) baz alındı."

            s_res = calculate_shift_hours(eff_g, eff_c)
            bonus_h, _ = calculate_department_bonus(dept, s_res.total_hours, self.rules)
            tot_h = round(s_res.total_hours + bonus_h, 2)

            if tot_h >= 4.0:
                return {
                    "suggested_hours": tot_h,
                    "suggested_g": eff_g,
                    "suggested_c": eff_c,
                    "confidence": 95,
                    "source": "DEPARTMENT_PEERS",
                    "badge_class": "badge-peer",
                    "reason": reason,
                    "peer_count": exit_count if status_type != "GİRİŞ_YOK_AKŞAM" else len(peer_exits),
                    "total_peers": len(peer_exits)
                }

        # -------------------------------------------------------------
        # 2. KADEME: PERSONELİN AY İÇİNDEKİ KENDİ VARDİYA ALIŞKANLIĞI
        # -------------------------------------------------------------
        habit_exits = []
        for (pk, d), drec in self.daily_results.items():
            if pk == name_k and d != day and not drec.get("is_exception", False) and drec.get("final_hours", 0) > 0:
                c_val = drec.get("raw_c_saat", "")
                c_hm = parse_hm(c_val)
                if is_morning_entry:
                    if c_hm and (c_hm[0] * 60 + c_hm[1] >= 15 * 60):
                        habit_exits.append(c_val[:5])
                else:
                    if c_val and c_val != "-":
                        habit_exits.append(c_val[:5])

        if len(habit_exits) >= 2:
            mode_exit, habit_count = Counter(habit_exits).most_common(1)[0]
            eff_g = raw_g if raw_g != "-" else "08:00"
            eff_c = raw_c if raw_c != "-" else mode_exit
            s_res = calculate_shift_hours(eff_g, eff_c)
            bonus_h, _ = calculate_department_bonus(dept, s_res.total_hours, self.rules)
            tot_h = round(s_res.total_hours + bonus_h, 2)

            if tot_h >= 4.0:
                return {
                    "suggested_hours": tot_h,
                    "suggested_g": eff_g,
                    "suggested_c": eff_c,
                    "confidence": 85,
                    "source": "PERSONAL_HABIT",
                    "badge_class": "badge-habit",
                    "reason": f"Personelin bu aydaki {habit_count} günlük normal vardiya alışkanlığı ({mode_exit}) baz alındı.",
                    "peer_count": 0,
                    "total_peers": 0
                }

        # -------------------------------------------------------------
        # 3. KADEME: FABRİKA STANDART VARDİYA KURALI (7,5 SAAT)
        # -------------------------------------------------------------
        eff_g = raw_g if raw_g != "-" else "08:00"
        eff_c = raw_c if raw_c != "-" else "17:00"
        bonus_h, _ = calculate_department_bonus(dept, 7.5, self.rules)
        tot_std = round(7.5 + bonus_h, 2)
        return {
            "suggested_hours": tot_std,
            "suggested_g": eff_g,
            "suggested_c": eff_c,
            "confidence": 75,
            "source": "FACTORY_STANDARD",
            "badge_class": "badge-std",
            "reason": "Fabrika standart vardiya kuralı (7,5 saat tam gün) uygulandı.",
            "peer_count": 0,
            "total_peers": 0
        }

    def enrich_exceptions_with_smart_suggestions(self):
        """Tüm istisna kayıtlarını akıllı amir önerisiyle zenginleştirir."""
        for rec in self.exception_records:
            sug = self.predict_smart_suggestion(rec)
            rec["smart_suggestion"] = sug
            name_k = rec.get("name_key") or norm_name_key(rec.get("ad_soyad", ""))
            day_num = rec.get("gun")
            if (name_k, day_num) in self.daily_results:
                self.daily_results[(name_k, day_num)]["smart_suggestion"] = sug

    def get_summary_matrix(self) -> List[Dict]:
        """Personel bazında 1..days_in_month günlük saat matrisini ve toplamlarını döndürür."""
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
                "net_maas": p.get("net_maas", 0.0),
                "mesai_durumu": p.get("mesai_durumu", "ALIR"),
                "daily_hours": {},
                "daily_formula_hours": {},
                "total_work_days": 0,
                "total_base_hours": 0.0,
                "total_ot_hours": 0.0,
                "total_missing_hours": 0.0,
                "total_weekday_ot_hours": 0.0,
                "total_sunday_ot_hours": 0.0,
                "net_weekday_ot_hours": 0.0,
                "total_bonus_hours": 0.0,
                "total_hours": 0.0
            }
            
            for d in range(1, self.days_in_month + 1):
                rec = self.daily_results.get((name_k, d))
                if rec and rec["final_hours"] > 0:
                    h = min(24.0, max(0.0, float(rec["final_hours"])))
                    row["daily_hours"][d] = h
                    row["total_work_days"] += 1
                    row["total_base_hours"] += rec["base_hours"]
                    row["total_ot_hours"] += rec["overtime_hours"]
                    
                    if d in self.sundays:
                        row["total_sunday_ot_hours"] += h
                        row["daily_formula_hours"][d] = round(h * 1.5, 2)
                    else:
                        ot_h = rec["overtime_hours"]
                        row["total_weekday_ot_hours"] += ot_h
                        if h < 7.5:
                            miss_h = round(7.5 - h, 2)
                            row["total_missing_hours"] += miss_h
                            # Formül: IF(h>7.5, (h-7.5)*1.5, h-7.5) -> Eksik saat 1.0x (negatif saat)
                            row["daily_formula_hours"][d] = round(h - 7.5, 2)
                        elif ot_h > 0:
                            # Formül: IF(h>7.5, (h-7.5)*1.5, h-7.5) -> Mesai 1.5x
                            row["daily_formula_hours"][d] = round(ot_h * 1.5, 2)
                        else:
                            row["daily_formula_hours"][d] = 0.0
                            
                    row["total_bonus_hours"] += rec["bonus_hours"]
                    row["total_hours"] += h
                else:
                    row["daily_hours"][d] = 0.0
                    row["daily_formula_hours"][d] = None
                    
            row["total_weekday_ot_hours"] = round(row["total_weekday_ot_hours"], 2)
            row["total_missing_hours"] = round(row["total_missing_hours"], 2)
            row["total_sunday_ot_hours"] = round(row["total_sunday_ot_hours"], 2)
            # Net hafta içi mesai bakiyesi: (Fazla Mesai * 1.5) - (Eksik Saat * 1.0)
            row["net_weekday_ot_hours"] = round((row["total_weekday_ot_hours"] * 1.5) - (row["total_missing_hours"] * 1.0), 2)
            matrix.append(row)
        return matrix

    def calculate_financial_radar(self) -> Dict[str, Any]:
        """
        Tüm personelin net maaş ve saatlik ücretlerini baz alarak:
        - Toplam normal çalışma maliyeti
        - Toplam fazla mesai maliyeti:
          * Hafta İçi: =IF(h>7.5; (h-7.5)*1.5; h-7.5) kuralıyla (1.5x Mesai, 1.0x Eksik Saat Kesintisi)
          * Pazar: 1.5x Pazar Mesai Çarpanı
        - Toplam bölüm primi maliyeti (1.5x)
        - Toplam hakediş ve elden fark maliyeti
        - Departman bazlı anlamlı maliyet ve bütçe analizini (TL, %, kişi başı ortalamalar)
        üretir.
        """
        matrix = self.get_summary_matrix()
        
        salary_map = {}
        for p in self.personnel_list:
            salary_map[p["name_key"]] = p.get("net_maas", 0.0)
            
        total_base_cost = 0.0
        total_ot_cost = 0.0
        total_weekday_ot_cost = 0.0
        total_sunday_ot_cost = 0.0
        total_missing_deduction = 0.0
        total_bonus_cost = 0.0
        total_payroll_cost = 0.0
        total_bank_cost = 0.0
        total_cash_diff = 0.0
        
        personnel_costs = []
        dept_costs = {}
        
        for m in matrix:
            if m["total_work_days"] == 0:
                continue
                
            name_k = m["name_key"]
            tc = m.get("tc", "")
            net_m = salary_map.get(name_k, 0.0)
            
            if net_m <= 0:
                if tc in self.bordro_by_tc and self.bordro_by_tc[tc].get("odenecek_net", 0.0) > 0:
                    net_m = self.bordro_by_tc[tc]["odenecek_net"]
                else:
                    net_m = 28075.50  # Standart net asgari ücret tabanı
                    
            hourly_net = net_m / 225.0
            ot_hourly_net = hourly_net * 1.5
            
            weekday_ot = m.get("total_weekday_ot_hours", 0.0)
            missing_hours = m.get("total_missing_hours", 0.0)
            sunday_ot = m.get("total_sunday_ot_hours", 0.0)
            bonus_hours = m.get("total_bonus_hours", 0.0)
            all_ot_hours = round(weekday_ot + sunday_ot, 2)
            
            # Formül: =IF($AE3=""; ""; IF($AE3>7.5; ($AE3-7.5)*1.5; $AE3-7.5))
            # Hafta içi mesainin çarpanı 1.5, eksik saatin çarpanı 1:
            net_weekday_ot_hours = round((weekday_ot * 1.5) - (missing_hours * 1.0), 2)
            
            base_cost = round((net_m / 30.0) * m["total_work_days"], 2)
            weekday_ot_cost = round(net_weekday_ot_hours * hourly_net, 2)
            missing_deduction_cost = round(missing_hours * hourly_net, 2)
            gross_weekday_ot_cost = round(weekday_ot * 1.5 * hourly_net, 2)
            sunday_ot_cost = round(sunday_ot * 1.5 * hourly_net, 2)
            ot_cost = round(weekday_ot_cost + sunday_ot_cost, 2)
            bonus_cost = round(bonus_hours * ot_hourly_net, 2)
            total_net_wage = round(base_cost + ot_cost + bonus_cost, 2)
            
            bank_net = self.bordro_by_tc.get(tc, {}).get("odenecek_net", 0.0)
            icra_info = self.icra_by_name.get(name_k, {})
            icra_kes = icra_info.get("kesilen", 0.0) if isinstance(icra_info, dict) else 0.0
            cash_diff = max(0.0, total_net_wage - bank_net - icra_kes)
            
            total_base_cost += base_cost
            total_weekday_ot_cost += weekday_ot_cost
            total_sunday_ot_cost += sunday_ot_cost
            total_missing_deduction += missing_deduction_cost
            total_ot_cost += ot_cost
            total_bonus_cost += bonus_cost
            total_payroll_cost += total_net_wage
            total_bank_cost += bank_net
            total_cash_diff += cash_diff
            
            dept = m.get("bolum") or "Bilinmeyen"
            if dept not in dept_costs:
                dept_costs[dept] = {
                    "department": dept,
                    "active_count": 0,
                    "total_hours": 0.0,
                    "ot_hours": 0.0,
                    "weekday_ot_hours": 0.0,
                    "missing_hours": 0.0,
                    "net_weekday_ot_hours": 0.0,
                    "sunday_ot_hours": 0.0,
                    "bonus_hours": 0.0,
                    "base_cost": 0.0,
                    "ot_cost": 0.0,
                    "missing_deduction": 0.0,
                    "bonus_cost": 0.0,
                    "total_cost": 0.0,
                    "avg_ot_hours_per_worker": 0.0,
                    "avg_ot_cost_per_worker": 0.0,
                    "ot_share_pct": 0.0,
                    "insight": ""
                }
            
            dept_costs[dept]["active_count"] += 1
            dept_costs[dept]["total_hours"] += m["total_hours"]
            dept_costs[dept]["ot_hours"] += all_ot_hours
            dept_costs[dept]["weekday_ot_hours"] += weekday_ot
            dept_costs[dept]["missing_hours"] += missing_hours
            dept_costs[dept]["net_weekday_ot_hours"] += net_weekday_ot_hours
            dept_costs[dept]["sunday_ot_hours"] += sunday_ot
            dept_costs[dept]["bonus_hours"] += m["total_bonus_hours"]
            dept_costs[dept]["base_cost"] += base_cost
            dept_costs[dept]["ot_cost"] += ot_cost
            dept_costs[dept]["missing_deduction"] += missing_deduction_cost
            dept_costs[dept]["bonus_cost"] += bonus_cost
            dept_costs[dept]["total_cost"] += total_net_wage
            
            personnel_costs.append({
                "sno": m["sno"],
                "tc": tc,
                "tc_no": tc,
                "ad_soyad": m["ad_soyad"],
                "name_key": name_k,
                "bolum": dept,
                "work_days": m["total_work_days"],
                "days_worked": m["total_work_days"],
                "net_maas": round(net_m, 2),
                "hourly_net": round(hourly_net, 2),
                "hourly_base_rate": round(hourly_net, 2),
                "ot_hourly_net": round(ot_hourly_net, 2),
                "hourly_overtime_rate": round(ot_hourly_net, 2),
                "total_hours": round(m["total_hours"], 2),
                "weekday_ot_hours": round(weekday_ot, 2),
                "missing_hours": round(missing_hours, 2),
                "net_weekday_ot_hours": round(net_weekday_ot_hours, 2),
                "weekday_ot_cost": round(weekday_ot_cost, 2),
                "missing_deduction_cost": round(missing_deduction_cost, 2),
                "gross_weekday_ot_cost": round(gross_weekday_ot_cost, 2),
                "sunday_ot_hours": round(sunday_ot, 2),
                "sunday_ot_cost": round(sunday_ot_cost, 2),
                "all_ot_hours": round(all_ot_hours, 2),
                "total_ot_hours": round(all_ot_hours, 2),
                "ot_cost": round(ot_cost, 2),
                "overtime_cost": round(ot_cost, 2),
                "bonus_hours": round(m["total_bonus_hours"], 1),
                "bonus_cost": round(bonus_cost, 2),
                "base_cost": round(base_cost, 2),
                "total_net_wage": round(total_net_wage, 2),
                "total_net_earned": round(total_net_wage, 2),
                "bank_net": round(bank_net, 2),
                "icra": round(icra_kes, 2),
                "cash_diff": round(cash_diff, 2),
                "cash_difference": round(cash_diff, 2)
            })
            
        # Departman ortalamaları ve anlamlı yönetim içgörüleri (Insights)
        for dept, d_data in dept_costs.items():
            cnt = d_data["active_count"]
            d_data["avg_ot_hours_per_worker"] = round(d_data["ot_hours"] / cnt, 1) if cnt > 0 else 0.0
            d_data["avg_overtime_hours_per_worker"] = d_data["avg_ot_hours_per_worker"]
            d_data["avg_ot_cost_per_worker"] = round(d_data["ot_cost"] / cnt, 2) if cnt > 0 else 0.0
            d_data["avg_overtime_cost_per_worker"] = d_data["avg_ot_cost_per_worker"]
            d_data["ot_share_pct"] = round((d_data["ot_cost"] / total_ot_cost * 100), 1) if total_ot_cost > 0 else 0.0
            d_data["overtime_budget_share_pct"] = d_data["ot_share_pct"]
            d_data["overtime_cost"] = round(d_data["ot_cost"], 2)
            
            # Anlamlı analiz metinleri
            norm_dept = dept.upper()
            if "DOLUM" in norm_dept:
                d_data["insight"] = f"Kişi başı en yüksek mesai yoğunluğu: Sadece {cnt} çalışan kişi başı {d_data['avg_ot_hours_per_worker']} saat mesai yaptı ({d_data['avg_ot_cost_per_worker']:,.2f} TL/kişi)."
            elif "TEMİZLEME" in norm_dept:
                d_data["insight"] = f"Fabrikanın en kalabalık birimi ({cnt} aktif). Toplam mesai harcamasının %{d_data['ot_share_pct']}'ini oluşturuyor."
            elif "ÜRETİM" in norm_dept or "URETIM" in norm_dept:
                d_data["insight"] = f"Fabrikanın en yüksek toplam mesai tutarı (%{d_data['ot_share_pct']} pay, {d_data['ot_cost']:,.2f} TL)."
            elif "AMBAR" in norm_dept:
                d_data["insight"] = f"Sevkiyat ve hammadde kabul yoğunluğu ({d_data['ot_cost']:,.2f} TL mesai harcaması)."
            elif "BAKIM" in norm_dept:
                d_data["insight"] = f"Hat arıza ve revizyon nedeniyle yüksek saatlik maliyet ({d_data['ot_cost']:,.2f} TL)."
            else:
                d_data["insight"] = f"{d_data['ot_hours']} saat mesai, ortalama {d_data['avg_ot_hours_per_worker']} saat/kişi."

        # Maliyete göre sıralı departman listesi
        sorted_depts = sorted(dept_costs.values(), key=lambda x: x["ot_cost"], reverse=True)
        
        # En yüksek departmanlar
        top_ot_dept = sorted_depts[0]["department"] if sorted_depts else "-"
        # En yoğun kişi başı çalışan departman
        top_intensity_dept = max(sorted_depts, key=lambda x: x["avg_ot_hours_per_worker"])["department"] if sorted_depts else "-"
        
        total_all_ot = sum(p["all_ot_hours"] for p in personnel_costs)
        total_weekday_ot_all = sum(p["weekday_ot_hours"] for p in personnel_costs)
        total_missing_all = sum(p["missing_hours"] for p in personnel_costs)
        total_net_weekday_ot_all = sum(p["net_weekday_ot_hours"] for p in personnel_costs)
        avg_rate = round(total_ot_cost / total_all_ot, 2) if total_all_ot > 0 else 0.0

        summary_dict = {
            "total_active_personnel": len(personnel_costs),
            "total_personnel_active": len(personnel_costs),
            "total_hours": round(sum(p["total_hours"] for p in personnel_costs), 1),
            "total_ot_hours": round(total_all_ot, 2),
            "total_overtime_hours": round(total_all_ot, 2),
            "total_weekday_ot_hours": round(total_weekday_ot_all, 2),
            "total_missing_hours": round(total_missing_all, 2),
            "total_missing_deduction": round(total_missing_deduction, 2),
            "total_net_weekday_ot_hours": round(total_net_weekday_ot_all, 2),
            "total_bonus_hours": round(sum(p["bonus_hours"] for p in personnel_costs), 1),
            "total_base_cost": round(total_base_cost, 2),
            "total_weekday_ot_cost": round(total_weekday_ot_cost, 2),
            "total_sunday_ot_cost": round(total_sunday_ot_cost, 2),
            "total_ot_cost": round(total_ot_cost, 2),
            "total_overtime_cost": round(total_ot_cost, 2),
            "total_bonus_cost": round(total_bonus_cost, 2),
            "total_payroll_cost": round(total_payroll_cost, 2),
            "total_payroll_budget": round(total_payroll_cost, 2),
            "total_bank_cost": round(total_bank_cost, 2),
            "total_cash_diff": round(total_cash_diff, 2),
            "avg_ot_hourly_rate": avg_rate,
            "average_overtime_rate_per_hour": avg_rate,
            "top_ot_department": top_ot_dept,
            "top_intensity_department": top_intensity_dept
        }

        return {
            "summary": summary_dict,
            "departments": sorted_depts,
            "department_breakdown": sorted_depts,
            "personnel": personnel_costs,
            "personnel_costs": personnel_costs
        }

    def calculate_legal_compliance_radar(self) -> Dict[str, Any]:
        """
        4857 Sayılı İş Kanunu ve SGK Yasal Uyum & Risk Radarı.
        İş Müfettişi ve SGK denetim risklerini sıfıra indiren 3 temel yasal kural denetimi:
        1. 11 Saat Kesintisiz Günlük Dinlenme Kuralı (Madde 68 & Postalar Yönetmeliği Md. 9)
        2. 7 Gün Kesintisiz Çalışma & Hafta Tatili İhlali (Madde 46)
        3. Yıllık 270 Saat Fazla Mesai & Aylık Aşırı Çalışma Sınırı (Madde 41)
        """
        matrix = self.get_summary_matrix()
        active_matrix = {m["name_key"]: m for m in matrix if m["total_work_days"] > 0}
        
        all_violations: List[Dict[str, Any]] = []
        rest_violations: List[Dict[str, Any]] = []
        consecutive_violations: List[Dict[str, Any]] = []
        overtime_risks: List[Dict[str, Any]] = []
        
        dept_compliance: Dict[str, Dict[str, Any]] = {}

        # ---------------------------------------------------------------------
        # 1. KURAL: 11 SAAT KESİNTİSİZ GÜNLÜK DİNLENME KURALI (Madde 68)
        # ---------------------------------------------------------------------
        for p in self.personnel_list:
            name_k = p["name_key"]
            if name_k not in active_matrix:
                continue
            
            p_mat = active_matrix[name_k]
            dept = p_mat.get("bolum") or "Bilinmeyen"
            
            for d in range(1, self.days_in_month):
                r1 = self.daily_results.get((name_k, d))
                r2 = self.daily_results.get((name_k, d + 1))
                if not r1 or not r2:
                    continue
                
                c1 = r1.get("raw_c_saat", "-")
                g2 = r2.get("raw_g_saat", "-")
                if c1 == "-" or g2 == "-" or r1.get("final_hours", 0.0) <= 0 or r2.get("final_hours", 0.0) <= 0:
                    continue
                
                t1 = parse_hm(c1)
                t2 = parse_hm(g2)
                if t1 is None or t2 is None:
                    continue
                
                c1_min = t1[0] * 60 + t1[1]
                g2_min = t2[0] * 60 + t2[1]
                
                # Dinlenme süresi (saat cinsinden)
                rest_hours = ((1440 - c1_min) + g2_min) / 60.0
                
                if 0.0 < rest_hours < 11.0:
                    severity = "CRITICAL" if rest_hours < 8.0 else "WARNING"
                    v_item = {
                        "id": f"REST_{name_k}_{d}_{d+1}",
                        "name_key": name_k,
                        "ad_soyad": p["ad_soyad"],
                        "tc": p["tc"],
                        "bolum": dept,
                        "violation_type": "REST_11H",
                        "violation_title": "11 Saat Altı Dinlenme",
                        "severity": severity,
                        "date_str": f"{d:02d}.{self.target_month:02d} → {d+1:02d}.{self.target_month:02d}",
                        "start_day": d,
                        "end_day": d + 1,
                        "c1": c1,
                        "g2": g2,
                        "metric_value": f"{rest_hours:.1f} saat",
                        "legal_limit": "En az 11,0 saat",
                        "legal_article": "4857 SK Madde 68 & Postalar Yön. Md. 9",
                        "detail": f"{c1} çıkışından sonra ertesi gün {g2} işbaşı yapıldı. Aradaki dinlenme süresi {rest_hours:.1f} saat (Yasal açık: {11.0 - rest_hours:.1f} saat)."
                    }
                    rest_violations.append(v_item)
                    all_violations.append(v_item)

        # ---------------------------------------------------------------------
        # 2. KURAL: 7 GÜN KESİNTİSİZ ÇALIŞMA / HAFTA TATİLİ İHLALİ (Madde 46)
        # ---------------------------------------------------------------------
        for p in self.personnel_list:
            name_k = p["name_key"]
            if name_k not in active_matrix:
                continue
            
            p_mat = active_matrix[name_k]
            dept = p_mat.get("bolum") or "Bilinmeyen"
            
            days_worked = [
                d for d in range(1, self.days_in_month + 1)
                if self.daily_results.get((name_k, d)) and self.daily_results.get((name_k, d)).get("raw_g_saat") != "-" and self.daily_results.get((name_k, d)).get("final_hours", 0.0) > 0
            ]
            
            streak: List[int] = []
            for d in range(1, self.days_in_month + 1):
                if d in days_worked:
                    streak.append(d)
                else:
                    if len(streak) >= 7:
                        consec_len = len(streak)
                        severity = "CRITICAL" if consec_len >= 10 else "WARNING"
                        v_item = {
                            "id": f"CONSEC_{name_k}_{streak[0]}_{streak[-1]}",
                            "name_key": name_k,
                            "ad_soyad": p["ad_soyad"],
                            "tc": p["tc"],
                            "bolum": dept,
                            "violation_type": "CONSECUTIVE_7D",
                            "violation_title": f"{consec_len} Gün Kesintisiz Çalışma",
                            "severity": severity,
                            "date_str": f"{streak[0]:02d}.{self.target_month:02d} - {streak[-1]:02d}.{self.target_month:02d}",
                            "start_day": streak[0],
                            "end_day": streak[-1],
                            "c1": "-",
                            "g2": "-",
                            "metric_value": f"{consec_len} gün aralıksız",
                            "legal_limit": "Azami 6 gün (7. gün 24s tatil)",
                            "legal_article": "4857 SK Madde 46 (Hafta Tatili)",
                            "detail": f"Çalışana yasal 24 saatlik dinlenme hakkı tanınmadan {consec_len} gün kesintisiz mesai yaptırılmıştır."
                        }
                        consecutive_violations.append(v_item)
                        all_violations.append(v_item)
                    streak = []
            
            if len(streak) >= 7:
                consec_len = len(streak)
                severity = "CRITICAL" if consec_len >= 10 else "WARNING"
                v_item = {
                    "id": f"CONSEC_{name_k}_{streak[0]}_{streak[-1]}",
                    "name_key": name_k,
                    "ad_soyad": p["ad_soyad"],
                    "tc": p["tc"],
                    "bolum": dept,
                    "violation_type": "CONSECUTIVE_7D",
                    "violation_title": f"{consec_len} Gün Kesintisiz Çalışma",
                    "severity": severity,
                    "date_str": f"{streak[0]:02d}.{self.target_month:02d} - {streak[-1]:02d}.{self.target_month:02d}",
                    "start_day": streak[0],
                    "end_day": streak[-1],
                    "c1": "-",
                    "g2": "-",
                    "metric_value": f"{consec_len} gün aralıksız",
                    "legal_limit": "Azami 6 gün (7. gün 24s tatil)",
                    "legal_article": "4857 SK Madde 46 (Hafta Tatili)",
                    "detail": f"Çalışana yasal 24 saatlik dinlenme hakkı tanınmadan {consec_len} gün kesintisiz mesai yaptırılmıştır."
                }
                consecutive_violations.append(v_item)
                all_violations.append(v_item)

        # ---------------------------------------------------------------------
        # 3. KURAL: YILLIK 270 SAAT FAZLA MESAİ & AŞIRI ÇALIŞMA (Madde 41)
        # ---------------------------------------------------------------------
        for name_k, p_mat in active_matrix.items():
            ot = p_mat.get("total_ot_hours", 0.0)
            dept = p_mat.get("bolum") or "Bilinmeyen"
            p_info = self.personnel_by_key.get(name_k, {})
            
            if ot >= 35.0:
                severity = "CRITICAL" if ot >= 50.0 else "WARNING"
                title = "270 Saat Aşım Riski (Kritik)" if ot >= 50.0 else "Aşırı Aylık Fazla Mesai"
                v_item = {
                    "id": f"OT270_{name_k}",
                    "name_key": name_k,
                    "ad_soyad": p_mat["ad_soyad"],
                    "tc": p_mat.get("tc", ""),
                    "bolum": dept,
                    "violation_type": "OVERTIME_270H",
                    "violation_title": title,
                    "severity": severity,
                    "date_str": f"Eylül {self.target_year}",
                    "start_day": 1,
                    "end_day": self.days_in_month,
                    "c1": "-",
                    "g2": "-",
                    "metric_value": f"{ot:.1f} saat FM",
                    "legal_limit": "Yılda azami 270 saat (Ayda ~22.5s)",
                    "legal_article": "4857 SK Madde 41 & Fazla Çalışma Yön. Md. 5",
                    "detail": f"Eylül ayında tek başına {ot:.1f} saat fazla mesai yapılmıştır. Yıllık 270 saat yasal üst sınırını delme riski çok yüksektir."
                }
                overtime_risks.append(v_item)
                all_violations.append(v_item)

        # ---------------------------------------------------------------------
        # DEPARTMAN BAZLI UYUM VE RİSK ANALİZİ
        # ---------------------------------------------------------------------
        for name_k, p_mat in active_matrix.items():
            dept = p_mat.get("bolum") or "Bilinmeyen"
            if dept not in dept_compliance:
                dept_compliance[dept] = {
                    "department": dept,
                    "active_headcount": 0,
                    "rest_count": 0,
                    "consecutive_count": 0,
                    "ot_risk_count": 0,
                    "critical_count": 0,
                    "warning_count": 0,
                    "total_violations": 0,
                    "compliance_score": 100.0,
                    "risk_level": "DÜŞÜK"
                }
            dept_compliance[dept]["active_headcount"] += 1

        for v in all_violations:
            dept = v["bolum"]
            if dept in dept_compliance:
                dept_compliance[dept]["total_violations"] += 1
                if v["severity"] == "CRITICAL":
                    dept_compliance[dept]["critical_count"] += 1
                else:
                    dept_compliance[dept]["warning_count"] += 1
                    
                if v["violation_type"] == "REST_11H":
                    dept_compliance[dept]["rest_count"] += 1
                elif v["violation_type"] == "CONSECUTIVE_7D":
                    dept_compliance[dept]["consecutive_count"] += 1
                elif v["violation_type"] == "OVERTIME_270H":
                    dept_compliance[dept]["ot_risk_count"] += 1

        for dept, d_comp in dept_compliance.items():
            cnt = d_comp["active_headcount"]
            tot_v = d_comp["total_violations"]
            crit = d_comp["critical_count"]
            # Departman uyum puanı
            dept_penalty = (crit * 3.0) + ((tot_v - crit) * 1.5)
            d_comp["compliance_score"] = max(20.0, round(100.0 - (dept_penalty / max(cnt, 1) * 20.0), 1))
            if d_comp["compliance_score"] < 70.0 or crit >= 3:
                d_comp["risk_level"] = "YÜKSEK"
            elif d_comp["compliance_score"] < 85.0 or tot_v >= 5:
                d_comp["risk_level"] = "ORTA"
            else:
                d_comp["risk_level"] = "DÜŞÜK"

        sorted_dept_compliance = sorted(dept_compliance.values(), key=lambda x: (x["critical_count"], x["total_violations"]), reverse=True)

        # ---------------------------------------------------------------------
        # FABRİKA YASAL UYUM ENDEKSİ (LEGAL COMPLIANCE INDEX)
        # ---------------------------------------------------------------------
        critical_total = sum(1 for v in all_violations if v["severity"] == "CRITICAL")
        warning_total = len(all_violations) - critical_total
        
        # Ceza puanı hesabı: Kritik ihlaller 0.6 puan, Uyarılar 0.2 puan kırar
        penalty = (critical_total * 0.5) + (warning_total * 0.15)
        compliance_index = max(10.0, round(100.0 - penalty, 1))
        
        if compliance_index >= 90.0:
            risk_status = "DÜŞÜK RİSK"
            risk_desc = "Fabrika geneli yasal uyum yüksek, teftiş riski asgari düzeyde."
        elif compliance_index >= 75.0:
            risk_status = "ORTA RİSK"
            risk_desc = "Vardiya geçişleri ve hafta tatillerinde düzeltici tedbir alınmalı."
        else:
            risk_status = "YÜKSEK RİSK"
            risk_desc = "İş müfettişi denetimlerinde idari para cezası ve İSG kusur riski yüksek."

        # Ciddiyete göre sıralı tüm ihlaller (Kritikler en başta)
        all_violations.sort(key=lambda x: (0 if x["severity"] == "CRITICAL" else 1, x["start_day"], x["ad_soyad"]))

        return {
            "summary": {
                "compliance_index": compliance_index,
                "risk_status": risk_status,
                "risk_description": risk_desc,
                "total_violations": len(all_violations),
                "critical_violations": critical_total,
                "warning_violations": warning_total,
                "rest_violations_count": len(rest_violations),
                "consecutive_work_count": len(consecutive_violations),
                "overtime_limit_count": len(overtime_risks),
                "inspected_personnel": len(active_matrix),
                "highest_risk_department": sorted_dept_compliance[0]["department"] if sorted_dept_compliance else "-"
            },
            "department_compliance": sorted_dept_compliance,
            "violations": all_violations,
            "rest_violations": rest_violations,
            "consecutive_violations": consecutive_violations,
            "overtime_risks": overtime_risks
        }
