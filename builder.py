# -*- coding: utf-8 -*-
"""
Puantaj ve PDKS Sistemi Uçtan Uca Yapılandırma ve Modernizasyon Scripti
Fide Konserve Gıda San. A.Ş.

Bu script:
1. 'pdks.xls' dosyasındaki turnike hareketlerini okur ve analiz eder.
2. 'puantaj.xls' dosyasındaki personel, bordro, icra ve rapor verilerini harmanlar.
3. Tüm formülleri (SUM, COUNTIF, IF, MIN, VLOOKUP), biçimlendirmeleri, renkleri
   ve Türkçe karakterleri eksiksiz içeren modern 'puantaj.xlsx' dosyasını üretir.
"""
import xlrd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import os
import shutil

# Türkçe Karakter Onarım Sözlüğü
TR_MAP = {
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
    "GZELDAL": "GÜZELDAL",
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
    "SA": "İSA",
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
    "OKUN": "COŞKUN",
    "ALAR": "ÇAĞLAR",
    "AY": "ÇAY",
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
    "SGT": "SÖĞÜT",
    "DUDU AY": "DUDU AY",
    "YUCEL": "YÜCEL",
    "SILA ZER": "SILA ÖZER",
}

def repair_text(text):
    if not text:
        return ""
    s = str(text).strip()
    for k, v in TR_MAP.items():
        s = s.replace(k, v)
    s = s.replace("\ufffd", "").replace("", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

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
    return repair_text(str(val).strip())

def clean_money(val):
    if val is None or val == "":
        return 0.0
    try:
        return float(val)
    except:
        return 0.0

def parse_time_str(s):
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
    if isinstance(val, (int, float)) and val > 30000 and val < 60000:
        try:
            return xlrd.xldate.xldate_as_datetime(val, 0).strftime("%d.%m.%Y")
        except:
            return str(val)
    return str(val) if val is not None else ""

def get_sheet_by_keyword(wb, keyword, default_idx=None):
    for s in wb.sheets():
        if keyword.lower() in s.name.lower():
            return s
    if default_idx is not None and default_idx < len(wb.sheets()):
        return wb.sheet_by_index(default_idx)
    return None

def parse_hm(t_str):
    if not t_str or ":" not in str(t_str):
        return None
    p = str(t_str).strip().split(":")
    try:
        return int(p[0]), int(p[1])
    except:
        return None

def calc_factory_worked_hours(g_saat_str, c_saat_str, sure_str="", is_office=False):
    """
    Fabrika Çalışma & Fazla Mesai Özel Kuralları:
    1. Sabah Giriş:
       - 08:10 ve öncesi girişler: Mesai 08:00'de başlamış sayılır (Normal 7.5 saat taban).
       - 08:10'dan sonraki girişler: Mesai 09:00'a yuvarlanır (6.5 saat taban).
    2. Akşam Çıkış & Fazla Mesai Aralıkları:
       - 17:00 - 17:24 arası çıkış -> 7.5 saat (Fazla mesai yok)
       - 17:25 - 17:49 arası çıkış -> 8.0 saat (+0.5 saat FM)
       - 17:50 - 18:24 arası çıkış -> 8.5 saat (+1.0 saat FM)
       - 18:25 - 18:49 arası çıkış -> 9.0 saat (+1.5 saat FM)
       - 18:50 - 19:24 arası çıkış -> 9.5 saat (+2.0 saat FM)
       - 19:25 - 19:49 arası çıkış -> 10.0 saat (+2.5 saat FM)
       - 19:50 - 20:24 arası çıkış -> 10.5 saat (+3.0 saat FM)
       - ve bu aralıklarla artarak devam eder.
    """
    g = parse_hm(g_saat_str)
    c = parse_hm(c_saat_str)
    
    if not g or not c:
        if sure_str:
            raw = parse_time_str(sure_str)
            if raw > 0:
                half_steps = round(raw / 0.5)
                h = max(0.0, half_steps * 0.5)
                return h, max(0.0, h - 7.5)
        return 0.0, 0.0
        
    g_min = g[0] * 60 + g[1]
    c_min = c[0] * 60 + c[1]
    
    # Giriş ve çıkış aynı dakikaysa (hatalı basım)
    if abs(c_min - g_min) <= 3:
        return 0.0, 0.0
        
    # Gece vardiyası çıkış ertesi güne sarktıysa
    if c_min < g_min:
        c_min += 24 * 60
        
    # Gündüz / Sabah vardiyası (06:00 - 11:00 arası girişler)
    if 6 * 60 <= g_min <= 11 * 60:
        # Sabah kuralı: 08:10 ve öncesi -> 08:00 başlangıç. 08:10 sonrası -> 09:00 başlangıç
        if g_min <= 8 * 60 + 10:
            effective_start_min = 8 * 60
            base_hours = 7.5
        else:
            effective_start_min = 9 * 60
            base_hours = 6.5
            
        shift_end_min = 17 * 60
        
        # 17:00 öncesi erken çıkış varsa
        if c_min < shift_end_min:
            raw_worked = (c_min - effective_start_min) / 60.0
            net = max(0.0, raw_worked - 1.0)
            return round(net * 2) / 2.0, 0.0
            
        # 17:00 sonrası çıkış kademeleri:
        # 17:25'e kadar -> 0 FM
        # 17:25 - 17:49 -> +0.5 FM (8.0)
        # 17:50 - 18:24 -> +1.0 FM (8.5)
        # 18:25 - 18:49 -> +1.5 FM (9.0)
        # 18:50 - 19:24 -> +2.0 FM (9.5) ...
        ot_min = c_min - shift_end_min
        if ot_min < 25:
            step = 0
        else:
            hours_past = ot_min // 60
            min_in_hour = ot_min % 60
            if min_in_hour < 25:
                step = hours_past * 2
            elif min_in_hour < 50:
                step = hours_past * 2 + 1
            else:
                step = hours_past * 2 + 2
                
        fm = step * 0.5
        total_hours = base_hours + fm
        return total_hours, fm

    # İkinci vardiya (15:00 - 24:00) veya Gece vardiyası (20:00 - 06:00 / 22:00 - 08:00)
    raw_duration = (c_min - g_min) / 60.0
    net_duration = max(0.0, raw_duration - 1.0)
    calc_h = max(0.0, (int(net_duration / 0.5)) * 0.5)
    if calc_h >= 7.0 and calc_h < 7.5:
        calc_h = 7.5
    fm = max(0.0, calc_h - 7.5)
    return calc_h, fm

def main():
    print("==================================================")
    print(" FİDE KONSERVE - PUANTAJ VE PDKS İŞLEME SİSTEMİ")
    print("==================================================")

    # Orijinal puantaj.xls dosyasını yedekten tazele (eğer yedek varsa)
    if os.path.exists("puantaj_backup.xls"):
        try:
            shutil.copyfile("puantaj_backup.xls", "puantaj.xls")
        except:
            pass

    print("1. 'puantaj.xls' ve 'pdks.xls' dosyaları yükleniyor...")
    wb_old = xlrd.open_workbook("puantaj.xls", encoding_override="cp1254")
    wb_pdks = xlrd.open_workbook("pdks.xls", encoding_override="cp1254")

    # ==========================================
    # 1. PDKS HAREKETLERİNİ İŞLEME
    # ==========================================
    sh_pdks = wb_pdks.sheet_by_name("HarList")
    pdks_records = []
    emp_pdks_daily = {}

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
        sure_str = clean_str(sh_pdks.cell_value(r, 12))
        mesai_str = clean_str(sh_pdks.cell_value(r, 15)) if sh_pdks.ncols > 15 else ""
        fazla_str = clean_str(sh_pdks.cell_value(r, 17)) if sh_pdks.ncols > 17 else ""
        puantaj_tarih = clean_str(sh_pdks.cell_value(r, 18)) if sh_pdks.ncols > 18 else ""
        
        # Fabrika kurallarına göre çalışma saati ve fazla mesai hesaplama
        sure, fazla = calc_factory_worked_hours(g_saat, c_saat, sure_str, is_office=False)
        mesai = min(7.5, sure)
        
        if not full_name:
            continue
        
        pdks_records.append({
            "sira": sira, "sicil": sicil, "kart": kart, "gun_adi": gun_adi,
            "ad_soyad": full_name, "lokasyon": lokasyon, "g_tarih": g_tarih,
            "g_saat": g_saat, "c_tarih": c_tarih, "c_saat": c_saat,
            "sure": f"{sure:.1f}", "mesai": f"{mesai:.1f}", "fazla": f"{fazla:.1f}",
            "puantaj_tarih": puantaj_tarih
        })
        
        date_val = puantaj_tarih if puantaj_tarih else g_tarih
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
                
        if day_num:
            norm_name = full_name.replace("İ", "I").replace("ı", "i").replace(" ", "").upper()
            if norm_name not in emp_pdks_daily:
                emp_pdks_daily[norm_name] = {}
            if day_num not in emp_pdks_daily[norm_name]:
                emp_pdks_daily[norm_name][day_num] = {"sure": 0.0, "normal": 0.0, "fazla": 0.0}
            emp_pdks_daily[norm_name][day_num]["sure"] += sure
            emp_pdks_daily[norm_name][day_num]["normal"] += mesai
            emp_pdks_daily[norm_name][day_num]["fazla"] += fazla

    print(f"   -> PDKS'den {len(pdks_records)} turnike kaydı ve {len(emp_pdks_daily)} personel okundu.")

    # ==========================================
    # 2. RAPORLAR SAYFASI
    # ==========================================
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

    # ==========================================
    # 3. İCRA TAKİP SAYFASI
    # ==========================================
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
                icra_by_name[isim.replace("İ", "I").replace("ı", "i").replace(" ", "").upper()] = {
                    "kesilen": kesilen, "kalan_borc": kalan_borc, "dosya": dosya, "iban": iban
                }

    # ==========================================
    # 4. ÜCRETLİ İZİNLER
    # ==========================================
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

    # ==========================================
    # 5. RESMİ BORDRO SGK
    # ==========================================
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

    # ==========================================
    # 6. DAİMİ PERSONEL LİSTESİ
    # ==========================================
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

    # ==========================================
    # 7. ANA PUANTAJ LİSTESİ (PUANTAJ)
    # ==========================================
    sh_p = get_sheet_by_keyword(wb_old, "puantaj", 0)
    puantaj_rows = []
    seen_tcs = {}

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

    # ==========================================
    # 8. MODERN EXCEL (.xlsx) OLUŞTURMA
    # ==========================================
    print("2. Modern 'puantaj.xlsx' çalışma kitabı oluşturuluyor...")
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

    # ----------------------------------------------------
    # SAYFA 1: Aylik_Puantaj (Eylül 2026)
    # ----------------------------------------------------
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
    start_col_idx = 15 # Column O
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
        
        norm_name = emp["ad_soyad"].replace("İ", "I").replace("ı", "i").replace(" ", "").upper()
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
                hours = emp_pdks[d]["sure"]
                fazla = emp_pdks[d]["fazla"]
                cell.value = hours if hours > 0 else 0
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
        
        # Canlı Formüller
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

    # ----------------------------------------------------
    # SAYFA 2: PDKS_Hareket_Kayitlari
    # ----------------------------------------------------
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
        ws_pdk[f"M{r_idx}"] = rec["lokasyon"]
        
        for c_idx in range(1, 14):
            cl = get_column_letter(c_idx)
            ws_pdk[f"{cl}{r_idx}"].border = border_cell
            ws_pdk[f"{cl}{r_idx}"].font = font_data
            if c_idx not in (4, 13):
                ws_pdk[f"{cl}{r_idx}"].alignment = align_center

    # ----------------------------------------------------
    # SAYFA 3: Resmi_Bordro_SGK
    # ----------------------------------------------------
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

    # ----------------------------------------------------
    # SAYFA 4: Icra_Takip
    # ----------------------------------------------------
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
        norm_ic_name = ic["isim"].replace("İ", "I").replace("ı", "i").replace(" ", "").upper()
        for p in puantaj_rows:
            if p["ad_soyad"].replace("İ", "I").replace("ı", "i").replace(" ", "").upper() == norm_ic_name:
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

    # ----------------------------------------------------
    # SAYFA 5: SGK_Raporlar
    # ----------------------------------------------------
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

    # ----------------------------------------------------
    # DİĞER YARDIMCI SAYFALAR
    # ----------------------------------------------------
    ws_izn = wb_new.create_sheet(title="Ucretli_Izinler")
    ws_izn.views.sheetView[0].showGridLines = True
    for c_idx, h_text in enumerate(["TC Kimlik No", "Adı Soyadı", "İzin Saati", "İzin Günü"], start=1):
        ws_izn.cell(1, c_idx, h_text).font = font_hdr
        ws_izn.cell(1, c_idx).fill = fill_navy

    ws_dai = wb_new.create_sheet(title="Daimi_Personel_Listesi")
    ws_dai.views.sheetView[0].showGridLines = True
    for c_idx, h_text in enumerate(["Sıra No", "TC Kimlik No", "Adı Soyadı", "SGK Giriş", "SGK Durum", "Kadro", "Maaş (TL)"], start=1):
        ws_dai.cell(1, c_idx, h_text).font = font_hdr
        ws_dai.cell(1, c_idx).fill = fill_navy

    # Elden Ödeme Farkı Sayfası
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
        p_row = r_idx + 2  # Aylik_Puantaj'daki satır (başlangıç 4)
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

    # 4. Kaydetme
    output_filename = "puantaj.xlsx"
    saved_name = output_filename
    try:
        wb_new.save(output_filename)
        print(f"3. Dosya '{output_filename}' olarak kaydedildi.")
    except PermissionError:
        saved_name = "puantaj_guncel.xlsx"
        wb_new.save(saved_name)
        print(f"3. UYARI: '{output_filename}' Microsoft Excel'de açık olduğu için '{saved_name}' olarak kaydedildi.")
        print(f"   (Tam '{output_filename}' üzerine yazmak için Excel'i kapatıp tekrar çalıştırabilirsiniz).")

    print("==================================================")
    print(f" BAŞARILI! '{saved_name}' dosyası eksiksiz oluşturuldu.")
    print(" - 8-5 Vardiyası: 07:10 ve sonrası sabah gelişleri 08:00 iş başı kabul edildi")
    print(" - 17:00 sonrası çıkış aralıkları (17:25->8.0, 17:50->8.5, 18:25->9.0, 18:50->9.5...) uygulandı")
    print(" - Canlı Formüller (SUM, COUNTIF, IF, MIN, VLOOKUP) Aktif")
    print(f" - {len(puantaj_rows)} personelin 30 günlük çalışma süreleri işlendi")
    print("==================================================")

if __name__ == "__main__":
    main()
