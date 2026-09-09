# -*- coding: utf-8 -*-
"""
Puantaj ve PDKS Verisi Eşleştirme ve İşleme Scripti
Fide Konserve Gıda San. A.Ş.

Bu script, 'pdks.xls' dosyasındaki turnike giriş-çıkış ve süre kayıtlarını okuyarak
doğrudan mevcut 'puantaj.xls' dosyasının 'PUANTAJ' sayfasına gün gün (1-30 Eylül) işler.
"""
import xlrd
from xlutils.copy import copy
import shutil
import re
import os

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

def normalize_name(name):
    if not name:
        return ""
    s = repair_text(name).upper()
    # Karakter eşitleme
    s = s.replace("İ", "I").replace("ı", "I").replace("Ş", "S").replace("Ğ", "G")
    s = s.replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

def clean_tc(tc_val):
    if tc_val is None:
        return ""
    if isinstance(tc_val, float):
        s = str(int(round(tc_val))).strip()
    else:
        s = str(tc_val).replace(".0", "").strip()
    s = re.sub(r"\D", "", s)
    if len(s) == 10:
        s = "0" + s
    return s

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

def process_pdks_to_puantaj(pdks_file="pdks.xls", puantaj_file="puantaj.xls", backup=True):
    print("==================================================")
    print(" PDKS -> PUANTAJ VERİ İŞLEME SİSTEMİ")
    print("==================================================")
    
    if not os.path.exists(pdks_file):
        print(f"HATA: '{pdks_file}' dosyası bulunamadı!")
        return
    if not os.path.exists(puantaj_file):
        print(f"HATA: '{puantaj_file}' dosyası bulunamadı!")
        return

    # 1. Yedekleme
    if backup:
        backup_file = "puantaj_yedek.xls"
        shutil.copyfile(puantaj_file, backup_file)
        print(f"[1/4] Güvenlik yedeği alındı: '{backup_file}'")

    # 2. PDKS Verilerini Okuma
    print(f"[2/4] '{pdks_file}' okunuyor...")
    wb_pdks = xlrd.open_workbook(pdks_file, encoding_override="cp1254")
    sh_pdks = wb_pdks.sheet_by_name("HarList")
    
    # Personel bazında günlük çalışma saatlerini topla
    emp_daily_hours = {}
    pdks_count = 0
    
    for r in range(1, sh_pdks.nrows):
        adi = str(sh_pdks.cell_value(r, 5)).strip()
        soyadi = str(sh_pdks.cell_value(r, 6)).strip()
        full_name = f"{adi} {soyadi}".strip()
        sicil = str(sh_pdks.cell_value(r, 2)).strip()
        
        g_tarih = str(sh_pdks.cell_value(r, 8)).strip()
        puantaj_tarih = str(sh_pdks.cell_value(r, 18)).strip() if sh_pdks.ncols > 18 else ""
        date_str = puantaj_tarih if puantaj_tarih else g_tarih
        
        sure_str = str(sh_pdks.cell_value(r, 12)).strip()
        hours = parse_time_str(sure_str)
        
        # Gün numarasını çıkar (1-30)
        day_num = None
        if "." in date_str:
            parts = date_str.split(".")
            try:
                day_num = int(parts[0])
            except:
                pass
                
        if full_name and day_num and (1 <= day_num <= 31):
            norm_key = normalize_name(full_name)
            if norm_key not in emp_daily_hours:
                emp_daily_hours[norm_key] = {"full_name": full_name, "days": {}}
            emp_daily_hours[norm_key]["days"][day_num] = emp_daily_hours[norm_key]["days"].get(day_num, 0.0) + hours
            pdks_count += 1
            
    print(f"      -> {pdks_count} adet turnike hareketi okundu.")
    print(f"      -> {len(emp_daily_hours)} tekil personel tespit edildi.")

    # 3. Puantaj Dosyasını Açma ve Eşleştirme
    print(f"[3/4] '{puantaj_file}' dosyasına veriler işleniyor...")
    rb_puantaj = xlrd.open_workbook(puantaj_file, formatting_info=True, encoding_override="cp1254")
    sh_puantaj = rb_puantaj.sheet_by_name("PUANTAJ")
    
    # Yazılabilir kopya oluştur
    wb_writable = copy(rb_puantaj)
    ws_puantaj = wb_writable.get_sheet(0)  # PUANTAJ sayfası
    
    # Gün sütunlarının başlangıcı: 
    # Col 33 = 1 Eylül, Col 34 = 2 Eylül, ..., Col 62 = 30 Eylül
    day_start_col = 32  # index 32 + gün_no = ilgili sütun
    
    matched_employees = 0
    total_written_cells = 0
    
    for r in range(2, sh_puantaj.nrows):
        raw_name = str(sh_puantaj.cell_value(r, 2)).strip()
        if not raw_name:
            continue
            
        norm_key = normalize_name(raw_name)
        
        if norm_key in emp_daily_hours:
            matched_employees += 1
            days_data = emp_daily_hours[norm_key]["days"]
            
            for day in range(1, 31):
                col_idx = day_start_col + day
                if day in days_data:
                    h_val = round(days_data[day], 2)
                    ws_puantaj.write(r, col_idx, h_val)
                    total_written_cells += 1
                else:
                    # Çalışma yoksa mevcut boş bırakılır veya 0 yazılabilir
                    pass

    # 4. Kaydetme
    wb_writable.save(puantaj_file)
    print(f"[4/4] BAŞARILI: '{puantaj_file}' dosyası doğrudan güncellendi!")
    print("--------------------------------------------------")
    print(f"  Toplam Eşleşen Personel Sayısı : {matched_employees}")
    print(f"  Puantaja Yazılan Gün/Saat Hücresi: {total_written_cells}")
    print("==================================================")

if __name__ == "__main__":
    process_pdks_to_puantaj("pdks.xls", "puantaj.xls", backup=True)
