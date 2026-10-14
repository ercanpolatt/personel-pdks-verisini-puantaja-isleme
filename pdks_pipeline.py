# -*- coding: utf-8 -*-
"""
===================================================================================================
FİDE KONSERVE - PDKS ZAMAN & DEVAM VERİ İŞLEME VE MESAİ HESAPLAMA MOTORU (pdks_pipeline.py)
===================================================================================================
Senior Python Software Engineer & Workforce Data Specialist Implementation

Bu modül; ham PDKS turnike verilerini (pdks.xls) okur, veri kalitesi temizliği (debouncing,
tarih/saat normalizasyonu, eksik basım tamamlama) yapar, vardiya sınıflandırmasını gerçekleştirir
ve fabrika çalışma kurallarına uygun olarak net çalışma saatlerini ve yasal gece çalışma sürelerini
deterministik olarak hesaplar.

Temel Özellikler:
1. Temiz Mimari: DataLoader, Cleaner, ShiftClassifier, HoursCalculator, ReportGenerator.
2. Deterministik Kural Motoru: Yapay çarpanlar ve şişirilmiş saatler kaldırılmış, net kurallar.
3. Yasal Gece Çalışması Takibi: 20:00 - 06:00 arası yasal sürelerin kesin hesabı (İş K. m.69).
4. Otomatik Eksik Basım Çözümlemesi: Tek giriş/çıkış hareketlerinin atanmış vardiyaya göre tamamlanması.
5. İki Sayfalı Excel Çıktısı: 'Daily_Summary' ve 'Audit_Alerts' sayfalarıyla zengin açık denetim.
===================================================================================================
"""

from __future__ import annotations

import os
import sys
import re
import logging
from dataclasses import dataclass, field
from datetime import datetime, date, time, timedelta
from typing import Optional, Dict, Any, List, Tuple, Union

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Loglama Konfigürasyonu
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger("PDKSPipeline")


# =================================================================================================
# 1. KONFİGÜRASYON VE VERİ SINIFLARI (Configuration & Dataclasses)
# =================================================================================================

@dataclass(frozen=True)
class ShiftSchedule:
    """Tek bir vardiya tipine ait çalışma penceresi ve mola kuralları."""
    code: str
    name: str
    scheduled_start: time
    scheduled_end: time
    scheduled_gross_hours: float
    standard_break_hours: float
    standard_net_hours: float
    is_rotating: bool = False
    is_overnight: bool = False
    description: str = ""


@dataclass
class PDKSConfig:
    """PDKS işleme motorunun tüm eşik değerlerini ve vardiya kurallarını içeren konfigürasyon."""
    # Debounce Eşiği: Aynı personelin peş peşe basımlarında filtrelenecek dakika farkı
    debounce_threshold_minutes: float = 5.0

    # Süre Denetim Eşiği: Ham Süre sütunu ile yeniden hesaplanan süre arasındaki izin verilen fark
    duration_mismatch_threshold_minutes: float = 5.0

    # Gece Çalışması Sınırları (4857 Sayılı İş Kanunu m.69 gereğince 20:00 - 06:00)
    night_start_hour: int = 20
    night_start_minute: int = 0
    night_end_hour: int = 6
    night_end_minute: int = 0
    max_legal_night_hours: float = 7.5

    # Günlük Maksimum Yasal Çalışma Sınırı (Guard)
    max_daily_presence_hours: float = 24.0
    excessive_presence_alert_hours: float = 12.0

    # Tanımlı Vardiyalar
    schedules: Dict[str, ShiftSchedule] = field(default_factory=lambda: {
        # A. Standart Gündüz Vardiyası (08:00 - 17:00 / 9 Saat Brüt)
        "DAY_08_17": ShiftSchedule(
            code="DAY_08_17",
            name="Standart Gündüz (08:00 - 17:00)",
            scheduled_start=time(8, 0),
            scheduled_end=time(17, 0),
            scheduled_gross_hours=9.0,
            standard_break_hours=1.5,
            standard_net_hours=7.5,
            is_rotating=False,
            is_overnight=False,
            description="9 saat brüt, 1.5 saat mola, 7.5 saat net mesai."
        ),
        # B. Dönen 3-Vardiya: 1. Vardiya Sabah (08:00 - 16:00 / 8 Saat Brüt)
        "SHIFT_08_16": ShiftSchedule(
            code="SHIFT_08_16",
            name="3-Vardiya Sabah (08:00 - 16:00)",
            scheduled_start=time(8, 0),
            scheduled_end=time(16, 0),
            scheduled_gross_hours=8.0,
            standard_break_hours=0.5,
            standard_net_hours=7.5,
            is_rotating=True,
            is_overnight=False,
            description="8 saat brüt, 0.5 saat mola, 7.5 saat net mesai."
        ),
        # B. Dönen 3-Vardiya: 2. Vardiya Akşam (16:00 - 24:00 / 8 Saat Brüt)
        "SHIFT_16_24": ShiftSchedule(
            code="SHIFT_16_24",
            name="3-Vardiya Akşam (16:00 - 24:00)",
            scheduled_start=time(16, 0),
            scheduled_end=time(0, 0),
            scheduled_gross_hours=8.0,
            standard_break_hours=0.5,
            standard_net_hours=7.5,
            is_rotating=True,
            is_overnight=True,
            description="8 saat brüt, 0.5 saat mola, 7.5 saat net mesai."
        ),
        # B. Dönen 3-Vardiya: 3. Vardiya Gece (00:00 - 08:00 / 24-08 / 8 Saat Brüt)
        "SHIFT_00_08": ShiftSchedule(
            code="SHIFT_00_08",
            name="3-Vardiya Gece (00:00 - 08:00)",
            scheduled_start=time(0, 0),
            scheduled_end=time(8, 0),
            scheduled_gross_hours=8.0,
            standard_break_hours=0.5,
            standard_net_hours=7.5,
            is_rotating=True,
            is_overnight=False,
            description="8 saat brüt, 0.5 saat mola, 7.5 saat net mesai."
        ),
        # C. Esnek / Geç Gündüz (10:00 - 19:00 / 9 Saat Brüt)
        "SHIFT_10_19": ShiftSchedule(
            code="SHIFT_10_19",
            name="Geç Gündüz (10:00 - 19:00)",
            scheduled_start=time(10, 0),
            scheduled_end=time(19, 0),
            scheduled_gross_hours=9.0,
            standard_break_hours=1.5,
            standard_net_hours=7.5,
            is_rotating=False,
            is_overnight=False,
            description="9 saat brüt, 1.5 saat mola, 7.5 saat net."
        ),
        # D. 12 Saatlik Gece Vardiyası (20:00 - 08:00)
        "SHIFT_20_08": ShiftSchedule(
            code="SHIFT_20_08",
            name="12-Saat Gece (20:00 - 08:00)",
            scheduled_start=time(20, 0),
            scheduled_end=time(8, 0),
            scheduled_gross_hours=12.0,
            standard_break_hours=1.5,
            standard_net_hours=10.5,
            is_rotating=True,
            is_overnight=True,
            description="12 saat brüt gece vardiyası."
        )
    })


# =================================================================================================
# 2. YARDIMCI ZAMAN VE TARİH FONKSİYONLARI (Time & Date Helpers)
# =================================================================================================

def parse_flexible_date(val: Any) -> Optional[date]:
    """
    Çeşitli biçimlerdeki tarih verilerini (DD.MM.YYYY, YYYY-MM-DD, datetime vb.)
    standart datetime.date nesnesine dönüştürür.
    """
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.date()
    if isinstance(val, date):
        return val

    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "nat", "-"):
        return None

    # Olası saat kısmını ayıkla (örn. '2026-09-02 00:00:00')
    if " " in s:
        s = s.split(" ")[0].strip()

    # 1. DD.MM.YYYY formatı
    if "." in s:
        parts = s.split(".")
        if len(parts) == 3:
            try:
                d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                if y < 100:
                    y += 2000
                return date(y, m, d)
            except ValueError:
                pass

    # 2. YYYY-MM-DD veya DD-MM-YYYY formatı
    if "-" in s:
        parts = s.split("-")
        if len(parts) == 3:
            try:
                if len(parts[0]) == 4:  # YYYY-MM-DD
                    y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                else:  # DD-MM-YYYY
                    d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                return date(y, m, d)
            except ValueError:
                pass

    # 3. Pandas genel parse
    try:
        dt = pd.to_datetime(s, dayfirst=True)
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass

    return None


def parse_flexible_time(val: Any) -> Optional[time]:
    """
    '08:15:00', '08:15', '8:15' veya time nesnelerini standart datetime.time nesnesine dönüştürür.
    """
    if val is None or pd.isna(val):
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.time()

    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "nat", "-"):
        return None

    if " " in s:
        s = s.split(" ")[-1].strip()

    parts = s.split(":")
    if len(parts) >= 2:
        try:
            h = int(parts[0])
            m = int(parts[1])
            sec = int(parts[2]) if len(parts) > 2 else 0
            if h == 24:  # 24:00:00 gösterimi gece yarısıdır
                h = 0
            if 0 <= h <= 23 and 0 <= m <= 59 and 0 <= sec <= 59:
                return time(h, m, sec)
        except ValueError:
            pass

    return None


def parse_duration_to_minutes(val: Any) -> Optional[float]:
    """'10:35' biçimindeki Süre metnini toplam dakikaya (float) çevirir."""
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "-"):
        return None
    parts = s.split(":")
    if len(parts) >= 2:
        try:
            h = int(parts[0])
            m = int(parts[1])
            return float(h * 60 + m)
        except ValueError:
            return None
    return None


def calculate_legal_night_hours(check_in_dt: datetime, check_out_dt: datetime) -> float:
    """
    4857 Sayılı İş Kanunu m.69 uyarınca strictly 20:00 ile 06:00 saatleri arasına denk
    gelen yasal gece çalışma süresini saat cinsinden (ondalık) hesaplar.
    Vardiyanın gün dönümünü (gece yarısını) geçmesi durumlarını kusursuz yönetir.
    """
    if check_in_dt >= check_out_dt:
        return 0.0

    # İlgili tarih aralığındaki tüm gece pencerelerini oluştur
    # Her takvim günü için gece penceresi: o gün saat 20:00'den ertesi gün 06:00'ya kadardır (10 saat).
    start_d = check_in_dt.date() - timedelta(days=1)
    end_d = check_out_dt.date() + timedelta(days=1)

    total_night_seconds = 0.0
    cur_d = start_d
    while cur_d <= end_d:
        night_window_start = datetime(cur_d.year, cur_d.month, cur_d.day, 20, 0, 0)
        night_window_end = datetime(cur_d.year, cur_d.month, cur_d.day, 6, 0, 0) + timedelta(days=1)

        # Kesişim hesabı (Intersection)
        overlap_start = max(check_in_dt, night_window_start)
        overlap_end = min(check_out_dt, night_window_end)

        if overlap_end > overlap_start:
            total_night_seconds += (overlap_end - overlap_start).total_seconds()

        cur_d += timedelta(days=1)

    return round(total_night_seconds / 3600.0, 2)


# =================================================================================================
# 3. VERİ YÜKLEYİCİ (DataLoader)
# =================================================================================================

class DataLoader:
    """
    Ham PDKS turnike Excel/CSV dışa aktarım dosyasını okur, olası karakter
    kodlama (encoding) bozulmalarını tolere eder ve sütunları standart isimlere eşler.
    """

    COLUMN_PATTERNS = {
        "employee_id": [r"sicil", r"personel.*no", r"emp.*id"],
        "card_no": [r"kart.*no", r"card.*no"],
        "day_name": [r"g[üu]n$"],
        "first_name": [r"^ad[ıi]?$"],
        "last_name": [r"^soyad[ıi]?$"],
        "location": [r"lokasyon", r"b[öo]lge"],
        "check_in_date": [r"giri[şs].*tarih"],
        "check_in_time": [r"giri[şs].*saat"],
        "check_out_date": [r"[çc][ıi]k[ıi][şs].*tarih"],
        "check_out_time": [r"[çc][ıi]k[ıi][şs].*saat"],
        "raw_duration": [r"s[üu]re$"],
        "shift_code": [r"vardiya.*kod"],
        "group_code": [r"grup.*kod"]
    }

    def __init__(self, config: Optional[PDKSConfig] = None):
        self.config = config or PDKSConfig()

    def load_file(self, file_path: str) -> pd.DataFrame:
        """Dosyayı okur ve sütunlarını kanonik isimlere dönüştürerek DataFrame döndürür."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"PDKS dosyası bulunamadı: {file_path}")

        logger.info(f"PDKS dosyası yükleniyor: {file_path}")
        df: Optional[pd.DataFrame] = None

        # 1. Excel Okuma (xls / xlsx)
        if file_path.endswith((".xls", ".xlsx")):
            try:
                # Legacy .xls için xlrd, modern .xlsx için openpyxl
                engine = "xlrd" if file_path.endswith(".xls") else "openpyxl"
                df = pd.read_excel(file_path, engine=engine)
            except Exception as e:
                logger.warning(f"Birincil Excel motoru ({engine}) hata verdi: {e}. Alternatif deneniyor...")
                df = pd.read_excel(file_path)

        # 2. CSV Okuma Fallback
        elif file_path.endswith(".csv"):
            for enc in ("utf-8", "utf-8-sig", "iso-8859-9", "cp1254"):
                try:
                    df = pd.read_csv(file_path, encoding=enc, sep=None, engine="python")
                    break
                except Exception:
                    continue

        if df is None:
            raise ValueError(f"Dosya okunamadı veya desteklenmeyen format: {file_path}")

        logger.info(f"Ham veri başarıyla yüklendi: {len(df)} satır, {len(df.columns)} sütun.")
        normalized_df = self._normalize_columns(df)
        return normalized_df

    def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Karakter bozulmalarından arındırarak kanonik sütun haritasını kurar."""
        rename_map = {}
        matched_standards = set()

        for col in df.columns:
            clean_col = str(col).strip().lower()
            # Bozuk Türkçe karakterleri temizle
            clean_col = clean_col.replace("", "").replace("?", "").strip()

            for std_name, patterns in self.COLUMN_PATTERNS.items():
                if std_name in matched_standards:
                    continue
                for pat in patterns:
                    if re.search(pat, clean_col, re.IGNORECASE):
                        rename_map[col] = std_name
                        matched_standards.add(std_name)
                        break

        df_renamed = df.rename(columns=rename_map)

        # Gerekli minimum sütunların varlığını doğrula
        essential_cols = ["employee_id", "check_in_date", "check_in_time"]
        missing = [c for c in essential_cols if c not in df_renamed.columns]
        if missing:
            logger.warning(f"Bazı temel sütunlar eşlenemedi: {missing}. Orijinal sütunlar: {df.columns.tolist()[:10]}")

        return df_renamed


# =================================================================================================
# 4. VERİ TEMİZLEYİCİ VE ANOMALİ ONARICI (Cleaner)
# =================================================================================================

class Cleaner:
    """
    Ham verilerde;
    1. Tarih ve saatleri native datetime nesnelerine çevirir.
    2. Debounce: 5 dakika içindeki mükerrer basımları ayıklar.
    3. Eksik giriş veya çıkışları vardiya takvimine göre tamamlar (auto-imputation).
    4. Süre farklarını (> 5 dk delta) tespit eder.
    """

    def __init__(self, config: Optional[PDKSConfig] = None):
        self.config = config or PDKSConfig()

    def clean_records(self, df: pd.DataFrame) -> pd.DataFrame:
        """Tüm temizlik ve veri zenginleştirme adımlarını uygular."""
        clean_df = df.copy()

        # 1. Personel Bilgilerini Düzelt
        if "employee_id" in clean_df.columns:
            clean_df["employee_id"] = clean_df["employee_id"].apply(self._clean_id)
        else:
            clean_df["employee_id"] = "UNKNOWN"

        if "first_name" in clean_df.columns:
            first_name = clean_df["first_name"].fillna("").astype(str)
        else:
            first_name = pd.Series([""] * len(clean_df), index=clean_df.index)

        if "last_name" in clean_df.columns:
            last_name = clean_df["last_name"].fillna("").astype(str)
        else:
            last_name = pd.Series([""] * len(clean_df), index=clean_df.index)

        clean_df["full_name"] = (first_name + " " + last_name).str.strip()

        # 2. Tarih ve Saatleri Parse Et
        if "check_in_date" in clean_df.columns:
            clean_df["cin_date_obj"] = clean_df["check_in_date"].apply(parse_flexible_date)
        else:
            clean_df["cin_date_obj"] = pd.Series([None] * len(clean_df), index=clean_df.index)

        if "check_in_time" in clean_df.columns:
            clean_df["cin_time_obj"] = clean_df["check_in_time"].apply(parse_flexible_time)
        else:
            clean_df["cin_time_obj"] = pd.Series([None] * len(clean_df), index=clean_df.index)

        if "check_out_date" in clean_df.columns:
            clean_df["cout_date_obj"] = clean_df["check_out_date"].apply(parse_flexible_date)
        else:
            clean_df["cout_date_obj"] = pd.Series([None] * len(clean_df), index=clean_df.index)

        if "check_out_time" in clean_df.columns:
            clean_df["cout_time_obj"] = clean_df["check_out_time"].apply(parse_flexible_time)
        else:
            clean_df["cout_time_obj"] = pd.Series([None] * len(clean_df), index=clean_df.index)

        # 3. İki basımı da boş olan geçersiz satırları filtrele
        both_null_mask = clean_df["cin_time_obj"].isna() & clean_df["cout_time_obj"].isna()
        if both_null_mask.any():
            dropped_count = both_null_mask.sum()
            logger.info(f"Hem giriş hem çıkış saati boş olan {dropped_count} satır elendi.")
            clean_df = clean_df[~both_null_mask].copy()

        # 4. Debounce: Aynı personelin 5 dakika içindeki mükerrer basımlarını temizle
        clean_df = self._apply_debounce(clean_df)

        # 5. Eksik Basımları Tamamla (Single Punch Imputation)
        clean_df = self._impute_missing_swipes(clean_df)

        # 6. Kesin Check-In ve Check-Out Datetime Nesnelerini Oluştur
        clean_df = self._build_canonical_datetimes(clean_df)

        # 7. Önceden var olan Süre ile karşılaştırma yap
        clean_df = self._audit_raw_duration(clean_df)

        return clean_df

    @staticmethod
    def _clean_id(val: Any) -> str:
        """Sicil / TC numarasını temiz float/str'den temiz string'e çevirir."""
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip()
        if s.endswith(".0"):
            s = s[:-2]
        return s

    def _apply_debounce(self, df: pd.DataFrame) -> pd.DataFrame:
        """Aynı personelin 5 dakika içerisindeki mükerrer basımlarını ayıklar."""
        if len(df) == 0:
            return df

        # Geçici timestamp oluştur
        def make_ts(row):
            d = row["cin_date_obj"] or row["cout_date_obj"]
            t = row["cin_time_obj"] or row["cout_time_obj"]
            if d and t:
                return datetime.combine(d, t)
            return pd.NaT

        df["_temp_swipe_ts"] = df.apply(make_ts, axis=1)
        valid_ts = df["_temp_swipe_ts"].notna()

        df_valid = df[valid_ts].sort_values(by=["employee_id", "_temp_swipe_ts"])
        df_valid["_prev_ts"] = df_valid.groupby("employee_id")["_temp_swipe_ts"].shift(1)
        df_valid["_diff_sec"] = (df_valid["_temp_swipe_ts"] - df_valid["_prev_ts"]).dt.total_seconds()

        threshold_sec = self.config.debounce_threshold_minutes * 60.0
        is_duplicate = df_valid["_diff_sec"] < threshold_sec

        debounced_count = is_duplicate.sum()
        if debounced_count > 0:
            logger.info(f"Debounce: 5 dakika içindeki {debounced_count} mükerrer basım elendi.")
            keep_indices = df_valid[~is_duplicate].index
            # Geçerli olmayan ts'li satırları da koru
            final_indices = keep_indices.union(df[~valid_ts].index)
            df = df.loc[final_indices].copy()

        df.drop(columns=["_temp_swipe_ts", "_prev_ts", "_diff_sec"], errors="ignore", inplace=True)
        return df

    def _impute_missing_swipes(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Tek giriş veya tek çıkış olan satırlarda karşı basımı vardiya saatine göre tamamlar.
        """
        df["imputed_entry"] = False
        df["imputed_exit"] = False
        df["audit_status"] = "NORMAL"

        for idx, row in df.iterrows():
            cin_t = row["cin_time_obj"]
            cout_t = row["cout_time_obj"]
            cin_d = row["cin_date_obj"]
            cout_d = row["cout_date_obj"]
            shift_c = str(row.get("shift_code", "")).strip().upper()

            # Durum 1: Yalnızca Giriş Var (Çıkış Yok)
            if cin_t is not None and cout_t is None:
                df.at[idx, "imputed_exit"] = True
                df.at[idx, "audit_status"] = "AUTO_IMPUTED_EXIT"

                ref_date = cin_d or date.today()
                df.at[idx, "cout_date_obj"] = ref_date

                # Vardiyaya göre çıkış saati belirle
                if "08-16" in shift_c:
                    df.at[idx, "cout_time_obj"] = time(16, 0)
                elif "16-24" in shift_c:
                    df.at[idx, "cout_time_obj"] = time(0, 0)
                    df.at[idx, "cout_date_obj"] = ref_date + timedelta(days=1)
                elif "24-08" in shift_c or "00-08" in shift_c:
                    df.at[idx, "cout_time_obj"] = time(8, 0)
                    df.at[idx, "cout_date_obj"] = ref_date + timedelta(days=1)
                elif "10-19" in shift_c:
                    df.at[idx, "cout_time_obj"] = time(19, 0)
                elif "20-08" in shift_c:
                    df.at[idx, "cout_time_obj"] = time(8, 0)
                    df.at[idx, "cout_date_obj"] = ref_date + timedelta(days=1)
                else:
                    # Giriş saatine göre dinamik eşleme
                    if cin_t.hour <= 11:
                        df.at[idx, "cout_time_obj"] = time(17, 0)  # Standart gün sonu
                    elif 14 <= cin_t.hour <= 17:
                        df.at[idx, "cout_time_obj"] = time(0, 0)
                        df.at[idx, "cout_date_obj"] = ref_date + timedelta(days=1)
                    else:
                        df.at[idx, "cout_time_obj"] = time(8, 0)
                        df.at[idx, "cout_date_obj"] = ref_date + timedelta(days=1)

            # Durum 2: Yalnızca Çıkış Var (Giriş Yok)
            elif cin_t is None and cout_t is not None:
                df.at[idx, "imputed_entry"] = True
                df.at[idx, "audit_status"] = "AUTO_IMPUTED_ENTRY"

                ref_date = cout_d or date.today()
                df.at[idx, "cin_date_obj"] = ref_date

                # Vardiyaya göre giriş saati belirle
                if "08-16" in shift_c or "08-17" in shift_c:
                    df.at[idx, "cin_time_obj"] = time(8, 0)
                elif "16-24" in shift_c:
                    df.at[idx, "cin_time_obj"] = time(16, 0)
                    if cout_t == time(0, 0) or cout_t.hour < 4:
                        df.at[idx, "cin_date_obj"] = ref_date - timedelta(days=1)
                elif "24-08" in shift_c or "00-08" in shift_c:
                    df.at[idx, "cin_time_obj"] = time(0, 0)
                elif "10-19" in shift_c:
                    df.at[idx, "cin_time_obj"] = time(10, 0)
                elif "20-08" in shift_c:
                    df.at[idx, "cin_time_obj"] = time(20, 0)
                    df.at[idx, "cin_date_obj"] = ref_date - timedelta(days=1)
                else:
                    # Çıkış saatine göre dinamik eşleme
                    if 15 <= cout_t.hour <= 20:
                        df.at[idx, "cin_time_obj"] = time(8, 0)
                    elif cout_t.hour == 0 or cout_t.hour <= 2:
                        df.at[idx, "cin_time_obj"] = time(16, 0)
                        df.at[idx, "cin_date_obj"] = ref_date - timedelta(days=1)
                    elif 6 <= cout_t.hour <= 10:
                        df.at[idx, "cin_time_obj"] = time(0, 0)
                    else:
                        df.at[idx, "cin_time_obj"] = time(8, 0)

        return df

    def _build_canonical_datetimes(self, df: pd.DataFrame) -> pd.DataFrame:
        """Kesin check_in_dt ve check_out_dt nesnelerini birleştirir ve gece geçişlerini yönetir."""
        check_in_dts = []
        check_out_dts = []
        is_overnight_flags = []

        for _, row in df.iterrows():
            d_in = row["cin_date_obj"]
            t_in = row["cin_time_obj"]
            d_out = row["cout_date_obj"]
            t_out = row["cout_time_obj"]

            if d_in and t_in:
                dt_in = datetime.combine(d_in, t_in)
            else:
                dt_in = pd.NaT

            if d_out and t_out:
                dt_out = datetime.combine(d_out, t_out)
            elif d_in and t_out:
                dt_out = datetime.combine(d_in, t_out)
            else:
                dt_out = pd.NaT

            # Gece yarısı geçiş kontrolü (Midnight Cross)
            is_overnight = False
            if pd.notna(dt_in) and pd.notna(dt_out):
                if dt_out < dt_in:
                    # Çıkış tarihi aynı gün yazılmış ama saat girişe göre küçükse -> ertesi güne taşır
                    dt_out += timedelta(days=1)
                    is_overnight = True
                elif dt_out.date() > dt_in.date():
                    is_overnight = True

            check_in_dts.append(dt_in)
            check_out_dts.append(dt_out)
            is_overnight_flags.append(is_overnight)

        df["check_in_dt"] = check_in_dts
        df["check_out_dt"] = check_out_dts
        df["is_overnight"] = is_overnight_flags
        return df

    def _audit_raw_duration(self, df: pd.DataFrame) -> pd.DataFrame:
        """Hesaplanan brüt süre ile dosyadaki 'Süre' sütununu karşılaştırır ve 5 dk üzeri farkları bayraklar."""
        df["duration_mismatch"] = False
        df["duration_delta_min"] = 0.0

        if "raw_duration" not in df.columns:
            return df

        for idx, row in df.iterrows():
            dt_in = row["check_in_dt"]
            dt_out = row["check_out_dt"]
            raw_dur_str = row.get("raw_duration")

            if pd.notna(dt_in) and pd.notna(dt_out) and pd.notna(raw_dur_str):
                calc_minutes = (dt_out - dt_in).total_seconds() / 60.0
                raw_minutes = parse_duration_to_minutes(raw_dur_str)

                if raw_minutes is not None:
                    delta = calc_minutes - raw_minutes
                    df.at[idx, "duration_delta_min"] = round(delta, 1)

                    if abs(delta) > self.config.duration_mismatch_threshold_minutes:
                        df.at[idx, "duration_mismatch"] = True
                        if df.at[idx, "audit_status"] == "NORMAL":
                            df.at[idx, "audit_status"] = "DURATION_MISMATCH"

        return df


# =================================================================================================
# 5. VARDİYA SINIFLANDIRICI (ShiftClassifier)
# =================================================================================================

class ShiftClassifier:
    """
    Çalışanın vardiya modelini (Standart Gündüz 08-17 vs Dönen 3-Vardiya 8 Saatlik Sistem)
    dinamik olarak sınıflandırır.
    
    KRİTİK GEREKSİNİM:
    - 08-16 dönen vardiya işçisi 8 saatini tamamladığında -> 7.5 saat net mesai alır (Mola = 0.5s).
    - 08-17 standart gündüz işçisi 1 saat erken çıkıp 8 saat kaldığında -> 6.5 saat net mesai alır (Mola = 1.5s).
    """

    def __init__(self, config: Optional[PDKSConfig] = None):
        self.config = config or PDKSConfig()

    def classify_shifts(self, df: pd.DataFrame) -> pd.DataFrame:
        """Her hareket satırı için shift_type ve shift_name değerlerini belirler."""
        shift_types = []
        shift_names = []

        for _, row in df.iterrows():
            shift_code = str(row.get("shift_code", "")).strip().upper()
            dt_in = row["check_in_dt"]
            dt_out = row["check_out_dt"]
            is_overnight = row.get("is_overnight", False)

            s_type, s_name = self._determine_shift(shift_code, dt_in, dt_out, is_overnight)
            shift_types.append(s_type)
            shift_names.append(s_name)

        df["shift_type"] = shift_types
        df["shift_name"] = shift_names
        return df

    def _determine_shift(
        self,
        shift_code: str,
        dt_in: Optional[datetime],
        dt_out: Optional[datetime],
        is_overnight: bool
    ) -> Tuple[str, str]:
        """Tek bir kayıt için vardiya eşleşmesi üretir."""
        # 1. Açık Kod Kontrolü (Explicit Code Matching)
        if "08-16" in shift_code:
            return "SHIFT_08_16", "3-Vardiya Sabah (08:00 - 16:00)"

        if "16-24" in shift_code:
            return "SHIFT_16_24", "3-Vardiya Akşam (16:00 - 24:00)"

        if "24-08" in shift_code or "00-08" in shift_code:
            return "SHIFT_00_08", "3-Vardiya Gece (00:00 - 08:00)"

        if "10-19" in shift_code:
            return "SHIFT_10_19", "Geç Gündüz (10:00 - 19:00)"

        if "20-08" in shift_code:
            return "SHIFT_20_08", "12-Saat Gece (20:00 - 08:00)"

        # 2. Dinamik Zaman Penceresi Kontrolü (Punches Window Matching)
        # Turnike yazılımı hatalı olarak '08-17' basmış olsa dahi personelin gerçek saatine göre ayırt edilir:
        if pd.notna(dt_in) and pd.notna(dt_out):
            t_in = dt_in.time()
            t_out = dt_out.time()
            gross_hours = (dt_out - dt_in).total_seconds() / 3600.0

            # Akşam Vardiyası (15:00 - 16:30 arası giriş, gece yarısına doğru çıkış)
            if 15 <= t_in.hour <= 17 and (is_overnight or t_out.hour == 0 or t_out.hour >= 23):
                return "SHIFT_16_24", "3-Vardiya Akşam (16:00 - 24:00)"

            # Gece Vardiyası (22:00 - 01:00 arası giriş, sabah 06:00 - 08:30 arası çıkış)
            if (t_in.hour >= 21 or t_in.hour <= 1) and (6 <= t_out.hour <= 9):
                return "SHIFT_00_08", "3-Vardiya Gece (00:00 - 08:00)"

            # Sabah 08:00 Girişi için İKİ KRİTİK DURUM:
            # - Eğer vardiya kodu açıkça '08-16' ise zaten yukarıda yakalandı.
            # - Eğer '08-17' kodlu ise ve ~16:00'da çıkmışsa (8 saat kalmışsa):
            #   Bu işçi erken çıkan bir Gündüz işçisidir!
            if 7 <= t_in.hour <= 9:
                if "08-17" in shift_code or "SERBEST" in shift_code:
                    return "DAY_08_17", "Standart Gündüz (08:00 - 17:00)"
                # Kod yoksa ama tam 16:00 civarı (7.7h - 8.3h) çıkmışsa ve akşam vardiyası ile hat değişimi varsa
                if 15 <= t_out.hour <= 16 and 7.5 <= gross_hours <= 8.3:
                    return "SHIFT_08_16", "3-Vardiya Sabah (08:00 - 16:00)"
                return "DAY_08_17", "Standart Gündüz (08:00 - 17:00)"

        # Varsayılan (Default)
        if "08-17" in shift_code:
            return "DAY_08_17", "Standart Gündüz (08:00 - 17:00)"

        return "DAY_08_17", "Standart Gündüz (08:00 - 17:00)"


# =================================================================================================
# 6. MESAİ VE MOLA HESAPLAYICI (HoursCalculator)
# =================================================================================================

class HoursCalculator:
    """
    Brüt süreyi, mola kesintilerini, net çalışma saatini ve yasal gece çalışma süresini
    deterministik olarak hesaplar.
    """

    def __init__(self, config: Optional[PDKSConfig] = None):
        self.config = config or PDKSConfig()

    def calculate_hours(self, df: pd.DataFrame) -> pd.DataFrame:
        """Tüm kayıtların brüt, mola, net ve gece saatlerini hesaplar."""
        gross_list = []
        break_list = []
        net_list = []
        night_list = []
        flags = []

        for idx, row in df.iterrows():
            dt_in = row["check_in_dt"]
            dt_out = row["check_out_dt"]
            shift_type = row.get("shift_type", "DAY_08_17")
            status = row.get("audit_status", "NORMAL")

            if pd.isna(dt_in) or pd.isna(dt_out):
                gross_list.append(0.0)
                break_list.append(0.0)
                net_list.append(0.0)
                night_list.append(0.0)
                flags.append(status)
                continue

            # 1. Brüt Süre Hesabı (Gross Hours)
            diff_hours = (dt_out - dt_in).total_seconds() / 3600.0
            # 24 saat tavan guardı
            gross_hours = min(diff_hours, self.config.max_daily_presence_hours)
            gross_hours = max(0.0, round(gross_hours, 2))

            # 2. Mola ve Net Süre Hesabı
            break_hours, net_hours = self._compute_break_and_net(gross_hours, shift_type)

            # 3. Yasal Gece Çalışma Süresi (20:00 - 06:00)
            night_hours = calculate_legal_night_hours(dt_in, dt_out)

            # 4. Denetim Bayrakları (Audit Flags)
            flag = self._evaluate_audit_flag(status, gross_hours, night_hours)

            gross_list.append(gross_hours)
            break_list.append(break_hours)
            net_list.append(net_hours)
            night_list.append(night_hours)
            flags.append(flag)

        df["gross_hours"] = gross_list
        df["break_hours"] = break_list
        df["net_hours"] = net_list
        df["night_hours"] = night_list
        df["audit_flag"] = flags

        return df

    def _compute_break_and_net(self, gross_hours: float, shift_type: str) -> Tuple[float, float]:
        """
        Vardiya modeline göre mola kesintisi ve net mesai süresini belirler.
        """
        if gross_hours <= 0.0:
            return 0.0, 0.0

        is_rotating = shift_type in ("SHIFT_08_16", "SHIFT_16_24", "SHIFT_00_08", "SHIFT_20_08")

        # ---------------------------------------------------------------------
        # MODEL 1: DÖNEN 3-VARDİYA SİSTEMİ (8 Saatlik Vardiyalar)
        # ---------------------------------------------------------------------
        if is_rotating:
            # 8.0 saatlik tam vardiya -> Mola 0.5 saat -> Net 7.5 saat
            if gross_hours >= 7.5:
                break_hours = 0.5
                net_hours = round(gross_hours - break_hours, 2)
            elif gross_hours > 4.0:
                break_hours = 0.5
                net_hours = round(gross_hours - break_hours, 2)
            else:  # <= 4.0 saat
                break_hours = 0.25  # 15 dk
                net_hours = round(max(0.0, gross_hours - break_hours), 2)

            return break_hours, net_hours

        # ---------------------------------------------------------------------
        # MODEL 2: STANDART GÜNDÜZ VARDİYASI (08:00 - 17:00 / 9 Saat Brüt)
        # ---------------------------------------------------------------------
        else:
            # Tam gün (9.0 saat) -> Mola 1.5 saat -> Net 7.5 saat
            # 1 saat erken çıkış (8.0 saat) -> Mola sabit 1.5 saat -> Net 6.5 saat!
            # Fazla mesai (10.0 saat) -> Mola 1.5 saat -> Net 8.5 saat
            if gross_hours >= 7.5:
                break_hours = 1.5
                net_hours = round(gross_hours - break_hours, 2)
            # Kısmi çalışma: ~5.0 saat bulunma -> 1.0 saat mola kesilir -> Net 4.0 saat
            elif gross_hours > 4.0:
                break_hours = 1.0
                net_hours = round(gross_hours - break_hours, 2)
            # Kısa çalışma: ~4.0 saat bulunma -> 0.5 saat mola kesilir -> Net 3.5 saat
            elif gross_hours > 0.5:
                break_hours = 0.5
                net_hours = round(gross_hours - break_hours, 2)
            else:
                break_hours = 0.0
                net_hours = gross_hours

            return break_hours, max(0.0, net_hours)

    def _evaluate_audit_flag(self, current_status: str, gross_hours: float, night_hours: float) -> str:
        """Denetim bayrağını öncelik sırasına göre belirler."""
        if current_status in ("AUTO_IMPUTED_ENTRY", "AUTO_IMPUTED_EXIT"):
            return current_status
        if current_status == "DURATION_MISMATCH":
            return "DURATION_MISMATCH"
        if night_hours > self.config.max_legal_night_hours:
            return "EXCESSIVE_NIGHT_WORK"
        if gross_hours >= self.config.excessive_presence_alert_hours:
            return "EXCESSIVE_DAILY_PRESENCE"
        return "NORMAL"


# =================================================================================================
# 7. RAPOR ÜRETİCİ (ReportGenerator)
# =================================================================================================

class ReportGenerator:
    """
    İşlenmiş puantaj verilerinden 'Daily_Summary' ve 'Audit_Alerts' sayfalarını
    içeren profesyonel, renkli ve biçimlendirilmiş bir Excel çalışma kitabı üretir.
    """

    FONT_FAMILY = "Segoe UI"

    # Renk Paleti (Modern Corporate Navy & Slate)
    COLOR_NAVY_DARK = "1E293B"     # Başlık Zemin
    COLOR_NAVY_HEADER = "0F172A"   # Sütun Başlıkları
    COLOR_TEXT_WHITE = "FFFFFF"
    COLOR_BORDER_GREY = "E2E8F0"
    COLOR_ZEBRA_ROW = "F8FAFC"
    COLOR_ALERT_WARN = "FEF3C7"    # Amber (İmputed)
    COLOR_ALERT_DANGER = "FEE2E2"  # Red (Mismatch / Excessive)
    COLOR_ALERT_SUCCESS = "ECFDF5" # Green (Normal)

    def __init__(self, config: Optional[PDKSConfig] = None):
        self.config = config or PDKSConfig()

    def generate_excel_report(self, df: pd.DataFrame, output_path: str) -> str:
        """Daily_Summary ve Audit_Alerts sayfalarını oluşturup dosyaya yazar."""
        logger.info(f"Excel çalışma kitabı oluşturuluyor: {output_path}")
        wb = openpyxl.Workbook()
        ws_default = wb.active

        # 1. SAYFA: Daily_Summary
        ws_summary = wb.create_sheet(title="Daily_Summary")
        self._build_daily_summary_sheet(ws_summary, df)

        # 2. SAYFA: Audit_Alerts
        ws_alerts = wb.create_sheet(title="Audit_Alerts")
        self._build_audit_alerts_sheet(ws_alerts, df)

        # Varsayılan boş sayfayı kaldır
        wb.remove(ws_default)

        # Dizini oluştur ve kaydet
        out_dir = os.path.dirname(output_path)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        wb.save(output_path)
        logger.info(f"Rapor başarıyla kaydedildi: {output_path}")
        return output_path

    def _build_daily_summary_sheet(self, ws: Any, df: pd.DataFrame) -> None:
        """Tüm günlük kayıtların listelendiği ana sayfa."""
        ws.views.sheetView[0].showGridLines = True

        # Başlık Blok
        ws.merge_cells("A1:K1")
        title_cell = ws["A1"]
        title_cell.value = "FİDE KONSERVE — GÜNLÜK PDKS VE VARDİYA MESAİ CETVELİ (DAILY SUMMARY)"
        title_cell.font = Font(name=self.FONT_FAMILY, size=13, bold=True, color=self.COLOR_TEXT_WHITE)
        title_cell.fill = PatternFill(start_color=self.COLOR_NAVY_DARK, end_color=self.COLOR_NAVY_DARK, fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # Sütun Başlıkları
        headers = [
            ("Employee ID", "Sicil No"),
            ("Name", "Adı Soyadı"),
            ("Date", "Tarih"),
            ("Check-In", "Giriş Saati"),
            ("Check-Out", "Çıkış Saati"),
            ("Shift Type", "Vardiya Türü"),
            ("Gross Hours", "Brüt Süre (s)"),
            ("Break Hours", "Mola Süresi (s)"),
            ("Net Hours", "Net Mesai (s)"),
            ("Night Hours", "Yasal Gece (s)"),
            ("Audit Flag", "Denetim Durumu")
        ]

        header_font = Font(name=self.FONT_FAMILY, size=9, bold=True, color=self.COLOR_TEXT_WHITE)
        header_fill = PatternFill(start_color=self.COLOR_NAVY_HEADER, end_color=self.COLOR_NAVY_HEADER, fill_type="solid")
        thin_border = Border(
            left=Side(style="thin", color=self.COLOR_BORDER_GREY),
            right=Side(style="thin", color=self.COLOR_BORDER_GREY),
            top=Side(style="thin", color=self.COLOR_BORDER_GREY),
            bottom=Side(style="thin", color=self.COLOR_BORDER_GREY)
        )

        ws.row_dimensions[2].height = 24
        for col_idx, (col_en, col_tr) in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_idx, value=f"{col_en}\n({col_tr})")
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        # Satır Verileri
        font_data = Font(name=self.FONT_FAMILY, size=9)
        font_mono = Font(name="Consolas", size=9)
        font_bold = Font(name=self.FONT_FAMILY, size=9, bold=True)

        row_idx = 3
        for _, r in df.iterrows():
            d_str = r["cin_date_obj"].strftime("%Y-%m-%d") if r.get("cin_date_obj") else (
                r["cout_date_obj"].strftime("%Y-%m-%d") if r.get("cout_date_obj") else ""
            )
            t_in_str = r["check_in_dt"].strftime("%H:%M:%S") if pd.notna(r.get("check_in_dt")) else "-"
            t_out_str = r["check_out_dt"].strftime("%H:%M:%S") if pd.notna(r.get("check_out_dt")) else "-"

            flag = str(r.get("audit_flag", "NORMAL"))
            row_fill = PatternFill(start_color=self.COLOR_ZEBRA_ROW, end_color=self.COLOR_ZEBRA_ROW, fill_type="solid") if row_idx % 2 == 0 else PatternFill(fill_type=None)

            # Audit durumuna göre satır vurgulama
            if flag.startswith("AUTO_IMPUTED"):
                flag_fill = PatternFill(start_color=self.COLOR_ALERT_WARN, end_color=self.COLOR_ALERT_WARN, fill_type="solid")
            elif flag in ("DURATION_MISMATCH", "EXCESSIVE_NIGHT_WORK", "EXCESSIVE_DAILY_PRESENCE"):
                flag_fill = PatternFill(start_color=self.COLOR_ALERT_DANGER, end_color=self.COLOR_ALERT_DANGER, fill_type="solid")
            else:
                flag_fill = row_fill

            vals = [
                (r.get("employee_id", ""), Alignment(horizontal="center"), font_mono, row_fill),
                (r.get("full_name", ""), Alignment(horizontal="left"), font_bold, row_fill),
                (d_str, Alignment(horizontal="center"), font_mono, row_fill),
                (t_in_str, Alignment(horizontal="center"), font_mono, row_fill),
                (t_out_str, Alignment(horizontal="center"), font_mono, row_fill),
                (r.get("shift_name", r.get("shift_type", "")), Alignment(horizontal="left"), font_data, row_fill),
                (r.get("gross_hours", 0.0), Alignment(horizontal="right"), font_mono, row_fill),
                (r.get("break_hours", 0.0), Alignment(horizontal="right"), font_mono, row_fill),
                (r.get("net_hours", 0.0), Alignment(horizontal="right"), font_bold, row_fill),
                (r.get("night_hours", 0.0), Alignment(horizontal="right"), font_mono, row_fill),
                (flag, Alignment(horizontal="center"), font_bold, flag_fill)
            ]

            ws.row_dimensions[row_idx].height = 19
            for c_idx, (val, align, font_style, fill_style) in enumerate(vals, 1):
                c = ws.cell(row=row_idx, column=c_idx, value=val)
                c.alignment = align
                c.font = font_style
                c.fill = fill_style
                c.border = thin_border
                if c_idx in (7, 8, 9, 10):
                    c.number_format = "0.00"

            row_idx += 1

        # Genel Toplam Satırı
        total_gross = df["gross_hours"].sum() if "gross_hours" in df.columns else 0.0
        total_break = df["break_hours"].sum() if "break_hours" in df.columns else 0.0
        total_net = df["net_hours"].sum() if "net_hours" in df.columns else 0.0
        total_night = df["night_hours"].sum() if "night_hours" in df.columns else 0.0

        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)
        tot_label = ws.cell(row=row_idx, column=1, value="GENEL TOPLAM:")
        tot_label.font = Font(name=self.FONT_FAMILY, size=9, bold=True)
        tot_label.alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row=row_idx, column=7, value=round(total_gross, 2)).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=8, value=round(total_break, 2)).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=9, value=round(total_net, 2)).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=10, value=round(total_night, 2)).number_format = "#,##0.00"

        tot_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        for c in range(1, 12):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = font_bold
            cell.fill = tot_fill
            cell.border = thin_border

        # Sütun Genişliklerini Otomatik Ayarla
        self._autofit_columns(ws, max_cols=11)
        ws.freeze_panes = "A3"

    def _build_audit_alerts_sheet(self, ws: Any, df: pd.DataFrame) -> None:
        """Yalnızca anomalili, otomatik tamamlanan veya şüpheli satırları listeleyen denetim sayfası."""
        ws.views.sheetView[0].showGridLines = True

        # Filtre: Sadece NORMAL olmayan kayıtlar
        alerts_df = df[df["audit_flag"] != "NORMAL"].copy()

        # Başlık Blok
        ws.merge_cells("A1:K1")
        title_cell = ws["A1"]
        title_cell.value = "FİDE KONSERVE — PDKS ANOMALİ VE DENETİM İNCELEME LİSTESİ (AUDIT ALERTS)"
        title_cell.font = Font(name=self.FONT_FAMILY, size=13, bold=True, color=self.COLOR_TEXT_WHITE)
        title_cell.fill = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")  # Crimson
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        headers = [
            ("Employee ID", "Sicil"),
            ("Name", "Adı Soyadı"),
            ("Date", "Tarih"),
            ("Check-In", "Giriş"),
            ("Check-Out", "Çıkış"),
            ("Shift Type", "Vardiya"),
            ("Gross Hours", "Brüt (s)"),
            ("Net Hours", "Net (s)"),
            ("Night Hours", "Gece (s)"),
            ("Alert Flag", "Uyarı Tipi"),
            ("Audit Reason & Detail", "İnceleme Gerekçesi / Açıklama")
        ]

        header_font = Font(name=self.FONT_FAMILY, size=9, bold=True, color=self.COLOR_TEXT_WHITE)
        header_fill = PatternFill(start_color="7F1D1D", end_color="7F1D1D", fill_type="solid")
        thin_border = Border(
            left=Side(style="thin", color=self.COLOR_BORDER_GREY),
            right=Side(style="thin", color=self.COLOR_BORDER_GREY),
            top=Side(style="thin", color=self.COLOR_BORDER_GREY),
            bottom=Side(style="thin", color=self.COLOR_BORDER_GREY)
        )

        ws.row_dimensions[2].height = 24
        for col_idx, (col_en, col_tr) in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_idx, value=f"{col_en}\n({col_tr})")
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        font_data = Font(name=self.FONT_FAMILY, size=9)
        font_mono = Font(name="Consolas", size=9)
        font_bold = Font(name=self.FONT_FAMILY, size=9, bold=True)

        row_idx = 3
        if len(alerts_df) == 0:
            ws.merge_cells("A3:K3")
            cell = ws["A3"]
            cell.value = "Tebrikler! İncelenen veri setinde hiçbir anomali veya eksik basım bulunamadı."
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name=self.FONT_FAMILY, size=11, bold=True, color="047857")
        else:
            for _, r in alerts_df.iterrows():
                d_str = r["cin_date_obj"].strftime("%Y-%m-%d") if r.get("cin_date_obj") else (
                    r["cout_date_obj"].strftime("%Y-%m-%d") if r.get("cout_date_obj") else ""
                )
                t_in_str = r["check_in_dt"].strftime("%H:%M:%S") if pd.notna(r.get("check_in_dt")) else "-"
                t_out_str = r["check_out_dt"].strftime("%H:%M:%S") if pd.notna(r.get("check_out_dt")) else "-"

                flag = str(r.get("audit_flag", "NORMAL"))
                reason = self._build_alert_explanation(r, flag)

                fill_alert = PatternFill(start_color=self.COLOR_ALERT_WARN, end_color=self.COLOR_ALERT_WARN, fill_type="solid") if "IMPUTED" in flag else PatternFill(start_color=self.COLOR_ALERT_DANGER, end_color=self.COLOR_ALERT_DANGER, fill_type="solid")

                vals = [
                    (r.get("employee_id", ""), Alignment(horizontal="center"), font_mono),
                    (r.get("full_name", ""), Alignment(horizontal="left"), font_bold),
                    (d_str, Alignment(horizontal="center"), font_mono),
                    (t_in_str, Alignment(horizontal="center"), font_mono),
                    (t_out_str, Alignment(horizontal="center"), font_mono),
                    (r.get("shift_name", r.get("shift_type", "")), Alignment(horizontal="left"), font_data),
                    (r.get("gross_hours", 0.0), Alignment(horizontal="right"), font_mono),
                    (r.get("net_hours", 0.0), Alignment(horizontal="right"), font_bold),
                    (r.get("night_hours", 0.0), Alignment(horizontal="right"), font_mono),
                    (flag, Alignment(horizontal="center"), font_bold),
                    (reason, Alignment(horizontal="left"), font_data)
                ]

                ws.row_dimensions[row_idx].height = 20
                for c_idx, (val, align, font_style) in enumerate(vals, 1):
                    c = ws.cell(row=row_idx, column=c_idx, value=val)
                    c.alignment = align
                    c.font = font_style
                    c.fill = fill_alert
                    c.border = thin_border
                    if c_idx in (7, 8, 9):
                        c.number_format = "0.00"

                row_idx += 1

        self._autofit_columns(ws, max_cols=11)
        ws.freeze_panes = "A3"

    @staticmethod
    def _build_alert_explanation(row: pd.Series, flag: str) -> str:
        """Denetim uyarısı için anlaşılır açıklama üretir."""
        if flag == "AUTO_IMPUTED_EXIT":
            return f"Çıkış basılmadı. Atanan vardiya bitiş saatine göre otomatik tamamlandı."
        elif flag == "AUTO_IMPUTED_ENTRY":
            return f"Giriş basılmadı. Atanan vardiya başlangıç saatine göre otomatik tamamlandı."
        elif flag == "DURATION_MISMATCH":
            delta = row.get("duration_delta_min", 0.0)
            return f"Turnike 'Süre' sütunu ile hesaplanan süre arasında {delta:+.1f} dakika fark tespit edildi."
        elif flag == "EXCESSIVE_NIGHT_WORK":
            n_h = row.get("night_hours", 0.0)
            return f"Yasal sınır (7.5s) aşıldı: Gece vardiyasında {n_h:.2f} saat çalışıldı (İş K. m.69)."
        elif flag == "EXCESSIVE_DAILY_PRESENCE":
            g_h = row.get("gross_hours", 0.0)
            return f"Yüksek günlük bulunma süresi: {g_h:.2f} saat."
        return "İnceleme önerilir."

    @staticmethod
    def _autofit_columns(ws: Any, max_cols: int = 11) -> None:
        """Sütun genişliklerini içerik uzunluğuna göre ayarlar."""
        for c in range(1, max_cols + 1):
            col_letter = get_column_letter(c)
            max_len = 0
            for r in range(2, min(ws.max_row + 1, 100)):  # İlk 100 satıra göre hızlı hesap
                val = ws.cell(row=r, column=c).value
                if val:
                    lines = str(val).split("\n")
                    max_len = max(max_len, max(len(line) for line in lines))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 11)


# =================================================================================================
# 8. ANA İŞ AKIŞI MOTORU (PDKSPipeline)
# =================================================================================================

class PDKSPipeline:
    """
    PDKS işleme motorunun uçtan uca yürütücüsü.
    DataLoader -> Cleaner -> ShiftClassifier -> HoursCalculator -> ReportGenerator
    adımlarını koordine eder.
    """

    def __init__(self, config: Optional[PDKSConfig] = None):
        self.config = config or PDKSConfig()
        self.loader = DataLoader(self.config)
        self.cleaner = Cleaner(self.config)
        self.classifier = ShiftClassifier(self.config)
        self.calculator = HoursCalculator(self.config)
        self.reporter = ReportGenerator(self.config)

    def process_dataframe(self, raw_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Ham DataFrame'i işler, (özet_df, alarmlar_df) tuple'ı döndürür."""
        logger.info("Adım 1: Veri temizleme ve anomali onarımı...")
        cleaned_df = self.cleaner.clean_records(raw_df)

        logger.info("Adım 2: Vardiya sınıflandırması...")
        classified_df = self.classifier.classify_shifts(cleaned_df)

        logger.info("Adım 3: Brüt, mola, net ve yasal gece saatleri hesabı...")
        calculated_df = self.calculator.calculate_hours(classified_df)

        alerts_df = calculated_df[calculated_df["audit_flag"] != "NORMAL"].copy()
        logger.info(f"İşlem tamamlandı. Toplam Kayıt: {len(calculated_df)}, Anomali/Alarm: {len(alerts_df)}")
        return calculated_df, alerts_df

    def run(self, input_path: str = "pdks.xls", output_path: str = "pdks_processed_report.xlsx") -> str:
        """Dosyadan okuyup Excel raporu üretir."""
        raw_df = self.loader.load_file(input_path)
        summary_df, _ = self.process_dataframe(raw_df)
        out_file = self.reporter.generate_excel_report(summary_df, output_path)
        return out_file


# =================================================================================================
# 9. KOMUT SATIRI ARAYÜZÜ (CLI Entry Point)
# =================================================================================================

def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="FİDE Konserve — Otomatik PDKS ve Mesai İşleme Hattı"
    )
    parser.add_argument(
        "-i", "--input", default="pdks.xls", help="Girdi PDKS Excel/CSV dosyası (varsayılan: pdks.xls)"
    )
    parser.add_argument(
        "-o", "--output", default="pdks_processed_report.xlsx", help="Çıktı Excel dosya yolu (varsayılan: pdks_processed_report.xlsx)"
    )
    args = parser.parse_args()

    pipeline = PDKSPipeline()
    try:
        out = pipeline.run(input_path=args.input, output_path=args.output)
        print(f"\n[BAŞARILI] İşlem tamamlandı. Rapor oluşturuldu: {out}")
    except Exception as e:
        logger.error(f"Hata oluştu: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
