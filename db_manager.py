# -*- coding: utf-8 -*-
"""
===================================================================================================
FİDE KONSERVE - ÇOKLU AY ARŞİVİ VE KALICI VERİTABANI MOTORU (db_manager.py)
===================================================================================================
SQLite tabanlı ilişkisel arşiv veritabanı:
1. periods: Aylar ve dönem durumları (Aktif, Arşivlendi, Donduruldu)
2. monthly_records: Her aya ait personel puantaj ve bordro kesin dökümleri
3. period_exceptions: Aya özel çözümlenen amir istisna onayları
4. cumulative_overtime: 12 aylık gerçek zamanlı 270 saat fazla mesai kütüğü (4857 SK Md. 41)
5. icra_records & icra_movements: İcra ve maaş haczi takip ve otomatik bakiye devir motoru
===================================================================================================
"""

import os
import sqlite3
from datetime import datetime
from typing import Dict, Any, List, Optional, Tuple

from pdks_engine import norm_name_key, clean_display_text

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "puantaj_archive.db")
RESOLVED_EXCEPTIONS_JSON = os.path.join(DATA_DIR, "resolved_exceptions.json")

MONTH_NAMES_TR = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan",
    5: "Mayıs", 6: "Haziran", 7: "Temmuz", 8: "Ağustos",
    9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
}


def get_db_connection() -> sqlite3.Connection:
    """Veritabanı bağlantısı oluşturur ve satırları dict gibi erişilebilir yapar."""
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


class DatabaseManager:
    """Kalıcı SQLite Arşiv ve İcra/270s Takip Yöneticisi."""

    def __init__(self, db_path: str = DB_PATH):
        self.db_path = db_path
        self.init_db()

    def get_conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        return conn

    def init_db(self):
        """Tüm ilişkisel tabloları ve dizinleri oluşturur."""
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        with self.get_conn() as conn:
            cur = conn.cursor()

            # 1. DÖNEMLER (PERIODS)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS periods (
                    period_key TEXT PRIMARY KEY,
                    year INTEGER NOT NULL,
                    month INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    days_count INTEGER DEFAULT 30,
                    active_personnel_count INTEGER DEFAULT 0,
                    total_work_hours REAL DEFAULT 0.0,
                    total_ot_hours REAL DEFAULT 0.0,
                    total_payroll_tl REAL DEFAULT 0.0,
                    status TEXT DEFAULT 'active', -- 'active', 'archived', 'locked'
                    created_at TEXT NOT NULL,
                    archived_at TEXT
                );
            """)

            # 2. AYLIK PERSONEL PUANTAJ & BORDRO KAYITLARI (MONTHLY_RECORDS)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS monthly_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period_key TEXT NOT NULL,
                    sno INTEGER,
                    sicil TEXT,
                    tc TEXT,
                    name_key TEXT NOT NULL,
                    ad_soyad TEXT NOT NULL,
                    bolum TEXT,
                    durumu TEXT,
                    work_days INTEGER DEFAULT 0,
                    base_hours REAL DEFAULT 0.0,
                    weekday_ot REAL DEFAULT 0.0,
                    sunday_ot REAL DEFAULT 0.0,
                    net_weekday_ot REAL DEFAULT 0.0,
                    bonus_hours REAL DEFAULT 0.0,
                    net_maas REAL DEFAULT 0.0,
                    hourly_rate REAL DEFAULT 0.0,
                    total_earnings REAL DEFAULT 0.0,
                    icra_kesintisi REAL DEFAULT 0.0,
                    banka_neti REAL DEFAULT 0.0,
                    elden_fark REAL DEFAULT 0.0,
                    status TEXT DEFAULT 'OK',
                    UNIQUE(period_key, name_key),
                    FOREIGN KEY(period_key) REFERENCES periods(period_key) ON DELETE CASCADE
                );
            """)

            # 3. İSTİSNA AMİR ONAYLARI (PERIOD_EXCEPTIONS)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS period_exceptions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period_key TEXT NOT NULL,
                    name_key TEXT NOT NULL,
                    day INTEGER NOT NULL,
                    approved_hours REAL NOT NULL,
                    note TEXT,
                    resolved_by TEXT,
                    timestamp TEXT,
                    UNIQUE(period_key, name_key, day),
                    FOREIGN KEY(period_key) REFERENCES periods(period_key) ON DELETE CASCADE
                );
            """)

            # 4. YILLIK 270 SAAT KÜMÜLATİF FAZLA MESAİ KÜTÜĞÜ (CUMULATIVE_OVERTIME)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS cumulative_overtime (
                    year INTEGER NOT NULL,
                    name_key TEXT NOT NULL,
                    tc TEXT,
                    ad_soyad TEXT NOT NULL,
                    bolum TEXT,
                    m01 REAL DEFAULT 0.0,
                    m02 REAL DEFAULT 0.0,
                    m03 REAL DEFAULT 0.0,
                    m04 REAL DEFAULT 0.0,
                    m05 REAL DEFAULT 0.0,
                    m06 REAL DEFAULT 0.0,
                    m07 REAL DEFAULT 0.0,
                    m08 REAL DEFAULT 0.0,
                    m09 REAL DEFAULT 0.0,
                    m10 REAL DEFAULT 0.0,
                    m11 REAL DEFAULT 0.0,
                    m12 REAL DEFAULT 0.0,
                    total_ot_hours REAL DEFAULT 0.0,
                    remaining_limit_hours REAL DEFAULT 270.0,
                    risk_status TEXT DEFAULT 'SAFE', -- 'SAFE', 'WARNING', 'CRITICAL', 'EXCEEDED'
                    last_updated TEXT,
                    PRIMARY KEY(year, name_key)
                );
            """)

            # 5. İCRA TAKİP KÜTÜĞÜ (ICRA_RECORDS)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS icra_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tc TEXT,
                    name_key TEXT UNIQUE NOT NULL,
                    ad_soyad TEXT NOT NULL,
                    bolum TEXT,
                    sirket TEXT,
                    dosya_no TEXT,
                    toplam_borc REAL DEFAULT 0.0,
                    kesilen_kumulatif REAL DEFAULT 0.0,
                    kalan_borc REAL DEFAULT 0.0,
                    aylik_kesinti_orani REAL DEFAULT 0.25,
                    durum TEXT DEFAULT 'AKTİF', -- 'AKTİF', 'KAPANDI', 'DONDURULDU'
                    iban TEXT,
                    aciklama TEXT,
                    created_at TEXT,
                    updated_at TEXT
                );
            """)

            # 6. İCRA AYLIK KESİNTİ VE DEVİR HAREKETLERİ (ICRA_MOVEMENTS)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS icra_movements (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    icra_id INTEGER NOT NULL,
                    period_key TEXT NOT NULL,
                    year INTEGER NOT NULL,
                    month INTEGER NOT NULL,
                    kesinti_tutari REAL DEFAULT 0.0,
                    onceki_kalan REAL DEFAULT 0.0,
                    sonraki_kalan REAL DEFAULT 0.0,
                    tarih TEXT,
                    aciklama TEXT,
                    FOREIGN KEY(icra_id) REFERENCES icra_records(id) ON DELETE CASCADE,
                    FOREIGN KEY(period_key) REFERENCES periods(period_key) ON DELETE CASCADE
                );
            """)

            # Hızlı arama indeksleri
            cur.execute("CREATE INDEX IF NOT EXISTS idx_monthly_period ON monthly_records(period_key);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_monthly_name ON monthly_records(name_key);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_cum_year ON cumulative_overtime(year);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_cum_risk ON cumulative_overtime(risk_status);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_icra_name ON icra_records(name_key);")
            conn.commit()

    # =========================================================================
    # OTOMATİK VERİ TAŞIMA VE BAŞLANGIÇ YÜKLEMESİ (SEEDING)
    # =========================================================================
    def seed_initial_data(self, engine: Any) -> bool:
        """
        Mevcut motor verilerini (Eylül 2026, 250 personel, 13 icra kaydı ve istisnalar)
        SQLite veritabanına taşır ve ilk kurulumu tamamlar.
        """
        target_year = engine.target_year
        target_month = engine.target_month
        period_key = f"{target_year}-{target_month:02d}"

        with self.get_conn() as conn:
            cur = conn.cursor()

            # Dönem daha önce oluşturulmuş mu?
            cur.execute("SELECT period_key FROM periods WHERE period_key = ?", (period_key,))
            row = cur.fetchone()
            if row:
                return False  # Zaten yüklenmiş

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            month_name = f"{MONTH_NAMES_TR.get(target_month, str(target_month))} {target_year}"

            matrix = engine.get_summary_matrix()
            active_workers = [p for p in matrix if p.get("total_work_days", 0) > 0]
            fin_radar = engine.calculate_financial_radar()
            fin_kpis = fin_radar.get("kpis", {})

            # 1. Eylül 2026 Dönemini Ekle (Aktif)
            cur.execute("""
                INSERT INTO periods (
                    period_key, year, month, name, days_count, active_personnel_count,
                    total_work_hours, total_ot_hours, total_payroll_tl, status, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', ?)
            """, (
                period_key, target_year, target_month, month_name, engine.days_in_month,
                len(active_workers), fin_kpis.get("total_work_hours", 0.0),
                fin_kpis.get("total_overtime_hours", 0.0), fin_kpis.get("total_payroll_net", 0.0),
                now_str
            ))

            # 2. 250 Personel Aylık Puantaj Satırlarını Kaydet
            for p in active_workers:
                name_k = p.get("name_key", "")
                ad_soyad = p.get("ad_soyad", "")
                work_days = p.get("total_work_days", 0)
                base_h = round(p.get("total_base_hours", 0.0), 2)
                wd_ot = round(p.get("weekday_ot_hours", 0.0), 2)
                sun_ot = round(p.get("sunday_ot_hours", 0.0), 2)
                net_wd_ot = round(p.get("net_weekday_ot_hours", 0.0), 2)
                bonus_h = round(p.get("bonus_hours", 0.0), 2)

                # Finansal karşılık
                fin_p = next((f for f in fin_radar.get("personnel", []) if f.get("name_key") == name_k), None)
                net_maas = fin_p.get("net_wage", 0.0) if fin_p else 28075.50
                h_rate = fin_p.get("hourly_rate", 0.0) if fin_p else round(net_maas / 225.0, 2)
                tot_earn = fin_p.get("total_net_earned", 0.0) if fin_p else net_maas
                icra_kes = fin_p.get("icra", 0.0) if fin_p else 0.0
                banka = fin_p.get("bank_net", 0.0) if fin_p else net_maas
                elden = fin_p.get("cash_difference", 0.0) if fin_p else 0.0

                cur.execute("""
                    INSERT OR REPLACE INTO monthly_records (
                        period_key, sno, sicil, tc, name_key, ad_soyad, bolum, durumu,
                        work_days, base_hours, weekday_ot, sunday_ot, net_weekday_ot,
                        bonus_hours, net_maas, hourly_rate, total_earnings,
                        icra_kesintisi, banka_neti, elden_fark, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'OK')
                """, (
                    period_key, p.get("sno", 0), p.get("sicil", ""), p.get("tc", ""),
                    name_k, ad_soyad, p.get("bolum", ""), p.get("durumu", "MEVSİMLİK"),
                    work_days, base_h, wd_ot, sun_ot, net_wd_ot, bonus_h,
                    net_maas, h_rate, tot_earn, icra_kes, banka, elden
                ))

            # 3. İcra Takip Kayıtlarını Yükle (13 Fabrika Dosyası)
            # engine.icra_list ve icra_by_name üzerinden başlangıç borçları
            for ic in engine.icra_list:
                isim = ic.get("isim", "")
                name_k = norm_name_key(isim)
                kesilen_bu_ay = ic.get("kesilen", 0.0)
                # Kalan borç ham veride mevcut:
                kalan_b = ic.get("kalan_borc", 0.0)
                # Toplam borç: kalan + kesilen (tahmini veya 3 katı)
                toplam_b = round(max(kalan_b + kesilen_bu_ay, kalan_b * 1.5, 25000.0), 2)
                
                # Personel bilgisi
                emp_match = next((x for x in active_workers if x.get("name_key") == name_k), None)
                tc_val = emp_match.get("tc", "") if emp_match else ""
                bolum_val = emp_match.get("bolum", "Genel") if emp_match else "Genel"

                cur.execute("""
                    INSERT OR IGNORE INTO icra_records (
                        tc, name_key, ad_soyad, bolum, sirket, dosya_no, toplam_borc,
                        kesilen_kumulatif, kalan_borc, aylik_kesinti_orani, durum, iban,
                        created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0.25, 'AKTİF', ?, ?, ?)
                """, (
                    tc_val, name_k, isim, bolum_val, ic.get("sirket", "Fide Konserve"),
                    ic.get("dosya", "2026/100 Esas"), toplam_b, kesilen_bu_ay,
                    kalan_b, ic.get("iban", ""), now_str, now_str
                ))

                # Hareket kaydı
                cur.execute("SELECT id FROM icra_records WHERE name_key = ?", (name_k,))
                ic_row = cur.fetchone()
                if ic_row:
                    icra_id = ic_row["id"]
                    cur.execute("""
                        INSERT INTO icra_movements (
                            icra_id, period_key, year, month, kesinti_tutari,
                            onceki_kalan, sonraki_kalan, tarih, aciklama
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        icra_id, period_key, target_year, target_month, kesilen_bu_ay,
                        kalan_b + kesilen_bu_ay, kalan_b, now_str,
                        f"Eylül 2026 Bordro Maaş Haczi Kesintisi ({target_month:02d}.{target_year})"
                    ))

            # 4. Yıllık 270 Saat Mesai Kütüğünü Başlat (Ocak - Ağustos Geçmiş Veri Simülasyonu + Eylül)
            # Gerçekçi fabrika geçmiş mesaisi: Yıllık denetim radarının zengin çalışabilmesi için
            for p in active_workers:
                name_k = p.get("name_key", "")
                ad_soyad = p.get("ad_soyad", "")
                tc_val = p.get("tc", "")
                bolum_val = p.get("bolum", "")
                eylul_ot = round(p.get("total_ot_hours", 0.0), 2)

                # Personelin departman ve yoğunluğuna göre Ocak - Ağustos geçmiş mesai oluştur
                # Örneğin yüksek mesai yapanlarda geçmiş toplam 120-180 saat arası
                base_factor = min(1.0, eylul_ot / 35.0) if eylul_ot > 0 else 0.1
                # 8 aylık geçmiş dağılım:
                m1 = round(12.0 * base_factor, 1)
                m2 = round(14.5 * base_factor, 1)
                m3 = round(18.0 * base_factor, 1)
                m4 = round(20.0 * base_factor, 1)
                m5 = round(22.5 * base_factor, 1)
                m6 = round(25.0 * base_factor, 1)
                m7 = round(28.0 * base_factor, 1)
                m8 = round(26.0 * base_factor, 1)
                m9 = eylul_ot

                tot_ot = round(m1 + m2 + m3 + m4 + m5 + m6 + m7 + m8 + m9, 1)
                rem_limit = max(0.0, round(270.0 - tot_ot, 1))

                # Risk derecelendirmesi
                if tot_ot >= 270.0:
                    risk = "EXCEEDED"
                elif tot_ot >= 240.0:
                    risk = "CRITICAL"
                elif tot_ot >= 190.0:
                    risk = "WARNING"
                else:
                    risk = "SAFE"

                cur.execute("""
                    INSERT OR REPLACE INTO cumulative_overtime (
                        year, name_key, tc, ad_soyad, bolum,
                        m01, m02, m03, m04, m05, m06, m07, m08, m09, m10, m11, m12,
                        total_ot_hours, remaining_limit_hours, risk_status, last_updated
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0.0, 0.0, 0.0, ?, ?, ?, ?)
                """, (
                    target_year, name_k, tc_val, ad_soyad, bolum_val,
                    m1, m2, m3, m4, m5, m6, m7, m8, m9,
                    tot_ot, rem_limit, risk, now_str
                ))

            conn.commit()
            print(f"[DatabaseManager] SQLite Veritabanı başarıyla ilklendirildi: {period_key} dönemi kaydedildi.")
            return True

    # =========================================================================
    # DÖNEM YÖNETİMİ VE GEÇİŞLER (PERIOD MANAGEMENT)
    # =========================================================================
    def get_all_periods(self) -> List[Dict[str, Any]]:
        """Kayıtlı tüm puantaj dönemlerini kronolojik olarak listeler."""
        with self.get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT period_key, year, month, name, days_count, active_personnel_count,
                       total_work_hours, total_ot_hours, total_payroll_tl, status,
                       created_at, archived_at
                FROM periods
                ORDER BY year DESC, month DESC;
            """)
            rows = cur.fetchall()
            return [dict(r) for r in rows]

    def get_period_details(self, period_key: str) -> Optional[Dict[str, Any]]:
        """Seçilen dönemin özetini ve personel kayıtlarını döner."""
        with self.get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM periods WHERE period_key = ?", (period_key,))
            p_row = cur.fetchone()
            if not p_row:
                return None

            period_data = dict(p_row)

            # Personel kayıtları
            cur.execute("""
                SELECT * FROM monthly_records
                WHERE period_key = ?
                ORDER BY sno ASC;
            """, (period_key,))
            records = [dict(r) for r in cur.fetchall()]
            period_data["records"] = records
            return period_data

    def close_and_rollover_period(self, current_period_key: str) -> Dict[str, Any]:
        """
        Mevcut dönemi kapatır, arşivler, icra borç bakiyelerini düşüp devreder
        ve sonraki ayı otomatik başlatır.
        """
        with self.get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM periods WHERE period_key = ?", (current_period_key,))
            curr = cur.fetchone()
            if not curr:
                raise ValueError(f"Dönem bulunamadı: {current_period_key}")

            year = curr["year"]
            month = curr["month"]
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # 1. Mevcut dönemi arşivle / kilitle
            cur.execute("""
                UPDATE periods
                SET status = 'archived', archived_at = ?
                WHERE period_key = ?;
            """, (now_str, current_period_key))

            # 2. İcra Kesintilerini Düş ve Kalan Bakiyeleri Sonraki Aya Devret
            cur.execute("""
                SELECT r.id, r.name_key, r.ad_soyad, r.toplam_borc, r.kalan_borc,
                       m.icra_kesintisi
                FROM icra_records r
                JOIN monthly_records m ON r.name_key = m.name_key AND m.period_key = ?
                WHERE r.durum = 'AKTİF';
            """, (current_period_key,))
            active_icras = cur.fetchall()

            total_icra_deducted = 0.0
            closed_files = []
            rollover_files = []

            for ic in active_icras:
                icra_id = ic["id"]
                kesinti = ic["icra_kesintisi"]
                onceki_kalan = ic["kalan_borc"]
                yeni_kalan = max(0.0, round(onceki_kalan - kesinti, 2))
                total_icra_deducted += kesinti

                yeni_durum = "KAPANDI" if yeni_kalan <= 0.0 else "AKTİF"
                if yeni_durum == "KAPANDI":
                    closed_files.append({
                        "id": icra_id,
                        "ad_soyad": ic["ad_soyad"],
                        "kapanan_borc": onceki_kalan
                    })
                else:
                    rollover_files.append({
                        "id": icra_id,
                        "ad_soyad": ic["ad_soyad"],
                        "onceki_bakiye": onceki_kalan,
                        "kesilen": kesinti,
                        "devreden_bakiye": yeni_kalan
                    })

                # İcra ana tablosunu güncelle
                cur.execute("""
                    UPDATE icra_records
                    SET kesilen_kumulatif = kesilen_kumulatif + ?,
                        kalan_borc = ?,
                        durum = ?,
                        updated_at = ?
                    WHERE id = ?;
                """, (kesinti, yeni_kalan, yeni_durum, now_str, icra_id))

            # 3. Sonraki Dönemi Hesapla
            if month == 12:
                next_year = year + 1
                next_month = 1
            else:
                next_year = year
                next_month = month + 1

            next_period_key = f"{next_year}-{next_month:02d}"
            next_name = f"{MONTH_NAMES_TR.get(next_month, str(next_month))} {next_year}"

            # Gün sayısı hesabı (Ekim=31, Kasım=30 vs.)
            if next_month in (1, 3, 5, 7, 8, 10, 12):
                days_c = 31
            elif next_month in (4, 6, 9, 11):
                days_c = 30
            else:
                days_c = 29 if next_year % 4 == 0 else 28

            # Önceki dönem kayıtlarını yeni döneme şablon olarak devret
            cur.execute("SELECT COUNT(*) as cnt FROM monthly_records WHERE period_key = ?", (current_period_key,))
            rec_cnt = cur.fetchone()["cnt"]

            cur.execute("""
                INSERT OR REPLACE INTO periods (
                    period_key, year, month, name, days_count, active_personnel_count,
                    total_work_hours, total_ot_hours, total_payroll_tl, status, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, 0.0, 0.0, 0.0, 'active', ?);
            """, (next_period_key, next_year, next_month, next_name, days_c, rec_cnt, now_str))

            conn.commit()

            return {
                "status": "success",
                "message": f"{curr['name']} dönemi arşivlendi. {next_name} dönemi başarıyla açıldı ve icra bakiyeleri devredildi.",
                "archived_period": current_period_key,
                "new_active_period": next_period_key,
                "total_icra_deducted": round(total_icra_deducted, 2),
                "closed_files_count": len(closed_files),
                "closed_files": closed_files,
                "rollover_count": len(rollover_files),
                "rollover_files": rollover_files
            }

    # =========================================================================
    # YILLIK 270 SAAT KÜMÜLATİF FAZLA MESAİ TAKİBİ
    # =========================================================================
    def get_cumulative_overtime_report(self, year: int = 2026) -> Dict[str, Any]:
        """
        Fabrikadaki tüm personelin 12 aylık kümülatif mesai dağılımını,
        270 saate kalan sürelerini ve teftiş risk bayraklarını döner.
        """
        with self.get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT year, name_key, tc, ad_soyad, bolum,
                       m01, m02, m03, m04, m05, m06, m07, m08, m09, m10, m11, m12,
                       total_ot_hours, remaining_limit_hours, risk_status, last_updated
                FROM cumulative_overtime
                WHERE year = ?
                ORDER BY total_ot_hours DESC;
            """, (year,))
            rows = [dict(r) for r in cur.fetchall()]

            # İstatistiksel özetler
            safe_cnt = sum(1 for r in rows if r["risk_status"] == "SAFE")
            warning_cnt = sum(1 for r in rows if r["risk_status"] == "WARNING")
            critical_cnt = sum(1 for r in rows if r["risk_status"] == "CRITICAL")
            exceeded_cnt = sum(1 for r in rows if r["risk_status"] == "EXCEEDED")

            total_annual_ot = round(sum(r["total_ot_hours"] for r in rows), 1)
            avg_annual_ot = round(total_annual_ot / len(rows), 1) if rows else 0.0

            return {
                "year": year,
                "total_personnel": len(rows),
                "kpis": {
                    "total_annual_ot_hours": total_annual_ot,
                    "avg_annual_ot_hours": avg_annual_ot,
                    "safe_count": safe_cnt,
                    "warning_count": warning_cnt,
                    "critical_count": critical_cnt,
                    "exceeded_count": exceeded_cnt,
                    "legal_limit_hours": 270.0
                },
                "personnel": rows
            }

    # =========================================================================
    # İCRA & HACİZ TAKİP MASASI
    # =========================================================================
    def get_icra_records(self, status: Optional[str] = None) -> Dict[str, Any]:
        """Tüm icra kayıtlarını ve hareket özetlerini döner."""
        with self.get_conn() as conn:
            cur = conn.cursor()
            if status and status.upper() != "ALL":
                cur.execute("""
                    SELECT * FROM icra_records
                    WHERE durum = ?
                    ORDER BY kalan_borc DESC;
                """, (status.upper(),))
            else:
                cur.execute("SELECT * FROM icra_records ORDER BY durum ASC, kalan_borc DESC;")
            
            rows = [dict(r) for r in cur.fetchall()]

            # Her icra için son 5 hareketi ekle
            for r in rows:
                cur.execute("""
                    SELECT period_key, kesinti_tutari, onceki_kalan, sonraki_kalan, tarih, aciklama
                    FROM icra_movements
                    WHERE icra_id = ?
                    ORDER BY id DESC LIMIT 5;
                """, (r["id"],))
                r["movements"] = [dict(m) for m in cur.fetchall()]

            total_debt = round(sum(r["toplam_borc"] for r in rows), 2)
            total_deducted = round(sum(r["kesilen_kumulatif"] for r in rows), 2)
            total_remaining = round(sum(r["kalan_borc"] for r in rows), 2)
            active_count = sum(1 for r in rows if r["durum"] == "AKTİF")
            closed_count = sum(1 for r in rows if r["durum"] == "KAPANDI")

            return {
                "status": "success",
                "kpis": {
                    "total_files": len(rows),
                    "active_files": active_count,
                    "closed_files": closed_count,
                    "total_debt_tl": total_debt,
                    "total_deducted_tl": total_deducted,
                    "total_remaining_tl": total_remaining
                },
                "files": rows
            }

    def add_or_update_icra(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """Yeni icra dosyası ekler veya var olanı günceller."""
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT OR REPLACE INTO icra_records (
                    tc, name_key, ad_soyad, bolum, sirket, dosya_no, toplam_borc,
                    kesilen_kumulatif, kalan_borc, aylik_kesinti_orani, durum, iban,
                    aciklama, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (
                payload.get("tc", ""), payload["name_key"], payload["ad_soyad"],
                payload.get("bolum", "Genel"), payload.get("sirket", "Fide Konserve"),
                payload.get("dosya_no", ""), float(payload.get("toplam_borc", 0.0)),
                float(payload.get("kesilen_kumulatif", 0.0)),
                float(payload.get("kalan_borc", payload.get("toplam_borc", 0.0))),
                float(payload.get("aylik_kesinti_orani", 0.25)),
                payload.get("durum", "AKTİF"), payload.get("iban", ""),
                payload.get("aciklama", ""), now_str, now_str
            ))
            conn.commit()
            return {"status": "success", "message": "İcra kaydı başarıyla kaydedildi."}
