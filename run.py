# -*- coding: utf-8 -*-
"""
===================================================================================================
FİDE KONSERVE - PUANTAJ VE PDKS SİSTEMİ ANA YÖNETİCİSİ (run.py)
===================================================================================================
Komut satırından parametrik çalıştırma, rapor üretimi ve birim test yürütme arayüzü.

Örnek Kullanımlar:
  python run.py                                (Tüm raporları Eylül 2026 için üretir)
  python run.py --month 10 --year 2026         (Ekim 2026 dönemi için çalıştırır)
  python run.py --mode test                    (60 birim testini çalıştırır)
  python run.py --mode puantaj                 (Sadece puantaj.xlsx üretir)
  python run.py --mode daily                   (Sadece günlük raporları üretir)
"""

import argparse
import sys
import subprocess

from builder import build_puantaj_workbook
from export_daily_hours import build_excel_report, export_csv_and_json
from export_simple_pdks import generate_simple_report
from pdks_engine import PDKSEngine, fmt_hours_tr, detect_dominant_period

def run_tests():
    print("==================================================")
    print(" BİRİM VE ENTEGRASYON TESTLERİ ÇALIŞTIRILIYOR")
    print("==================================================")
    result = subprocess.run([sys.executable, "-m", "pytest", "-v"])
    sys.exit(result.returncode)

def main():
    parser = argparse.ArgumentParser(description="FİDE Konserve Puantaj ve PDKS İşleme Sistemi")
    parser.add_argument("--month", type=int, default=None, help="Hedef ay (1-12, varsayılan: pdks.xls üzerinden otomatik algılama)")
    parser.add_argument("--year", type=int, default=None, help="Hedef yıl (varsayılan: pdks.xls üzerinden otomatik algılama)")
    parser.add_argument("--pdks", type=str, default="pdks.xls", help="PDKS hareketleri dosyası")
    parser.add_argument("--puantaj", type=str, default="puantaj.xls", help="Puantaj personel dosyası")
    parser.add_argument("--mode", type=str, default="all", choices=["all", "puantaj", "daily", "simple", "test", "web"],
                        help="Çalışma modu: all, puantaj, daily, simple, test, web (varsayılan: all)")
    parser.add_argument("--host", type=str, default="127.0.0.1", help="Web sunucu adresi (varsayılan: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8000, help="Web sunucu portu (varsayılan: 8000)")

    args = parser.parse_args()

    if args.mode == "test":
        run_tests()
        return

    # Dönem Otomatik Algılama
    if args.month is None or args.year is None:
        det_m, det_y = detect_dominant_period(args.pdks)
        if args.month is None:
            args.month = det_m
        if args.year is None:
            args.year = det_y

    if args.mode == "web":
        import uvicorn
        print("===================================================================")
        print(f" FİDE KONSERVE WEB DASHBOARD BAŞLATILIYOR: http://{args.host}:{args.port}")
        print("===================================================================")
        uvicorn.run("app:app", host=args.host, port=args.port, reload=True)
        return

    print("===================================================================")
    print(f" FİDE KONSERVE PUANTAJ SİSTEMİ BAŞLATILIYOR ({args.month:02d}.{args.year})")
    print("===================================================================")

    if args.mode in ("all", "puantaj"):
        print("\n>>> 1. 10 Sayfalı Canlı Formüllü Puantaj (puantaj.xlsx) Üretiliyor...")
        build_puantaj_workbook(
            pdks_path=args.pdks,
            puantaj_path=args.puantaj,
            target_month=args.month,
            target_year=args.year,
            output_filename="puantaj.xlsx"
        )

    if args.mode in ("all", "daily"):
        print("\n>>> 2. Günlük Çalışma ve Denetim İzi Raporları Üretiliyor...")
        engine = PDKSEngine(
            pdks_path=args.pdks,
            puantaj_path=args.puantaj,
            target_month=args.month,
            target_year=args.year
        )
        engine.load_personnel()
        engine.load_and_process_pdks()
        build_excel_report(engine, "gunluk_calisma_raporu.xlsx")
        export_csv_and_json(engine, "gunluk_calismalar.csv", "gunluk_calismalar.json")
        print("   -> 'gunluk_calisma_raporu.xlsx', CSV ve JSON hazırlandı.")

    if args.mode in ("all", "simple"):
        print("\n>>> 3. Alfabetik Günlük Çalışma Cetveli Üretiliyor...")
        tr_month_slugs = ["", "ocak", "subat", "mart", "nisan", "mayis", "haziran", "temmuz", "agustos", "eylul", "ekim", "kasim", "aralik"]
        slug = tr_month_slugs[args.month] if 1 <= args.month <= 12 else "aylik"
        simple_xlsx = f"pdks_{slug}_gunluk_calisma_saatleri.xlsx"
        simple_csv = f"pdks_{slug}_gunluk_calisma_saatleri.csv"
        generate_simple_report(
            output_xlsx=simple_xlsx,
            output_csv=simple_csv,
            month=args.month,
            year=args.year
        )
        print(f"   -> Alfabetik çalışma raporları hazırlandı ({simple_xlsx}).")

    print("\n===================================================================")
    print(" İŞLEM BAŞARIYLA TAMAMLANDI! TÜM RAPORLAR GÜNCEL VE DENETLENDİ.")
    print("===================================================================")

if __name__ == "__main__":
    main()
