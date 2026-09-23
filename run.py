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
from pdks_engine import PDKSEngine, fmt_hours_tr

def run_tests():
    print("==================================================")
    print(" BİRİM VE ENTEGRASYON TESTLERİ ÇALIŞTIRILIYOR")
    print("==================================================")
    result = subprocess.run([sys.executable, "-m", "pytest", "-v"])
    sys.exit(result.returncode)

def main():
    parser = argparse.ArgumentParser(description="FİDE Konserve Puantaj ve PDKS İşleme Sistemi")
    parser.add_argument("--month", type=int, default=9, help="Hedef ay (1-12, varsayılan: 9)")
    parser.add_argument("--year", type=int, default=2026, help="Hedef yıl (varsayılan: 2026)")
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

    if args.mode == "web":
        import uvicorn
        print("===================================================================")
        print(f" FİDE KONSERVE WEB DASHBOARD BAŞLATILIYOR: http://{args.host}:{args.port}")
        print("===================================================================")
        uvicorn.run("app:app", host=args.host, port=args.port, reload=False)
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
        generate_simple_report(
            output_xlsx="pdks_eylul_gunluk_calisma_saatleri.xlsx",
            output_csv="pdks_eylul_gunluk_calisma_saatleri.csv",
            month=args.month,
            year=args.year
        )
        print("   -> Alfabetik çalışma raporları hazırlandı.")

    print("\n===================================================================")
    print(" İŞLEM BAŞARIYLA TAMAMLANDI! TÜM RAPORLAR GÜNCEL VE DENETLENDİ.")
    print("===================================================================")

if __name__ == "__main__":
    main()
