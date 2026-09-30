# -*- coding: utf-8 -*-
"""
Uçtan Uca Entegrasyon ve Bordro Doğrulama Testi
(Gerçek puantaj.xls ve pdks.xls veri setleri üzerinden tam denetim)
"""

import os
import pytest
from pdks_engine import PDKSEngine

@pytest.fixture(scope="module")
def loaded_engine():
    if not os.path.exists("puantaj.xls") or not os.path.exists("pdks.xls"):
        pytest.skip("puantaj.xls veya pdks.xls bulunamadığı için entegrasyon testi atlandı.")
    engine = PDKSEngine(pdks_path="pdks.xls", puantaj_path="puantaj.xls", target_month=9, target_year=2026)
    engine.load_personnel()
    engine.load_and_process_pdks()
    engine.load_payroll_sheets()
    return engine

class TestEndToEndEngine:
    def test_personnel_count(self, loaded_engine):
        """Puantaj tablosundan 493 personel ve PDKS'den gelen 7 ek çalışanla toplam 500 personel olmalıdır."""
        assert len(loaded_engine.personnel_list) == 500

    def test_audit_records_count(self, loaded_engine):
        """PDKS turnikesinden 3318 günlük çalışma kaydı işlenmelidir."""
        assert len(loaded_engine.audit_records) == 3318

    def test_active_personnel_count(self, loaded_engine):
        """Kart basan aktif çalışan sayısı 254 olmalıdır."""
        matrix = loaded_engine.get_summary_matrix()
        active = [m for m in matrix if m["total_work_days"] > 0]
        assert len(active) == 254

    def test_exception_records_count(self, loaded_engine):
        """Amir incelemesine sunulan eksik/çoklu basım sayısı 366 olmalıdır."""
        assert len(loaded_engine.exception_records) == 366

    def test_payroll_sheets_loaded(self, loaded_engine):
        """Ek bordro sayfaları eksiksiz okunmalıdır."""
        assert len(loaded_engine.rapor_list) == 19
        assert len(loaded_engine.icra_list) == 13
        assert len(loaded_engine.izin_list) == 19
        assert len(loaded_engine.bordro_list) == 444
        assert len(loaded_engine.daimi_list) == 71
