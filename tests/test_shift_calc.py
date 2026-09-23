# -*- coding: utf-8 -*-
"""
Vardiya ve Fazla Mesai Hesaplama Motoru Birim Testleri
(08:20 toleransı, 1.5 saat mola, kademeli FM, 8-4, 16-24, 24-8 vardiyaları)
"""

import pytest
from pdks_engine import calculate_shift_hours, calc_factory_worked_hours

class TestDayShiftStandard:
    def test_standard_full_day(self):
        """08:00 - 17:00 tam gün çalışması: 7.5 saat net, 0 FM."""
        res = calculate_shift_hours("08:00", "17:00")
        assert res.total_hours == 7.5
        assert res.base_hours == 7.5
        assert res.overtime_hours == 0.0
        assert res.break_hours == 1.5

    def test_tolerance_inside(self):
        """08:15 gelişi 08:20 toleransı içinde olduğundan 08:00 sayılır (7.5 saat)."""
        res = calculate_shift_hours("08:15", "17:00")
        assert res.total_hours == 7.5
        assert res.effective_g == "08:00"

    def test_tolerance_exact_edge(self):
        """08:20 tam sınır: 7.5 saat tam verilir."""
        res = calculate_shift_hours("08:20", "17:00")
        assert res.total_hours == 7.5
        assert res.effective_g == "08:00"

    def test_tolerance_first_cut(self):
        """08:21 gelişi toleransı aştığı için 30 dk ceza ile 08:30 başlar (7.0 saat)."""
        res = calculate_shift_hours("08:21", "17:00")
        assert res.total_hours == 7.0
        assert res.effective_g == "08:30"

    def test_tolerance_second_cut(self):
        """08:51 gelişi 60 dk ceza ile 09:00 başlar (6.5 saat)."""
        res = calculate_shift_hours("08:51", "17:00")
        assert res.total_hours == 6.5
        assert res.effective_g == "09:00"


class TestEarlyDeparturesAndBreaks:
    def test_departure_before_lunch(self):
        """Öğle öncesi çıkış (08:00 - 11:30): Mola kesilmez, tam 3.5 saat fiili süre."""
        res = calculate_shift_hours("08:00", "11:30")
        assert res.total_hours == 3.5
        assert res.break_hours == 0.0

    def test_departure_exact_lunch_start(self):
        """12:00 çıkışı: Mola kesilmez, tam 4.0 saat."""
        res = calculate_shift_hours("08:00", "12:00")
        assert res.total_hours == 4.0
        assert res.break_hours == 0.0

    def test_departure_during_lunch(self):
        """Öğle esnasında çıkış (08:00 - 12:30): 12:00'ye kadar olan 4.0 saat verilir."""
        res = calculate_shift_hours("08:00", "12:30")
        assert res.total_hours == 4.0

    def test_departure_after_lunch(self):
        """Öğleden sonra çıkış (08:00 - 14:00): 6 saat brüt - 1.5 saat mola = 4.5 saat net."""
        res = calculate_shift_hours("08:00", "14:00")
        assert res.total_hours == 4.5
        assert res.break_hours == 1.5

    def test_departure_after_lunch_1500(self):
        """08:00 - 15:00 çıkışı: 7 saat brüt - 1.5 saat mola = 5.5 saat net."""
        res = calculate_shift_hours("08:00", "15:00")
        assert res.total_hours == 5.5
        assert res.break_hours == 1.5


class TestEightFourShift:
    def test_eight_four_exact(self):
        """8-4 vardiyası (08:00 - 16:00): 7.5 saat tam verilir."""
        res = calculate_shift_hours("08:00", "16:00")
        assert res.total_hours == 7.5

    def test_eight_four_tolerance_window(self):
        """8-4 vardiyası (08:00 - 15:55 aralığı): 7.5 saat tam verilir."""
        res = calculate_shift_hours("08:00", "15:55")
        assert res.total_hours == 7.5


class TestStepwiseOvertime:
    def test_ot_grace_period(self):
        """17:00 - 17:25 arası çıkış: Fazla mesai yok (7.5 saat)."""
        res = calculate_shift_hours("08:00", "17:25")
        assert res.total_hours == 7.5
        assert res.overtime_hours == 0.0

    def test_ot_step_1(self):
        """17:26 çıkışı: 26 dk FM -> +0.5 saat FM (Toplam 8.0 saat)."""
        res = calculate_shift_hours("08:00", "17:26")
        assert res.total_hours == 8.0
        assert res.overtime_hours == 0.5

    def test_ot_step_1_edge(self):
        """17:49 çıkışı: 49 dk FM -> +0.5 saat FM (Toplam 8.0 saat)."""
        res = calculate_shift_hours("08:00", "17:49")
        assert res.total_hours == 8.0
        assert res.overtime_hours == 0.5

    def test_ot_step_2(self):
        """17:50 çıkışı: 50 dk FM -> +1.0 saat FM (Toplam 8.5 saat)."""
        res = calculate_shift_hours("08:00", "17:50")
        assert res.total_hours == 8.5
        assert res.overtime_hours == 1.0

    def test_ot_step_2_edge(self):
        """18:25 çıkışı: 85 dk FM -> +1.0 saat FM (Toplam 8.5 saat)."""
        res = calculate_shift_hours("08:00", "18:25")
        assert res.total_hours == 8.5
        assert res.overtime_hours == 1.0

    def test_ot_step_3(self):
        """18:26 çıkışı: 86 dk FM -> +1.5 saat FM (Toplam 9.0 saat)."""
        res = calculate_shift_hours("08:00", "18:26")
        assert res.total_hours == 9.0
        assert res.overtime_hours == 1.5

    def test_ot_step_4(self):
        """18:50 çıkışı: 110 dk FM -> +2.0 saat FM (Toplam 9.5 saat)."""
        res = calculate_shift_hours("08:00", "18:50")
        assert res.total_hours == 9.5
        assert res.overtime_hours == 2.0


class TestOtherShiftsAndEdges:
    def test_evening_shift(self):
        """16:00 - 24:00 Akşam vardiyası: 7.5 saat net."""
        res = calculate_shift_hours("16:00", "24:00")
        assert res.total_hours == 7.5

    def test_night_shift(self):
        """24:00 - 08:00 Gece vardiyası: 7.5 saat net."""
        res = calculate_shift_hours("24:00", "08:00")
        assert res.total_hours == 7.5

    def test_instant_double_punch(self):
        """Turnikeye peş peşe 2 dakika arayla basılmışsa (çift basım): 7.5 saat verilir."""
        res = calculate_shift_hours("08:01", "08:03")
        assert res.total_hours == 7.5

    def test_backwards_compatibility_tuple(self):
        """calc_factory_worked_hours geriye dönük uyumlu (total, fm) döndürmelidir."""
        tot, fm = calc_factory_worked_hours("08:00", "18:26")
        assert tot == 9.0
        assert fm == 1.5
