# -*- coding: utf-8 -*-
"""
Vardiya ve Fazla Mesai Hesaplama Motoru Birim Testleri
(08:20 toleransı, 1.5 saat mola, kademeli FM, 8-4, 16-24, 24-8 vardiyaları)
"""

import pytest
from pdks_engine import calculate_shift_hours, calc_factory_worked_hours, calculate_multiple_shifts

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
        """17:00 - 17:20 arası çıkış: Fazla mesai yok (7.5 saat)."""
        res = calculate_shift_hours("08:00", "17:20")
        assert res.total_hours == 7.5
        assert res.overtime_hours == 0.0

    def test_ot_step_1(self):
        """17:21 çıkışı: 21 dk FM -> +0.5 saat FM (Toplam 8.0 saat)."""
        res = calculate_shift_hours("08:00", "17:21")
        assert res.total_hours == 8.0
        assert res.overtime_hours == 0.5

    def test_ot_step_1_edge(self):
        """17:44 çıkışı: 44 dk FM -> +0.5 saat FM (Toplam 8.0 saat)."""
        res = calculate_shift_hours("08:00", "17:44")
        assert res.total_hours == 8.0
        assert res.overtime_hours == 0.5

    def test_ot_step_2(self):
        """17:45 çıkışı: 45 dk FM -> +1.0 saat FM (Toplam 8.5 saat)."""
        res = calculate_shift_hours("08:00", "17:45")
        assert res.total_hours == 8.5
        assert res.overtime_hours == 1.0

    def test_ot_step_2_edge(self):
        """18:20 çıkışı: 80 dk FM -> +1.0 saat FM (Toplam 8.5 saat)."""
        res = calculate_shift_hours("08:00", "18:20")
        assert res.total_hours == 8.5
        assert res.overtime_hours == 1.0

    def test_ot_step_3(self):
        """18:21 çıkışı: 81 dk FM -> +1.5 saat FM (Toplam 9.0 saat)."""
        res = calculate_shift_hours("08:00", "18:21")
        assert res.total_hours == 9.0
        assert res.overtime_hours == 1.5

    def test_ot_step_3_edge(self):
        """18:44 çıkışı: 104 dk FM -> +1.5 saat FM (Toplam 9.0 saat)."""
        res = calculate_shift_hours("08:00", "18:44")
        assert res.total_hours == 9.0
        assert res.overtime_hours == 1.5

    def test_ot_step_4(self):
        """18:45 çıkışı: 105 dk FM -> +2.0 saat FM (Toplam 9.5 saat)."""
        res = calculate_shift_hours("08:00", "18:45")
        assert res.total_hours == 9.5
        assert res.overtime_hours == 2.0

    def test_evening_shift_overtime(self):
        """16:00-24:00 Akşam vardiyasında 00:21 çıkışı 0.5 saat, 00:45 çıkışı 1.0 saat FM."""
        res_grace = calculate_shift_hours("16:00", "00:20")
        assert res_grace.overtime_hours == 0.0
        res_half = calculate_shift_hours("16:00", "00:21")
        assert res_half.overtime_hours == 0.5
        res_full = calculate_shift_hours("16:00", "00:45")
        assert res_full.overtime_hours == 1.0

    def test_night_shift_overtime(self):
        """20:00-04:00 Gece vardiyasında 04:21 çıkışı 0.5 saat, 04:45 çıkışı 1.0 saat FM, 08:00 (12 saat çalışma) çıkışı 3.0 saat FM ve 10.5 saat toplam."""
        res_grace = calculate_shift_hours("20:00", "04:20")
        assert res_grace.overtime_hours == 0.0
        res_half = calculate_shift_hours("20:00", "04:21")
        assert res_half.overtime_hours == 0.5
        res_full = calculate_shift_hours("20:00", "04:45")
        assert res_full.overtime_hours == 1.0
        res_12h = calculate_shift_hours("20:00", "08:00")
        assert res_12h.overtime_hours == 3.0
        assert res_12h.total_hours == 10.5
        assert res_12h.break_hours == 1.5


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


class TestMultipleShiftsAndShortWork:
    def test_split_shifts_user_case(self):
        """08:00-10:00 (2 saat) ve 14:00-15:00 (1 saat) parçalı çalışma -> Toplam tam 3.0 saat."""
        res = calculate_multiple_shifts([("08:00", "10:00"), ("14:00", "15:00")])
        assert res.total_hours == 3.0
        assert res.base_hours == 3.0
        assert res.overtime_hours == 0.0
        assert res.break_hours == 0.0

    def test_split_shifts_with_lunch_break_outside(self):
        """08:02-12:29 ve 13:04-19:07 (Ahmet Beyhan örneği): 9.5 saat toplam (7.5 temel + 2.0 FM)."""
        res = calculate_multiple_shifts([("08:02", "12:29"), ("13:04", "19:07")])
        assert res.total_hours == 9.5
        assert res.base_hours == 7.5
        assert res.overtime_hours == 2.0

    def test_fourteen_hour_gross_work_day(self):
        """08:00 - 22:00 (14 saat brüt çalışma): 1.5 saat mola kesilerek net 12.5 saat yazılır (7.5 temel + 5.0 FM)."""
        res = calculate_shift_hours("08:00", "22:00")
        assert res.total_hours == 12.5
        assert res.base_hours == 7.5
        assert res.overtime_hours == 5.0
        assert res.break_hours == 1.5

    def test_fourteen_hour_gross_work_night(self):
        """20:00 - 10:00 (14 saat brüt gece çalışma): 1.5 saat mola kesilerek net 12.5 saat yazılır (7.5 temel + 5.0 FM)."""
        res = calculate_shift_hours("20:00", "10:00")
        assert res.total_hours == 12.5
        assert res.base_hours == 7.5
        assert res.overtime_hours == 5.0
        assert res.break_hours == 1.5

    def test_ercument_aygul_september_10(self):
        """Ercüment Aygül 10.09.2026: 08:15 giriş - 15:49 çıkış (PDKS Süre 07:34) -> Net 6.5 saat (asla 19 saat olamaz!)."""
        res = calculate_shift_hours("08:15", "15:49")
        assert res.total_hours == 6.5
        assert res.base_hours == 6.5
        assert res.overtime_hours == 0.0
        assert res.gross_hours == 7.57
        assert res.break_hours == 1.5


class TestMidnightShiftEarlyArrivalAndDateAssignment:
    """24:00 - 08:00 gece vardiyasına erken gelenler (örn: 23:46) 00:00'a yuvarlanır ve ertesi güne (çıkış gününe) yazılır."""

    def test_midnight_shift_exact_00_to_08(self):
        """00:00 - 08:00 vardiyası tam 7.5 saat net hak ediş üretir."""
        res = calculate_shift_hours("00:00", "08:00")
        assert res.total_hours == 7.5
        assert res.base_hours == 7.5
        assert res.overtime_hours == 0.0
        assert res.break_hours == 0.5

    def test_midnight_shift_early_arrival_rounding(self):
        """23:46'da gelen kişi 00:00'a yuvarlanarak 00:00-08:00 çalışması olarak 7.5 saat alır."""
        res = calculate_shift_hours("00:00", "08:00")
        assert res.total_hours == 7.5
        assert res.effective_g == "00:00"
        assert res.effective_c == "08:00"

    def test_midnight_shift_with_overtime_1016(self):
        """24:00 vardiyası sabah 10:16'ya kadar mesaiye kalmışsa: 7.5 normal + 2.0 FM = 9.5 saat."""
        res = calculate_shift_hours("00:00", "10:16")
        assert res.total_hours == 9.5
        assert res.base_hours == 7.5
        assert res.overtime_hours == 2.0

    def test_midnight_shift_engine_assignment_to_exit_day(self):
        """PDKSEngine içinde 16.10 23:46 giriş - 17.10 08:00 çıkış kaydı 17.10 tarihine işlenmelidir."""
        from pdks_engine import PDKSEngine
        engine = PDKSEngine("pdks.xls", "puantaj.xls", target_month=10, target_year=2026)
        engine.load_personnel()
        engine.load_and_process_pdks()
        # Erdin Tetik 16.10 23:46 -> 17.10 08:00
        res_17 = engine.daily_results.get(("ERDINTETIK", 17))
        assert res_17 is not None, "17.10 tarihinde Erdin Tetik için kayıt bulunmalı"
        assert res_17["raw_g_saat"] == "00:00", "Giriş saati 00:00'a yuvarlanmış olmalı"
        assert res_17["final_hours"] == 7.5, "Hak ediş 7.5 saat olmalı"


class TestRapidPunchAndExitReentryGuard:
    """
    Kullanıcı Kuralı:
    1. 'mesaiyi son basımdan hesaplayacağız zaten' -> Mesai son basıma (c) kadar hesaplanır.
    2. 'giriş ve çıkış yapmış ve çıktıktan 10 dk içerisinde tekrar giriş yapmışsa önlemini alalım'
       - Çıkıştan sonra 10 dk içinde girip çıkış basmamışsa (örn: Mahsun Görgün 18.10: 17:28 -> 17:31):
         Çıkış 17:28'de dondurulur, sahte açık vardiya engellenir, kırmızı alarm verilir.
       - Çıkıştan sonra 10 dk içinde girip çalışmaya devam etmişse (örn: Kadir Koç 12.10: 20:04 -> 20:06 -> 07:53):
         Son basıma kadar hak ediş hesaplanır (21.0s), kırmızı alarm verilir.
    """

    def test_mahsun_gorgun_unclosed_reentry(self):
        """18.10 Mahsun Görgün: 17:28 çıkışından 3 dk sonra 17:31 giriş basılmış ve çıkış basılmamış."""
        from pdks_engine import PDKSEngine
        engine = PDKSEngine("pdks.xls", "puantaj.xls", target_month=10, target_year=2026)
        engine.load_personnel()
        engine.load_and_process_pdks()

        rec = engine.daily_results.get(("MAHSUNGIRGIN", 18))
        assert rec is not None, "Mahsun Görgün 18.10 kaydı bulunmalı"
        assert rec["is_rapid_punch"] is True, "Hızlı basım (10 dk içi) tespit edilmeli"
        assert rec["status_type"] == "ŞÜPHELİ_HIZLI_BASIM", "Durum türü ŞÜPHELİ_HIZLI_BASIM olmalı"
        assert rec["is_exception"] is True, "Denetimci onay cetveline istisna olarak düşmeli"
        assert rec["raw_c_saat"] == "17:28", "Çıkış 17:28'de dondurulmuş olmalı (17:31 açık bırakılmamalı)"
        assert rec["final_hours"] == 8.0, "Çalışma saati 07:42 - 17:28 arası 8.0 saat olmalı"
        assert "17:28 çıkışından 3 dk sonra tekrar giriş yapıldı (17:31) ve çıkış basılmadı!" in rec["audit_note"]
        assert "Kırmızı Alarm" in rec["audit_note"]

    def test_kadir_koc_continuous_double_shift(self):
        """12.10 Kadir Koç: 20:04 çıkışından 2 dk sonra 20:06 tekrar giriş ve 07:53 son çıkış."""
        from pdks_engine import PDKSEngine
        engine = PDKSEngine("pdks.xls", "puantaj.xls", target_month=10, target_year=2026)
        engine.load_personnel()
        engine.load_and_process_pdks()

        rec = engine.daily_results.get(("KADIRKOC", 12))
        assert rec is not None, "Kadir Koç 12.10 kaydı bulunmalı"
        assert rec["is_rapid_punch"] is True, "20:04 -> 20:06 hızlı basım olarak algılanmalı"
        assert rec["status_type"] == "ŞÜPHELİ_HIZLI_BASIM"
        assert rec["is_exception"] is True
        # Mesai son basıma (07:53) kadar hesaplanmış olmalı
        assert rec["final_hours"] == 21.0, "Son basıma kadar hak ediş (21.0s) korunmalı"
        assert "20:04 çıkışından 2 dk sonra tekrar giriş yapıldı (20:06)!" in rec["audit_note"]
        assert "Kırmızı Alarm" in rec["audit_note"]

    def test_turnstile_double_swipe_preserves_last_punch(self):
        """Çıkışta 18:12 ve 18:21 peş peşe basıldığında mesai son basıma (18:21) göre hesaplanmalıdır."""
        from pdks_engine import extract_valid_intervals, calculate_shift_hours
        rows = [
            {"g_saat": "08:00:00", "c_saat": "18:12:00", "sure": "10:12"},
            {"g_saat": "", "c_saat": "18:21:00", "sure": ""}
        ]
        intervals = extract_valid_intervals(rows)
        assert intervals == [("08:00:00", "18:21:00")], "Son çıkış basımı (18:21) korunmalı"
        res = calculate_shift_hours(intervals[0][0], intervals[0][1])
        assert res.effective_c == "18:21"
        assert res.overtime_hours == 1.5, "18:21 son çıkışına göre 1.5 saat mesai verilmeli"
        assert res.total_hours == 9.0




