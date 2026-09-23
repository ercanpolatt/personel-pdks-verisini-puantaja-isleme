# -*- coding: utf-8 -*-
"""
Dinamik Takvim Motoru Birim Testleri
(Ayın gün sayıları, artık yıllar, Pazar günleri tespiti, Türkçe gün adları)
"""

import pytest
from pdks_engine import PDKSEngine

class TestDynamicCalendar:
    def test_september_2026(self):
        """Eylül 2026: 30 gün çeker, Pazar günleri {6, 13, 20, 27}'dir."""
        engine = PDKSEngine(target_month=9, target_year=2026)
        assert engine.days_in_month == 30
        assert engine.sundays == {6, 13, 20, 27}
        assert engine.is_sunday(6) is True
        assert engine.is_sunday(7) is False
        assert engine.get_day_name_for_day(1) == "Salı"
        assert engine.get_day_name_for_day(6) == "Pazar"

    def test_october_2026(self):
        """Ekim 2026: 31 gün çeker, 31. gün kaybolmamalıdır."""
        engine = PDKSEngine(target_month=10, target_year=2026)
        assert engine.days_in_month == 31
        assert engine.sundays == {4, 11, 18, 25}
        assert engine.is_sunday(4) is True
        assert engine.get_day_name_for_day(1) == "Perşembe"

    def test_february_leap_year_2024(self):
        """Şubat 2024 artık yıldır ve 29 gün çekmelidir."""
        engine = PDKSEngine(target_month=2, target_year=2024)
        assert engine.days_in_month == 29

    def test_february_regular_year_2026(self):
        """Şubat 2026 normal yıldır ve 28 gün çekmelidir."""
        engine = PDKSEngine(target_month=2, target_year=2026)
        assert engine.days_in_month == 28

    def test_short_day_names(self):
        """Kısa gün adları Excel başlıkları için test edilir."""
        engine = PDKSEngine(target_month=9, target_year=2026)
        assert engine.get_day_name_for_day(1, short=True) == "Sal"
        assert engine.get_day_name_for_day(6, short=True) == "PAZAR"
