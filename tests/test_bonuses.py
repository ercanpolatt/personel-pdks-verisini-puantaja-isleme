# -*- coding: utf-8 -*-
"""
Bölüm Primleri Kontrol Testleri:
Kullanıcı Kuralı: Üretim, Balık Dolum, Balık Kesim gibi ekstra primler olmayacak.
Herkes standart aynı şekilde puantaja yazılacak (Prim = 0.0 saat).
"""

import pytest
from config.settings import calculate_department_bonus

class TestStandardNoDepartmentBonuses:
    def test_balik_dolum_no_bonus(self):
        """Balık Dolum için ekstra prim verilmez (0.0 saat)."""
        bonus_h, note = calculate_department_bonus("BALIK DOLUM", 8.0)
        assert bonus_h == 0.0
        assert note == ""

    def test_balik_kesim_no_bonus(self):
        """Balık Kesim için ekstra prim verilmez (0.0 saat)."""
        bonus_h, note = calculate_department_bonus("BALIK KESİM", 7.5)
        assert bonus_h == 0.0
        assert note == ""

    def test_uretim_no_bonus(self):
        """Üretim ekibi için ekstra prim verilmez (0.0 saat)."""
        bonus_h, note = calculate_department_bonus("ÜRETİM", 12.0)
        assert bonus_h == 0.0
        assert note == ""

    def test_other_departments_no_bonus(self):
        """Kalite Kontrol ve Depo gibi bölümler için bonus verilmez."""
        bonus_h, note = calculate_department_bonus("KALİTE KONTROL", 12.0)
        assert bonus_h == 0.0
        assert note == ""
        bonus_h, note = calculate_department_bonus("DEPO", 8.0)
        assert bonus_h == 0.0
        assert note == ""
