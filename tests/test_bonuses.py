# -*- coding: utf-8 -*-
"""
Özel Bölüm Primleri Birim Testleri
- Balık Dolum ve Kesim: Çalışılan her güne (saat fark etmeksizin) +2.0 saat prim (örn: 8s -> 10s, 9s -> 11s)
- Üretim Ekibi: 12.0 saat ve üzeri çalışmalara +4.0 saat prim
- Diğer bölümler: Bonus yok
"""

import pytest
from config.settings import calculate_department_bonus

class TestFishFillingCuttingBonuses:
    def test_balik_dolum_8_hours(self):
        """Kullanıcı Kuralı: Balık Dolum 8 saat çalışsa da +2 saat prim alır (8s -> 10s)."""
        bonus_h, note = calculate_department_bonus("BALIK DOLUM", 8.0)
        assert bonus_h == 2.0
        assert "Bonuslu 10,0s" in note

    def test_balik_dolum_9_hours(self):
        """Kullanıcı Kuralı: Balık Dolum 9 saat çalışınca 11 saat yazılır (9s -> 11s)."""
        bonus_h, note = calculate_department_bonus("BALIK DOLUM", 9.0)
        assert bonus_h == 2.0
        assert "Bonuslu 11,0s" in note

    def test_balik_kesim_7_5_hours(self):
        """Balık Kesim standart 7.5 saat çalışınca +2.0 saat prim alır (7.5s -> 9.5s)."""
        bonus_h, note = calculate_department_bonus("BALIK KESİM", 7.5)
        assert bonus_h == 2.0
        assert "Bonuslu 9,5s" in note

    def test_balik_kesim_overtime(self):
        """Balık Kesim 11.5 saat çalışınca +2.0 saat prim alır (11.5s -> 13.5s)."""
        bonus_h, note = calculate_department_bonus("BALIK KESIM ELEMANI", 11.5)
        assert bonus_h == 2.0
        assert "Bonuslu 13,5s" in note


class TestProductionBonuses:
    def test_uretim_under_threshold(self):
        """Üretim ekibi 11.5 saat çalışırsa eşiğin (12.0s) altında kaldığı için prim alamaz."""
        bonus_h, note = calculate_department_bonus("ÜRETİM", 11.5)
        assert bonus_h == 0.0
        assert note == ""

    def test_uretim_exact_threshold(self):
        """Üretim ekibi 12.0 saat tam eşikte çalışırsa +4.0 saat prim alır (12s -> 16s)."""
        bonus_h, note = calculate_department_bonus("KONSERVE URETİM ELEMANI", 12.0)
        assert bonus_h == 4.0
        assert "Bonuslu 16,0s" in note

    def test_uretim_over_threshold(self):
        """Üretim ekibi 13.5 saat çalışırsa +4.0 saat prim alır (13.5s -> 17.5s)."""
        bonus_h, note = calculate_department_bonus("URETIM", 13.5)
        assert bonus_h == 4.0
        assert "Bonuslu 17,5s" in note


class TestOtherDepartmentsNoBonus:
    def test_quality_control_no_bonus(self):
        """Kalite Kontrol 12 saat çalışsa dahi bölüm primi alamaz."""
        bonus_h, note = calculate_department_bonus("KALİTE KONTROL", 12.0)
        assert bonus_h == 0.0
        assert note == ""

    def test_warehouse_no_bonus(self):
        """Depo elemanı prim alamaz."""
        bonus_h, note = calculate_department_bonus("DEPO", 13.0)
        assert bonus_h == 0.0
        assert note == ""

    def test_zero_hours_no_bonus(self):
        """Çalışılmayan gün için bonus hesaplanmaz."""
        bonus_h, note = calculate_department_bonus("BALIK DOLUM", 0.0)
        assert bonus_h == 0.0
