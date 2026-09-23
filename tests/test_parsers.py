# -*- coding: utf-8 -*-
"""
Veri Ayrıştırıcı ve Temizleme Fonksiyonları Birim Testleri
(clean_tc, clean_money, norm_name_key, clean_display_text, parse_hm)
"""

import pytest
from pdks_engine import (
    clean_tc,
    clean_money,
    norm_name_key,
    clean_display_text,
    parse_hm,
    parse_d_hm,
    fmt_hm,
    fmt_hours_tr
)

class TestCleanTC:
    def test_clean_tc_standard(self):
        assert clean_tc("12345678901") == "12345678901"

    def test_clean_tc_from_float(self):
        assert clean_tc(12345678901.0) == "12345678901"

    def test_clean_tc_ten_digits_adds_leading_zero(self):
        """10 haneli TC gelirse başına 0 eklenmelidir."""
        assert clean_tc("1234567890") == "01234567890"

    def test_clean_tc_with_spaces_or_chars(self):
        assert clean_tc(" 123 456 789 01 ") == "12345678901"

    def test_clean_tc_empty(self):
        assert clean_tc("") == ""
        assert clean_tc(None) == ""


class TestCleanMoney:
    def test_clean_money_string_with_tl(self):
        assert clean_money("33.030,00 TL") == 33030.0
        assert clean_money("5.827,30 tl") == 5827.3

    def test_clean_money_numeric(self):
        assert clean_money(28075.5) == 28075.5
        assert clean_money(15000) == 15000.0

    def test_clean_money_empty(self):
        assert clean_money("") == 0.0
        assert clean_money(None) == 0.0


class TestNormNameKey:
    def test_norm_name_basic(self):
        assert norm_name_key("Ahmet Yılmaz") == "AHMETYILMAZ"

    def test_norm_name_removes_parentheses(self):
        assert norm_name_key("Mehmet Demir (Şoför)") == "MEHMETDEMIR"

    def test_norm_name_removes_dash_suffix(self):
        assert norm_name_key("Ayşe Yılmaz-MKP") == "AYSEYILMAZ"

    def test_norm_name_aliases(self):
        assert norm_name_key("Selahattin Söğüt") == "SELAHATTINSOGUT"
        assert norm_name_key("Selahattin Sıt") == "SELAHATTINSOGUT"
        assert norm_name_key("Büşra Yavrutürk") == "BUSRAYAVRUKURT"


class TestCleanDisplayText:
    def test_clean_display_text_broken_chars(self):
        """Bozuk cp1254 karakterlerini onarır."""
        assert "DAİMİ" in clean_display_text("DA\ufffdM\ufffd")
        assert "MEVSİMLİK" in clean_display_text("MEVS\ufffdML\ufffdK")
        assert "ÇETİN" in clean_display_text("ETN")

    def test_clean_display_text_spaces(self):
        assert clean_display_text("  Ali   Veli  ") == "Ali Veli"


class TestTimeParsers:
    def test_parse_hm(self):
        assert parse_hm("08:30") == (8, 30)
        assert parse_hm("17:45") == (17, 45)
        assert parse_hm("") is None
        assert parse_hm("invalid") is None

    def test_fmt_hm(self):
        assert fmt_hm(510) == "08:30"
        assert fmt_hm(1020) == "17:00"

    def test_fmt_hours_tr(self):
        assert fmt_hours_tr(7.5) == "7,5"
        assert fmt_hours_tr(10.0) == "10,0"
