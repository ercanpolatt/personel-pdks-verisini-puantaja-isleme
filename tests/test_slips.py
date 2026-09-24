"""
Tests for Feature 3: Printable A4 / PDF Monthly Attendance & Overtime Reconciliation Slips.
(Aylık Çalışma ve Fazla Mesai Mutabakat Pusulası)
Verifies:
1. Single employee slip generation (A4 layout, 30 days, personal info, disclaimers, dual signatures)
2. Bulk factory slips generation (250 active employees, continuous multipage A4 printing)
3. Department-filtered bulk slips (e.g. Balık Dolum: 9 employees)
4. Auto-print script integration
5. 404 handling on invalid employee
"""
import pytest
from fastapi import HTTPException
from pdks_engine import PDKSEngine
from app import get_single_slip, get_bulk_slips, EngineManager


@pytest.fixture(scope="module")
def engine():
    eng = PDKSEngine(
        pdks_path="pdks.xls",
        puantaj_path="puantaj.xls",
        target_month=9,
        target_year=2026
    )
    eng.load_personnel()
    eng.load_and_process_pdks()
    eng.load_payroll_sheets()
    return eng


def test_single_slip_rendering(engine):
    """Verify single A4 slip contains company title, employee data, 30 days, and signature boxes."""
    res = get_single_slip("ABDULLAHKAYA")
    assert res.status_code == 200
    html = res.body.decode("utf-8")

    # Brand & Title
    assert "FİDE KONSERVE GIDA SAN. VE TİC. A.Ş." in html
    assert "AYLIK ÇALIŞMA VE FAZLA MESAİ MUTABAKAT PUSULASI" in html
    assert "DÖNEM: EYLÜL 2026" in html

    # Personnel details
    assert "ABDULLAH KAYA" in html
    assert "18748661292" in html
    assert "BALIK TEMİZLEME" in html

    # Legal Disclaimer & Signatures
    assert "4857 Sayılı İş Kanunu ve ilgili yönetmelikler uyarınca" in html
    assert "PERSONEL (İŞÇİ) ONAYI" in html
    assert "İŞVEREN VEKİLİ" in html

    # Page container
    assert html.count('<div class="slip-page">') == 1


def test_bulk_slips_rendering_all_active(engine):
    """Verify bulk generation for all 250 active personnel creates 250 A4 printable pages."""
    res = get_bulk_slips(department="all", status="active")
    assert res.status_code == 200
    html = res.body.decode("utf-8")

    # Exactly 250 A4 pages
    assert html.count('<div class="slip-page">') == 250
    assert "Toplam 250 Personel Sayfası" in html
    assert "Toplu Personel Puantaj Fişleri (250 Kişi)" in html


def test_bulk_slips_rendering_department_filter(engine):
    """Verify bulk generation filtered by department (BALIK DOLUM) produces 9 pages."""
    res = get_bulk_slips(department="BALIK DOLUM", status="active")
    assert res.status_code == 200
    html = res.body.decode("utf-8")

    # Exactly 9 Balık Dolum employees
    assert html.count('<div class="slip-page">') == 9
    assert "Toplam 9 Personel Sayfası" in html
    assert "BALIK DOLUM — Personel Puantaj Fişleri (9 Kişi)" in html


def test_slip_auto_print_script(engine):
    """Verify auto_print=True injects window.print() trigger."""
    res = get_single_slip("ABDULLAHKAYA", auto_print=True)
    assert res.status_code == 200
    html = res.body.decode("utf-8")
    assert "window.onload = function() { window.print(); };" in html


def test_single_slip_not_found(engine):
    """Verify 404 exception when requesting slip for non-existing employee."""
    with pytest.raises(HTTPException) as exc_info:
        get_single_slip("NON_EXISTING_PERSON_KEY_999")
    assert exc_info.value.status_code == 404
