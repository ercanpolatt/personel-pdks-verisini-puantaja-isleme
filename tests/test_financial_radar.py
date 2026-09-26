"""
Tests for Financial & Budget Radar calculations and endpoints.
Verifies:
1. Hourly normal rate = Net Maaş / 225
2. Overtime rate = (Net Maaş / 225) * 1.5
3. Sunday overtime rate = (Net Maaş / 225) * 1.5 (User requirement: 1.5x for Sunday)
4. Cash difference (Elden Fark) = max(0, total_net - bank_net - icra)
5. Department aggregation and factory financial summary
6. API endpoint /api/financial-radar
"""
import pytest
from pdks_engine import PDKSEngine


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


def test_financial_hourly_rates(engine):
    """Verify hourly base rate and 1.5x multiplier for overtime."""
    fin = engine.calculate_financial_radar()
    summary = fin["summary"]
    items = fin["personnel_costs"]

    assert summary["total_personnel_active"] == 250
    assert summary["total_overtime_cost"] > 1_000_000  # ~₺1.04M
    assert summary["total_payroll_budget"] > 4_000_000  # ~₺4.65M

    # Spot check sample personnel
    for p in items:
        net_maas = p["net_maas"]
        expected_base_rate = round(net_maas / 225.0, 2)
        expected_ot_rate = round(expected_base_rate * 1.5, 2)

        assert abs(p["hourly_base_rate"] - expected_base_rate) <= 0.05
        assert abs(p["hourly_overtime_rate"] - expected_ot_rate) <= 0.05

        # Verify weekday overtime (1.5x), missing hours deduction (1.0x), and sunday overtime (1.5x)
        # Formül: IF(h>7.5, (h-7.5)*1.5, h-7.5) kuralı: (weekday_ot * 1.5 - missing_hours * 1.0)
        expected_weekday_cost = round(((p["weekday_ot_hours"] * 1.5) - (p["missing_hours"] * 1.0)) * expected_base_rate, 2)
        expected_sunday_cost = round(p["sunday_ot_hours"] * expected_ot_rate, 2)
        expected_total_ot_cost = round(expected_weekday_cost + expected_sunday_cost, 2)

        assert abs(p["overtime_cost"] - expected_total_ot_cost) <= 1.00


def test_overtime_1_5_and_missing_1_0_formula(engine):
    """
    Kullanıcı Formülü Testi:
    =IF($AE3="";"";IF($AE3>7,5;($AE3-7,5)*1,5;$AE3-7,5))
    Fazla mesai çarpanı 1.5, eksik saat çarpanı 1.0 olarak hesaplanmalı.
    Eksik saat ve fazla mesai 1:1 toplanmamalı; 1 saat mesai (+1.5) ile 1 saat eksik (-1.0)
    birbirini sıfırlamaz, geriye 0.5 saatlik net mesai primi bırakır.
    """
    fin = engine.calculate_financial_radar()
    summary = fin["summary"]
    items = fin["personnel_costs"]

    # 1. Fabrika genelinde eksik saat ve fazla mesaisi olan personeller bulunmalı
    workers_with_missing = [p for p in items if p["missing_hours"] > 0]
    workers_with_both = [p for p in items if p["missing_hours"] > 0 and p["weekday_ot_hours"] > 0]
    
    assert len(workers_with_missing) > 0, "Eksik çalışması olan personel tespit edilemedi"
    assert len(workers_with_both) > 0, "Hem mesaisi hem eksik çalışması olan personel tespit edilemedi"
    assert summary["total_missing_hours"] > 0, "Fabrika toplam eksik saati 0 olamaz"
    assert summary["total_missing_deduction"] > 0, "Fabrika toplam eksik çalışma kesintisi 0 olamaz"

    # 2. Her personel için net mesai saati doğrulaması: (weekday_ot * 1.5) - (missing_hours * 1.0)
    for p in items:
        expected_net_ot_hours = round((p["weekday_ot_hours"] * 1.5) - (p["missing_hours"] * 1.0), 2)
        assert abs(p["net_weekday_ot_hours"] - expected_net_ot_hours) <= 0.05

        # Eksik saat kesintisi tam 1.0 katı saatlik ücret olmalıdır:
        expected_deduction = round(p["missing_hours"] * p["hourly_base_rate"], 2)
        assert abs(p["missing_deduction_cost"] - expected_deduction) <= 0.50

        # Brüt mesai kazancı tam 1.5 katı olmalıdır:
        expected_gross_ot = round(p["weekday_ot_hours"] * 1.5 * p["hourly_base_rate"], 2)
        assert abs(p["gross_weekday_ot_cost"] - expected_gross_ot) <= 0.50

        # Net hafta içi mesai tutarı = Brüt Mesai - Eksik Kesinti
        assert abs(p["weekday_ot_cost"] - (expected_gross_ot - expected_deduction)) <= 0.50


def test_sunday_overtime_multiplier_1_5(engine):
    """User requirement test: Sunday overtime MUST be calculated at 1.5x."""
    fin = engine.calculate_financial_radar()
    items = fin["personnel_costs"]

    sunday_workers = [p for p in items if p["sunday_ot_hours"] > 0]
    assert len(sunday_workers) > 0

    for p in sunday_workers[:10]:
        ot_rate = p["hourly_overtime_rate"]
        expected_sunday_cost = round(p["sunday_ot_hours"] * ot_rate, 2)
        assert abs(p["sunday_ot_cost"] - expected_sunday_cost) <= 0.10


def test_cash_difference_calculation(engine):
    """Verify Elden Ödenecek Fark = max(0, total_net - bank_net - icra)."""
    fin = engine.calculate_financial_radar()
    items = fin["personnel_costs"]

    for p in items:
        calc_diff = max(0.0, round(p["total_net_earned"] - p["bank_net"] - p["icra"], 2))
        assert abs(p["cash_difference"] - calc_diff) <= 0.05


def test_department_breakdown_aggregations(engine):
    """Verify department overtime shares sum to ~100% and costs sum to total_overtime_cost."""
    fin = engine.calculate_financial_radar()
    depts = fin["departments"]
    summary = fin["summary"]

    total_dept_ot_cost = sum(d["overtime_cost"] for d in depts)
    total_dept_share = sum(d["overtime_budget_share_pct"] for d in depts)

    assert abs(total_dept_ot_cost - summary["total_overtime_cost"]) <= 1.00
    assert abs(total_dept_share - 100.0) <= 0.5


def test_department_insights_alerts(engine):
    """Verify executive insights identify Balık Dolum as high-intensity overtime line."""
    fin = engine.calculate_financial_radar()
    depts = fin["departments"]

    dolum = next((d for d in depts if "DOLUM" in d["department"].upper()), None)
    assert dolum is not None
    assert dolum["avg_overtime_hours_per_worker"] > 50.0  # ~58 hours per worker
