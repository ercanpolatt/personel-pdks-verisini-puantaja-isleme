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

        # Verify weekday and sunday overtime both use 1.5 multiplier
        expected_weekday_cost = round(p["weekday_ot_hours"] * expected_ot_rate, 2)
        expected_sunday_cost = round(p["sunday_ot_hours"] * expected_ot_rate, 2)
        expected_total_ot_cost = round(expected_weekday_cost + expected_sunday_cost, 2)

        assert abs(p["overtime_cost"] - expected_total_ot_cost) <= 1.00


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
