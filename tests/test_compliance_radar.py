"""
Tests for 4857 Sayılı İş Kanunu ve SGK Uyum / Risk Radarı (Legal Compliance Radar).
Verifies:
1. 11-Hour Minimum Rest Rule (Madde 68 & Postalar Yönetmeliği Md. 9): < 11.0h detection, < 8.0h critical
2. Weekly Rest & Consecutive Work Rule (Madde 46): 7+ days continuous work detection, 10+ days critical
3. Annual 270-Hour Overtime Limit & Monthly Pace (Madde 41): >= 35.0h warning, >= 50.0h critical
4. Factory Legal Compliance Index (LCI) & Department Risk Matrix
5. REST API endpoints and Excel Inspection Report generation
"""
import os
import pytest
from pdks_engine import PDKSEngine
from app import get_compliance_radar, generate_compliance_excel, EngineManager


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


def test_rest_11h_violations_detection(engine):
    """Verify detection of shift transitions with less than 11 hours of uninterrupted rest."""
    comp = engine.calculate_legal_compliance_radar()
    summary = comp["summary"]
    rest_items = comp["rest_violations"]

    assert summary["rest_violations_count"] == 20
    assert len(rest_items) == 20

    for item in rest_items:
        assert item["violation_type"] == "REST_11H"
        assert item["legal_article"] == "4857 SK Madde 68 & Postalar Yön. Md. 9"
        assert item["end_day"] == item["start_day"] + 1

        # Check metric value parse
        hours_str = item["metric_value"].split()[0].replace(",", ".")
        rest_h = float(hours_str)
        assert 0.0 < rest_h < 11.0
        assert item["severity"] in ("CRITICAL", "WARNING")

    # Verify specific real case: Kemal Durmuş (7.6 hours rest)
    kemal = next((v for v in rest_items if "KEMAL DURMUŞ" in v["ad_soyad"].upper()), None)
    assert kemal is not None
    assert kemal["severity"] == "CRITICAL"
    assert "7.6 saat" in kemal["metric_value"]


def test_consecutive_7d_work_violations(engine):
    """Verify detection of 7+ consecutive work days without 24-hour weekly rest (Hafta Tatili)."""
    comp = engine.calculate_legal_compliance_radar()
    summary = comp["summary"]
    consec_items = comp["consecutive_violations"]

    assert summary["consecutive_work_count"] == 62
    assert len(consec_items) == 62

    for item in consec_items:
        assert item["violation_type"] == "CONSECUTIVE_7D"
        assert "4857 SK Madde 46" in item["legal_article"]

        days_count = int(item["metric_value"].split()[0])
        assert days_count >= 7

        if days_count >= 10:
            assert item["severity"] == "CRITICAL"
        else:
            assert item["severity"] == "WARNING"

    # Verify long streaks: Cemalettin Bedizci (14 days), Ayşe Şenecioğlu (13 days)
    bedizci = next((v for v in consec_items if "BEDİZCİ" in v["ad_soyad"].upper() or "BEDIZCI" in v["ad_soyad"].upper()), None)
    assert bedizci is not None
    assert int(bedizci["metric_value"].split()[0]) == 14
    assert bedizci["severity"] == "CRITICAL"


def test_overtime_270h_risk_detection(engine):
    """Verify detection of employees on track to breach the annual 270-hour overtime ceiling."""
    comp = engine.calculate_legal_compliance_radar()
    summary = comp["summary"]
    ot_items = comp["overtime_risks"]

    assert summary["overtime_limit_count"] == 26
    assert len(ot_items) == 26

    for item in ot_items:
        assert item["violation_type"] == "OVERTIME_270H"
        assert "4857 SK Madde 41" in item["legal_article"]

        ot_h = float(item["metric_value"].split()[0].replace(",", "."))
        assert ot_h >= 35.0

        if ot_h >= 50.0:
            assert item["severity"] == "CRITICAL"
        else:
            assert item["severity"] == "WARNING"

    # Verify highest overtime workers: Edip Ekinci (116.0s FM)
    edip = next((v for v in ot_items if "EDİP" in v["ad_soyad"].upper() or "EDIP" in v["ad_soyad"].upper()), None)
    assert edip is not None
    assert edip["severity"] == "CRITICAL"
    assert "116.0" in edip["metric_value"]


def test_compliance_index_and_department_matrix(engine):
    """Verify factory Legal Compliance Index (LCI) and department breakdown."""
    comp = engine.calculate_legal_compliance_radar()
    summary = comp["summary"]
    depts = comp["department_compliance"]

    assert summary["compliance_index"] == 62.1
    assert summary["risk_status"] == "YÜKSEK RİSK"
    assert summary["total_violations"] == 108
    assert summary["critical_violations"] == 62
    assert summary["warning_violations"] == 46
    assert summary["inspected_personnel"] == 254

    # Ensure department list is ranked with high risk depts first
    assert len(depts) > 0
    top_dept = depts[0]
    assert top_dept["total_violations"] > 0
    assert top_dept["risk_level"] in ("YÜKSEK", "ORTA")


def test_api_compliance_radar_filtering(engine):
    """Verify get_compliance_radar logic and query parameter filtering."""
    # 1. Default request
    data = get_compliance_radar()
    assert data["total_count"] == 108
    assert len(data["items"]) == 50  # default page_size
    assert data["summary"]["compliance_index"] == 62.1

    # 2. Filter by REST_11H
    data_rest = get_compliance_radar(violation_type="REST_11H")
    assert data_rest["total_count"] == 20
    assert all(item["violation_type"] == "REST_11H" for item in data_rest["items"])

    # 3. Filter by CONSECUTIVE_7D
    data_consec = get_compliance_radar(violation_type="CONSECUTIVE_7D")
    assert data_consec["total_count"] == 62
    assert all(item["violation_type"] == "CONSECUTIVE_7D" for item in data_consec["items"])

    # 4. Filter by OVERTIME_270H
    data_ot = get_compliance_radar(violation_type="OVERTIME_270H")
    assert data_ot["total_count"] == 26
    assert all(item["violation_type"] == "OVERTIME_270H" for item in data_ot["items"])

    # 5. Filter by severity CRITICAL
    data_crit = get_compliance_radar(severity="CRITICAL")
    assert data_crit["total_count"] == 62
    assert all(item["severity"] == "CRITICAL" for item in data_crit["items"])


def test_compliance_excel_generation(engine, tmp_path):
    """Verify generation of multi-sheet Fide_Konserve_Yasal_Uyum_ve_Risk_Radari.xlsx."""
    comp = engine.calculate_legal_compliance_radar()
    out_file = str(tmp_path / "test_compliance.xlsx")
    generate_compliance_excel(comp, out_file)

    assert os.path.exists(out_file)
    file_size = os.path.getsize(out_file)
    assert file_size > 5000  # valid multi-sheet Excel file with 3 sheets and openpyxl formatting
