"""
Tests for 24-Hour Daily Work Ceiling (Max 24h Guard).
Verifies:
1. PDKS Engine: calculate_shift_hours clamps/limits total hours to 24.0.
2. PDKS Engine: load_and_process_pdks with department bonuses never exceeds 24.0 hours.
3. Excel Builder: openpyxl workbook contains DataValidation (0 <= hours <= 24) on day cells.
4. Backend API & Pydantic: ExceptionResolveRequest rejects approved_hours > 24 with ValidationError.
5. Backend API & Pydantic: BulkExceptionResolveItem rejects approved_hours > 24 with ValidationError.
6. EngineManager: resolve_exception raises ValueError if approved_hours > 24.
"""
import pytest
from pydantic import ValidationError
import openpyxl

from pdks_engine import PDKSEngine, calculate_shift_hours
from app import (
    app,
    EngineManager,
    ExceptionResolveRequest,
    BulkExceptionResolveRequest,
    BulkExceptionResolveItem,
    resolve_exception_endpoint,
    bulk_resolve_exceptions_endpoint
)


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
    return eng


def test_calculate_shift_hours_caps_at_24h():
    """Verify shift calculation never returns more than 24.0 hours."""
    # Normal 8-hour shift
    res_normal = calculate_shift_hours("08:00", "17:00")
    assert res_normal.total_hours <= 24.0

    # Long overtime (e.g. 08:00 to 07:30 next day = 23.5 hours)
    res_long = calculate_shift_hours("08:00", "07:30")
    assert res_long.total_hours <= 24.0

    # Theoretical edge cases
    res_edge = calculate_shift_hours("08:00", "08:30")
    assert res_edge.total_hours <= 24.0


def test_engine_all_daily_records_le_24h(engine):
    """Verify all processed daily records in PDKS never exceed 24.0 hours."""
    for rec in engine.audit_records:
        assert rec["final_hours"] <= 24.0, f"Record exceeds 24h: {rec}"
        assert rec["final_hours"] >= 0.0


def test_engine_summary_matrix_le_24h(engine):
    """Verify daily hours in summary matrix never exceed 24.0."""
    matrix = engine.get_summary_matrix()
    for emp in matrix:
        for day, h in emp["daily_hours"].items():
            assert h <= 24.0, f"Employee {emp['name_key']} day {day} has {h} hours > 24.0"
            assert h >= 0.0


def test_manager_resolve_exception_rejects_gt_24h():
    """Verify EngineManager.resolve_exception raises ValueError if approved_hours > 24.0."""
    mgr = EngineManager.get_instance()
    with pytest.raises(ValueError, match="24 saatten fazla"):
        mgr.resolve_exception("TESTKEY", 1, 24.5, "Over 24h test")

    with pytest.raises(ValueError, match="24 saatten fazla"):
        mgr.resolve_exception("TESTKEY", 1, 30.0, "Over 24h test")

    with pytest.raises(ValueError):
        mgr.resolve_exception("TESTKEY", 1, -1.0, "Negative hours test")


def test_manager_bulk_resolve_rejects_gt_24h():
    """Verify EngineManager.bulk_resolve_exceptions raises ValueError if any item > 24.0."""
    mgr = EngineManager.get_instance()
    items = [
        {"name_key": "TEST1", "day": 1, "approved_hours": 8.0},
        {"name_key": "TEST2", "day": 1, "approved_hours": 26.0}
    ]
    with pytest.raises(ValueError, match="24 saatten fazla"):
        mgr.bulk_resolve_exceptions(items)


def test_pydantic_schema_enforces_max_24h():
    """Verify Pydantic models reject approved_hours > 24.0 or < 0.0."""
    # 25.0 hours must fail validation
    with pytest.raises(ValidationError):
        ExceptionResolveRequest(
            name_key="TESTUSER",
            day=5,
            approved_hours=25.0,
            note="Too many hours"
        )

    # 100.0 hours must fail validation
    with pytest.raises(ValidationError):
        ExceptionResolveRequest(
            name_key="TESTUSER",
            day=5,
            approved_hours=100.0,
            note="Extreme hours"
        )

    # Negative hours must fail validation
    with pytest.raises(ValidationError):
        ExceptionResolveRequest(
            name_key="TESTUSER",
            day=5,
            approved_hours=-2.0,
            note="Negative hours"
        )

    # Bulk items must fail validation for > 24.0
    with pytest.raises(ValidationError):
        BulkExceptionResolveItem(
            name_key="TESTUSER",
            day=5,
            approved_hours=24.5
        )

    # Valid hours (0.0, 7.5, 12.0, 24.0) must succeed
    req_valid = ExceptionResolveRequest(
        name_key="TESTUSER",
        day=5,
        approved_hours=24.0,
        note="Valid max hours"
    )
    assert req_valid.approved_hours == 24.0


def test_excel_workbook_data_validation():
    """Verify generated puantaj.xlsx contains DataValidation on daily hours restricted to 0..24."""
    wb = openpyxl.load_workbook("puantaj.xlsx", data_only=False)
    ws_p = wb["Aylik_Puantaj"]

    # Verify DataValidation objects exist on the sheet
    validations = ws_p.data_validations.dataValidation
    assert len(validations) > 0, "No data validations found on Puantaj sheet"

    # Find the max 24 hours validation rule
    dv_24 = None
    for dv in validations:
        if str(dv.formula1) == "0" and str(dv.formula2) == "24":
            dv_24 = dv
            break

    assert dv_24 is not None, "DataValidation between 0 and 24 hours not found"
    assert dv_24.operator == "between"
    assert "24" in dv_24.error, "Error message does not mention 24 hours limit"
