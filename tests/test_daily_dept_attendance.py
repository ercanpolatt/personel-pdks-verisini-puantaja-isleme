# -*- coding: utf-8 -*-
"""
Test Suite: Günlük Bölüm Bazında Çalışan Personel Analizi & Patron Masası Testleri
"""

import os
import pytest
from app import (
    EngineManager,
    generate_daily_department_attendance_excel,
    get_daily_department_attendance,
    download_daily_department_excel
)
from pdks_engine import PDKSEngine


@pytest.fixture(scope="module")
def engine_data():
    """PDKS motorunu yükler ve günlük bölüm verisini üretir."""
    mgr = EngineManager.get_instance()
    eng = mgr.engine
    data = eng.calculate_daily_department_attendance()
    return eng, data


def test_calculate_daily_department_attendance_structure(engine_data):
    eng, data = engine_data
    
    assert "kpis" in data
    assert "days" in data
    assert "department_matrix" in data
    assert "all_departments" in data
    
    # 30 günlük takvim doğrulaması
    assert len(data["days"]) == eng.days_in_month
    assert len(data["all_departments"]) > 0
    
    kpis = data["kpis"]
    assert kpis["total_departments"] >= 20
    assert kpis["avg_daily_working_day_workers"] > 100
    assert kpis["total_man_days"] > 3000
    assert kpis["peak_day"]["workers"] > 150
    assert kpis["busiest_department"]["total_man_days"] > 500


def test_daily_attendance_day_breakdown(engine_data):
    eng, data = engine_data
    
    day_1 = data["days"][0]
    assert day_1["day"] == 1
    assert day_1["total_workers"] > 0
    assert len(day_1["departments"]) > 0
    
    top_dept = day_1["departments"][0]
    assert "department" in top_dept
    assert top_dept["count"] > 0
    assert top_dept["hours"] > 0
    assert "workers" in top_dept
    assert len(top_dept["workers"]) == top_dept["count"]


def test_department_matrix_counts(engine_data):
    eng, data = engine_data
    
    matrix = data["department_matrix"]
    assert len(matrix) > 0
    
    for row in matrix:
        assert len(row["daily_counts"]) == eng.days_in_month
        assert row["total_man_days"] == sum(row["daily_counts"])
        assert row["peak_workers"] == max(row["daily_counts"])


def test_excel_generation(engine_data, tmp_path):
    _, data = engine_data
    
    excel_path = str(tmp_path / "test_gunluk_bolum_calisan.xlsx")
    res_path = generate_daily_department_attendance_excel(data, excel_path)
    
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 10000


def test_api_daily_attendance_endpoint():
    res = get_daily_department_attendance()
    assert res["status"] == "success"
    assert "kpis" in res
    assert "days" in res
    assert "department_matrix" in res


def test_api_daily_attendance_day_filter():
    res = get_daily_department_attendance(day=1)
    assert res["status"] == "success"
    assert "selected_day" in res
    assert res["selected_day"]["day"] == 1
    assert res["selected_day"]["total_workers"] > 0


def test_api_download_excel():
    response = download_daily_department_excel()
    assert response is not None
    assert response.media_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert os.path.exists(response.path)
    assert os.path.getsize(response.path) > 10000

