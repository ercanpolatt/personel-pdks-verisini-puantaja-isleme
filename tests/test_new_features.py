# -*- coding: utf-8 -*-
"""
Test Suite: Yeni Kurumsal Özellikler (Servis Radarı, Vardiya Planlama, İcra Brifingi, DB Yedekleme)
"""

import os
import pytest
from app import (
    EngineManager,
    get_shuttle_logistics,
    get_shift_planner,
    assign_shift,
    get_shift_comparison,
    get_executive_briefing,
    trigger_database_backup,
    get_database_backups,
    ShiftAssignRequest
)
from db_manager import DatabaseManager


@pytest.fixture(scope="module")
def engine_instance():
    mgr = EngineManager.get_instance()
    return mgr.engine


def test_shuttle_logistics_engine(engine_instance):
    eng = engine_instance
    data = eng.calculate_shuttle_logistics()

    assert "kpis" in data
    assert "routes_summary" in data
    assert "days" in data
    assert len(data["days"]) == eng.days_in_month

    kpis = data["kpis"]
    assert kpis["total_routes"] >= 5
    assert kpis["total_assigned_personnel"] > 0
    assert kpis["total_passenger_trips"] > 1000

    # Güzergah özetleri kontrolü
    routes = data["routes_summary"]
    route_names = [r["route_name"] for r in routes]
    assert any("Bergama" in name for name in route_names)
    assert any("Dikili" in name for name in route_names)

    # 1. Gün servis dökümü
    day1 = data["days"][0]
    assert day1["day"] == 1
    assert day1["total_passengers"] > 0
    assert day1["total_vehicles"] > 0
    assert len(day1["routes"]) == len(routes)


def test_api_shuttle_logistics_endpoint():
    res = get_shuttle_logistics(day=1)
    assert res["status"] == "success"
    assert "days" in res
    assert len(res["days"]) == 1
    assert res["selected_day"] == 1

    # Tüm ay sorgusu
    res_all = get_shuttle_logistics(day=None)
    assert res_all["status"] == "success"
    assert len(res_all["days"]) == 30


def test_shift_planner_and_deviation_radar(engine_instance):
    eng = engine_instance
    period_key = f"{eng.target_year}-{eng.target_month:02d}"

    # 1. Personele V2 vardiyası ata
    test_person = eng.personnel_list[0]
    nk = test_person["name_key"]

    req = ShiftAssignRequest(
        period_key=period_key,
        day=5,
        name_key=nk,
        shift_code="V2",
        note="Özel Deneme Vardiyası"
    )
    assign_res = assign_shift(req)
    assert assign_res["status"] == "success"

    # 2. Vardiya çizelgesini oku
    planner_res = get_shift_planner(period_key=period_key, day=5)
    assert planner_res["status"] == "success"
    assert "schedules" in planner_res
    sched_map = {s["name_key"]: s for s in planner_res["schedules"]}
    assert nk in sched_map
    assert sched_map[nk]["shift_code"] == "V2"

    # 3. Gerçekleşenle karşılaştırma yap
    comp_res = get_shift_comparison(period_key=period_key, day=5)
    assert comp_res["status"] == "success"
    assert "compliance_rate" in comp_res
    assert "counts" in comp_res
    assert "deviations" in comp_res
    assert "compliant" in comp_res


def test_executive_briefing():
    # Sabah brifingi
    morning = get_executive_briefing(briefing_type="morning", day=1)
    assert morning["status"] == "success"
    assert "SABAH YÖNETİM" in morning["title"]
    assert "Sahadaki Personel" in morning["markdown"]
    assert "Servis Yolcusu" in morning["markdown"]

    # Akşam brifingi
    evening = get_executive_briefing(briefing_type="evening", day=1)
    assert evening["status"] == "success"
    assert "AKŞAM KAPANIŞ" in evening["title"]
    assert "Fazla Mesai" in evening["markdown"]


def test_database_backup():
    # Anlık yedek al
    backup_res = trigger_database_backup()
    assert backup_res["success"] is True
    assert "puantaj_archive_" in backup_res["filename"]
    assert backup_res["size_bytes"] > 0
    assert os.path.exists(backup_res["path"])

    # Yedek listesini kontrol et
    list_res = get_database_backups()
    assert list_res["status"] == "success"
    assert list_res["count"] >= 1
    found = any(b["filename"] == backup_res["filename"] for b in list_res["backups"])
    assert found is True
