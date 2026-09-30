# -*- coding: utf-8 -*-
"""
Tests for Multi-Month Archive, SQLite Persistence, Cumulative 270h Overtime, and Icra Rollover Engine.
"""
import os
import tempfile
import pytest
from starlette.requests import Request

from db_manager import DatabaseManager, MONTH_NAMES_TR
from app import (
    EngineManager,
    PeriodRolloverRequest,
    IcraCreateUpdateRequest,
    list_periods_endpoint,
    get_period_details_endpoint,
    close_and_rollover_period_endpoint,
    get_cumulative_overtime_endpoint,
    get_icra_records_endpoint,
    add_update_icra_endpoint
)


def make_mock_request(username: str = "admin") -> Request:
    """Belirtilen kullanıcı rolüyle sahte bir HTTP Request nesnesi üretir."""
    headers = [(b"x-user-role", username.encode("utf-8"))]
    scope = {
        "type": "http",
        "method": "GET",
        "path": "/",
        "headers": headers
    }
    return Request(scope)


@pytest.fixture
def temp_db():
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    db = DatabaseManager(db_path=path)
    yield db
    if os.path.exists(path):
        try:
            os.remove(path)
        except Exception:
            pass


def test_db_init_and_tables(temp_db):
    """Tüm SQLite ilişkisel tabloları ve indeksleri eksiksiz oluşturulmalıdır."""
    tables = [
        "periods", "monthly_records", "period_exceptions",
        "cumulative_overtime", "icra_records", "icra_movements"
    ]
    with temp_db.get_conn() as conn:
        cursor = conn.cursor()
        for t in tables:
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (t,))
            assert cursor.fetchone() is not None, f"Tablo {t} oluşturulamadı!"


def test_seed_initial_data(temp_db):
    """PDKS motorundaki 254 personel, icra listesi ve kümülatif mesai SQLite'a aktarılmalıdır."""
    mgr = EngineManager.get_instance()
    mgr.load()

    # İlk aktarım başarılı olmalıdır
    success = temp_db.seed_initial_data(mgr.engine)
    assert success is True

    # İkinci aktarım tekrarlanmamalıdır (idempotent)
    success_second = temp_db.seed_initial_data(mgr.engine)
    assert success_second is False

    # Aktif dönem kontrolü
    periods = temp_db.get_all_periods()
    assert len(periods) >= 1
    assert periods[0]["period_key"] == "2026-09"
    assert periods[0]["status"] == "active"
    assert periods[0]["active_personnel_count"] == 254

    # Personel kayıtları kontrolü
    details = temp_db.get_period_details("2026-09")
    assert details is not None
    assert len(details["records"]) == 254


def test_cumulative_overtime_270h_rules(temp_db):
    """4857 SK Md. 41 kapsamındaki 270 saat fazla mesai kütüğü ve risk dağılımı doğrulanmalıdır."""
    mgr = EngineManager.get_instance()
    mgr.load()
    temp_db.seed_initial_data(mgr.engine)

    report = temp_db.get_cumulative_overtime_report(year=2026)
    assert report["year"] == 2026
    assert report["total_personnel"] == 254

    kpis = report["kpis"]
    assert kpis["legal_limit_hours"] == 270.0
    assert "safe_count" in kpis
    assert "warning_count" in kpis
    assert "critical_count" in kpis
    assert "exceeded_count" in kpis
    assert (kpis["safe_count"] + kpis["warning_count"] + kpis["critical_count"] + kpis["exceeded_count"]) == 254

    # Personel kayıtlarında 12 ay sütunları ve kalan süre kontrolü
    personnel = report["personnel"]
    assert len(personnel) == 254
    for p in personnel[:10]:
        assert "m01" in p and "m09" in p and "m12" in p
        assert p["remaining_limit_hours"] >= 0.0
        assert p["risk_status"] in ("SAFE", "WARNING", "CRITICAL", "EXCEEDED")


def test_icra_records_and_rollover(temp_db):
    """İcra dosyalarının otomatik düşümü, bakiye devri ve dönem kapatma motoru doğrulanmalıdır."""
    mgr = EngineManager.get_instance()
    mgr.load()
    temp_db.seed_initial_data(mgr.engine)

    # İcra kayıtlarını oku
    icra_data = temp_db.get_icra_records(status="ALL")
    assert icra_data["status"] == "success"
    assert icra_data["kpis"]["total_files"] >= 10
    assert icra_data["kpis"]["active_files"] >= 10
    assert icra_data["kpis"]["total_debt_tl"] > 500000

    # Eylül 2026 dönemini kapat ve Ekim 2026'ya devret
    rollover_res = temp_db.close_and_rollover_period("2026-09")
    assert rollover_res["status"] == "success"
    assert rollover_res["archived_period"] == "2026-09"
    assert rollover_res["new_active_period"] == "2026-10"
    assert rollover_res["total_icra_deducted"] > 0
    assert len(rollover_res["rollover_files"]) > 0

    # Dönemlerin durumunu kontrol et
    periods = temp_db.get_all_periods()
    period_map = {p["period_key"]: p["status"] for p in periods}
    assert period_map["2026-09"] == "archived"
    assert period_map["2026-10"] == "active"


def test_api_archive_endpoints():
    """Arşiv, 270s mesai kütüğü ve icra takip API endpointleri doğru çalışmalıdır."""
    req_admin = make_mock_request("admin")
    req_acc = make_mock_request("muhasebe")

    # 1. Dönemler Listesi API
    res_periods = list_periods_endpoint(req_admin)
    assert res_periods["status"] == "success"
    assert len(res_periods["periods"]) >= 1

    # 2. Belirli Dönem Detayı API
    res_detail = get_period_details_endpoint("2026-09", req_admin)
    assert res_detail["status"] == "success"
    assert res_detail["period"]["period_key"] == "2026-09"
    assert len(res_detail["period"]["records"]) == 254

    # 3. Yıllık 270 Saat Mesai Kütüğü API
    res_ot = get_cumulative_overtime_endpoint(year=2026, request=req_admin)
    assert res_ot["year"] == 2026
    assert res_ot["total_personnel"] == 254
    assert "kpis" in res_ot

    # 4. İcra Takip Masası API
    res_icra = get_icra_records_endpoint(status="ALL", request=req_acc)
    assert res_icra["status"] == "success"
    assert res_icra["kpis"]["total_files"] >= 10

    # 5. Yeni İcra Kaydı Ekleme API
    new_icra_payload = IcraCreateUpdateRequest(
        tc="99988877766",
        name_key="yeni_test_personel",
        ad_soyad="Yeni Test Personel",
        bolum="Konserve",
        sirket="İzmir 3. İcra Dairesi",
        dosya_no="2026/8888 E.",
        toplam_borc=60000.0,
        kesilen_kumulatif=0.0,
        kalan_borc=60000.0,
        aylik_kesinti_orani=0.25,
        durum="AKTİF",
        aciklama="Test İcra Açıklaması"
    )
    res_add_icra = add_update_icra_endpoint(new_icra_payload, req_acc)
    assert res_add_icra["status"] == "success"

    # Kaydın eklendiğini teyit et
    res_icra_updated = get_icra_records_endpoint(status="ALL", request=req_acc)
    files = res_icra_updated["files"]
    assert any(f["name_key"] == "yeni_test_personel" for f in files)
