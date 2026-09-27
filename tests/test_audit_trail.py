# -*- coding: utf-8 -*-
"""
Tests for Audit Trail (Değişiklik ve Güvenlik Denetim İzi) Engine.
Validates SQLite audit_logs table, immutable logging, search & filters,
summary KPIs, openpyxl corporate Excel generation, A4 printable HTML report,
and automated event instrumentation across all change endpoints.
"""
import os
import io
import tempfile
import pytest
from starlette.requests import Request
import openpyxl

from db_manager import DatabaseManager
from app import (
    EngineManager,
    USERS,
    ExceptionResolveRequest,
    BulkExceptionResolveRequest,
    BulkExceptionResolveItem,
    RuleUpdateRequest,
    IcraCreateUpdateRequest,
    PeriodRolloverRequest,
    resolve_exception_endpoint,
    bulk_resolve_exceptions_endpoint,
    update_rules,
    add_update_icra_endpoint,
    close_and_rollover_period_endpoint,
    get_audit_logs_endpoint,
    export_audit_excel_endpoint,
    get_audit_print_html_endpoint,
    generate_audit_print_html
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


def test_audit_db_table_and_indexes(temp_db):
    """audit_logs tablosu ve indeksleri eksiksiz oluşturulmalıdır."""
    with temp_db.get_conn() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='audit_logs'")
        assert cursor.fetchone() is not None, "audit_logs tablosu oluşturulmadı!"

        # Sütunları kontrol et
        cursor.execute("PRAGMA table_info(audit_logs)")
        cols = {row[1] for row in cursor.fetchall()}
        expected_cols = {
            "id", "timestamp", "username", "user_name", "user_role",
            "action_type", "target_tc", "target_name", "target_dept",
            "day", "old_value", "new_value", "reason", "ip_address", "metadata_json"
        }
        assert expected_cols.issubset(cols), f"Eksik sütunlar var: {expected_cols - cols}"

        # İndeksleri kontrol et
        cursor.execute("SELECT name FROM sqlite_master WHERE type='index'")
        indexes = {row[0] for row in cursor.fetchall()}
        assert "idx_audit_time" in indexes
        assert "idx_audit_action" in indexes
        assert "idx_audit_tc" in indexes
        assert "idx_audit_user" in indexes


def test_db_log_action_and_retrieval(temp_db):
    """db.log_action ile eklenen kayıtlar eksiksiz ve doğru değerlerle okunmalıdır."""
    log_id = temp_db.log_action(
        username="ik_yonetici",
        user_name="Ayşe Yılmaz",
        user_role="hr",
        action_type="EXCEPTION_RESOLVE",
        target_tc="12345678901",
        target_name="AHMET YILMAZ",
        target_dept="Balık Dolum",
        day=14,
        old_value="7.5s",
        new_value="9.0s",
        reason="Vardiya amiri ek mesai tutanağı",
        ip_address="192.168.1.45",
        metadata={"unit": "hours", "multiplier": 1.5}
    )
    assert log_id > 0

    res = temp_db.get_audit_logs(page=1, page_size=10)
    assert res["total"] >= 1
    items = res["items"]
    match = next((it for it in items if it["id"] == log_id), None)
    assert match is not None
    assert match["username"] == "ik_yonetici"
    assert match["user_name"] == "Ayşe Yılmaz"
    assert match["user_role"] == "hr"
    assert match["action_type"] == "EXCEPTION_RESOLVE"
    assert match["target_name"] == "AHMET YILMAZ"
    assert match["day"] == 14
    assert match["old_value"] == "7.5s"
    assert match["new_value"] == "9.0s"
    assert match["reason"] == "Vardiya amiri ek mesai tutanağı"
    assert match["ip_address"] == "192.168.1.45"
    assert match["metadata"]["multiplier"] == 1.5


def test_audit_filters_and_pagination(temp_db):
    """Arama, işlem türü filtresi, kullanıcı filtresi ve sayfalama doğru çalışmalıdır."""
    # 5 farklı log ekle
    temp_db.log_action(
        username="ik_yonetici", user_name="Ayşe", user_role="hr",
        action_type="EXCEPTION_RESOLVE", target_name="MEHMET DEMİR",
        reason="Kart basımı unutuldu", ip_address="10.0.0.1"
    )
    temp_db.log_action(
        username="muhasebe", user_name="Burak", user_role="accounting",
        action_type="ICRA_ADD", target_name="FATMA KAYA",
        reason="İcra dairesi 2026/102 dosya", ip_address="10.0.0.2"
    )
    temp_db.log_action(
        username="admin", user_name="Sistem", user_role="admin",
        action_type="RULE_UPDATE", target_name="Fabrika Kuralları",
        reason="Sabah toleransı güncellendi", ip_address="10.0.0.3"
    )

    # 1. Arama filtresi: MEHMET
    res_search = temp_db.get_audit_logs(search="MEHMET")
    assert res_search["total"] == 1
    assert res_search["items"][0]["target_name"] == "MEHMET DEMİR"

    # 2. İşlem türü filtresi: ICRA_ADD
    res_action = temp_db.get_audit_logs(action_type="ICRA_ADD")
    assert any(it["action_type"] == "ICRA_ADD" for it in res_action["items"])
    assert all(it["action_type"] == "ICRA_ADD" for it in res_action["items"])

    # 3. Kullanıcı filtresi: admin
    res_user = temp_db.get_audit_logs(username="admin")
    assert all(it["username"] == "admin" for it in res_user["items"])

    # 4. Sayfalama
    res_p1 = temp_db.get_audit_logs(page=1, page_size=2)
    assert len(res_p1["items"]) <= 2
    assert res_p1["page"] == 1
    assert res_p1["page_size"] == 2


def test_audit_summary_kpis(temp_db):
    """Denetim izi özet KPI'ları (toplam, bugün, saat revizyonu, en aktif yetkili) doğru hesaplanmalıdır."""
    temp_db.log_action(
        username="ik_yonetici", user_name="Ayşe Yılmaz", user_role="hr",
        action_type="EXCEPTION_RESOLVE", target_name="P1", old_value="0s", new_value="7.5s"
    )
    temp_db.log_action(
        username="ik_yonetici", user_name="Ayşe Yılmaz", user_role="hr",
        action_type="BULK_RESOLVE", target_name="P2", old_value="0s", new_value="8s"
    )
    temp_db.log_action(
        username="admin", user_name="Sistem Yöneticisi", user_role="admin",
        action_type="RULE_UPDATE", target_name="Kurallar"
    )

    res = temp_db.get_audit_logs()
    summary = res["summary"]
    assert summary["total_logs"] >= 3
    assert summary["hours_revisions"] >= 2
    assert summary["today_logs"] >= 3
    assert summary["most_active_user"] != ""


def test_audit_excel_export(temp_db):
    """Denetim izi resmi Excel çalışma kitabı oluşturulup açılabilmelidir."""
    temp_db.log_action(
        username="ik_yonetici", user_name="Ayşe Yılmaz", user_role="hr",
        action_type="EXCEPTION_RESOLVE", target_tc="99887766554",
        target_name="TEST PERSONEL", target_dept="Konserve",
        day=5, old_value="7.5s", new_value="10.0s",
        reason="İş Müfettişi Testi", ip_address="127.0.0.1"
    )

    excel_bytes = temp_db.export_audit_excel()
    assert len(excel_bytes) > 1000

    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    assert "İdari Değişiklik Denetim İzi" in wb.sheetnames
    ws = wb["İdari Değişiklik Denetim İzi"]

    # Başlık ve şirket adı
    assert "FİDE KONSERVE" in str(ws["A1"].value)
    # Tutanak başlığı
    assert "DENETİM" in str(ws["A2"].value)


def test_audit_api_endpoints():
    """FastAPI denetim izi endpointleri (logs, export/excel, print/html) doğru yanıt vermelidir."""
    req_admin = make_mock_request("admin")

    # 1. GET /api/audit/logs
    res_logs = get_audit_logs_endpoint(page=1, page_size=20, request=req_admin)
    assert "items" in res_logs
    assert "summary" in res_logs
    assert res_logs["total"] >= 1

    # 2. GET /api/audit/export/excel
    res_excel = export_audit_excel_endpoint(request=req_admin)
    assert res_excel.status_code == 200
    assert "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" in res_excel.media_type
    assert len(res_excel.body) > 1000

    # 3. GET /api/audit/print/html
    res_html = get_audit_print_html_endpoint(request=req_admin)
    assert res_html.status_code == 200
    assert "text/html" in res_html.media_type
    html_str = res_html.body.decode("utf-8")
    assert "FİDE KONSERVE GIDA SAN. VE TİC. A.Ş." in html_str
    assert "RESMİ İŞ MÜFETTİŞLİĞİ" in html_str
    assert "DÜZENLEYEN" in html_str
    assert "FABRİKA MÜDÜRÜ" in html_str


def test_event_instrumentation_on_exception_resolve():
    """resolve_exception_endpoint çağrıldığında audit_logs tablosuna EXCEPTION_RESOLVE kaydı yazılmalıdır."""
    mgr = EngineManager.get_instance()
    eng = mgr.engine
    p = eng.personnel_list[0]
    name_key = p["name_key"]

    req_hr = make_mock_request("ik_yonetici")
    payload = ExceptionResolveRequest(
        name_key=name_key,
        day=3,
        approved_hours=9.0,
        note="Test Audit Loglama",
        resolved_by="Ayşe Yılmaz"
    )

    resolve_exception_endpoint(payload, request=req_hr)

    db = DatabaseManager()
    res = db.get_audit_logs(search="Test Audit Loglama")
    assert res["total"] >= 1
    item = res["items"][0]
    assert item["action_type"] == "EXCEPTION_RESOLVE"
    assert item["username"] == "ik_yonetici"
    assert item["new_value"] == "9.0s"


def test_event_instrumentation_on_bulk_resolve():
    """bulk_resolve_exceptions_endpoint çağrıldığında audit_logs tablosuna BULK_RESOLVE kaydı yazılmalıdır."""
    mgr = EngineManager.get_instance()
    eng = mgr.engine
    p1 = eng.personnel_list[0]
    p2 = eng.personnel_list[1]

    req_hr = make_mock_request("ik_yonetici")
    items = [
        BulkExceptionResolveItem(name_key=p1["name_key"], day=4, approved_hours=7.5, note="Toplu 1"),
        BulkExceptionResolveItem(name_key=p2["name_key"], day=4, approved_hours=7.5, note="Toplu 2")
    ]
    payload = BulkExceptionResolveRequest(
        items=items,
        action_type="standard_7_5",
        global_note="Otomatik Toplu Onay Denetim Testi"
    )

    bulk_resolve_exceptions_endpoint(payload, request=req_hr)

    db = DatabaseManager()
    res = db.get_audit_logs(search="Otomatik Toplu Onay Denetim Testi")
    assert res["total"] >= 1
    item = res["items"][0]
    assert item["action_type"] == "BULK_RESOLVE"
    assert item["username"] == "ik_yonetici"


def test_event_instrumentation_on_rule_update():
    """update_rules çağrıldığında RULE_UPDATE denetim kaydı oluşturulmalıdır."""
    req_admin = make_mock_request("admin")
    from config.settings import load_rules
    current = load_rules()

    payload = RuleUpdateRequest(rules=current)
    update_rules(payload, request=req_admin)

    db = DatabaseManager()
    res = db.get_audit_logs(action_type="RULE_UPDATE")
    assert res["total"] >= 1
    item = res["items"][0]
    assert item["action_type"] == "RULE_UPDATE"
    assert item["username"] == "admin"


def test_event_instrumentation_on_icra_add():
    """add_update_icra_endpoint çağrıldığında ICRA_ADD denetim kaydı oluşturulmalıdır."""
    req_acc = make_mock_request("muhasebe")
    payload = IcraCreateUpdateRequest(
        tc="11223344556",
        name_key="TEST_AUDIT_ICRA",
        ad_soyad="TEST AUDIT ICRA",
        bolum="Muhasebe",
        sirket="Test İcra Dairesi",
        dosya_no="2026/9999",
        toplam_borc=15000.0,
        aciklama="Test İcra Tebligatı"
    )

    add_update_icra_endpoint(payload, request=req_acc)

    db = DatabaseManager()
    res = db.get_audit_logs(search="TEST AUDIT ICRA")
    assert res["total"] >= 1
    item = res["items"][0]
    assert item["action_type"] == "ICRA_ADD"
    assert item["username"] == "muhasebe"
