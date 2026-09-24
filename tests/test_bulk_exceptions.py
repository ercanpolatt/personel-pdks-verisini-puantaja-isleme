# -*- coding: utf-8 -*-
"""
Toplu Onay ve Akıllı Amir Öneri Motoru Testleri
"""

import os
import pytest
from pdks_engine import PDKSEngine
from app import (
    app,
    EngineManager,
    get_exceptions,
    bulk_resolve_exceptions_endpoint,
    BulkExceptionResolveRequest,
    BulkExceptionResolveItem
)

@pytest.fixture(scope="module")
def loaded_engine():
    if not os.path.exists("puantaj.xls") or not os.path.exists("pdks.xls"):
        pytest.skip("puantaj.xls veya pdks.xls bulunamadığı için test atlandı.")
    engine = PDKSEngine(pdks_path="pdks.xls", puantaj_path="puantaj.xls", target_month=9, target_year=2026)
    engine.load_personnel()
    engine.load_and_process_pdks()
    return engine

def test_smart_suggestions_populated(loaded_engine):
    """Her istisna kaydının eksiksiz akıllı amir önerisine sahip olduğu doğrulanmalıdır."""
    assert len(loaded_engine.exception_records) > 0
    for exc in loaded_engine.exception_records:
        assert "smart_suggestion" in exc
        sug = exc["smart_suggestion"]
        assert "suggested_hours" in sug
        assert sug["suggested_hours"] > 0
        assert "suggested_g" in sug
        assert "suggested_c" in sug
        assert "confidence" in sug
        assert sug["confidence"] in (75, 85, 95)
        assert sug["source"] in ("DEPARTMENT_PEERS", "PERSONAL_HABIT", "FACTORY_STANDARD")
        assert "reason" in sug
        assert len(sug["reason"]) > 5

def test_department_peer_heuristic_dominant(loaded_engine):
    """Fabrika ortak mesai kuralı gereği istisnaların büyük çoğunluğu (%65+) aynı gün çalışan mesai arkadaşlarından tahmin edilmelidir."""
    sources = [exc["smart_suggestion"]["source"] for exc in loaded_engine.exception_records]
    peer_count = sources.count("DEPARTMENT_PEERS")
    total = len(sources)
    peer_ratio = peer_count / total
    assert peer_ratio >= 0.65, f"Beklenen peer oranı >= 0.65, gerçekleşen: {peer_ratio:.2f}"

def test_api_exceptions_endpoint_contains_suggestions():
    """get_exceptions yanıtındaki tüm öğeler smart_suggestion içermelidir."""
    data = get_exceptions(status="all", search=None)
    assert "items" in data
    assert len(data["items"]) > 0
    first = data["items"][0]
    assert "smart_suggestion" in first
    assert first["smart_suggestion"]["confidence"] >= 75

def test_api_bulk_resolve_flow():
    """bulk_resolve_exceptions_endpoint fonksiyonu seçilen kayıtları başarıyla toplu onaylamalıdır."""
    test_payload = BulkExceptionResolveRequest(
        items=[
            BulkExceptionResolveItem(
                name_key="TESTPERSONELA",
                day=5,
                approved_hours=8.0,
                note="Toplu Test Onayı 1"
            ),
            BulkExceptionResolveItem(
                name_key="TESTPERSONELB",
                day=6,
                approved_hours=7.5,
                note="Toplu Test Onayı 2"
            )
        ],
        action_type="smart",
        global_note="Toplu Amir Akıllı Onayı",
        resolved_by="Test Amiri"
    )

    body = bulk_resolve_exceptions_endpoint(test_payload)
    assert body["status"] == "success"
    assert body["resolved_count"] == 2

    # Temizlik: test anahtarlarını sil ve dosyayı güncelle
    mgr = EngineManager.get_instance()
    for k in ["TESTPERSONELA:5", "TESTPERSONELB:6"]:
        mgr.resolved_exceptions.pop(k, None)
    from app import save_resolved_exceptions
    save_resolved_exceptions(mgr.resolved_exceptions)
