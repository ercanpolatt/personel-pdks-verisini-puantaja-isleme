# -*- coding: utf-8 -*-
"""
Rol Bazlı Erişim ve Çoklu Kullanıcı Girişi (RBAC) Testleri
"""

import pytest
from starlette.requests import Request
from fastapi import HTTPException

from app import (
    USERS,
    LoginRequest,
    RuleUpdateRequest,
    ExceptionResolveRequest,
    login_endpoint,
    get_current_user_endpoint,
    get_available_users_endpoint,
    get_financial_radar,
    update_rules,
    resolve_exception_endpoint
)

def make_mock_request(username: str = None) -> Request:
    """Belirtilen kullanıcı rolüyle sahte bir HTTP Request nesnesi üretir."""
    headers = []
    if username:
        headers.append((b"x-user-role", username.encode("utf-8")))
    scope = {
        "type": "http",
        "method": "GET",
        "path": "/",
        "headers": headers
    }
    return Request(scope)

def test_rbac_roles_defined():
    """Tüm fabrika rolleri (İK, Muhasebe, Fabrika Müdürü, Admin) eksiksiz tanımlanmalıdır."""
    assert "ik_yonetici" in USERS
    assert "muhasebe" in USERS
    assert "genel_mudur" in USERS
    assert "admin" in USERS

    assert USERS["ik_yonetici"]["role"] == "hr"
    assert USERS["muhasebe"]["role"] == "accounting"
    assert USERS["genel_mudur"]["role"] == "plant_manager"
    assert USERS["admin"]["role"] == "admin"

def test_auth_login_success():
    """Geçerli kullanıcı adı ve şifre ile başarılı giriş yapılabilmelidir."""
    res = login_endpoint(LoginRequest(username="ik_yonetici", password="123"))
    assert res["status"] == "success"
    assert res["user"]["role"] == "hr"
    assert "Ayşe Yılmaz" in res["user"]["name"]

def test_auth_login_invalid():
    """Hatalı şifre veya bilinmeyen kullanıcı için 401 Unauthorized dönmelidir."""
    with pytest.raises(HTTPException) as exc1:
        login_endpoint(LoginRequest(username="ik_yonetici", password="yanlis_sifre"))
    assert exc1.value.status_code == 401

    with pytest.raises(HTTPException) as exc2:
        login_endpoint(LoginRequest(username="bilinmeyen_kullanici", password="123"))
    assert exc2.value.status_code == 401

def test_auth_me_endpoint():
    """X-User-Role başlığı ile aktif profil doğru çözümlenmelidir."""
    req_hr = make_mock_request("ik_yonetici")
    user_hr = get_current_user_endpoint(req_hr)
    assert user_hr["role"] == "hr"

    req_acc = make_mock_request("muhasebe")
    user_acc = get_current_user_endpoint(req_acc)
    assert user_acc["role"] == "accounting"

def test_hr_permission_enforcement():
    """İnsan Kaynakları: İstisna onaylayabilir, fakat Finans Radarı ve Kural Güncellemeden 403 almalıdır."""
    req_hr = make_mock_request("ik_yonetici")

    # Finans radarına erişememeli (403)
    with pytest.raises(HTTPException) as exc_fin:
        get_financial_radar(request=req_hr)
    assert exc_fin.value.status_code == 403

    # Kural düzenleyememeli (403)
    with pytest.raises(HTTPException) as exc_rule:
        update_rules(payload=RuleUpdateRequest(rules={}), request=req_hr)
    assert exc_rule.value.status_code == 403

def test_accounting_permission_enforcement():
    """Muhasebe: Finans Radarını görebilir, fakat İstisna Onaylama ve Kural Güncellemeden 403 almalıdır."""
    req_acc = make_mock_request("muhasebe")

    # Finans radarını başarıyla görebilmeli
    fin_res = get_financial_radar(request=req_acc, page=1, page_size=10)
    assert "summary" in fin_res
    assert len(fin_res["items"]) > 0

    # İstisna onaylayamamalı (403)
    with pytest.raises(HTTPException) as exc_res:
        resolve_exception_endpoint(
            payload=ExceptionResolveRequest(name_key="TEST", day=1, approved_hours=7.5, note="test"),
            request=req_acc
        )
    assert exc_res.value.status_code == 403

def test_plant_manager_permission_enforcement():
    """Fabrika Müdürü: Finans ve KPI'ları görebilir, operasyonel istisna onaylama yetkisi yoktur (403)."""
    req_mgr = make_mock_request("genel_mudur")

    # Finans radarını başarıyla görebilmeli
    fin_res = get_financial_radar(request=req_mgr, page=1, page_size=10)
    assert "summary" in fin_res

    # İstisna onaylayamamalı (403)
    with pytest.raises(HTTPException) as exc_res:
        resolve_exception_endpoint(
            payload=ExceptionResolveRequest(name_key="TEST", day=1, approved_hours=7.5, note="test"),
            request=req_mgr
        )
    assert exc_res.value.status_code == 403

def test_admin_unrestricted_access():
    """Admin: Tüm modüllere, finansa, istisna onayına ve kural ayarlarına kısıtlamasız erişir."""
    req_admin = make_mock_request("admin")
    fin_res = get_financial_radar(request=req_admin, page=1, page_size=5)
    assert "summary" in fin_res
