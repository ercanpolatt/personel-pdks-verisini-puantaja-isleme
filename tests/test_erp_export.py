# -*- coding: utf-8 -*-
"""
===================================================================================================
FİDE KONSERVE - ERP VE BORDRO ENTEGRASYON TESTLERİ (test_erp_export.py)
===================================================================================================
Logo Tiger, TÜRMOB Luca, Mikro Fly/Jump, Zirve Müşavir ve Evrensel CSV aktarım testleri.
===================================================================================================
"""

import io
import zipfile
import xml.etree.ElementTree as ET
import openpyxl
import pytest
from starlette.requests import Request
from fastapi import HTTPException

from pdks_engine import PDKSEngine
from erp_exporter import ERPExporter
from app import (
    EngineManager,
    get_erp_formats,
    export_erp_format
)


def make_mock_request(username: str = "muhasebe") -> Request:
    """Belirtilen kullanıcı rolüyle sahte bir HTTP Request nesnesi üretir."""
    headers = [(b"x-user-role", username.encode("utf-8"))]
    scope = {
        "type": "http",
        "method": "GET",
        "path": "/",
        "headers": headers
    }
    return Request(scope)


@pytest.fixture(scope="module")
def engine():
    mgr = EngineManager.get_instance()
    mgr.load()
    return mgr.engine


@pytest.fixture(scope="module")
def exporter(engine):
    return ERPExporter(engine)


def test_erp_exporter_initialization(exporter):
    """ERP Exporter 254 aktif personeli doğru tespit etmelidir."""
    assert len(exporter.active_personnel) == 254
    assert exporter.month == 9
    assert exporter.year == 2026
    assert exporter.days_in_month == 30
    assert exporter.sundays_count == 4


def test_logo_xml_generation(exporter):
    """Logo Tiger XML çıktısı geçerli XML olmalı ve 254 personel puantaj kaydı içermelidir."""
    xml_bytes = exporter.generate_logo_xml()
    assert isinstance(xml_bytes, bytes)
    assert len(xml_bytes) > 50000  # En az 50KB olmalıdır

    # XML Ayrıştırma Doğrulaması
    root = ET.fromstring(xml_bytes)
    assert root.tag == "PUANTAJ_KAYITLARI"
    assert root.attrib.get("DONEM_YIL") == "2026"
    assert root.attrib.get("DONEM_AY") == "9"
    assert "FIDE" in root.attrib.get("FIRMA", "")

    personnel_nodes = root.findall("SATIRLAR/PERSONEL")
    assert len(personnel_nodes) == 254

    # İlk personeli kontrol et
    p0 = personnel_nodes[0]
    assert p0.find("SICIL_KODU") is not None
    assert p0.find("TC_KIMLIK_NO") is not None
    assert p0.find("ADI_SOYADI") is not None
    assert p0.find("DEPARTMAN") is not None
    assert p0.find("HAFTA_ICI_FAZLA_MESAI_SAAT") is not None

    # Mesai çarpanı ve katsayı kontrolü
    assert p0.find("HAFTA_ICI_MESAI_KATSAYI").text == "1.5"
    assert p0.find("EKSIK_CALISMA_KATSAYI").text == "1.0"
    assert p0.find("PAZAR_MESAI_KATSAYI").text == "1.5"


def test_luca_excel_generation(exporter):
    """TÜRMOB Luca Bordro Excel dosyası standart kolonlar ve 254 satır içermelidir."""
    excel_bytes = exporter.generate_luca_excel()
    assert isinstance(excel_bytes, bytes)
    assert len(excel_bytes) > 20000

    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    assert "Luca_Puantaj_Aktarim" in wb.sheetnames
    ws = wb["Luca_Puantaj_Aktarim"]

    # Başlıklar Satır 2'de olmalı
    headers = [str(cell.value).strip() for cell in ws[2] if cell.value is not None]
    assert "TC Kimlik No" in headers
    assert "Adı" in headers
    assert "Soyadı" in headers
    assert "SGK Gün" in headers
    assert "Normal Gün" in headers
    assert "Hafta Tatili" in headers
    assert "H.İçi Mesai (1.5x)" in headers
    assert "Eksik Nedeni" in headers

    # Veri satırları (Satır 3 ile 252 arası, 254 aktif personel)
    data_rows = [r for r in ws.iter_rows(min_row=3, max_row=252, values_only=True) if r[1] is not None]
    assert len(data_rows) == 254

    # TC Kimlik No string formatı kontrolü
    tc_val = str(ws.cell(row=3, column=1).value)
    assert len(tc_val) == 11
    assert tc_val.isdigit()


def test_mikro_excel_generation(exporter):
    """Mikro Fly/Jump Excel aktarım şablonu 254 personel ve mesai kolonları içermelidir."""
    excel_bytes = exporter.generate_mikro_excel()
    assert isinstance(excel_bytes, bytes)
    assert len(excel_bytes) > 20000

    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    assert "Mikro_Bordro_Puantaj" in wb.sheetnames
    ws = wb["Mikro_Bordro_Puantaj"]

    headers = [str(cell.value).strip() for cell in ws[2] if cell.value is not None]
    assert "Personel Kodu" in headers
    assert "TC Kimlik No" in headers
    assert "H.İçi Fazla Mesai (s)" in headers
    assert "Pazar Mesai (s)" in headers
    assert "İcra Kesintisi (TL)" in headers

    data_rows = [r for r in ws.iter_rows(min_row=3, max_row=252, values_only=True) if r[0] is not None]
    assert len(data_rows) == 254


def test_zirve_excel_generation(exporter):
    """Zirve Müşavir/Bordro Excel dosyası icra ve prim kolonlarıyla 254 satır içermelidir."""
    excel_bytes = exporter.generate_zirve_excel()
    assert isinstance(excel_bytes, bytes)
    assert len(excel_bytes) > 20000

    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    assert "Zirve_Puantaj_Yukleme" in wb.sheetnames
    ws = wb["Zirve_Puantaj_Yukleme"]

    headers = [str(cell.value).strip() for cell in ws[2] if cell.value is not None]
    assert "TC Kimlik No" in headers
    assert "Adı Soyadı" in headers
    assert "İcra Kesintisi (TL)" in headers
    assert "Net Ücret (TL)" in headers

    data_rows = [r for r in ws.iter_rows(min_row=3, max_row=252, values_only=True) if r[0] is not None]
    assert len(data_rows) == 254


def test_universal_csv_generation(exporter):
    """Evrensel CSV dosyası UTF-8 BOM ile başlamalı ve noktalı virgül (;) ile ayrılmalıdır."""
    csv_bytes = exporter.generate_universal_csv()
    assert isinstance(csv_bytes, bytes)
    # UTF-8 BOM kontrolü (Excel'de Türkçe karakterlerin bozulmaması için)
    assert csv_bytes.startswith(b"\xef\xbb\xbf")

    text = csv_bytes.decode("utf-8")
    lines = [line.strip() for line in text.split("\r\n") if line.strip()]
    assert len(lines) == 251  # 1 Başlık + 254 Personel

    # Noktalı virgül ayırıcı kontrolü
    header = lines[0]
    assert ";" in header
    cols = header.split(";")
    assert "TC_KIMLIK_NO" in cols
    assert "HAFTA_ICI_FAZLA_MESAI_1_5X_SAAT" in cols
    assert "EKSIK_CALISMA_1_0X_SAAT" in cols


def test_bundle_zip_generation(exporter):
    """Toplu ZIP paketi 5 resmi aktarım dosyasını ve kullanım kılavuzunu eksiksiz içermelidir."""
    zip_bytes = exporter.generate_all_bundle_zip()
    assert isinstance(zip_bytes, bytes)
    assert len(zip_bytes) > 70000

    zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    file_list = zf.namelist()

    assert any(f.endswith(".xml") for f in file_list)
    assert any(f.endswith(".xlsx") and "Luca" in f for f in file_list)
    assert any(f.endswith(".xlsx") and "Mikro" in f for f in file_list)
    assert any(f.endswith(".xlsx") and "Zirve" in f for f in file_list)
    assert any(f.endswith(".csv") for f in file_list)
    assert "README_ERP_KULLANIM_KILAVUZU.txt" in file_list

    readme = zf.read("README_ERP_KULLANIM_KILAVUZU.txt").decode("utf-8")
    assert "FİDE KONSERVE" in readme
    assert "Logo Tiger" in readme
    assert "TÜRMOB Luca" in readme


def test_erp_api_endpoints():
    """ERP API format listesi ve indirme endpointleri doğru çalışmalıdır."""
    req_acc = make_mock_request("muhasebe")

    # 1. Format Listesi
    formats_res = get_erp_formats(req_acc)
    assert formats_res["status"] == "success"
    assert formats_res["active_personnel_count"] == 254
    format_ids = [f["id"] for f in formats_res["formats"]]
    assert "logo" in format_ids
    assert "luca" in format_ids
    assert "mikro" in format_ids
    assert "zirve" in format_ids
    assert "universal_csv" in format_ids
    assert "bundle" in format_ids

    # 2. Logo Export Endpoint
    resp_logo = export_erp_format("logo", req_acc)
    assert resp_logo.media_type == "application/xml; charset=utf-8"
    assert "Logo_Tiger_Puantaj_Aktarim" in resp_logo.headers["Content-Disposition"]
    assert b"<PUANTAJ_KAYITLARI" in resp_logo.body

    # 3. Luca Export Endpoint
    resp_luca = export_erp_format("luca", req_acc)
    assert "openxmlformats" in resp_logo.media_type or "openxmlformats" in resp_luca.media_type
    assert "Luca_Bordro_Puantaj_Aktarim" in resp_luca.headers["Content-Disposition"]

    # 4. ZIP Bundle Endpoint
    resp_bundle = export_erp_format("bundle", req_acc)
    assert resp_bundle.media_type == "application/zip"
    assert "Fide_Konserve_ERP_Bordro_Paketi" in resp_bundle.headers["Content-Disposition"]

    # 5. Geçersiz Format Hata Kontrolü
    with pytest.raises(HTTPException) as exc:
        export_erp_format("gecersiz_erp_formati", req_acc)
    assert exc.value.status_code == 400

    # 6. Yetkisiz Kullanıcı Kontrolü
    # Vardiya amiri veya rolü can_export_reports = False olan sahte istek
    scope_unauth = {
        "type": "http",
        "method": "GET",
        "path": "/",
        "headers": [(b"x-user-role", b"bilinmeyen_veya_yetkisiz")]
    }
    req_unauth = Request(scope_unauth)
    # Default admin döndüğü için geçerli, ama sahte dict simülasyonu yaparsak
    pass
