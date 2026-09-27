# -*- coding: utf-8 -*-
"""
===================================================================================================
FİDE KONSERVE - PUANTAJ & PDKS YÖNETİM SİSTEMİ WEB SERVİSİ (FastAPI)
===================================================================================================
Kurumsal PDKS, Vardiya, Mesai ve Puantaj Takip Web API'si ve Arayüz Sunucusu.
"""

import os
import json
import shutil
from typing import Optional, Dict, Any, List
from datetime import datetime

from fastapi import FastAPI, HTTPException, UploadFile, File, Query, Body, Request
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

from pdks_engine import PDKSEngine, fmt_hours_tr, norm_name_key, clean_display_text
from slip_builder import generate_single_slip_html, render_slips_document
from erp_exporter import ERPExporter

def _clean_str(val: Any) -> Optional[str]:
    if val is None or not isinstance(val, str):
        return None
    val = val.strip()
    return val if val else None

def _matches_search(haystack: str, query: str) -> bool:
    if not query or not haystack:
        return False
    return norm_name_key(query) in norm_name_key(haystack)
from builder import build_puantaj_workbook
from export_daily_hours import build_excel_report, export_csv_and_json
from export_simple_pdks import generate_simple_report
from config.settings import load_rules, save_rules, RULES_FILE_PATH

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# SABİTLER VE DİZİN YAPISI
# -----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
DATA_DIR = os.path.join(BASE_DIR, "data")
RESOLVED_EXCEPTIONS_FILE = os.path.join(DATA_DIR, "resolved_exceptions.json")

def generate_financial_excel(fin_data: Dict[str, Any], output_path: str = "maas_ve_maliyet_radari.xlsx") -> str:
    """Yönetim ve Finans için renkli ve formüllü Maaş & Mesai Maliyet Raporu üretir."""
    wb = Workbook()
    
    font_title = Font(name="Arial", size=13, bold=True, color="FFFFFF")
    font_head = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Arial", size=9, bold=True)
    font_data = Font(name="Arial", size=9)
    fill_navy = PatternFill("solid", fgColor="1F4E79")
    fill_total = PatternFill("solid", fgColor="E7E6E6")
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    border_total = Border(
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )
    
    # 1. Departman Maliyet Özeti Sayfası
    ws_dept = wb.active
    ws_dept.title = "Departman_Maliyet_Ozeti"
    ws_dept.views.sheetView[0].showGridLines = True
    
    ws_dept.merge_cells("A1:J1")
    ws_dept["A1"] = "FİDE KONSERVE — DEPARTMAN İŞÇİLİK VE FAZLA MESAİ BÜTÇE ANALİZİ"
    ws_dept["A1"].font = font_title
    ws_dept["A1"].fill = fill_navy
    ws_dept["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_dept.row_dimensions[1].height = 28
    
    headers_dept = [
        ("Departman", 22), ("Aktif Personel", 14), ("Toplam Çalışma (Saat)", 20),
        ("Toplam FM (Saat)", 18), ("FM Bütçe Payı (%)", 18), ("Toplam FM Maliyeti (TL)", 22),
        ("Normal Maaş Maliyeti (TL)", 24), ("Toplam İşçilik Maliyeti (TL)", 26),
        ("Kişi Başı Ortalama FM (Saat)", 24), ("Yönetim Analiz Notu", 45)
    ]
    
    ws_dept.row_dimensions[3].height = 24
    for c_idx, (h_text, w) in enumerate(headers_dept, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_dept[f"{col_let}3"]
        cell.value = h_text
        cell.font = font_head
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws_dept.column_dimensions[col_let].width = w
        
    for r_idx, d in enumerate(fin_data["departments"], start=4):
        ws_dept[f"A{r_idx}"] = d["department"]
        ws_dept[f"B{r_idx}"] = d["active_count"]
        ws_dept[f"C{r_idx}"] = d["total_hours"]
        ws_dept[f"D{r_idx}"] = d["ot_hours"]
        ws_dept[f"E{r_idx}"] = d["ot_share_pct"] / 100.0
        ws_dept[f"F{r_idx}"] = d["ot_cost"]
        ws_dept[f"G{r_idx}"] = d["base_cost"]
        ws_dept[f"H{r_idx}"] = d["total_cost"]
        ws_dept[f"I{r_idx}"] = d["avg_ot_hours_per_worker"]
        ws_dept[f"J{r_idx}"] = d["insight"]
        
        ws_dept[f"A{r_idx}"].font = font_bold
        for c in range(1, 11):
            cl = get_column_letter(c)
            cell = ws_dept[f"{cl}{r_idx}"]
            cell.border = border_thin
            if c not in (1, 10):
                cell.font = font_data
            if c == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c == 2:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in (3, 4, 9):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.0"
            elif c == 5:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "0.0%"
            elif c in (6, 7, 8):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.00 TL"
            elif c == 10:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.font = Font(name="Arial", size=8, italic=True, color="595959")
                
    tot_row = len(fin_data["departments"]) + 4
    ws_dept[f"A{tot_row}"] = "GENEL TOPLAM"
    ws_dept[f"B{tot_row}"] = f"=SUM(B4:B{tot_row-1})"
    ws_dept[f"C{tot_row}"] = f"=SUM(C4:C{tot_row-1})"
    ws_dept[f"D{tot_row}"] = f"=SUM(D4:D{tot_row-1})"
    ws_dept[f"E{tot_row}"] = f"=SUM(E4:E{tot_row-1})"
    ws_dept[f"F{tot_row}"] = f"=SUM(F4:F{tot_row-1})"
    ws_dept[f"G{tot_row}"] = f"=SUM(G4:G{tot_row-1})"
    ws_dept[f"H{tot_row}"] = f"=SUM(H4:H{tot_row-1})"
    ws_dept[f"I{tot_row}"] = f"=D{tot_row}/B{tot_row}"
    ws_dept[f"J{tot_row}"] = "Fabrika Bütçe Toplamı"
    
    for c in range(1, 11):
        cl = get_column_letter(c)
        cell = ws_dept[f"{cl}{tot_row}"]
        cell.font = font_bold
        cell.fill = fill_total
        cell.border = border_total
        if c in (6, 7, 8):
            cell.number_format = "#,##0.00 TL"
        elif c in (3, 4, 9):
            cell.number_format = "#,##0.0"
        elif c == 5:
            cell.number_format = "0.0%"

    # 2. Personel Detay Sayfası
    ws_p = wb.create_sheet(title="Personel_Maas_ve_Mesai_Detayi")
    ws_p.views.sheetView[0].showGridLines = True
    
    headers_p = [
        ("Sıra", 6), ("TC Kimlik No", 14), ("Adı Soyadı", 24), ("Bölüm", 20),
        ("Çalışılan Gün", 13), ("Net Maaş (TL)", 14), ("Saatlik Ücret (TL)", 15),
        ("Saatlik Mesai Ücreti (%50 Zam)", 18), ("Hafta İçi FM (Saat)", 16),
        ("Eksik Çalışma (Saat - 1.0x)", 18), ("Pazar FM (Saat)", 14),
        ("H.İçi Net Mesai Saati", 16), ("FM Tutarı (TL)", 16),
        ("Eksik Çalışma Kesintisi (TL)", 18),
        ("Bölüm Primi (Saat)", 15), ("Bölüm Primi Tutarı (TL)", 18),
        ("Normal Maaş Hakedişi (TL)", 18), ("Toplam Net Hakediş (TL)", 18),
        ("SGK Banka Neti (TL)", 16), ("Elden Ödenecek Net Fark (TL)", 18)
    ]
    
    ws_p.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(headers_p, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_p[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_head
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws_p.column_dimensions[col_let].width = w
        
    for r_idx, p in enumerate(fin_data["personnel"], start=2):
        ws_p[f"A{r_idx}"] = p["sno"]
        ws_p[f"B{r_idx}"] = p["tc"]
        ws_p[f"C{r_idx}"] = p["ad_soyad"]
        ws_p[f"D{r_idx}"] = p["bolum"]
        ws_p[f"E{r_idx}"] = p["work_days"]
        ws_p[f"F{r_idx}"] = p["net_maas"]
        ws_p[f"G{r_idx}"] = p["hourly_net"]
        ws_p[f"H{r_idx}"] = p["ot_hourly_net"]
        ws_p[f"I{r_idx}"] = p.get("weekday_ot_hours", 0.0)
        ws_p[f"J{r_idx}"] = p.get("missing_hours", 0.0)
        ws_p[f"K{r_idx}"] = p.get("sunday_ot_hours", 0.0)
        ws_p[f"L{r_idx}"] = p.get("net_weekday_ot_hours", 0.0)
        ws_p[f"M{r_idx}"] = p["ot_cost"]
        ws_p[f"N{r_idx}"] = p.get("missing_deduction_cost", 0.0)
        ws_p[f"O{r_idx}"] = p["bonus_hours"]
        ws_p[f"P{r_idx}"] = p["bonus_cost"]
        ws_p[f"Q{r_idx}"] = p["base_cost"]
        ws_p[f"R{r_idx}"] = p["total_net_wage"]
        ws_p[f"S{r_idx}"] = p["bank_net"]
        ws_p[f"T{r_idx}"] = p["cash_diff"]
        
        for c in range(1, 21):
            cl = get_column_letter(c)
            cell = ws_p[f"{cl}{r_idx}"]
            cell.border = border_thin
            cell.font = font_data
            if c in (6, 7, 8, 13, 14, 16, 17, 18, 19, 20):
                cell.number_format = "#,##0.00 TL"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in (9, 10, 11, 12, 15):
                cell.number_format = "#,##0.0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in (1, 2, 5):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
    wb.save(output_path)
    return output_path

def generate_compliance_excel(comp_data: Dict[str, Any], output_path: str):
    """3 Sayfalı Resmi İş Teftişi ve SGK Yasal Uyum Denetim Çalışma Kitabı."""
    wb = Workbook()
    
    font_title = Font(name="Arial", size=13, bold=True, color="1E3A8A")
    font_head = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Arial", size=9, bold=True)
    font_data = Font(name="Arial", size=9)
    font_critical = Font(name="Arial", size=9, bold=True, color="991B1B")
    font_warning = Font(name="Arial", size=9, bold=True, color="92400E")
    
    fill_navy = PatternFill("solid", fgColor="1E293B")
    fill_crimson = PatternFill("solid", fgColor="FEE2E2")
    fill_amber = PatternFill("solid", fgColor="FEF3C7")
    
    border_thin = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    
    # -------------------------------------------------------------------------
    # SAYFA 1: UYUM VE RİSK ÖZETİ
    # -------------------------------------------------------------------------
    ws_sum = wb.active
    ws_sum.title = "Yasal_Uyum_ve_Risk_Ozeti"
    ws_sum.views.sheetView[0].showGridLines = True
    
    ws_sum.merge_cells("A1:H1")
    ws_sum["A1"] = "FİDE KONSERVE — 4857 SAYILI İŞ KANUNU VE SGK YASAL UYUM & RİSK ANALİZİ (EYLÜL 2026)"
    ws_sum["A1"].font = font_title
    ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 28
    
    s = comp_data["summary"]
    kpis = [
        ("Fabrika Yasal Uyum Endeksi", f"%{s['compliance_index']:.1f}", f"Risk Seviyesi: {s['risk_status']}"),
        ("İncelenen Aktif Çalışan", f"{s['inspected_personnel']} Personel", "PDKS kart basan tüm çalışanlar"),
        ("Toplam Yasal İhlal Sayısı", f"{s['total_violations']} Olay", f"Kritik: {s['critical_violations']}, Uyarı: {s['warning_violations']}"),
        ("11 Saat Dinlenme İhlalleri (Md. 68)", f"{s['rest_violations_count']} Olay", "Postalar Halinde Çalışma Yön. Md. 9"),
        ("7+ Gün Kesintisiz Çalışma (Md. 46)", f"{s['consecutive_work_count']} Dönem", "Hafta tatili verilmeyen çalışma blokları"),
        ("270 Saat Aşım Riski (Md. 41)", f"{s['overtime_limit_count']} Personel", "Aylık 35 saat ve üzeri fazla mesai yapanlar"),
        ("En Yüksek Riskli Departman", f"{s['highest_risk_department']}", "İhlal ve aşırı mesai yoğunluğu")
    ]
    
    ws_sum["A3"] = "YASAL DENETİM VE RİSK GÖSTERGELERİ (KPI)"
    ws_sum["A3"].font = Font(name="Arial", size=10, bold=True, color="0F172A")
    
    for idx, (label, val, desc) in enumerate(kpis, start=4):
        ws_sum[f"A{idx}"] = label
        ws_sum[f"B{idx}"] = val
        ws_sum[f"C{idx}"] = desc
        ws_sum[f"A{idx}"].font = font_bold
        ws_sum[f"B{idx}"].font = Font(name="Arial", size=10, bold=True, color="1E3A8A")
        ws_sum[f"C{idx}"].font = Font(name="Arial", size=9, italic=True, color="64748B")
        for col in ("A", "B", "C"):
            ws_sum[f"{col}{idx}"].border = border_thin
            
    ws_sum.column_dimensions["A"].width = 38
    ws_sum.column_dimensions["B"].width = 24
    ws_sum.column_dimensions["C"].width = 45
    
    start_r = 13
    ws_sum[f"A{start_r}"] = "DEPARTMAN BAZLI YASAL UYUM VE RİSK MATRİSİ"
    ws_sum[f"A{start_r}"].font = Font(name="Arial", size=10, bold=True, color="0F172A")
    
    dept_heads = [
        ("Departman", 22), ("Aktif Kadro", 13), ("11s Dinlenme", 14), ("Hafta Tatilsiz", 14),
        ("270s FM Riski", 14), ("Toplam Olay", 13), ("Uyum Skoru (%)", 15), ("Risk Seviyesi", 15)
    ]
    
    ws_sum.row_dimensions[start_r + 1].height = 24
    for c_idx, (h_text, w) in enumerate(dept_heads, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_sum[f"{col_let}{start_r + 1}"]
        cell.value = h_text
        cell.font = font_head
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws_sum.column_dimensions[col_let].width = w
        
    for r_idx, d in enumerate(comp_data["department_compliance"], start=start_r + 2):
        ws_sum[f"A{r_idx}"] = d["department"]
        ws_sum[f"B{r_idx}"] = d["active_headcount"]
        ws_sum[f"C{r_idx}"] = d["rest_count"]
        ws_sum[f"D{r_idx}"] = d["consecutive_count"]
        ws_sum[f"E{r_idx}"] = d["ot_risk_count"]
        ws_sum[f"F{r_idx}"] = d["total_violations"]
        ws_sum[f"G{r_idx}"] = d["compliance_score"] / 100.0
        ws_sum[f"H{r_idx}"] = d["risk_level"]
        
        ws_sum[f"A{r_idx}"].font = font_bold
        for c in range(1, 9):
            cl = get_column_letter(c)
            cell = ws_sum[f"{cl}{r_idx}"]
            cell.border = border_thin
            if c not in (1, 8):
                cell.font = font_data
            if c == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c in (2, 3, 4, 5, 6):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c == 7:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "0.0%"
            elif c == 8:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = font_critical if d["risk_level"] == "YÜKSEK" else (font_warning if d["risk_level"] == "ORTA" else font_bold)
                cell.fill = fill_crimson if d["risk_level"] == "YÜKSEK" else (fill_amber if d["risk_level"] == "ORTA" else PatternFill(fill_type=None))
                
    # -------------------------------------------------------------------------
    # SAYFA 2: 11 SAAT DİNLENME İHLALLERİ (MADDE 68)
    # -------------------------------------------------------------------------
    ws_rest = wb.create_sheet(title="11_Saat_Dinlenme_Ihlalleri")
    ws_rest.views.sheetView[0].showGridLines = True
    
    rest_heads = [
        ("Sıra", 6), ("TC Kimlik No", 14), ("Adı Soyadı", 24), ("Bölüm", 20),
        ("Vardiya Geçiş Tarihi", 18), ("Çıkış Saati", 13), ("İşbaşı Saati", 13),
        ("Dinlenme Süresi", 16), ("Yasal Limit", 14), ("Ciddiyet", 12),
        ("Yasal Dayanak", 32), ("Denetim Açıklaması", 48)
    ]
    
    ws_rest.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(rest_heads, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_rest[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_head
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws_rest.column_dimensions[col_let].width = w
        
    for r_idx, v in enumerate(comp_data["rest_violations"], start=2):
        ws_rest[f"A{r_idx}"] = r_idx - 1
        ws_rest[f"B{r_idx}"] = v["tc"]
        ws_rest[f"C{r_idx}"] = v["ad_soyad"]
        ws_rest[f"D{r_idx}"] = v["bolum"]
        ws_rest[f"E{r_idx}"] = v["date_str"]
        ws_rest[f"F{r_idx}"] = v["c1"]
        ws_rest[f"G{r_idx}"] = v["g2"]
        ws_rest[f"H{r_idx}"] = v["metric_value"]
        ws_rest[f"I{r_idx}"] = v["legal_limit"]
        ws_rest[f"J{r_idx}"] = "KRİTİK" if v["severity"] == "CRITICAL" else "UYARI"
        ws_rest[f"K{r_idx}"] = v["legal_article"]
        ws_rest[f"L{r_idx}"] = v["detail"]
        
        ws_rest[f"C{r_idx}"].font = font_bold
        for c in range(1, 13):
            cl = get_column_letter(c)
            cell = ws_rest[f"{cl}{r_idx}"]
            cell.border = border_thin
            if c not in (3, 10):
                cell.font = font_data
            if c in (1, 5, 6, 7, 8, 9):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c == 10:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = font_critical if v["severity"] == "CRITICAL" else font_warning
                cell.fill = fill_crimson if v["severity"] == "CRITICAL" else fill_amber
                
    # -------------------------------------------------------------------------
    # SAYFA 3: HAFTA TATİLİ VE 270 SAAT MESAİ RİSKLERİ
    # -------------------------------------------------------------------------
    ws_other = wb.create_sheet(title="Hafta_Tatili_ve_270s_Riskleri")
    ws_other.views.sheetView[0].showGridLines = True
    
    other_heads = [
        ("Sıra", 6), ("TC Kimlik No", 14), ("Adı Soyadı", 24), ("Bölüm", 20),
        ("İhlal / Risk Türü", 26), ("Dönem / Tarih", 18), ("Tespit Edilen Değer", 18),
        ("Yasal Sınır", 22), ("Risk Derecesi", 13), ("Yasal Dayanak", 32),
        ("Denetim ve İdari Risk Notu", 48)
    ]
    
    ws_other.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(other_heads, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_other[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_head
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws_other.column_dimensions[col_let].width = w
        
    other_list = comp_data["consecutive_violations"] + comp_data["overtime_risks"]
    for r_idx, v in enumerate(other_list, start=2):
        ws_other[f"A{r_idx}"] = r_idx - 1
        ws_other[f"B{r_idx}"] = v["tc"]
        ws_other[f"C{r_idx}"] = v["ad_soyad"]
        ws_other[f"D{r_idx}"] = v["bolum"]
        ws_other[f"E{r_idx}"] = v["violation_title"]
        ws_other[f"F{r_idx}"] = v["date_str"]
        ws_other[f"G{r_idx}"] = v["metric_value"]
        ws_other[f"H{r_idx}"] = v["legal_limit"]
        ws_other[f"I{r_idx}"] = "KRİTİK" if v["severity"] == "CRITICAL" else "UYARI"
        ws_other[f"J{r_idx}"] = v["legal_article"]
        ws_other[f"K{r_idx}"] = v["detail"]
        
        ws_other[f"C{r_idx}"].font = font_bold
        for c in range(1, 12):
            cl = get_column_letter(c)
            cell = ws_other[f"{cl}{r_idx}"]
            cell.border = border_thin
            if c not in (3, 9):
                cell.font = font_data
            if c in (1, 6, 7, 8):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c == 9:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = font_critical if v["severity"] == "CRITICAL" else font_warning
                cell.fill = fill_crimson if v["severity"] == "CRITICAL" else fill_amber

    wb.save(output_path)
    return output_path

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

# -----------------------------------------------------------------------------
# İSTİSNA ÇÖZÜM VERİ YÖNETİMİ
# -----------------------------------------------------------------------------
def load_resolved_exceptions() -> Dict[str, Any]:
    if not os.path.exists(RESOLVED_EXCEPTIONS_FILE):
        return {}
    try:
        with open(RESOLVED_EXCEPTIONS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_resolved_exceptions(data: Dict[str, Any]):
    with open(RESOLVED_EXCEPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# -----------------------------------------------------------------------------
# ENGINE YÖNETİCİSİ (SINGLETON CACHE)
# -----------------------------------------------------------------------------
class EngineManager:
    _instance: Optional["EngineManager"] = None

    def __init__(self, pdks_path="pdks.xls", puantaj_path="puantaj.xls", target_month=9, target_year=2026):
        self.pdks_path = pdks_path
        self.puantaj_path = puantaj_path
        self.target_month = target_month
        self.target_year = target_year
        self.engine: Optional[PDKSEngine] = None
        self.matrix: List[Dict[str, Any]] = []
        self.resolved_exceptions: Dict[str, Any] = load_resolved_exceptions()
        self.load()

    @classmethod
    def get_instance(cls) -> "EngineManager":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def load(self):
        """PDKS ve Puantaj motorunu belleğe yükler ve analiz eder."""
        print(f"[EngineManager] PDKS Motoru yükleniyor ({self.target_month:02d}.{self.target_year})...")
        self.engine = PDKSEngine(
            pdks_path=self.pdks_path,
            puantaj_path=self.puantaj_path,
            target_month=self.target_month,
            target_year=self.target_year
        )
        self.engine.load_personnel()
        self.engine.load_and_process_pdks()
        self.apply_resolved_exceptions()
        self.matrix = self.engine.get_summary_matrix()
        print(f"[EngineManager] Başarıyla yüklendi: {len(self.engine.personnel_list)} personel, {len(self.engine.audit_records)} hareket.")

    def apply_resolved_exceptions(self):
        """Kayıtlı amir onaylarını motor sonuçlarına uygular."""
        self.resolved_exceptions = load_resolved_exceptions()
        for key, res in self.resolved_exceptions.items():
            # key: "name_key:day"
            if ":" not in key:
                continue
            name_k, day_str = key.split(":", 1)
            try:
                day_num = int(day_str)
            except ValueError:
                continue
            rec = self.engine.daily_results.get((name_k, day_num))
            if rec:
                rec["is_resolved"] = True
                rec["resolved_by"] = res.get("resolved_by", "Sistem / Amir")
                rec["resolved_note"] = res.get("note", "")
                rec["resolved_at"] = res.get("timestamp", "")
                rec["approved_hours"] = res.get("approved_hours", rec["final_hours"])
                # Eğer amir saat onayladıysa final_hours'ı güncelle (Maks. 24 saat kuralı)
                if "approved_hours" in res and res["approved_hours"] is not None:
                    clean_h = min(24.0, max(0.0, float(res["approved_hours"])))
                    rec["approved_hours"] = clean_h
                    rec["final_hours"] = clean_h
                    if clean_h <= 7.5:
                        rec["base_hours"] = clean_h
                        rec["overtime_hours"] = 0.0
                    else:
                        rec["base_hours"] = 7.5
                        rec["overtime_hours"] = round(clean_h - 7.5, 2)

    def resolve_exception(self, name_key: str, day: int, approved_hours: float, note: str, resolved_by: str = "Vardiya Amiri"):
        approved_hours = float(approved_hours)
        if approved_hours < 0.0 or approved_hours > 24.0:
            raise ValueError("Bir günde 24 saatten fazla çalışma süresi yazılamaz! (Maksimum 24.0 saat)")
        approved_hours = min(24.0, max(0.0, approved_hours))
        
        key = f"{name_key}:{day}"
        self.resolved_exceptions[key] = {
            "name_key": name_key,
            "day": day,
            "approved_hours": approved_hours,
            "note": note,
            "resolved_by": resolved_by,
            "timestamp": datetime.now().strftime("%d.%m.%Y %H:%M")
        }
        save_resolved_exceptions(self.resolved_exceptions)
        self.apply_resolved_exceptions()
        self.matrix = self.engine.get_summary_matrix()
        return self.resolved_exceptions[key]

    def bulk_resolve_exceptions(self, items: List[Dict[str, Any]], resolved_by: str = "Vardiya Amiri / Toplu Onay") -> int:
        """Birden çok istisna kaydını tek işlemde amir onayına alır ve puantajı günceller (Maks. 24s korumalı)."""
        now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
        count = 0
        for item in items:
            name_key = item.get("name_key")
            day = item.get("day")
            if not name_key or day is None:
                continue
            key = f"{name_key}:{day}"
            approved_hours = float(item.get("approved_hours", 7.5))
            if approved_hours < 0.0 or approved_hours > 24.0:
                raise ValueError(f"Geçersiz süre ({approved_hours}s): Bir günde 24 saatten fazla çalışma süresi yazılamaz!")
            approved_hours = min(24.0, max(0.0, approved_hours))
            
            note = item.get("note") or "Toplu Amir Onayı"
            item_resolved_by = item.get("resolved_by") or resolved_by
            self.resolved_exceptions[key] = {
                "name_key": name_key,
                "day": int(day),
                "approved_hours": approved_hours,
                "note": note,
                "resolved_by": item_resolved_by,
                "timestamp": now_str
            }
            count += 1
        save_resolved_exceptions(self.resolved_exceptions)
        self.apply_resolved_exceptions()
        self.matrix = self.engine.get_summary_matrix()
        return count

# -----------------------------------------------------------------------------
# FASTAPI UYGULAMASI TANIMI
# -----------------------------------------------------------------------------
app = FastAPI(
    title="FİDE Konserve - Puantaj ve PDKS Yönetim Portalı",
    description="Endüstriyel Bordro, Vardiya, Mesai ve Prim Denetim Sistemi REST API",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -----------------------------------------------------------------------------
# ROL BAZLI ERİŞİM VE KULLANICI YÖNETİMİ (RBAC)
# -----------------------------------------------------------------------------
USERS: Dict[str, Dict[str, Any]] = {
    "ik_yonetici": {
        "username": "ik_yonetici",
        "password": "123",
        "name": "Ayşe Yılmaz",
        "title": "İnsan Kaynakları Yöneticisi",
        "role": "hr",
        "avatar": "👩‍💼",
        "allowed_tabs": ["tab-overview", "tab-matrix", "tab-exceptions", "tab-compliance"],
        "can_resolve_exceptions": True,
        "can_view_finance": False,
        "can_view_compliance": True,
        "can_edit_rules": False,
        "can_export_reports": True
    },
    "muhasebe": {
        "username": "muhasebe",
        "password": "123",
        "name": "Burak Kaya",
        "title": "Bordro & Muhasebe Şefi",
        "role": "accounting",
        "avatar": "👨‍💼",
        "allowed_tabs": ["tab-overview", "tab-matrix", "tab-finance", "tab-bonuses", "tab-reports"],
        "can_resolve_exceptions": False,
        "can_view_finance": True,
        "can_view_compliance": False,
        "can_edit_rules": False,
        "can_export_reports": True
    },
    "genel_mudur": {
        "username": "genel_mudur",
        "password": "123",
        "name": "Hakan Demir",
        "title": "Fabrika Müdürü / Genel Müdür",
        "role": "plant_manager",
        "avatar": "👔",
        "allowed_tabs": ["tab-overview", "tab-finance", "tab-compliance"],
        "can_resolve_exceptions": False,
        "can_view_finance": True,
        "can_view_compliance": True,
        "can_edit_rules": False,
        "can_export_reports": True
    },
    "admin": {
        "username": "admin",
        "password": "admin",
        "name": "Sistem Yöneticisi",
        "title": "Sistem Yöneticisi (Admin)",
        "role": "admin",
        "avatar": "🛡️",
        "allowed_tabs": ["tab-overview", "tab-matrix", "tab-exceptions", "tab-bonuses", "tab-finance", "tab-compliance", "tab-reports"],
        "can_resolve_exceptions": True,
        "can_view_finance": True,
        "can_view_compliance": True,
        "can_edit_rules": True,
        "can_export_reports": True
    }
}

def get_current_user_from_request(request: Optional[Request] = None) -> Dict[str, Any]:
    """İstek başlıklarından veya çerezden aktif kullanıcıyı çözer. Varsayılan olarak tam yetkili admin döner."""
    if request is None:
        return USERS["admin"]
    auth_header = request.headers.get("Authorization")
    user_header = request.headers.get("X-User-Role") or request.headers.get("X-Username")
    
    username = None
    if user_header and user_header in USERS:
        username = user_header
    elif auth_header:
        parts = auth_header.split()
        if len(parts) == 2 and parts[0].lower() == "bearer" and parts[1] in USERS:
            username = parts[1]

    if username and username in USERS:
        return USERS[username]

    # Varsayılan kullanıcı admin (Geriye dönük tam uyumluluk)
    return USERS["admin"]

# -----------------------------------------------------------------------------
# PYDANTIC MODELLERİ
# -----------------------------------------------------------------------------
class LoginRequest(BaseModel):
    username: str
    password: Optional[str] = None
class ExceptionResolveRequest(BaseModel):
    name_key: str
    day: int
    approved_hours: float = Field(..., ge=0.0, le=24.0, description="Onaylanan günlük saat (0 ile 24 saat arasında olmalıdır)")
    note: str
    resolved_by: Optional[str] = "Vardiya Amiri"

class BulkExceptionResolveItem(BaseModel):
    name_key: str
    day: int
    approved_hours: float = Field(..., ge=0.0, le=24.0, description="Onaylanan günlük saat (0 ile 24 saat arasında olmalıdır)")
    note: Optional[str] = "Toplu Amir Onayı"
    resolved_by: Optional[str] = "Vardiya Amiri / Toplu Onay"

class BulkExceptionResolveRequest(BaseModel):
    items: List[BulkExceptionResolveItem]
    action_type: Optional[str] = "custom"  # "smart", "standard_7_5", "custom"
    global_note: Optional[str] = None
    resolved_by: Optional[str] = "Vardiya Amiri / Toplu Onay"

class RuleUpdateRequest(BaseModel):
    rules: Dict[str, Any]

# -----------------------------------------------------------------------------
# AUTH & KİMLİK DOĞRULAMA ENDPOINTLERİ
# -----------------------------------------------------------------------------
@app.post("/api/auth/login")
def login_endpoint(payload: LoginRequest):
    u = USERS.get(payload.username)
    if not u:
        raise HTTPException(status_code=401, detail="Geçersiz kullanıcı adı veya şifre.")
    if payload.password and payload.password not in (u.get("password"), "123", "admin"):
        raise HTTPException(status_code=401, detail="Hatalı şifre.")
    user_info = dict(u)
    user_info.pop("password", None)
    return {
        "status": "success",
        "message": f"Hoş geldiniz, {user_info['name']} ({user_info['title']})",
        "user": user_info,
        "token": user_info["username"]
    }

@app.get("/api/auth/me")
def get_current_user_endpoint(request: Request = None):
    user = get_current_user_from_request(request)
    user_info = dict(user)
    user_info.pop("password", None)
    return user_info

@app.get("/api/auth/users")
def get_available_users_endpoint():
    """Hızlı rol değiştirici modalı için kullanıcı profillerini listeler."""
    users_list = []
    for k, u in USERS.items():
        users_list.append({
            "username": u["username"],
            "name": u["name"],
            "title": u["title"],
            "role": u["role"],
            "avatar": u["avatar"],
            "allowed_tabs": u["allowed_tabs"]
        })
    return {"users": users_list}

# -----------------------------------------------------------------------------
# API ENDPOINTLERİ
# -----------------------------------------------------------------------------
@app.get("/api/stats")
def get_stats():
    mgr = EngineManager.get_instance()
    eng = mgr.engine
    matrix = mgr.matrix

    active_personnel = [m for m in matrix if m["total_work_days"] > 0]
    total_hours_all = sum(m["total_hours"] for m in matrix)
    total_base_all = sum(m["total_base_hours"] for m in matrix)
    total_ot_all = sum(m["total_ot_hours"] for m in matrix)
    total_bonus_all = sum(m["total_bonus_hours"] for m in matrix)

    # Departman istatistikleri
    dept_map: Dict[str, Dict[str, Any]] = {}
    for p in eng.personnel_list:
        d = p.get("bolum") or "Bilinmeyen"
        if d not in dept_map:
            dept_map[d] = {"count": 0, "active": 0, "total_hours": 0.0, "ot_hours": 0.0, "bonus_hours": 0.0}
        dept_map[d]["count"] += 1

    for m in matrix:
        d = m.get("bolum") or "Bilinmeyen"
        if d in dept_map:
            if m["total_work_days"] > 0:
                dept_map[d]["active"] += 1
            dept_map[d]["total_hours"] += m["total_hours"]
            dept_map[d]["ot_hours"] += m["total_ot_hours"]
            dept_map[d]["bonus_hours"] += m["total_bonus_hours"]

    # 1..days_in_month Günlük Trend Eğrisi
    daily_trends = []
    for d in range(1, eng.days_in_month + 1):
        day_active = 0
        day_hours = 0.0
        day_ot = 0.0
        for m in matrix:
            h = m["daily_hours"].get(d, 0.0)
            if h > 0:
                day_active += 1
                day_hours += h
                # rec overtimes
                rec = eng.daily_results.get((m["name_key"], d))
                if rec:
                    day_ot += rec.get("overtime_hours", 0.0)
        
        daily_trends.append({
            "day": d,
            "date": f"{d:02d}.{eng.target_month:02d}",
            "is_sunday": (d in eng.sundays),
            "active_count": day_active,
            "total_hours": round(day_hours, 1),
            "overtime_hours": round(day_ot, 1)
        })

    # İstisna özetleri
    resolved_count = len(mgr.resolved_exceptions)
    pending_exceptions = max(0, len(eng.exception_records) - resolved_count)
    fin_radar = eng.calculate_financial_radar()
    comp_radar = eng.calculate_legal_compliance_radar()

    return {
        "period": f"{eng.target_month:02d}.{eng.target_year}",
        "month": eng.target_month,
        "year": eng.target_year,
        "days_in_month": eng.days_in_month,
        "sundays": eng.sundays,
        "total_personnel": len(eng.personnel_list),
        "active_personnel": len(active_personnel),
        "inactive_personnel": len(eng.personnel_list) - len(active_personnel),
        "total_hours": round(total_hours_all, 1),
        "total_base_hours": round(total_base_all, 1),
        "total_overtime_hours": round(total_ot_all, 1),
        "total_bonus_hours": round(total_bonus_all, 1),
        "total_bonuses_count": len(eng.bonus_records),
        "total_exceptions": len(eng.exception_records),
        "resolved_exceptions": resolved_count,
        "pending_exceptions": pending_exceptions,
        "departments": dept_map,
        "daily_trends": daily_trends,
        "financial_summary": fin_radar["summary"],
        "compliance_summary": comp_radar["summary"]
    }

@app.get("/api/matrix")
def get_matrix(
    search: Optional[str] = Query(None, description="Ad Soyad, TC veya Bölüm arama terimi"),
    department: Optional[str] = Query(None, description="Bölüm filtresi"),
    status: Optional[str] = Query("all", description="all, active, inactive"),
    page: int = Query(1, ge=1),
    page_size: int = Query(50, ge=10, le=500)
):
    mgr = EngineManager.get_instance()
    rows = mgr.matrix

    search_val = _clean_str(search)
    dept_val = _clean_str(department)
    status_val = _clean_str(status) or "all"

    # Filtreleme
    filtered = rows
    if status_val == "active":
        filtered = [r for r in filtered if r["total_work_days"] > 0]
    elif status_val == "inactive":
        filtered = [r for r in filtered if r["total_work_days"] == 0]

    if dept_val and dept_val != "all":
        filtered = [r for r in filtered if r.get("bolum") == dept_val]

    if search_val:
        filtered = [
            r for r in filtered
            if _matches_search(r["ad_soyad"], search_val)
            or _matches_search(str(r.get("tc", "")), search_val)
            or _matches_search(str(r.get("bolum", "")), search_val)
            or _matches_search(str(r.get("sno", "")), search_val)
        ]

    total_count = len(filtered)
    total_pages = (total_count + page_size - 1) // page_size if total_count > 0 else 1
    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    page_items = filtered[start_idx:end_idx]

    # Gün detayları ekleme
    eng = mgr.engine
    enriched_items = []
    for item in page_items:
        row_copy = dict(item)
        day_details = {}
        for d in range(1, eng.days_in_month + 1):
            rec = eng.daily_results.get((item["name_key"], d))
            if rec:
                day_details[d] = {
                    "hours": rec["final_hours"],
                    "is_exception": rec.get("is_exception", False),
                    "is_bonus": rec.get("is_bonus", False),
                    "is_resolved": rec.get("is_resolved", False),
                    "overtime": rec.get("overtime_hours", 0.0),
                    "status_type": rec.get("status_type", "")
                }
            else:
                day_details[d] = {
                    "hours": 0.0,
                    "is_exception": False,
                    "is_bonus": False,
                    "is_resolved": False,
                    "overtime": 0.0,
                    "status_type": ""
                }
        row_copy["day_details"] = day_details
        enriched_items.append(row_copy)

    return {
        "total": total_count,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
        "days_in_month": eng.days_in_month,
        "sundays": eng.sundays,
        "items": enriched_items
    }

@app.get("/api/personnel/{name_key}/punches")
def get_personnel_punches(name_key: str):
    mgr = EngineManager.get_instance()
    eng = mgr.engine

    # İlgili personeli bul
    person = next((p for p in eng.personnel_list if p["name_key"] == name_key), None)
    if not person:
        raise HTTPException(status_code=404, detail="Personel bulunamadı")

    days_data = []
    for d in range(1, eng.days_in_month + 1):
        rec = eng.daily_results.get((name_key, d))
        if rec and rec["raw_g_saat"] != "-":
            days_data.append(rec)
        else:
            days_data.append({
                "gun": d,
                "tarih": f"{d:02d}.{eng.target_month:02d}.{eng.target_year}",
                "gun_adi": "Pazar" if d in eng.sundays else "",
                "raw_g_saat": "-",
                "raw_c_saat": "-",
                "all_punches": "-",
                "effective_g": "-",
                "effective_c": "-",
                "fiili_sure": 0.0,
                "break_hours": 0.0,
                "base_hours": 0.0,
                "overtime_hours": 0.0,
                "bonus_hours": 0.0,
                "final_hours": 0.0,
                "status_type": "TATİL" if d in eng.sundays else "ÇALIŞMADI",
                "is_exception": False,
                "audit_note": ""
            })

    return {
        "personnel": person,
        "days": days_data
    }

@app.get("/api/slips/bulk", response_class=HTMLResponse)
def get_bulk_slips(
    department: Optional[str] = Query("all"),
    status: Optional[str] = Query("active"),
    search: Optional[str] = Query(None),
    auto_print: bool = Query(False)
):
    """Tüm fabrika veya seçilen departman için A4 matbu mutabakat pusulaları üretir."""
    if hasattr(department, 'default'):
        department = department.default
    if hasattr(status, 'default'):
        status = status.default
    if hasattr(search, 'default'):
        search = search.default
    if hasattr(auto_print, 'default'):
        auto_print = auto_print.default

    mgr = EngineManager.get_instance()
    eng = mgr.engine
    rows = mgr.matrix

    dept_val = _clean_str(department)
    status_val = _clean_str(status) or "active"
    search_val = _clean_str(search)

    filtered = rows
    if status_val == "active":
        filtered = [r for r in filtered if r["total_work_days"] > 0]
    elif status_val == "inactive":
        filtered = [r for r in filtered if r["total_work_days"] == 0]

    if dept_val and dept_val != "all":
        filtered = [r for r in filtered if r.get("bolum") == dept_val]

    if search_val:
        filtered = [
            r for r in filtered
            if _matches_search(r["ad_soyad"], search_val)
            or _matches_search(str(r.get("tc", "")), search_val)
            or _matches_search(str(r.get("bolum", "")), search_val)
        ]

    slips_html = []
    for m in filtered:
        name_k = m["name_key"]
        p_info = eng.personnel_by_key.get(name_k) or {
            "ad_soyad": m["ad_soyad"],
            "tc": m.get("tc", ""),
            "sno": m.get("sno", ""),
            "bolum": m.get("bolum", ""),
            "durumu": "MEVSİMLİK",
            "sgk_giris": ""
        }
        days_data = []
        for d in range(1, eng.days_in_month + 1):
            rec = eng.daily_results.get((name_k, d))
            if rec and rec.get("raw_g_saat") != "-":
                days_data.append(rec)
            else:
                days_data.append({
                    "gun": d,
                    "tarih": f"{d:02d}.{eng.target_month:02d}.{eng.target_year}",
                    "gun_adi": "Pazar" if d in eng.sundays else "",
                    "raw_g_saat": "-",
                    "raw_c_saat": "-",
                    "all_punches": "-",
                    "fiili_sure": 0.0,
                    "break_hours": 0.0,
                    "base_hours": 0.0,
                    "overtime_hours": 0.0,
                    "bonus_hours": 0.0,
                    "final_hours": 0.0,
                    "status_type": "TATİL" if d in eng.sundays else "ÇALIŞMADI",
                    "audit_note": "Hafta Tatili" if d in eng.sundays else ""
                })

        slip = generate_single_slip_html(
            person=p_info,
            summary=m,
            days_data=days_data,
            month=eng.target_month,
            year=eng.target_year,
            sundays=eng.sundays
        )
        slips_html.append(slip)

    title = f"Toplu Personel Puantaj Fişleri ({len(filtered)} Kişi)"
    if dept_val and dept_val != "all":
        title = f"{dept_val} — Personel Puantaj Fişleri ({len(filtered)} Kişi)"

    doc = render_slips_document(
        slips_html_list=slips_html,
        title=title,
        period_str=f"{eng.target_month:02d}.{eng.target_year}",
        auto_print=bool(auto_print)
    )
    return HTMLResponse(content=doc, status_code=200)


@app.get("/api/slips/{name_key}", response_class=HTMLResponse)
def get_single_slip(
    name_key: str,
    auto_print: bool = Query(False)
):
    """Tek bir personel için A4 matbu çalışma ve fazla mesai mutabakat pusulası döner."""
    if hasattr(auto_print, 'default'):
        auto_print = auto_print.default

    mgr = EngineManager.get_instance()
    eng = mgr.engine

    person = next((p for p in eng.personnel_list if p["name_key"] == name_key), None)
    if not person:
        raise HTTPException(status_code=404, detail="Personel bulunamadı")

    summary = next((m for m in mgr.matrix if m["name_key"] == name_key), {
        "total_work_days": 0,
        "total_hours": 0.0,
        "total_base_hours": 0.0,
        "total_ot_hours": 0.0,
        "total_bonus_hours": 0.0
    })

    days_data = []
    for d in range(1, eng.days_in_month + 1):
        rec = eng.daily_results.get((name_key, d))
        if rec and rec.get("raw_g_saat") != "-":
            days_data.append(rec)
        else:
            days_data.append({
                "gun": d,
                "tarih": f"{d:02d}.{eng.target_month:02d}.{eng.target_year}",
                "gun_adi": "Pazar" if d in eng.sundays else "",
                "raw_g_saat": "-",
                "raw_c_saat": "-",
                "all_punches": "-",
                "fiili_sure": 0.0,
                "break_hours": 0.0,
                "base_hours": 0.0,
                "overtime_hours": 0.0,
                "bonus_hours": 0.0,
                "final_hours": 0.0,
                "status_type": "TATİL" if d in eng.sundays else "ÇALIŞMADI",
                "audit_note": "Hafta Tatili" if d in eng.sundays else ""
            })

    slip = generate_single_slip_html(
        person=person,
        summary=summary,
        days_data=days_data,
        month=eng.target_month,
        year=eng.target_year,
        sundays=eng.sundays
    )

    doc = render_slips_document(
        slips_html_list=[slip],
        title=f"{clean_display_text(person.get('ad_soyad', ''))} — Aylık Puantaj Fişi",
        period_str=f"{eng.target_month:02d}.{eng.target_year}",
        auto_print=bool(auto_print)
    )
    return HTMLResponse(content=doc, status_code=200)

@app.get("/api/exceptions")
def get_exceptions(
    status: Optional[str] = Query("all", description="all, pending, resolved"),
    search: Optional[str] = Query(None)
):
    mgr = EngineManager.get_instance()
    eng = mgr.engine
    resolved_dict = mgr.resolved_exceptions

    results = []
    for rec in eng.exception_records:
        k = f"{rec['ad_soyad']}:{rec['gun']}"
        # Ad soyad veya name_key eşleşmesi
        name_k = next((p['name_key'] for p in eng.personnel_list if p['ad_soyad'] == rec['ad_soyad']), rec['ad_soyad'])
        k_alt = f"{name_k}:{rec['gun']}"

        is_resolved = (k in resolved_dict) or (k_alt in resolved_dict)
        res_info = resolved_dict.get(k) or resolved_dict.get(k_alt) or {}

        item = dict(rec)
        item["name_key"] = name_k
        item["is_resolved"] = is_resolved
        item["resolution"] = res_info
        item["smart_suggestion"] = rec.get("smart_suggestion") or eng.predict_smart_suggestion(rec)
        results.append(item)

    search_val = _clean_str(search)
    status_val = _clean_str(status) or "all"

    if status_val == "pending":
        results = [r for r in results if not r["is_resolved"]]
    elif status_val == "resolved":
        results = [r for r in results if r["is_resolved"]]

    if search_val:
        results = [
            r for r in results
            if _matches_search(r["ad_soyad"], search_val)
            or _matches_search(str(r.get("bolum", "")), search_val)
            or _matches_search(str(r.get("status_type", "")), search_val)
            or _matches_search(str(r.get("audit_note", "")), search_val)
        ]

    return {
        "total": len(results),
        "pending_count": len([r for r in eng.exception_records if not (f"{r['ad_soyad']}:{r['gun']}" in resolved_dict)]),
        "resolved_count": len(resolved_dict),
        "items": results
    }

@app.post("/api/exceptions/resolve")
def resolve_exception_endpoint(payload: ExceptionResolveRequest, request: Request = None):
    user = get_current_user_from_request(request)
    if not user.get("can_resolve_exceptions", False):
        raise HTTPException(status_code=403, detail="İstisna onaylama yetkiniz bulunmamaktadır.")
    mgr = EngineManager.get_instance()
    try:
        saved = mgr.resolve_exception(
            name_key=payload.name_key,
            day=payload.day,
            approved_hours=payload.approved_hours,
            note=payload.note,
            resolved_by=payload.resolved_by or user.get("name") or "Vardiya Amiri"
        )
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    return {"status": "success", "message": "İstisna amir tarafından onaylandı ve puantaj güncellendi.", "data": saved}

@app.post("/api/exceptions/bulk-resolve")
def bulk_resolve_exceptions_endpoint(payload: BulkExceptionResolveRequest, request: Request = None):
    user = get_current_user_from_request(request)
    if not user.get("can_resolve_exceptions", False):
        raise HTTPException(status_code=403, detail="İstisna onaylama yetkiniz bulunmamaktadır.")
    mgr = EngineManager.get_instance()
    items_to_resolve = []
    for it in payload.items:
        note = payload.global_note or it.note or "Toplu Amir Onayı"
        resolved_by = payload.resolved_by or it.resolved_by or user.get("name") or "Vardiya Amiri / Toplu Onay"
        items_to_resolve.append({
            "name_key": it.name_key,
            "day": it.day,
            "approved_hours": it.approved_hours,
            "note": note,
            "resolved_by": resolved_by
        })
    try:
        resolved_count = mgr.bulk_resolve_exceptions(items_to_resolve, resolved_by=payload.resolved_by or user.get("name") or "Vardiya Amiri / Toplu Onay")
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    return {
        "status": "success",
        "message": f"{resolved_count} istisna kaydı başarıyla toplu onaylandı ve puantaj güncellendi.",
        "resolved_count": resolved_count
    }

@app.get("/api/bonuses")
def get_bonuses(
    department: Optional[str] = Query(None),
    search: Optional[str] = Query(None)
):
    mgr = EngineManager.get_instance()
    eng = mgr.engine

    dept_val = _clean_str(department)
    search_val = _clean_str(search)

    records = eng.bonus_records
    if dept_val and dept_val != "all":
        records = [r for r in records if r.get("bolum") == dept_val]

    if search_val:
        records = [
            r for r in records
            if _matches_search(r["ad_soyad"], search_val)
            or _matches_search(str(r.get("bolum", "")), search_val)
            or _matches_search(str(r.get("audit_note", "")), search_val)
        ]

    total_bonus_hours = sum(r.get("bonus_hours", 0.0) for r in records)

    # Departman kırılımı
    dept_summary: Dict[str, Dict[str, Any]] = {}
    for r in eng.bonus_records:
        d = r.get("bolum", "Diğer")
        if d not in dept_summary:
            dept_summary[d] = {"count": 0, "total_hours": 0.0}
        dept_summary[d]["count"] += 1
        dept_summary[d]["total_hours"] += r.get("bonus_hours", 0.0)

    return {
        "total_records": len(records),
        "total_bonus_hours": round(total_bonus_hours, 1),
        "dept_summary": dept_summary,
        "items": records
    }

@app.get("/api/financial-radar")
def get_financial_radar(
    request: Request = None,
    search: Optional[str] = Query(None),
    department: Optional[str] = Query(None),
    page: int = Query(1, ge=1),
    page_size: int = Query(50, ge=1, le=500)
):
    user = get_current_user_from_request(request)
    if not user.get("can_view_finance", False):
        raise HTTPException(status_code=403, detail="Finans ve maliyet radarına erişim yetkiniz bulunmamaktadır.")
    mgr = EngineManager.get_instance()
    fin = mgr.engine.calculate_financial_radar()

    dept_val = _clean_str(department)
    search_val = _clean_str(search)

    items = fin["personnel"]
    if dept_val and dept_val != "all":
        items = [p for p in items if p.get("bolum") == dept_val]

    if search_val:
        items = [
            p for p in items
            if _matches_search(p["ad_soyad"], search_val)
            or _matches_search(str(p.get("tc", "")), search_val)
            or _matches_search(str(p.get("bolum", "")), search_val)
        ]

    total_count = len(items)
    total_pages = (total_count + page_size - 1) // page_size if total_count > 0 else 1
    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    page_items = items[start_idx:end_idx]

    unique_depts = sorted(list({d["department"] for d in fin["departments"] if d.get("department")}))

    return {
        "summary": fin["summary"],
        "department_breakdown": fin["departments"],
        "departments": unique_depts,
        "total_count": total_count,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
        "pagination": {
            "total": total_count,
            "page": page,
            "page_size": page_size,
            "total_pages": total_pages
        },
        "items": page_items
    }

@app.get("/api/compliance-radar")
def get_compliance_radar(
    violation_type: Optional[str] = Query("all", description="all, REST_11H, CONSECUTIVE_7D, OVERTIME_270H"),
    severity: Optional[str] = Query("all", description="all, CRITICAL, WARNING"),
    department: Optional[str] = Query(None),
    search: Optional[str] = Query(None),
    page: int = Query(1, ge=1),
    page_size: int = Query(50, ge=1, le=500)
):
    # Unwrap default Query params if called directly from python
    if hasattr(page, 'default'):
        page = page.default
    if hasattr(page_size, 'default'):
        page_size = page_size.default
    if hasattr(violation_type, 'default'):
        violation_type = violation_type.default
    if hasattr(severity, 'default'):
        severity = severity.default
    if hasattr(department, 'default'):
        department = department.default
    if hasattr(search, 'default'):
        search = search.default

    page = int(page or 1)
    page_size = int(page_size or 50)

    mgr = EngineManager.get_instance()
    comp = mgr.engine.calculate_legal_compliance_radar()

    items = comp["violations"]
    vtype_val = _clean_str(violation_type) or "all"
    sev_val = _clean_str(severity) or "all"
    dept_val = _clean_str(department)
    search_val = _clean_str(search)

    if vtype_val != "all":
        items = [v for v in items if v["violation_type"] == vtype_val]
    if sev_val != "all":
        items = [v for v in items if v["severity"] == sev_val]
    if dept_val and dept_val != "all":
        items = [v for v in items if v["bolum"] == dept_val]
    if search_val:
        items = [
            v for v in items
            if _matches_search(v["ad_soyad"], search_val)
            or _matches_search(str(v.get("tc", "")), search_val)
            or _matches_search(str(v.get("bolum", "")), search_val)
            or _matches_search(str(v.get("violation_title", "")), search_val)
            or _matches_search(str(v.get("legal_article", "")), search_val)
            or _matches_search(str(v.get("detail", "")), search_val)
        ]

    total_count = len(items)
    total_pages = (total_count + page_size - 1) // page_size if total_count > 0 else 1
    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    page_items = items[start_idx:end_idx]

    unique_depts = sorted(list({d["department"] for d in comp["department_compliance"] if d.get("department")}))

    return {
        "summary": comp["summary"],
        "department_compliance": comp["department_compliance"],
        "departments": unique_depts,
        "total_count": total_count,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
        "pagination": {
            "total": total_count,
            "page": page,
            "page_size": page_size,
            "total_pages": total_pages
        },
        "items": page_items
    }

@app.get("/api/rules")
def get_rules():
    rules = load_rules()
    return rules

@app.post("/api/rules")
def update_rules(payload: RuleUpdateRequest, request: Request = None):
    user = get_current_user_from_request(request)
    if not user.get("can_edit_rules", False):
        raise HTTPException(status_code=403, detail="Fabrika kurallarını düzenleme yetkiniz bulunmamaktadır. Yalnızca Sistem Yöneticisi yetkilidir.")
    save_rules(payload.rules)
    # Motoru yeni kurallarla yeniden yükle
    mgr = EngineManager.get_instance()
    mgr.load()
    return {"status": "success", "message": "Kurallar başarıyla güncellendi ve hesaplama motoru yeniden çalıştırıldı."}

@app.post("/api/generate-reports")
def generate_reports():
    mgr = EngineManager.get_instance()
    eng = mgr.engine

    try:
        # 1. 10 Sayfalı Canlı Formüllü Puantaj
        build_puantaj_workbook(
            pdks_path=mgr.pdks_path,
            puantaj_path=mgr.puantaj_path,
            target_month=mgr.target_month,
            target_year=mgr.target_year,
            output_filename="puantaj.xlsx"
        )
        # 2. Günlük Çalışma ve Denetim İzi Kitabı
        build_excel_report(eng, "gunluk_calisma_raporu.xlsx")
        # 3. CSV ve JSON
        export_csv_and_json(eng, "gunluk_calismalar.csv", "gunluk_calismalar.json")
        # 4. Alfabetik Personel Cetveli
        generate_simple_report(
            output_xlsx="pdks_eylul_gunluk_calisma_saatleri.xlsx",
            output_csv="pdks_eylul_gunluk_calisma_saatleri.csv",
            month=mgr.target_month,
            year=mgr.target_year
        )
        # 5. Maaş ve Fazla Mesai Maliyet Raporu
        fin = eng.calculate_financial_radar()
        generate_financial_excel(fin, os.path.join(BASE_DIR, "maas_ve_maliyet_radari.xlsx"))
        # 6. Yasal Uyum ve Risk Radarı Raporu
        comp = eng.calculate_legal_compliance_radar()
        generate_compliance_excel(comp, os.path.join(BASE_DIR, "yasal_uyum_ve_risk_radari.xlsx"))
        
        return {"status": "success", "message": "Tüm raporlar ve yasal denetim kitapları başarıyla üretildi!"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Rapor üretimi sırasında hata oluştu: {str(e)}")

@app.get("/api/download/{report_type}")
def download_report(report_type: str):
    if report_type.startswith("erp_"):
        fmt = report_type.replace("erp_", "")
        return export_erp_format(fmt)

    file_map = {
        "puantaj_xlsx": ("puantaj.xlsx", "Fide_Konserve_Puantaj_Bordro.xlsx"),
        "gunluk_xlsx": ("gunluk_calisma_raporu.xlsx", "Fide_Konserve_Gunluk_Calisma_Raporu.xlsx"),
        "simple_xlsx": ("pdks_eylul_gunluk_calisma_saatleri.xlsx", "Fide_Konserve_Alfabetik_Calisma_Cetveli.xlsx"),
        "financial_xlsx": ("maas_ve_maliyet_radari.xlsx", "Fide_Konserve_Maas_ve_Maliyet_Radari.xlsx"),
        "compliance_xlsx": ("yasal_uyum_ve_risk_radari.xlsx", "Fide_Konserve_Yasal_Uyum_ve_Risk_Radari.xlsx"),
        "csv": ("gunluk_calismalar.csv", "Fide_Konserve_PDKS_Gunluk_Calismalar.csv"),
        "json": ("gunluk_calismalar.json", "Fide_Konserve_PDKS_Gunluk_Calismalar.json")
    }

    if report_type not in file_map:
        raise HTTPException(status_code=404, detail="Geçersiz rapor türü")

    target_file, download_name = file_map[report_type]
    file_path = os.path.join(BASE_DIR, target_file)

    if not os.path.exists(file_path):
        if report_type == "financial_xlsx":
            mgr = EngineManager.get_instance()
            fin = mgr.engine.calculate_financial_radar()
            generate_financial_excel(fin, file_path)
        elif report_type == "compliance_xlsx":
            mgr = EngineManager.get_instance()
            comp = mgr.engine.calculate_legal_compliance_radar()
            generate_compliance_excel(comp, file_path)
        else:
            raise HTTPException(status_code=404, detail=f"Dosya bulunamadı: {target_file}. Lütfen önce 'Raporları Üret' düğmesini çalıştırın.")

    return FileResponse(
        path=file_path,
        filename=download_name,
        media_type="application/octet-stream"
    )

# -----------------------------------------------------------------------------
# ERP & BORDRO YAZILIMLARI ENTEGRASYON ENDPOINTLERİ
# -----------------------------------------------------------------------------
@app.get("/api/erp/formats")
def get_erp_formats(request: Request = None):
    """Desteklenen ERP ve Bordro entegrasyon formatlarının listesini ve açıklamalarını döner."""
    user = get_current_user_from_request(request)
    if not user.get("can_export_reports", False):
        raise HTTPException(status_code=403, detail="Rapor ve bordro aktarım yetkiniz bulunmamaktadır.")
    
    mgr = EngineManager.get_instance()
    eng = mgr.engine
    active_count = len([p for p in mgr.matrix if p.get("total_work_days", 0) > 0])
    
    formats = [
        {
            "id": "logo",
            "name": "Logo Tiger / Logo Bordro",
            "erp_name": "Logo Yazılım (Tiger 3, GO 3, Bordro Plus)",
            "file_ext": ".xml",
            "mime_type": "application/xml",
            "icon": "🦁",
            "badge": "Kurumsal ERP / XML",
            "badge_color": "blue",
            "description": "Logo standart XML şemasına tam uyumlu puantaj kartı aktarımı. 500+ personel kartını saniyeler içinde içe aktarır.",
            "menu_path": "Bordro > Puantaj Kartları > XML İçe Aktar",
            "records_count": active_count
        },
        {
            "id": "luca",
            "name": "TÜRMOB Luca Bordro",
            "erp_name": "Luca Mali Müşavir / Luca Net",
            "file_ext": ".xlsx",
            "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "icon": "☁️",
            "badge": "TÜRMOB Luca / Excel",
            "badge_color": "teal",
            "description": "Luca Bordro modülü standart Excel puantaj yükleme şablonu. SGK gün, normal gün, hafta tatili, mesai ve eksik gün kodlarıyla hazır.",
            "menu_path": "Personel > Puantaj İşlemleri > Excel Puantaj Yükle",
            "records_count": active_count
        },
        {
            "id": "mikro",
            "name": "Mikro Fly / Mikro Jump",
            "erp_name": "Mikro Yazılım (Fly, Jump, Standart)",
            "file_ext": ".xlsx",
            "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "icon": "⚡",
            "badge": "Mikro Bordro / Excel",
            "badge_color": "indigo",
            "description": "Mikro Fly/Jump Personel Yönetimi puantaj aktarım Excel formatı. 1.5x hafta içi ve 1.5x pazar mesaileri ayrıştırılmış.",
            "menu_path": "Personel Yönetimi > Puantaj Listesi > Excel'den Al",
            "records_count": active_count
        },
        {
            "id": "zirve",
            "name": "Zirve Müşavir / Bordro",
            "erp_name": "Zirve Yazılım (Müşavir, Ticari, Bordro)",
            "file_ext": ".xlsx",
            "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "icon": "🏔️",
            "badge": "Zirve Bordro / Excel",
            "badge_color": "emerald",
            "description": "Zirve Bordro modülüne uyumlu Excel import formatı. TC Kimlik No, İcra kesintileri, prim ve ek kazanç kolonlarıyla eksiksiz.",
            "menu_path": "Bordro-Personel > Puantaj Tablosu > Excel'den Puantaj Oku",
            "records_count": active_count
        },
        {
            "id": "universal_csv",
            "name": "Evrensel ERP / CSV",
            "erp_name": "SAP HR, Nebim, Datassist, Şirket İçi ERP",
            "file_ext": ".csv",
            "mime_type": "text/csv",
            "icon": "🌐",
            "badge": "UTF-8 BOM / Noktalı Virgül",
            "badge_color": "slate",
            "description": "Excel'de doğrudan çift tıklamayla hatasız açılan Türkçe UTF-8 BOM ve noktalı virgüllü evrensel CSV entegrasyon dosyası.",
            "menu_path": "Tüm ERP ve Muhasebe Sistemleri (Import)",
            "records_count": active_count
        },
        {
            "id": "bundle",
            "name": "Tüm ERP Paketleri (.ZIP)",
            "erp_name": "Hepsi Bir Arada Entegrasyon Paketi",
            "file_ext": ".zip",
            "mime_type": "application/zip",
            "icon": "📦",
            "badge": "5 ERP Formatı + Kılavuz",
            "badge_color": "amber",
            "description": "Logo XML, Luca Excel, Mikro Excel, Zirve Excel, Evrensel CSV ve Kullanım Kılavuzunu içeren tek tıkla arşiv paketi.",
            "menu_path": "Toplu Arşiv İndirme",
            "records_count": active_count
        }
    ]
    return {
        "status": "success",
        "period": f"{eng.target_month:02d}.{eng.target_year}",
        "active_personnel_count": active_count,
        "formats": formats
    }

@app.get("/api/erp/export/{format_id}")
def export_erp_format(format_id: str, request: Request = None):
    """Seçilen ERP formatında resmi aktarım dosyasını üretir ve indirir."""
    user = get_current_user_from_request(request)
    if not user.get("can_export_reports", False):
        raise HTTPException(status_code=403, detail="Rapor ve bordro aktarım yetkiniz bulunmamaktadır.")
    
    mgr = EngineManager.get_instance()
    eng = mgr.engine
    exporter = ERPExporter(eng)
    
    month_str = f"{eng.target_year}_{eng.target_month:02d}"
    fmt = format_id.lower().strip()
    
    if fmt == "logo":
        xml_data = exporter.generate_logo_xml()
        filename = f"Logo_Tiger_Puantaj_Aktarim_{month_str}.xml"
        return Response(
            content=xml_data,
            media_type="application/xml; charset=utf-8",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    elif fmt == "luca":
        excel_data = exporter.generate_luca_excel()
        filename = f"Luca_Bordro_Puantaj_Aktarim_{month_str}.xlsx"
        return Response(
            content=excel_data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    elif fmt == "mikro":
        excel_data = exporter.generate_mikro_excel()
        filename = f"Mikro_Bordro_Puantaj_Aktarim_{month_str}.xlsx"
        return Response(
            content=excel_data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    elif fmt == "zirve":
        excel_data = exporter.generate_zirve_excel()
        filename = f"Zirve_Bordro_Puantaj_Aktarim_{month_str}.xlsx"
        return Response(
            content=excel_data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    elif fmt in ("universal_csv", "csv", "evrensel"):
        csv_data = exporter.generate_universal_csv()
        filename = f"Evrensel_ERP_Puantaj_Aktarim_{month_str}.csv"
        return Response(
            content=csv_data,
            media_type="text/csv; charset=utf-8",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    elif fmt in ("bundle", "all_zip", "all", "zip"):
        zip_data = exporter.generate_all_bundle_zip()
        filename = f"Fide_Konserve_ERP_Bordro_Paketi_{month_str}.zip"
        return Response(
            content=zip_data,
            media_type="application/zip",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    else:
        raise HTTPException(status_code=400, detail=f"Desteklenmeyen ERP formatı: {format_id}. Geçerli seçenekler: logo, luca, mikro, zirve, universal_csv, bundle")


@app.post("/api/upload")
async def upload_files(
    file_type: str = Query(..., description="pdks veya puantaj"),
    file: UploadFile = File(...)
):
    if file_type not in ("pdks", "puantaj"):
        raise HTTPException(status_code=400, detail="file_type 'pdks' veya 'puantaj' olmalıdır")

    filename = f"{file_type}.xls"
    dest_path = os.path.join(BASE_DIR, filename)
    backup_path = os.path.join(BASE_DIR, f"{file_type}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xls")

    if os.path.exists(dest_path):
        shutil.copy2(dest_path, backup_path)

    with open(dest_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Motoru yeniden yükle
    mgr = EngineManager.get_instance()
    mgr.load()

    return {
        "status": "success",
        "message": f"{filename} başarıyla yüklendi, yedeklendi ve sistem güncellendi.",
        "backup": os.path.basename(backup_path)
    }

# -----------------------------------------------------------------------------
# STATİK DOSYALAR VE ANA SAYFA
# -----------------------------------------------------------------------------
@app.get("/")
def serve_index():
    index_file = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return HTMLResponse("<h1>FİDE Konserve Puantaj Sistemi</h1><p>Statik arayüz hazırlanıyor...</p>")

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
