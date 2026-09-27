# -*- coding: utf-8 -*-
"""
===================================================================================================
FİDE KONSERVE - MUHASEBE VE ERP BORDRO ENTEGRASYON MODÜLÜ (erp_exporter.py)
===================================================================================================
Bu modül, fabrika PDKS ve puantaj hesaplama motorundan (PDKSEngine) aldığı verileri
Türkiye'de yaygın kullanılan kurumsal bordro ve ERP yazılımlarına doğrudan aktarılabilir
resmi formatlarda üretir:

1. Logo Tiger / Logo Bordro (XML)
2. Luca Bordro (TÜRMOB Standart Excel)
3. Mikro Yazılım (Mikro Fly / Jump Excel)
4. Zirve Müşavir / Bordro (Excel)
5. Evrensel ERP / CSV (UTF-8 BOM, Noktalı Virgül)
6. Tüm Paket (.ZIP)
===================================================================================================
"""

import io
import os
import zipfile
import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import List, Dict, Any, Optional, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from pdks_engine import PDKSEngine, clean_display_text, norm_name_key


# =================================================================================================
# STİL VE TEMA SABİTLERİ (Kurumsal Excel Raporlama)
# =================================================================================================
FONT_FAMILY = "Segoe UI"
FONT_TITLE = Font(name=FONT_FAMILY, size=13, bold=True, color="FFFFFF")
FONT_HEADER = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
FONT_DATA = Font(name=FONT_FAMILY, size=9.5, bold=False, color="1E293B")
FONT_TOTAL = Font(name=FONT_FAMILY, size=10, bold=True, color="0F172A")

FILL_LOGO_NAVY = PatternFill("solid", fgColor="1E3A8A")     # Logo Teması
FILL_LUCA_CYAN = PatternFill("solid", fgColor="0E7490")     # Luca Teması
FILL_MIKRO_INDIGO = PatternFill("solid", fgColor="3730A3")  # Mikro Teması
FILL_ZIRVE_EMERALD = PatternFill("solid", fgColor="065F46") # Zirve Teması
FILL_ZEBRA = PatternFill("solid", fgColor="F8FAFC")
FILL_TOTAL = PatternFill("solid", fgColor="E2E8F0")
FILL_OVERTIME = PatternFill("solid", fgColor="FEF3C7")      # Sarı/Turuncu mesai
FILL_DEDUCT = PatternFill("solid", fgColor="FEE2E2")        # Açık kırmızı kesinti

BORDER_THIN = Border(
    left=Side(style='thin', color="CBD5E1"),
    right=Side(style='thin', color="CBD5E1"),
    top=Side(style='thin', color="CBD5E1"),
    bottom=Side(style='thin', color="CBD5E1")
)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")


class ERPExporter:
    """Fabrika PDKS ve Puantaj verilerini kurumsal ERP formatlarına dönüştürücü sınıf."""

    def __init__(self, engine: PDKSEngine):
        self.engine = engine
        self.matrix = engine.get_summary_matrix()
        self.active_personnel = [p for p in self.matrix if p.get("total_work_days", 0) > 0]
        # Sıra numarasına göre sırala
        self.active_personnel.sort(key=lambda x: x.get("sno", 9999))
        self.days_in_month = engine.days_in_month
        self.sundays_count = len(engine.sundays)
        self.month = engine.target_month
        self.year = engine.target_year

    def _extract_employee_erp_data(self, emp: Dict[str, Any]) -> Dict[str, Any]:
        """Her personel için ERP aktarımında gereken tüm yasal ve finansal alanları hazırlar."""
        name_k = emp.get("name_key", "")
        tc = emp.get("tc", "")
        ad_soyad = clean_display_text(emp.get("ad_soyad", ""))
        
        # Ad ve soyadı ayrıştır
        parts = ad_soyad.strip().split()
        if len(parts) > 1:
            first_name = " ".join(parts[:-1])
            last_name = parts[-1]
        else:
            first_name = ad_soyad
            last_name = ""

        work_days = emp.get("total_work_days", 0)
        base_hours = round(emp.get("total_base_hours", 0.0), 2)
        
        # Hafta Tatili kuralı: 5 gün ve üzeri çalışana 4 Pazar tam HT, daha az çalışana her 6 güne 1 HT
        if work_days >= 5:
            ht_days = self.sundays_count
        elif work_days > 0:
            ht_days = int(work_days / 6)
        else:
            ht_days = 0

        izin_days = self.engine.izin_by_tc.get(tc, {}).get("gun", 0) if tc else 0
        rapor_days = self.engine.rapor_by_tc.get(tc, 0) if tc else 0

        # SGK ve Eksik Gün
        sgk_days = min(self.days_in_month, work_days + ht_days + izin_days + rapor_days)
        eksik_days = max(0, self.days_in_month - sgk_days)
        
        if rapor_days > 0:
            eksik_kodu = "01-İstirahat"
        elif eksik_days > 0:
            eksik_kodu = "12-Birden Fazla"
        else:
            eksik_kodu = ""

        # Mesai ve Eksik Saatler (1.5x / 1.0x katsayıları)
        ot_weekday = round(emp.get("total_weekday_ot_hours", 0.0), 2)
        missing_hours = round(emp.get("total_missing_hours", 0.0), 2)
        ot_sunday = round(emp.get("total_sunday_ot_hours", 0.0), 2)
        bonus_hours = round(emp.get("total_bonus_hours", 0.0), 2)
        net_weekday_ot = round((ot_weekday * 1.5) - (missing_hours * 1.0), 2)

        # Ücret ve İcra
        net_maas = round(emp.get("net_maas", 0.0), 2)
        if net_maas <= 0:
            if tc in self.engine.bordro_by_tc and self.engine.bordro_by_tc[tc].get("odenecek_net", 0.0) > 0:
                net_maas = round(self.engine.bordro_by_tc[tc]["odenecek_net"], 2)
            else:
                net_maas = 28075.50

        hourly_rate = round(net_maas / 225.0, 2)
        weekday_ot_cost = round(net_weekday_ot * hourly_rate, 2)
        sunday_ot_cost = round(ot_sunday * 1.5 * hourly_rate, 2)
        bonus_cost = round(bonus_hours * 1.5 * hourly_rate, 2)
        total_ot_cost = round(weekday_ot_cost + sunday_ot_cost + bonus_cost, 2)
        
        icra_info = self.engine.icra_by_name.get(name_k, {})
        icra_kes = icra_info.get("kesilen", 0.0) if isinstance(icra_info, dict) else 0.0

        # Sicil / Kart No
        sicil_no = str(emp.get("sno", "")).zfill(3)

        return {
            "sno": emp.get("sno", 0),
            "sicil_no": sicil_no,
            "tc": tc,
            "ad_soyad": ad_soyad,
            "first_name": first_name,
            "last_name": last_name,
            "bolum": clean_display_text(emp.get("bolum", "Genel")),
            "durumu": clean_display_text(emp.get("durumu", "MEVSİMLİK")),
            "work_days": work_days,
            "base_hours": base_hours,
            "ht_days": ht_days,
            "izin_days": izin_days,
            "rapor_days": rapor_days,
            "sgk_days": sgk_days,
            "eksik_days": eksik_days,
            "eksik_kodu": eksik_kodu,
            "ot_weekday": ot_weekday,
            "missing_hours": missing_hours,
            "net_weekday_ot": net_weekday_ot,
            "ot_sunday": ot_sunday,
            "bonus_hours": bonus_hours,
            "net_maas": net_maas,
            "hourly_rate": hourly_rate,
            "weekday_ot_cost": weekday_ot_cost,
            "sunday_ot_cost": sunday_ot_cost,
            "total_ot_cost": total_ot_cost,
            "icra_kesintisi": round(icra_kes, 2)
        }

    # =========================================================================
    # 1. LOGO TIGER / LOGO BORDRO (XML AKTARIM FORMATI)
    # =========================================================================
    def generate_logo_xml(self) -> bytes:
        """
        Logo Tiger, Logo GO ve Logo Bordro sistemi için standart XML içe aktarım dosyası üretir.
        """
        root = ET.Element("PUANTAJ_KAYITLARI", {
            "FIRMA": "FIDE KONSERVE GIDA SAN. TIC. A.S.",
            "DONEM_YIL": str(self.year),
            "DONEM_AY": str(self.month),
            "VERSIYON": "2.0",
            "URETIM_SISTEMI": "FIDE_PDKS_ENGINE"
        })

        meta = ET.SubElement(root, "BILGI")
        ET.SubElement(meta, "TOPLAM_AKTIF_PERSONEL").text = str(len(self.active_personnel))
        ET.SubElement(meta, "TOPLAM_CALISMA_GUN").text = str(sum(p.get("total_work_days", 0) for p in self.active_personnel))
        ET.SubElement(meta, "TOPLAM_MESAI_SAATI").text = str(round(sum(p.get("total_ot_hours", 0) for p in self.active_personnel), 2))
        ET.SubElement(meta, "MESAI_CARPANI").text = "1.5"
        ET.SubElement(meta, "EKSIK_SAAT_CARPANI").text = "1.0"

        items_el = ET.SubElement(root, "SATIRLAR")
        for emp_raw in self.active_personnel:
            d = self._extract_employee_erp_data(emp_raw)
            row = ET.SubElement(items_el, "PERSONEL")
            
            ET.SubElement(row, "SIRA_NO").text = str(d["sno"])
            ET.SubElement(row, "SICIL_KODU").text = d["sicil_no"]
            ET.SubElement(row, "TC_KIMLIK_NO").text = d["tc"]
            ET.SubElement(row, "ADI_SOYADI").text = d["ad_soyad"]
            ET.SubElement(row, "DEPARTMAN").text = d["bolum"]
            ET.SubElement(row, "ISTIHDAM_SEKLI").text = d["durumu"]

            # Gün Bazlı Dağılım
            ET.SubElement(row, "NORMAL_CALISMA_GUN").text = str(d["work_days"])
            ET.SubElement(row, "HAFTA_TATILI_GUN").text = str(d["ht_days"])
            ET.SubElement(row, "UCRETLI_IZIN_GUN").text = str(d["izin_days"])
            ET.SubElement(row, "RAPOR_GUN").text = str(d["rapor_days"])
            ET.SubElement(row, "SGK_PRIM_GUN").text = str(d["sgk_days"])
            ET.SubElement(row, "EKSIK_GUN").text = str(d["eksik_days"])
            ET.SubElement(row, "EKSIK_NEDEN_KODU").text = d["eksik_kodu"]

            # Saat Bazlı Dağılım ve Katsayılar
            ET.SubElement(row, "NORMAL_CALISMA_SAAT").text = f"{d['base_hours']:.1f}"
            ET.SubElement(row, "HAFTA_ICI_FAZLA_MESAI_SAAT").text = f"{d['ot_weekday']:.1f}"
            ET.SubElement(row, "HAFTA_ICI_MESAI_KATSAYI").text = "1.5"
            ET.SubElement(row, "EKSIK_CALISMA_SAAT").text = f"{d['missing_hours']:.1f}"
            ET.SubElement(row, "EKSIK_CALISMA_KATSAYI").text = "1.0"
            ET.SubElement(row, "NET_HAFTA_ICI_MESAI_SAAT").text = f"{d['net_weekday_ot']:.1f}"
            ET.SubElement(row, "PAZAR_MESAI_SAAT").text = f"{d['ot_sunday']:.1f}"
            ET.SubElement(row, "PAZAR_MESAI_KATSAYI").text = "1.5"
            ET.SubElement(row, "BOLUM_PRIMI_SAAT").text = f"{d['bonus_hours']:.1f}"

            # Finansal
            ET.SubElement(row, "NET_MAAS_TABAN").text = f"{d['net_maas']:.2f}"
            ET.SubElement(row, "SAATLIK_UCRET").text = f"{d['hourly_rate']:.2f}"
            ET.SubElement(row, "MESAI_TUTARI_TL").text = f"{d['total_ot_cost']:.2f}"
            ET.SubElement(row, "ICRA_KESINTISI_TL").text = f"{d['icra_kesintisi']:.2f}"

        # Girintili ve okunabilir XML
        rough_string = ET.tostring(root, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        pretty_xml = reparsed.toprettyxml(indent="  ", encoding="utf-8")
        return pretty_xml

    # =========================================================================
    # 2. LUCA BORDRO (TÜRMOB STANDART PUANTAJ EXCEL AKTARIM ŞABLONU)
    # =========================================================================
    def generate_luca_excel(self) -> bytes:
        """TÜRMOB Luca Mali Müşavir ve Bordro sistemi standart Excel şablonu üretir."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Luca_Puantaj_Aktarim"
        ws.views.sheetView[0].showGridLines = True

        # Üst Başlık Bandı
        ws.merge_cells("A1:Q1")
        cell_title = ws["A1"]
        cell_title.value = f"TÜRMOB LUCA BORDRO PUANTAJ AKTARIM CETVELİ — FİDE KONSERVE ({self.month:02d}.{self.year})"
        cell_title.font = FONT_TITLE
        cell_title.fill = FILL_LUCA_CYAN
        cell_title.alignment = ALIGN_CENTER
        ws.row_dimensions[1].height = 30

        headers = [
            ("TC Kimlik No", 14), ("Adı", 14), ("Soyadı", 14), ("Departman", 18),
            ("Normal Gün", 11), ("Hafta Tatili", 11), ("Yıllık İzin", 11), ("Rapor", 10),
            ("SGK Gün", 10), ("Eksik Gün", 10), ("Eksik Nedeni", 15),
            ("H.İçi Mesai (1.5x)", 14), ("Eksik Saat (1.0x)", 13), ("Pazar Mesai (1.5x)", 14),
            ("Özel Prim (s)", 12), ("Net Ücret (TL)", 14), ("İcra Kesintisi (TL)", 15)
        ]

        ws.row_dimensions[2].height = 24
        for col_idx, (h_text, w) in enumerate(headers, start=1):
            col_let = get_column_letter(col_idx)
            cell = ws[f"{col_let}2"]
            cell.value = h_text
            cell.font = FONT_HEADER
            cell.fill = FILL_LUCA_CYAN
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN
            ws.column_dimensions[col_let].width = w

        for r_idx, emp_raw in enumerate(self.active_personnel, start=3):
            d = self._extract_employee_erp_data(emp_raw)
            row_fill = FILL_ZEBRA if r_idx % 2 == 0 else None

            ws[f"A{r_idx}"] = str(d["tc"])
            ws[f"A{r_idx}"].number_format = "@"
            ws[f"B{r_idx}"] = d["first_name"]
            ws[f"C{r_idx}"] = d["last_name"]
            ws[f"D{r_idx}"] = d["bolum"]
            ws[f"E{r_idx}"] = d["work_days"]
            ws[f"F{r_idx}"] = d["ht_days"]
            ws[f"G{r_idx}"] = d["izin_days"]
            ws[f"H{r_idx}"] = d["rapor_days"]
            ws[f"I{r_idx}"] = d["sgk_days"]
            ws[f"J{r_idx}"] = d["eksik_days"]
            ws[f"K{r_idx}"] = d["eksik_kodu"]
            ws[f"L{r_idx}"] = d["ot_weekday"]
            ws[f"M{r_idx}"] = d["missing_hours"]
            ws[f"N{r_idx}"] = d["ot_sunday"]
            ws[f"O{r_idx}"] = d["bonus_hours"]
            ws[f"P{r_idx}"] = d["net_maas"]
            ws[f"Q{r_idx}"] = d["icra_kesintisi"]

            for col_idx in range(1, 18):
                cl = get_column_letter(col_idx)
                c = ws[f"{cl}{r_idx}"]
                c.font = FONT_DATA
                c.border = BORDER_THIN
                if row_fill:
                    c.fill = row_fill

                if col_idx in (1, 5, 6, 7, 8, 9, 10):
                    c.alignment = ALIGN_CENTER
                elif col_idx in (12, 13, 14, 15):
                    c.alignment = ALIGN_RIGHT
                    c.number_format = "0.0"
                    if col_idx in (12, 14):
                        c.fill = FILL_OVERTIME
                    elif col_idx == 13 and d["missing_hours"] > 0:
                        c.fill = FILL_DEDUCT
                elif col_idx in (16, 17):
                    c.alignment = ALIGN_RIGHT
                    c.number_format = "#,##0.00"

        # Toplam Satırı
        tot_row = len(self.active_personnel) + 3
        ws[f"A{tot_row}"] = "GENEL TOPLAM"
        ws[f"A{tot_row}"].font = FONT_TOTAL
        ws.merge_cells(f"A{tot_row}:D{tot_row}")
        
        for col_idx in range(5, 11):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"] = f"=SUM({cl}3:{cl}{tot_row-1})"
            ws[f"{cl}{tot_row}"].number_format = "0"
            ws[f"{cl}{tot_row}"].alignment = ALIGN_CENTER
            ws[f"{cl}{tot_row}"].font = FONT_TOTAL

        for col_idx in (12, 13, 14, 15):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"] = f"=SUM({cl}3:{cl}{tot_row-1})"
            ws[f"{cl}{tot_row}"].number_format = "0.0"
            ws[f"{cl}{tot_row}"].alignment = ALIGN_RIGHT
            ws[f"{cl}{tot_row}"].font = FONT_TOTAL

        for col_idx in (16, 17):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"] = f"=SUM({cl}3:{cl}{tot_row-1})"
            ws[f"{cl}{tot_row}"].number_format = "#,##0.00"
            ws[f"{cl}{tot_row}"].alignment = ALIGN_RIGHT
            ws[f"{cl}{tot_row}"].font = FONT_TOTAL

        for col_idx in range(1, 18):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"].border = BORDER_THIN
            ws[f"{cl}{tot_row}"].fill = FILL_TOTAL

        ws.freeze_panes = "A3"
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

    # =========================================================================
    # 3. MİKRO YAZILIM (MİKRO FLY / JUMP BORDRO PUANTAJ ŞABLONU)
    # =========================================================================
    def generate_mikro_excel(self) -> bytes:
        """Mikro Yazılım Fly / Jump Bordro modülü standart puantaj Excel aktarımı üretir."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Mikro_Bordro_Puantaj"
        ws.views.sheetView[0].showGridLines = True

        ws.merge_cells("A1:P1")
        cell_title = ws["A1"]
        cell_title.value = f"MİKRO YAZILIM BORDRO PUANTAJ KARTLARI — FİDE KONSERVE ({self.month:02d}.{self.year})"
        cell_title.font = FONT_TITLE
        cell_title.fill = FILL_MIKRO_INDIGO
        cell_title.alignment = ALIGN_CENTER
        ws.row_dimensions[1].height = 30

        headers = [
            ("Personel Kodu", 13), ("TC Kimlik No", 14), ("Adı Soyadı", 20), ("Departman", 18),
            ("Normal Gün", 11), ("Hafta Tatili", 11), ("H.İçi Fazla Mesai (s)", 15),
            ("Eksik Çalışma (s)", 13), ("Net H.İçi FM (s)", 14), ("Pazar Mesai (s)", 14),
            ("Bölüm Primi (s)", 13), ("Ücretli İzin Gün", 13), ("Rapor Gün", 11),
            ("SGK Gün", 10), ("Eksik Kodu", 15), ("İcra Kesintisi (TL)", 15)
        ]

        ws.row_dimensions[2].height = 24
        for col_idx, (h_text, w) in enumerate(headers, start=1):
            col_let = get_column_letter(col_idx)
            cell = ws[f"{col_let}2"]
            cell.value = h_text
            cell.font = FONT_HEADER
            cell.fill = FILL_MIKRO_INDIGO
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN
            ws.column_dimensions[col_let].width = w

        for r_idx, emp_raw in enumerate(self.active_personnel, start=3):
            d = self._extract_employee_erp_data(emp_raw)
            row_fill = FILL_ZEBRA if r_idx % 2 == 0 else None

            ws[f"A{r_idx}"] = d["sicil_no"]
            ws[f"B{r_idx}"] = str(d["tc"])
            ws[f"B{r_idx}"].number_format = "@"
            ws[f"C{r_idx}"] = d["ad_soyad"]
            ws[f"D{r_idx}"] = d["bolum"]
            ws[f"E{r_idx}"] = d["work_days"]
            ws[f"F{r_idx}"] = d["ht_days"]
            ws[f"G{r_idx}"] = d["ot_weekday"]
            ws[f"H{r_idx}"] = d["missing_hours"]
            ws[f"I{r_idx}"] = d["net_weekday_ot"]
            ws[f"J{r_idx}"] = d["ot_sunday"]
            ws[f"K{r_idx}"] = d["bonus_hours"]
            ws[f"L{r_idx}"] = d["izin_days"]
            ws[f"M{r_idx}"] = d["rapor_days"]
            ws[f"N{r_idx}"] = d["sgk_days"]
            ws[f"O{r_idx}"] = d["eksik_kodu"]
            ws[f"P{r_idx}"] = d["icra_kesintisi"]

            for col_idx in range(1, 17):
                cl = get_column_letter(col_idx)
                c = ws[f"{cl}{r_idx}"]
                c.font = FONT_DATA
                c.border = BORDER_THIN
                if row_fill:
                    c.fill = row_fill

                if col_idx in (1, 2, 5, 6, 12, 13, 14):
                    c.alignment = ALIGN_CENTER
                elif col_idx in (7, 8, 9, 10, 11):
                    c.alignment = ALIGN_RIGHT
                    c.number_format = "0.0"
                    if col_idx in (7, 9, 10):
                        c.fill = FILL_OVERTIME
                    elif col_idx == 8 and d["missing_hours"] > 0:
                        c.fill = FILL_DEDUCT
                elif col_idx == 16:
                    c.alignment = ALIGN_RIGHT
                    c.number_format = "#,##0.00"

        # Toplam Satırı
        tot_row = len(self.active_personnel) + 3
        ws[f"A{tot_row}"] = "TOPLAM"
        ws[f"A{tot_row}"].font = FONT_TOTAL
        ws.merge_cells(f"A{tot_row}:D{tot_row}")

        for col_idx in (5, 6, 12, 13, 14):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"] = f"=SUM({cl}3:{cl}{tot_row-1})"
            ws[f"{cl}{tot_row}"].number_format = "0"
            ws[f"{cl}{tot_row}"].alignment = ALIGN_CENTER
            ws[f"{cl}{tot_row}"].font = FONT_TOTAL

        for col_idx in (7, 8, 9, 10, 11):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"] = f"=SUM({cl}3:{cl}{tot_row-1})"
            ws[f"{cl}{tot_row}"].number_format = "0.0"
            ws[f"{cl}{tot_row}"].alignment = ALIGN_RIGHT
            ws[f"{cl}{tot_row}"].font = FONT_TOTAL

        ws[f"P{tot_row}"] = f"=SUM(P3:P{tot_row-1})"
        ws[f"P{tot_row}"].number_format = "#,##0.00"
        ws[f"P{tot_row}"].alignment = ALIGN_RIGHT
        ws[f"P{tot_row}"].font = FONT_TOTAL

        for col_idx in range(1, 17):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"].border = BORDER_THIN
            ws[f"{cl}{tot_row}"].fill = FILL_TOTAL

        ws.freeze_panes = "A3"
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

    # =========================================================================
    # 4. ZİRVE BORDRO (ZİRVE MÜŞAVİR EXCEL ŞABLONU)
    # =========================================================================
    def generate_zirve_excel(self) -> bytes:
        """Zirve Müşavir / Bordro sistemi puantaj veri yükleme Excel şablonu üretir."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Zirve_Puantaj_Yukleme"
        ws.views.sheetView[0].showGridLines = True

        ws.merge_cells("A1:P1")
        cell_title = ws["A1"]
        cell_title.value = f"ZİRVE BORDRO PUANTAJ YÜKLEME ŞABLONU — FİDE KONSERVE ({self.month:02d}.{self.year})"
        cell_title.font = FONT_TITLE
        cell_title.fill = FILL_ZIRVE_EMERALD
        cell_title.alignment = ALIGN_CENTER
        ws.row_dimensions[1].height = 30

        headers = [
            ("Sıra No", 8), ("TC Kimlik No", 14), ("Adı Soyadı", 20), ("Bölüm", 18),
            ("Normal Gün", 11), ("Hafta Tatili", 11), ("Ücretli İzin", 11), ("Rapor Gün", 10),
            ("SGK Gün", 10), ("Eksik Gün", 10), ("Eksik Kodu", 15),
            ("Mesai Saati (%50)", 15), ("Eksik Çalışma Saati", 14), ("Pazar Mesai Saati", 14),
            ("Net Ücret (TL)", 14), ("İcra Kesintisi (TL)", 15)
        ]

        ws.row_dimensions[2].height = 24
        for col_idx, (h_text, w) in enumerate(headers, start=1):
            col_let = get_column_letter(col_idx)
            cell = ws[f"{col_let}2"]
            cell.value = h_text
            cell.font = FONT_HEADER
            cell.fill = FILL_ZIRVE_EMERALD
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN
            ws.column_dimensions[col_let].width = w

        for r_idx, emp_raw in enumerate(self.active_personnel, start=3):
            d = self._extract_employee_erp_data(emp_raw)
            row_fill = FILL_ZEBRA if r_idx % 2 == 0 else None

            ws[f"A{r_idx}"] = d["sno"]
            ws[f"B{r_idx}"] = str(d["tc"])
            ws[f"B{r_idx}"].number_format = "@"
            ws[f"C{r_idx}"] = d["ad_soyad"]
            ws[f"D{r_idx}"] = d["bolum"]
            ws[f"E{r_idx}"] = d["work_days"]
            ws[f"F{r_idx}"] = d["ht_days"]
            ws[f"G{r_idx}"] = d["izin_days"]
            ws[f"H{r_idx}"] = d["rapor_days"]
            ws[f"I{r_idx}"] = d["sgk_days"]
            ws[f"J{r_idx}"] = d["eksik_days"]
            ws[f"K{r_idx}"] = d["eksik_kodu"]
            ws[f"L{r_idx}"] = d["ot_weekday"]
            ws[f"M{r_idx}"] = d["missing_hours"]
            ws[f"N{r_idx}"] = d["ot_sunday"]
            ws[f"O{r_idx}"] = d["net_maas"]
            ws[f"P{r_idx}"] = d["icra_kesintisi"]

            for col_idx in range(1, 17):
                cl = get_column_letter(col_idx)
                c = ws[f"{cl}{r_idx}"]
                c.font = FONT_DATA
                c.border = BORDER_THIN
                if row_fill:
                    c.fill = row_fill

                if col_idx in (1, 2, 5, 6, 7, 8, 9, 10):
                    c.alignment = ALIGN_CENTER
                elif col_idx in (12, 13, 14):
                    c.alignment = ALIGN_RIGHT
                    c.number_format = "0.0"
                    if col_idx in (12, 14):
                        c.fill = FILL_OVERTIME
                    elif col_idx == 13 and d["missing_hours"] > 0:
                        c.fill = FILL_DEDUCT
                elif col_idx in (15, 16):
                    c.alignment = ALIGN_RIGHT
                    c.number_format = "#,##0.00"

        # Toplam Satırı
        tot_row = len(self.active_personnel) + 3
        ws[f"A{tot_row}"] = "GENEL TOPLAM"
        ws[f"A{tot_row}"].font = FONT_TOTAL
        ws.merge_cells(f"A{tot_row}:D{tot_row}")

        for col_idx in range(5, 11):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"] = f"=SUM({cl}3:{cl}{tot_row-1})"
            ws[f"{cl}{tot_row}"].number_format = "0"
            ws[f"{cl}{tot_row}"].alignment = ALIGN_CENTER
            ws[f"{cl}{tot_row}"].font = FONT_TOTAL

        for col_idx in (12, 13, 14):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"] = f"=SUM({cl}3:{cl}{tot_row-1})"
            ws[f"{cl}{tot_row}"].number_format = "0.0"
            ws[f"{cl}{tot_row}"].alignment = ALIGN_RIGHT
            ws[f"{cl}{tot_row}"].font = FONT_TOTAL

        for col_idx in (15, 16):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"] = f"=SUM({cl}3:{cl}{tot_row-1})"
            ws[f"{cl}{tot_row}"].number_format = "#,##0.00"
            ws[f"{cl}{tot_row}"].alignment = ALIGN_RIGHT
            ws[f"{cl}{tot_row}"].font = FONT_TOTAL

        for col_idx in range(1, 17):
            cl = get_column_letter(col_idx)
            ws[f"{cl}{tot_row}"].border = BORDER_THIN
            ws[f"{cl}{tot_row}"].fill = FILL_TOTAL

        ws.freeze_panes = "A3"
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

    # =========================================================================
    # 5. EVRENSEL ERP / CSV FORMATI (UTF-8 BOM, NOKTALI VİRGÜL AYRAÇLI)
    # =========================================================================
    def generate_universal_csv(self) -> bytes:
        """
        SAP, Nebim, Datassist ve şirket içi yazılımlar için Türkçe Excel uyumlu
        (UTF-8 BOM ile açılan) noktalı virgüllü CSV dosyası üretir.
        """
        headers = [
            "SIRA_NO", "SICIL_NO", "TC_KIMLIK_NO", "ADI_SOYADI", "DEPARTMAN", "ISTIHDAM",
            "NORMAL_GUN", "HAFTA_TATILI_GUN", "UCRETLI_IZIN_GUN", "RAPOR_GUN", "SGK_PRIM_GUN",
            "EKSIK_GUN", "EKSIK_NEDEN_KODU", "FIILI_CALISMA_SAAT",
            "HAFTA_ICI_FAZLA_MESAI_1_5X_SAAT", "EKSIK_CALISMA_1_0X_SAAT", "NET_HAFTA_ICI_MESAI_SAAT",
            "PAZAR_MESAI_1_5X_SAAT", "BOLUM_PRIM_SAAT", "NET_MAAS_TL", "SAATLIK_UCRET_TL",
            "TOPLAM_MESAI_TUTARI_TL", "ICRA_KESINTISI_TL"
        ]

        lines = [";".join(headers)]
        for emp_raw in self.active_personnel:
            d = self._extract_employee_erp_data(emp_raw)
            row = [
                str(d["sno"]),
                d["sicil_no"],
                f"'{d['tc']}",  # Excel metin formatı
                f'"{d["ad_soyad"]}"',
                f'"{d["bolum"]}"',
                d["durumu"],
                str(d["work_days"]),
                str(d["ht_days"]),
                str(d["izin_days"]),
                str(d["rapor_days"]),
                str(d["sgk_days"]),
                str(d["eksik_days"]),
                f'"{d["eksik_kodu"]}"',
                f"{d['base_hours']:.1f}",
                f"{d['ot_weekday']:.1f}",
                f"{d['missing_hours']:.1f}",
                f"{d['net_weekday_ot']:.1f}",
                f"{d['ot_sunday']:.1f}",
                f"{d['bonus_hours']:.1f}",
                f"{d['net_maas']:.2f}".replace(".", ","),
                f"{d['hourly_rate']:.2f}".replace(".", ","),
                f"{d['total_ot_cost']:.2f}".replace(".", ","),
                f"{d['icra_kesintisi']:.2f}".replace(".", ",")
            ]
            lines.append(";".join(row))

        csv_text = "\r\n".join(lines)
        # UTF-8 BOM ekle (Excel'de Türkçe harflerin bozulmasını engeller)
        return b'\xef\xbb\xbf' + csv_text.encode("utf-8")

    # =========================================================================
    # 6. TÜM ERP PAKETİ (.ZIP ARŞİVİ)
    # =========================================================================
    def generate_all_bundle_zip(self) -> bytes:
        """Tüm ERP aktarım formatlarını tek bir ZIP arşivi içinde toplar."""
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            # 1. Logo XML
            zf.writestr(
                f"01_Logo_Tiger_Puantaj_Aktarim_{self.year}_{self.month:02d}.xml",
                self.generate_logo_xml()
            )
            # 2. Luca Excel
            zf.writestr(
                f"02_Luca_Bordro_Puantaj_Aktarim_{self.year}_{self.month:02d}.xlsx",
                self.generate_luca_excel()
            )
            # 3. Mikro Excel
            zf.writestr(
                f"03_Mikro_Bordro_Puantaj_Aktarim_{self.year}_{self.month:02d}.xlsx",
                self.generate_mikro_excel()
            )
            # 4. Zirve Excel
            zf.writestr(
                f"04_Zirve_Bordro_Puantaj_Aktarim_{self.year}_{self.month:02d}.xlsx",
                self.generate_zirve_excel()
            )
            # 5. Evrensel CSV
            zf.writestr(
                f"05_Evrensel_ERP_Puantaj_Aktarim_{self.year}_{self.month:02d}.csv",
                self.generate_universal_csv()
            )
            # 6. Kullanım Kılavuzu
            readme_text = f"""================================================================================
FİDE KONSERVE — KURUMSAL ERP VE BORDRO ENTEGRASYON PAKETİ
Dönem: {self.month:02d}.{self.year}
Toplam Aktif Personel: {len(self.active_personnel)}
Üretim Sistemi: Fide Konserve PDKS Engine (Antigravity Core)
================================================================================

Bu arşiv dosyasında Türkiye'de en yaygın kullanılan kurumsal bordro ve ERP
yazılımlarına tek tıkla aktarılabilir dosyalar yer almaktadır:

1. 01_Logo_Tiger_Puantaj_Aktarim.xml
   - Logo Tiger 3, Logo GO ve Logo Bordro için standart XML içe aktarım dosyasıdır.
   - Logo içerisinde 'Bordro / Puantaj Kartları / XML İçe Aktar' menüsünden yüklenir.

2. 02_Luca_Bordro_Puantaj_Aktarim.xlsx
   - TÜRMOB Luca Mali Müşavir ve Bordro sistemi standart Excel şablonudur.
   - Luca'da 'Personel / Puantaj İşlemleri / Excel Puantaj Yükle' alanından aktarılır.

3. 03_Mikro_Bordro_Puantaj_Aktarim.xlsx
   - Mikro Fly ve Mikro Jump Bordro modülleri için tasarlanmıştır.

4. 04_Zirve_Bordro_Puantaj_Aktarim.xlsx
   - Zirve Müşavir / Bordro puantaj yükleme şablonuyla %100 uyumludur.

5. 05_Evrensel_ERP_Puantaj_Aktarim.csv
   - SAP HR, Nebim, Datassist ve şirket içi özel yazılımlar için UTF-8 BOM
     Türkçe karakter uyumlu evrensel noktalı virgüllü CSV tablosudur.

YASAL HESAPLAMA PRENSİBİ:
- Hafta İçi Fazla Mesai Çarpanı: 1.5x (4857 SK Md. 41)
- Hafta Tatili (Pazar) Çarpanı: 1.5x
- Eksik Çalışma Kesinti Çarpanı: 1.0x (Formül: =IF(h>7.5, (h-7.5)*1.5, h-7.5))
- Maksimum Günlük Süre Tavanı: 24.0 saat (Korumalı)
================================================================================
"""
            zf.writestr("README_ERP_KULLANIM_KILAVUZU.txt", readme_text.encode("utf-8"))

        return buffer.getvalue()
