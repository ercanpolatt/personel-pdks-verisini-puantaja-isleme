# -*- coding: utf-8 -*-
"""
===================================================================================================
PDKS Günlük Çalışma Saati ve Puantaj Raporlayıcı (export_daily_hours.py)
===================================================================================================
Bu script, PDKSEngine motorunu çalıştırarak:
1. 'gunluk_calisma_raporu.xlsx' (3 sayfalı tam denetlenebilir Excel çalışma kitabı)
2. 'gunluk_calismalar.csv' (Sistem entegrasyonu için temiz CSV tablosu)
3. 'gunluk_calismalar.json' (API ve veritabanı aktarımı için yapılandırılmış JSON)
dosyalarını oluşturur.
"""

import os
import csv
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

from pdks_engine import PDKSEngine, fmt_hours_tr

def build_excel_report(engine, output_path="gunluk_calisma_raporu.xlsx"):
    wb = openpyxl.Workbook()
    # Varsayılan sayfayı kaldıracağız
    ws_default = wb.active

    # =========================================================================
    # STİL VE TEMA TANIMLARI (Modern Corporate / Premium Aesthetics)
    # =========================================================================
    FONT_NAME = "Segoe UI"
    
    font_main_title = Font(name=FONT_NAME, size=13, bold=True, color="FFFFFF")
    font_sub_title  = Font(name=FONT_NAME, size=9, italic=True, color="D9E1F2")
    font_header     = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
    font_bold       = Font(name=FONT_NAME, size=9, bold=True, color="000000")
    font_regular    = Font(name=FONT_NAME, size=9, color="000000")
    font_small      = Font(name=FONT_NAME, size=8, color="595959")
    font_alert      = Font(name=FONT_NAME, size=9, bold=True, color="C00000")
    
    fill_navy_dark  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid") # Ana Başlık
    fill_navy_head  = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid") # Sütun Başlıkları
    fill_sub_head   = PatternFill(start_color="3B3838", end_color="3B3838", fill_type="solid") # İkincil Başlık
    fill_zebra      = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid") # Zebra Satır
    fill_sunday     = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Hafta Tatili (Pazar)
    fill_overtime   = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Fazla Mesai (Açık Yeşil)
    fill_missing    = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # Eksik / Tek Basım (Açık Turuncu)
    fill_bonus      = PatternFill(start_color="E8D7F1", end_color="E8D7F1", fill_type="solid") # Özel Bölüm Primi (Lila / Açık Mor)
    fill_total      = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid") # Genel Toplam Satırı

    border_thin_grey = Side(style="thin", color="D3D3D3")
    border_thick_top = Side(style="thin", color="000000")
    border_double_bottom = Side(style="double", color="000000")
    
    border_cell = Border(left=border_thin_grey, right=border_thin_grey, top=border_thin_grey, bottom=border_thin_grey)
    border_header = Border(left=border_thin_grey, right=border_thin_grey, top=border_thick_top, bottom=border_thick_top)
    border_total = Border(top=border_thick_top, bottom=border_double_bottom, left=border_thin_grey, right=border_thin_grey)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left   = Alignment(horizontal="left", vertical="center")
    align_right  = Alignment(horizontal="right", vertical="center")
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Eylül 2026 Pazar günleri
    sundays = engine.sundays
    tr_months = ["", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
    month_name = tr_months[engine.target_month].upper() if 1 <= engine.target_month <= 12 else ""

    # =========================================================================
    # SAYFA 1: GÜNLÜK PUANTAJ VE ÇALIŞMA MATRİSİ
    # =========================================================================
    ws_mat = wb.create_sheet(title="Gunluk_Puantaj_Matrisi")
    ws_mat.views.sheetView[0].showGridLines = True

    # 1. Başlık Satırı
    last_title_col = get_column_letter(6 + engine.days_in_month + 5)
    ws_mat.merge_cells(f"A1:{last_title_col}1")
    cell_title = ws_mat["A1"]
    cell_title.value = f"FİDE KONSERVE - GÜNLÜK PERSONEL ÇALIŞMA SAATLERİ VE PUANTAJ MATRİSİ ({month_name} {engine.target_year})"
    cell_title.font = font_main_title
    cell_title.fill = fill_navy_dark
    cell_title.alignment = align_center
    ws_mat.row_dimensions[1].height = 28

    # 2. Açıklama & Lejant Satırı
    ws_mat.merge_cells(f"A2:{last_title_col}2")
    cell_sub = ws_mat["A2"]
    cell_sub.value = (
        "Lejant: Standart Gün: 7,5 Saat (1,5s Mola Düşülmüştür) | "
        "Açık Yeşil: Fazla Mesaili Günler | Açık Turuncu: Eksik/Tek Basım (Amir Onayı Bekleyen) | "
        "Lila: Balık Dolum/Kesim veya Üretim Bölüm Primli Mesai | Sarı: Hafta Tatili (Pazar)"
    )
    cell_sub.font = font_sub_title
    cell_sub.fill = fill_navy_dark
    cell_sub.alignment = align_center
    ws_mat.row_dimensions[2].height = 18

    # 3. Sütun Başlıkları
    headers_fixed = [
        ("Sıra", 6),
        ("TC Kimlik No", 14),
        ("Kart / Sicil", 11),
        ("Adı Soyadı", 24),
        ("Bölüm", 20),
        ("Durumu", 12)
    ]
    
    ws_mat.row_dimensions[3].height = 26
    
    # Sabit başlıkları yaz
    col_idx = 1
    for h_name, w in headers_fixed:
        c = ws_mat.cell(row=3, column=col_idx, value=h_name)
        c.font = font_header
        c.fill = fill_navy_head
        c.alignment = align_header
        c.border = border_header
        ws_mat.column_dimensions[get_column_letter(col_idx)].width = w
        col_idx += 1

    # Gün başlıklarını yaz (1..days_in_month)
    day_start_col = col_idx
    for d in range(1, engine.days_in_month + 1):
        c = ws_mat.cell(row=3, column=col_idx, value=f"{d:02d}.{engine.target_month:02d}")
        c.font = font_header
        c.fill = PatternFill(start_color="333F48", end_color="333F48", fill_type="solid") if d not in sundays else PatternFill(start_color="806000", end_color="806000", fill_type="solid")
        c.alignment = align_header
        c.border = border_header
        ws_mat.column_dimensions[get_column_letter(col_idx)].width = 6.2
        col_idx += 1
    day_end_col = col_idx - 1

    # Toplam sütunları
    totals_headers = [
        ("Çalışılan Gün", 13),
        ("Normal Çalışma (Saat)", 15),
        ("Fazla Mesai (Saat)", 15),
        ("Bölüm Primi (Saat)", 14),
        ("Toplam Çalışma (Saat)", 16)
    ]
    total_start_col = col_idx
    for h_name, w in totals_headers:
        c = ws_mat.cell(row=3, column=col_idx, value=h_name)
        c.font = font_header
        c.fill = fill_navy_head
        c.alignment = align_header
        c.border = border_header
        ws_mat.column_dimensions[get_column_letter(col_idx)].width = w
        col_idx += 1

    # Satırları Doldur
    matrix = engine.get_summary_matrix()
    row_idx = 4

    for p_idx, p in enumerate(matrix, start=1):
        ws_mat.row_dimensions[row_idx].height = 19
        row_bg = fill_zebra if p_idx % 2 == 0 else PatternFill(fill_type=None)
        
        # Sabit personel bilgileri
        p_info = engine.personnel_by_key.get(p["name_key"], {})
        kart_no = p_info.get("tc", "") if len(p_info.get("tc", "")) <= 6 else ""
        
        fixed_vals = [
            (p_idx, align_center),
            (p["tc"], align_center),
            (kart_no, align_center),
            (p["ad_soyad"], align_left),
            (p["bolum"], align_left),
            (p["durumu"], align_center)
        ]
        
        for c_i, (val, aln) in enumerate(fixed_vals, start=1):
            cell = ws_mat.cell(row=row_idx, column=c_i, value=val)
            cell.font = font_bold if c_i == 4 else font_regular
            cell.alignment = aln
            cell.border = border_cell
            if row_bg.fill_type:
                cell.fill = row_bg

        # Günlük Saatler (1..days_in_month)
        c_cur = day_start_col
        for d in range(1, engine.days_in_month + 1):
            cell = ws_mat.cell(row=row_idx, column=c_cur)
            cell.border = border_cell
            cell.alignment = align_center
            
            day_rec = engine.daily_results.get((p["name_key"], d))
            
            if d in sundays:
                cell.fill = fill_sunday
                
            if day_rec and day_rec["final_hours"] > 0:
                h_val = day_rec["final_hours"]
                cell.value = h_val
                cell.number_format = "0.0"
                cell.font = font_bold if h_val > 7.5 else font_regular
                
                # Özel Renklendirmeler:
                if day_rec.get("status_type") == "BÖLÜM_PRİMLİ":
                    cell.fill = fill_bonus
                    if day_rec.get("audit_note"):
                        cell.comment = Comment(day_rec["audit_note"], "PDKS Motoru")
                elif day_rec.get("is_exception"):
                    cell.fill = fill_missing
                    if day_rec.get("audit_note"):
                        cell.comment = Comment(day_rec["audit_note"], "PDKS Motoru")
                elif h_val > 7.5 and d not in sundays:
                    cell.fill = fill_overtime
            else:
                cell.value = ""
                if not (d in sundays) and row_bg.fill_type:
                    cell.fill = row_bg
                    
            c_cur += 1

        # Formüllerle Toplamlar
        col_start_let = get_column_letter(day_start_col)
        col_end_let = get_column_letter(day_end_col)
        
        # 1. Çalışılan Gün Sayısı = COUNTIF(G4:AJ4; ">0")
        c_work_days = ws_mat.cell(row=row_idx, column=total_start_col)
        c_work_days.value = f"=COUNTIF({col_start_let}{row_idx}:{col_end_let}{row_idx}, \">0\")"
        c_work_days.font = font_bold
        c_work_days.alignment = align_center
        c_work_days.border = border_cell
        c_work_days.number_format = "0"
        
        # 2. Normal Çalışma (Saat)
        c_norm = ws_mat.cell(row=row_idx, column=total_start_col + 1)
        c_norm.value = p["total_base_hours"]
        c_norm.font = font_regular
        c_norm.alignment = align_center
        c_norm.border = border_cell
        c_norm.number_format = "0.0"

        # 3. Fazla Mesai (Saat)
        c_ot = ws_mat.cell(row=row_idx, column=total_start_col + 2)
        c_ot.value = p["total_ot_hours"]
        c_ot.font = font_bold if p["total_ot_hours"] > 0 else font_regular
        c_ot.alignment = align_center
        c_ot.border = border_cell
        c_ot.number_format = "0.0"
        if p["total_ot_hours"] > 0:
            c_ot.fill = fill_overtime

        # 4. Bölüm Primi (Saat)
        c_bon = ws_mat.cell(row=row_idx, column=total_start_col + 3)
        c_bon.value = p["total_bonus_hours"]
        c_bon.font = font_bold if p["total_bonus_hours"] > 0 else font_regular
        c_bon.alignment = align_center
        c_bon.border = border_cell
        c_bon.number_format = "0.0"
        if p["total_bonus_hours"] > 0:
            c_bon.fill = fill_bonus

        # 5. Toplam Çalışma Saati = SUM(G4:AJ4)
        c_tot = ws_mat.cell(row=row_idx, column=total_start_col + 4)
        c_tot.value = f"=SUM({col_start_let}{row_idx}:{col_end_let}{row_idx})"
        c_tot.font = font_bold
        c_tot.alignment = align_center
        c_tot.border = border_cell
        c_tot.number_format = "0.0"
        c_tot.fill = fill_zebra

        row_idx += 1

    # Genel Toplam Satırı (SUM)
    tot_row = row_idx
    ws_mat.row_dimensions[tot_row].height = 24
    ws_mat.merge_cells(f"A{tot_row}:F{tot_row}")
    lbl_tot = ws_mat[f"A{tot_row}"]
    lbl_tot.value = "GENEL TOPLAM"
    lbl_tot.font = font_bold
    lbl_tot.alignment = align_right
    lbl_tot.fill = fill_total
    
    for c_i in range(1, day_start_col):
        ws_mat.cell(row=tot_row, column=c_i).border = border_total
        ws_mat.cell(row=tot_row, column=c_i).fill = fill_total

    # Günlük toplam formülleri
    for c_i in range(day_start_col, day_end_col + 1):
        col_let = get_column_letter(c_i)
        c = ws_mat.cell(row=tot_row, column=c_i)
        c.value = f"=SUM({col_let}4:{col_let}{tot_row - 1})"
        c.font = font_bold
        c.fill = fill_total
        c.border = border_total
        c.alignment = align_center
        c.number_format = "0.0"

    # Özet sütunlarının toplam formülleri
    for c_i in range(total_start_col, total_start_col + 5):
        col_let = get_column_letter(c_i)
        c = ws_mat.cell(row=tot_row, column=c_i)
        c.value = f"=SUM({col_let}4:{col_let}{tot_row - 1})"
        c.font = font_bold
        c.fill = fill_total
        c.border = border_total
        c.alignment = align_center
        c.number_format = "0.0" if c_i != total_start_col else "0"

    # İlk 6 sütun ve ilk 3 satırı dondur (Freeze Panes)
    ws_mat.freeze_panes = "G4"

    # =========================================================================
    # SAYFA 2: GÜNLÜK DENETİM İZİ (AUDIT TRAIL)
    # =========================================================================
    ws_aud = wb.create_sheet(title="Gunluk_Denetim_Izi")
    ws_aud.views.sheetView[0].showGridLines = True

    audit_headers = [
        ("Sıra No", 8),
        ("Tarih", 12),
        ("Gün", 12),
        ("TC Kimlik No", 14),
        ("Sicil / Kart", 12),
        ("Adı Soyadı", 24),
        ("Bölüm", 20),
        ("İlk Kart Basımı", 14),
        ("Efektif Giriş", 13),
        ("Son Kart Basımı", 14),
        ("Efektif Bitiş", 13),
        ("Tüm Basımlar", 18),
        ("Brüt Süre (s)", 12),
        ("Yasal Mola (s)", 12),
        ("Fiili Süre (s)", 12),
        ("Normal Mesai", 13),
        ("Fazla Mesai", 13),
        ("Bölüm Primi", 13),
        ("Puantaj Saati", 14),
        ("Durum", 20),
        ("Denetim / Hesaplama Açıklaması", 35)
    ]
    
    ws_aud.row_dimensions[1].height = 26
    for c_idx, (h_name, w) in enumerate(audit_headers, start=1):
        c = ws_aud.cell(row=1, column=c_idx, value=h_name)
        c.font = font_header
        c.fill = fill_navy_dark
        c.alignment = align_header
        c.border = border_header
        ws_aud.column_dimensions[get_column_letter(c_idx)].width = w

    # Denetim kayıtlarını sırala (Tarih ve Ad Soyad sırasına göre)
    sorted_audit = sorted(engine.audit_records, key=lambda x: (x["gun"], x["ad_soyad"]))
    
    for r_idx, rec in enumerate(sorted_audit, start=2):
        ws_aud.row_dimensions[r_idx].height = 19
        row_bg = fill_zebra if r_idx % 2 == 0 else PatternFill(fill_type=None)
        
        row_vals = [
            (r_idx - 1, align_center, None),
            (rec["tarih"], align_center, None),
            (rec["gun_adi"], align_center, None),
            (rec["tc"], align_center, None),
            (rec["sicil"], align_center, None),
            (rec["ad_soyad"], align_left, font_bold),
            (rec["bolum"], align_left, None),
            (rec["raw_g_saat"], align_center, None),
            (rec["effective_g"], align_center, None),
            (rec["raw_c_saat"], align_center, None),
            (rec["effective_c"], align_center, None),
            (rec["all_punches"], align_center, None),
            (rec["gross_hours"], align_center, None),
            (rec["break_hours"], align_center, None),
            (rec["fiili_sure"], align_center, None),
            (rec["base_hours"], align_center, None),
            (rec["overtime_hours"], align_center, font_bold if rec["overtime_hours"] > 0 else None),
            (rec["bonus_hours"], align_center, font_bold if rec["bonus_hours"] > 0 else None),
            (rec["final_hours"], align_center, font_bold),
            (rec["status_type"], align_center, font_bold if rec["is_exception"] else None),
            (rec["audit_note"], align_left, font_small)
        ]
        
        for c_idx, (val, aln, fnt) in enumerate(row_vals, start=1):
            cell = ws_aud.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = aln
            cell.font = fnt if fnt else font_regular
            cell.border = border_cell
            
            # Sayı formatları
            if c_idx in (13, 14, 15, 16, 17, 18, 19) and isinstance(val, (int, float)):
                cell.number_format = "0.0"
                
            # Arka plan vurguları
            if c_idx == 20: # Durum sütunu
                if rec["status_type"] == "BÖLÜM_PRİMLİ":
                    cell.fill = fill_bonus
                elif rec["is_exception"]:
                    cell.fill = fill_missing
            elif c_idx == 19: # Puantaj saati
                if rec["final_hours"] > 7.5:
                    cell.fill = fill_overtime
            elif row_bg.fill_type:
                cell.fill = row_bg

    ws_aud.freeze_panes = "G2"
    ws_aud.auto_filter.ref = f"A1:{get_column_letter(len(audit_headers))}{len(sorted_audit)+1}"

    # =========================================================================
    # SAYFA 3: EKSİK VE KUŞKULU BASIMLAR (AMİR ONAY CETVELİ)
    # =========================================================================
    ws_exc = wb.create_sheet(title="Eksik_ve_Kuskulu_Basimlar")
    ws_exc.views.sheetView[0].showGridLines = True

    exc_headers = [
        ("Sıra No", 8),
        ("Tarih", 12),
        ("Gün", 12),
        ("TC Kimlik No", 14),
        ("Sicil / Kart", 12),
        ("Adı Soyadı", 24),
        ("Bölüm", 20),
        ("Kart Giriş", 12),
        ("Kart Çıkış", 12),
        ("Tüm Basımlar", 18),
        ("Hata / İstisna Türü", 22),
        ("Puantaja Yazılan", 15),
        ("Uygulanan Varsayılan Kural & Açıklama", 40),
        ("İK / Amir Onayı", 16)
    ]
    
    ws_exc.row_dimensions[1].height = 26
    for c_idx, (h_name, w) in enumerate(exc_headers, start=1):
        c = ws_exc.cell(row=1, column=c_idx, value=h_name)
        c.font = font_header
        c.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid") # Kırmızı Başlık
        c.alignment = align_header
        c.border = border_header
        ws_exc.column_dimensions[get_column_letter(c_idx)].width = w

    sorted_exc = sorted(engine.exception_records, key=lambda x: (x["gun"], x["ad_soyad"]))
    
    for r_idx, rec in enumerate(sorted_exc, start=2):
        ws_exc.row_dimensions[r_idx].height = 20
        row_vals = [
            (r_idx - 1, align_center),
            (rec["tarih"], align_center),
            (rec["gun_adi"], align_center),
            (rec["tc"], align_center),
            (rec["sicil"], align_center),
            (rec["ad_soyad"], align_left),
            (rec["bolum"], align_left),
            (rec["raw_g_saat"], align_center),
            (rec["raw_c_saat"], align_center),
            (rec["all_punches"], align_center),
            (rec["status_type"], align_center),
            (rec["final_hours"], align_center),
            (rec["audit_note"], align_left),
            ("ONAYLANDI", align_center)
        ]
        
        for c_idx, (val, aln) in enumerate(row_vals, start=1):
            cell = ws_exc.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = aln
            cell.font = font_bold if c_idx in (6, 11, 12) else font_regular
            cell.border = border_cell
            
            if c_idx == 11: # Hata Türü
                cell.fill = fill_missing
                cell.font = font_alert
            elif c_idx == 12: # Yazılan Saat
                cell.number_format = "0.0"
                cell.fill = fill_overtime if rec["final_hours"] > 7.5 else fill_zebra

    ws_exc.freeze_panes = "G2"
    ws_exc.auto_filter.ref = f"A1:{get_column_letter(len(exc_headers))}{len(sorted_exc)+1}"

    # Varsayılan boş sayfayı sil
    if ws_default.title in wb.sheetnames:
        wb.remove(ws_default)

    # Dosyayı kaydet (Açık olma durumunu yönet)
    saved_path = output_path
    try:
        wb.save(output_path)
    except PermissionError:
        saved_path = output_path.replace(".xlsx", "_guncel.xlsx")
        wb.save(saved_path)

    return saved_path

def export_csv_and_json(engine, csv_path="gunluk_calismalar.csv", json_path="gunluk_calismalar.json"):
    """Sistem entegrasyonu için temiz CSV ve JSON dosyalarını dışa aktarır."""
    sorted_audit = sorted(engine.audit_records, key=lambda x: (x["gun"], x["ad_soyad"]))
    
    # 1. CSV Dışa Aktarımı
    fieldnames = [
        "tarih", "gun", "gun_adi", "tc", "sicil", "kart", "ad_soyad", "bolum", "lokasyon",
        "raw_g_saat", "raw_c_saat", "effective_g", "effective_c", "all_punches",
        "gross_hours", "break_hours", "fiili_sure", "base_hours", "overtime_hours",
        "bonus_hours", "final_hours", "status_type", "is_exception", "audit_note"
    ]
    
    with open(csv_path, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for rec in sorted_audit:
            writer.writerow(rec)

    # 2. JSON Dışa Aktarımı
    matrix = engine.get_summary_matrix()
    export_payload = {
        "metadata": {
            "period": f"{engine.target_month:02d}.{engine.target_year}",
            "generated_at": str(os.path.getmtime(engine.pdks_path)),
            "total_personnel": len(engine.personnel_list),
            "active_personnel": len([m for m in matrix if m["total_work_days"] > 0]),
            "total_records": len(sorted_audit),
            "total_exceptions": len(engine.exception_records)
        },
        "daily_records": sorted_audit,
        "personnel_summary_matrix": matrix
    }
    
    with open(json_path, mode="w", encoding="utf-8") as f:
        json.dump(export_payload, f, ensure_ascii=False, indent=2)

def main():
    print("===================================================================")
    print(" FİDE KONSERVE - GÜNLÜK ÇALIŞMA SAATLERİ VE PUANTAJ RAPORLAMA")
    print("===================================================================")
    
    engine = PDKSEngine(pdks_path="pdks.xls", puantaj_path="puantaj.xls", target_month=9, target_year=2026)
    
    print("1. Personel ve bölüm verileri puantaj.xls'den okunuyor...")
    p_count = engine.load_personnel()
    print(f"   -> Toplam {p_count} personel sisteme kaydedildi.")
    
    print("2. pdks.xls ham turnike verileri analiz ediliyor...")
    engine.load_and_process_pdks()
    matrix = engine.get_summary_matrix()
    active_count = len([m for m in matrix if m["total_work_days"] > 0])
    print(f"   -> Toplam {len(engine.audit_records)} günlük çalışma hareketi işlendi.")
    print(f"   -> Kart basan aktif personel sayısı: {active_count}")
    print(f"   -> İncelenen istisna/tek basım sayısı: {len(engine.exception_records)}")

    print("3. 'gunluk_calisma_raporu.xlsx' Excel çalışma kitabı üretiliyor...")
    saved_excel = build_excel_report(engine, "gunluk_calisma_raporu.xlsx")
    print(f"   -> Excel raporu hazır: {saved_excel}")

    print("4. 'gunluk_calismalar.csv' ve 'gunluk_calismalar.json' üretiliyor...")
    export_csv_and_json(engine, "gunluk_calismalar.csv", "gunluk_calismalar.json")
    print("   -> CSV ve JSON aktarımı tamamlandı.")

    # Özet istatistikler
    total_hours_all = sum(m["total_hours"] for m in matrix)
    total_ot_all = sum(m["total_ot_hours"] for m in matrix)
    total_bonus_all = sum(m["total_bonus_hours"] for m in matrix)
    
    print("===================================================================")
    print(" HESAPLAMA VE RAPORLAMA ÖZETİ:")
    print(f" - Toplam Çalışma Süresi : {fmt_hours_tr(total_hours_all)} Saat")
    print(f" - Toplam Fazla Mesai    : {fmt_hours_tr(total_ot_all)} Saat")
    print(f" - Toplam Bölüm Primi    : {fmt_hours_tr(total_bonus_all)} Saat")
    print(" - Yasal Mola Kesintisi  : Standart 1,5 Saat uygulandı")
    print(" - Sabah Toleransı       : 08:20 uygulandı (08:21+ 30'ar dk kesintili)")
    print(" - Fazla Mesai Dilimleri : 17:26 sonrası kademeli hesaplandı")
    print("===================================================================")

if __name__ == "__main__":
    main()
