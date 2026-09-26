# -*- coding: utf-8 -*-
"""
===================================================================================================
FİDE KONSERVE
Puantaj ve PDKS Entegrasyonu, Akıllı Vardiya Analizi ve Modern Excel Raporlama Sistemi (builder.py)
===================================================================================================

Bu modül, çekirdek hesaplama motoru (PDKSEngine) üzerinden ham PDKS turnike verilerini ve
bordro/personel tablolarını uçtan uca işleyerek hatasız, canlı formüllü ve 10 sayfalı
modern bir Excel (puantaj.xlsx) çalışma kitabı üretir.
"""

import os
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation

from pdks_engine import (
    PDKSEngine,
    clean_str,
    clean_tc,
    clean_money,
    norm_name_key,
    fmt_hours_tr
)

def build_puantaj_workbook(
    pdks_path: str = "pdks.xls",
    puantaj_path: str = "puantaj.xls",
    target_month: int = 9,
    target_year: int = 2026,
    output_filename: str = "puantaj.xlsx"
) -> str:
    print("==================================================")
    print(" FİDE KONSERVE - PUANTAJ VE PDKS İŞLEME SİSTEMİ")
    print("==================================================")

    # Güvenli yedek kontrolü: Orijinal puantaj.xls yoksa yedekten kopyala (asla üzerine yazma)
    if os.path.exists("puantaj_backup.xls") and not os.path.exists(puantaj_path):
        try:
            shutil.copyfile("puantaj_backup.xls", puantaj_path)
            print("   -> 'puantaj.xls' bulunamadı, 'puantaj_backup.xls' üzerinden oluşturuldu.")
        except Exception as e:
            print(f"   -> Yedek kopyalama uyarısı: {e}")

    # =========================================================================
    # 1. ÇEKİRDEK HESAPLAMA MOTORUNUN ÇALIŞTIRILMASI
    # =========================================================================
    print(f"1. PDKSEngine motoru başlatılıyor ({target_month:02d}.{target_year})...")
    engine = PDKSEngine(
        pdks_path=pdks_path,
        puantaj_path=puantaj_path,
        target_month=target_month,
        target_year=target_year
    )

    p_count = engine.load_personnel()
    print(f"   -> Puantaj tablosundan {p_count} personel yüklendi.")

    engine.load_and_process_pdks()
    print(f"   -> PDKS hareketleri işlendi ({len(engine.audit_records)} hareket, "
          f"{len(engine.exception_records)} istisna, {len(engine.bonus_records)} bölüm primi).")

    engine.load_payroll_sheets()
    print(f"   -> Bordro ek tabloları okundu ({len(engine.rapor_list)} rapor, "
          f"{len(engine.icra_list)} icra, {len(engine.izin_list)} izin, {len(engine.bordro_list)} SGK bordro).")

    puantaj_rows = engine.personnel_list
    emp_pdks_daily = engine.daily_results

    # =========================================================================
    # 2. STİL VE TEMA TANIMLARI (Modern Corporate / Premium Aesthetics)
    # =========================================================================
    print(f"2. Modern '{output_filename}' çalışma kitabı oluşturuluyor...")
    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)

    FONT_FAMILY = "Segoe UI"
    font_title = Font(name=FONT_FAMILY, size=13, bold=True, color="1F4E79")
    font_hdr = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
    font_data = Font(name=FONT_FAMILY, size=9, bold=False, color="000000")
    font_total = Font(name=FONT_FAMILY, size=9, bold=True, color="000000")

    fill_navy = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_pazar = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_overtime = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_net = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_deduct = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_missing = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # Turuncu (Eksik Basım)
    fill_bonus = PatternFill(start_color="E1D5E7", end_color="E1D5E7", fill_type="solid")   # Açık Mor/Lila (Bölüm Primi)
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_total = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_hdr = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border_side = Side(border_style="thin", color="D3D3D3")
    thick_bottom = Side(border_style="medium", color="1F4E79")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_header = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thick_bottom)
    border_total = Border(top=thin_border_side, bottom=Side(border_style="double", color="000000"))

    tr_months = ["", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
    month_name = tr_months[engine.target_month].upper() if 1 <= engine.target_month <= 12 else ""

    # =========================================================================
    # SAYFA 1: Aylik_Puantaj (Dinamik Ay ve Canlı Formüller)
    # =========================================================================
    ws_p = wb_new.create_sheet(title="Aylik_Puantaj")
    ws_p.views.sheetView[0].showGridLines = True

    ws_p.merge_cells("A1:N1")
    ws_p["A1"] = f"FİDE KONSERVE GIDA SAN. VE TİC. A.Ş. - {month_name} {engine.target_year} AYLIK PUANTAJ VE HAKEDİŞ CETVELİ"
    ws_p["A1"].font = font_title
    ws_p["A1"].alignment = Alignment(horizontal="left", vertical="center")

    info_headers = [
        ("A", "Sıra No", 7),
        ("B", "TC Kimlik No", 14),
        ("C", "Adı Soyadı", 22),
        ("D", "Cinsiyet", 7),
        ("E", "İşletme Giriş", 11),
        ("F", "SGK Giriş", 11),
        ("G", "SGK Çıkış", 11),
        ("H", "SGK Statü", 11),
        ("I", "Kadro", 11),
        ("J", "Bölüm / İş", 16),
        ("K", "İkamet / Güzergah", 15),
        ("L", "Anlaşılan Net Maaş (TL)", 15),
        ("M", "Şirket", 13),
        ("N", "Mesai Durumu", 11),
    ]

    day_cols = []
    start_col_idx = 15
    for d in range(1, engine.days_in_month + 1):
        col_letter = get_column_letter(start_col_idx + d - 1)
        day_name = engine.get_day_name_for_day(d, short=True)
        is_pazar = engine.is_sunday(d)
        day_cols.append((col_letter, f"{d:02d}.{engine.target_month:02d}\n{day_name}", 6, is_pazar, d))

    calc_headers = [
        ("Toplam Çalışma Saati", 13, fill_overtime),
        ("Fiili Çalışılan Gün", 11, fill_navy),
        ("Hafta Tatili Günü", 11, fill_pazar),
        ("Ücretli İzin (Gün)", 11, fill_navy),
        ("Raporlu Gün", 11, fill_deduct),
        ("Toplam SGK Günü", 11, fill_net),
        ("Eksik Gün Sayısı", 11, fill_deduct),
        ("Eksik Gün Nedeni", 15, fill_navy),
        ("Hafta İçi Fazla Mesai (Saat)", 13, fill_overtime),
        ("Eksik Çalışma (Saat)", 12, fill_deduct),
        ("Pazar Mesai (Saat)", 13, fill_pazar),
        ("Maaş Hakedişi (TL)", 14, fill_navy),
        ("H.İçi Mesai Tutarı (TL)", 15, fill_overtime),
        ("Pazar Mesai Tutarı (TL)", 15, fill_pazar),
        ("Toplam Mesai Tutarı (TL)", 16, fill_overtime),
        ("Toplam Net Hakediş (TL)", 16, fill_net),
        ("Resmi Bordro SGK Net (TL)", 15, fill_navy),
        ("İcra Kesintisi (TL)", 13, fill_deduct),
        ("Elden Ödenecek Fark (TL)", 15, fill_net),
        ("Ödenecek Net Tutar (TL)", 15, fill_net),
    ]

    header_row = 3
    ws_p.row_dimensions[header_row].height = 28
    ws_p.row_dimensions[1].height = 24

    for col_let, text, width in info_headers:
        cell = ws_p[f"{col_let}{header_row}"]
        cell.value = text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_p.column_dimensions[col_let].width = width

    for col_let, text, width, is_pazar, d in day_cols:
        cell = ws_p[f"{col_let}{header_row}"]
        cell.value = text
        cell.font = font_hdr if not is_pazar else Font(name=FONT_FAMILY, size=8, bold=True, color="7F6000")
        cell.fill = fill_navy if not is_pazar else fill_pazar
        cell.alignment = align_hdr
        cell.border = border_header
        ws_p.column_dimensions[col_let].width = width

    cur_col_idx = start_col_idx + engine.days_in_month
    calc_col_letters = []
    for text, width, hfill in calc_headers:
        col_let = get_column_letter(cur_col_idx)
        calc_col_letters.append((col_let, text))
        cell = ws_p[f"{col_let}{header_row}"]
        cell.value = text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_p.column_dimensions[col_let].width = width
        cur_col_idx += 1

    first_day_col = day_cols[0][0]
    last_day_col = day_cols[-1][0]

    # Personel Satırlarını Doldur
    start_data_row = 4
    for idx, emp in enumerate(puantaj_rows):
        r_idx = start_data_row + idx
        ws_p.row_dimensions[r_idx].height = 19
        is_zebra = (idx % 2 == 1)
        row_fill = fill_zebra if is_zebra else None
        
        ws_p[f"A{r_idx}"] = emp["sno"]
        ws_p[f"B{r_idx}"] = emp["tc"]
        ws_p[f"C{r_idx}"] = emp["ad_soyad"]
        ws_p[f"D{r_idx}"] = emp.get("cinsiyet", "")
        ws_p[f"E{r_idx}"] = emp.get("isletme_giris", "")
        ws_p[f"F{r_idx}"] = emp.get("sgk_giris", "")
        ws_p[f"G{r_idx}"] = emp.get("sgk_cikis", "")
        ws_p[f"H{r_idx}"] = emp.get("sgk_durumu", "NORMAL")
        ws_p[f"I{r_idx}"] = emp.get("durumu", "MEVSİMLİK")
        ws_p[f"J{r_idx}"] = emp.get("bolum", "")
        ws_p[f"K{r_idx}"] = emp.get("ikamet", "")
        ws_p[f"L{r_idx}"] = emp.get("net_maas", 0.0)
        ws_p[f"M{r_idx}"] = emp.get("sirket", "FİDE KONSERVE")
        ws_p[f"N{r_idx}"] = emp.get("mesai_durumu", "ALIR")
        
        ws_p[f"A{r_idx}"].alignment = align_center
        ws_p[f"B{r_idx}"].alignment = align_center
        ws_p[f"B{r_idx}"].number_format = "@"
        ws_p[f"C{r_idx}"].alignment = align_left
        ws_p[f"D{r_idx}"].alignment = align_center
        ws_p[f"E{r_idx}"].alignment = align_center
        ws_p[f"F{r_idx}"].alignment = align_center
        ws_p[f"G{r_idx}"].alignment = align_center
        ws_p[f"H{r_idx}"].alignment = align_center
        ws_p[f"I{r_idx}"].alignment = align_center
        ws_p[f"J{r_idx}"].alignment = align_left
        ws_p[f"K{r_idx}"].alignment = align_left
        ws_p[f"L{r_idx}"].alignment = align_right
        ws_p[f"L{r_idx}"].number_format = "#,##0.00"
        ws_p[f"M{r_idx}"].alignment = align_center
        ws_p[f"N{r_idx}"].alignment = align_center
        
        name_k = emp["name_key"]
        emp_ot_weekday = 0.0
        emp_missing_hours = 0.0
        emp_ot_sunday = 0.0
        
        for col_let, text, width, is_pazar, d in day_cols:
            cell = ws_p[f"{col_let}{r_idx}"]
            cell.alignment = align_center
            cell.number_format = "0.0"
            if is_pazar:
                cell.fill = fill_pazar
            elif row_fill:
                cell.fill = row_fill
                
            rec = emp_pdks_daily.get((name_k, d))
            if rec and rec["final_hours"] > 0:
                hours = min(24.0, max(0.0, float(rec["final_hours"])))
                fazla = min(24.0, max(0.0, float(rec["overtime_hours"])))
                cell.value = hours
                
                # Görsel Renklendirme ve Hücre Notları
                if not is_pazar:
                    if rec.get("is_bonus"): # Bölüm Primi Alan Günler (Açık Mor/Lila)
                        cell.fill = fill_bonus
                        if rec.get("audit_note"):
                            cell.comment = Comment(rec["audit_note"], "PDKS Sistemi")
                    elif rec.get("is_exception"): # Eksik / Hatalı Basımlar (Turuncu)
                        cell.fill = fill_missing
                        if rec.get("audit_note"):
                            cell.comment = Comment(rec["audit_note"], "PDKS Sistemi")
                            
                if is_pazar:
                    emp_ot_sunday += hours
                else:
                    emp_ot_weekday += fazla
                    if hours < 7.5:
                        emp_missing_hours += round(7.5 - hours, 2)
            else:
                cell.value = 0
                
        c_tot_hours = calc_col_letters[0][0]
        c_work_days = calc_col_letters[1][0]
        c_ht_days = calc_col_letters[2][0]
        c_izin_days = calc_col_letters[3][0]
        c_rap_days = calc_col_letters[4][0]
        c_sgk_days = calc_col_letters[5][0]
        c_eksik_days = calc_col_letters[6][0]
        c_eksik_neden = calc_col_letters[7][0]
        c_ot_weekday = calc_col_letters[8][0]
        c_missing_hours = calc_col_letters[9][0]
        c_ot_pazar = calc_col_letters[10][0]
        c_maas_tl = calc_col_letters[11][0]
        c_ot_hi_tl = calc_col_letters[12][0]
        c_ot_hs_tl = calc_col_letters[13][0]
        c_ot_tot_tl = calc_col_letters[14][0]
        c_hakedis_tl = calc_col_letters[15][0]
        c_bordro_net = calc_col_letters[16][0]
        c_icra_kes = calc_col_letters[17][0]
        c_elden_fark = calc_col_letters[18][0]
        c_net_odenecek = calc_col_letters[19][0]
        
        # Dinamik Canlı Excel Formülleri
        sundays_count = len(engine.sundays)
        ws_p[f"{c_tot_hours}{r_idx}"] = f"=SUM({first_day_col}{r_idx}:{last_day_col}{r_idx})"
        ws_p[f"{c_work_days}{r_idx}"] = f'=COUNTIF({first_day_col}{r_idx}:{last_day_col}{r_idx}, ">0")'
        ws_p[f"{c_ht_days}{r_idx}"] = f"=IF({c_work_days}{r_idx}>=5, {sundays_count}, IF({c_work_days}{r_idx}>0, INT({c_work_days}{r_idx}/6), 0))"
        
        tc_val = emp["tc"]
        izin_val = engine.izin_by_tc.get(tc_val, {}).get("gun", 0) if tc_val else 0
        rap_val = engine.rapor_by_tc.get(tc_val, 0) if tc_val else 0
        
        ws_p[f"{c_izin_days}{r_idx}"] = izin_val
        ws_p[f"{c_rap_days}{r_idx}"] = rap_val
        
        ws_p[f"{c_sgk_days}{r_idx}"] = f"=MIN({engine.days_in_month}, {c_work_days}{r_idx}+{c_ht_days}{r_idx}+{c_izin_days}{r_idx}+{c_rap_days}{r_idx})"
        ws_p[f"{c_eksik_days}{r_idx}"] = f"={engine.days_in_month}-{c_sgk_days}{r_idx}"
        ws_p[f"{c_eksik_neden}{r_idx}"] = f'=IF({c_rap_days}{r_idx}>0, "01-İstirahat", IF({c_eksik_days}{r_idx}>0, "12-Birden Fazla", ""))'
        
        ws_p[f"{c_ot_weekday}{r_idx}"] = emp_ot_weekday
        ws_p[f"{c_missing_hours}{r_idx}"] = emp_missing_hours
        ws_p[f"{c_ot_pazar}{r_idx}"] = emp_ot_sunday
        
        # Finansal Ücret ve Fazla Mesai Formülleri (Mesai 1.5x, Eksik Saat 1.0x katsayılı)
        # Formül: IF(h>7.5, (h-7.5)*1.5, h-7.5) kuralına uygun: (Mesai * 1.5) - (Eksik * 1.0)
        ws_p[f"{c_maas_tl}{r_idx}"] = f"=(L{r_idx}/30)*{c_work_days}{r_idx}"
        ws_p[f"{c_ot_hi_tl}{r_idx}"] = f"=(L{r_idx}/225)*(({c_ot_weekday}{r_idx}*1.5)-({c_missing_hours}{r_idx}*1.0))"
        ws_p[f"{c_ot_hs_tl}{r_idx}"] = f"=(L{r_idx}/225)*1.5*{c_ot_pazar}{r_idx}"
        ws_p[f"{c_ot_tot_tl}{r_idx}"] = f"={c_ot_hi_tl}{r_idx}+{c_ot_hs_tl}{r_idx}"
        ws_p[f"{c_hakedis_tl}{r_idx}"] = f"={c_maas_tl}{r_idx}+{c_ot_tot_tl}{r_idx}"
        
        ws_p[f"{c_bordro_net}{r_idx}"] = f'=IFERROR(VLOOKUP(B{r_idx}, Resmi_Bordro_SGK!B:M, 11, FALSE), 0)'
        ws_p[f"{c_icra_kes}{r_idx}"] = f'=IFERROR(VLOOKUP(B{r_idx}, Icra_Takip!A:E, 4, FALSE), 0)'
        ws_p[f"{c_elden_fark}{r_idx}"] = f'=IF({c_hakedis_tl}{r_idx}>{c_bordro_net}{r_idx}, {c_hakedis_tl}{r_idx}-{c_bordro_net}{r_idx}-{c_icra_kes}{r_idx}, 0)'
        ws_p[f"{c_net_odenecek}{r_idx}"] = f'={c_hakedis_tl}{r_idx}-{c_icra_kes}{r_idx}'
        
        for col_info in [
            (c_tot_hours, "0.0", align_center, fill_overtime),
            (c_work_days, "0", align_center, None),
            (c_ht_days, "0", align_center, fill_pazar),
            (c_izin_days, "0", align_center, None),
            (c_rap_days, "0", align_center, fill_deduct if rap_val > 0 else None),
            (c_sgk_days, "0", align_center, fill_net),
            (c_eksik_days, "0", align_center, fill_deduct),
            (c_eksik_neden, "@", align_center, None),
            (c_ot_weekday, "0.0", align_center, fill_overtime),
            (c_missing_hours, "0.0", align_center, fill_deduct if emp_missing_hours > 0 else None),
            (c_ot_pazar, "0.0", align_center, fill_pazar),
            (c_maas_tl, "#,##0.00", align_right, None),
            (c_ot_hi_tl, "#,##0.00", align_right, fill_overtime),
            (c_ot_hs_tl, "#,##0.00", align_right, fill_pazar),
            (c_ot_tot_tl, "#,##0.00", align_right, fill_overtime),
            (c_hakedis_tl, "#,##0.00", align_right, fill_net),
            (c_bordro_net, "#,##0.00", align_right, None),
            (c_icra_kes, "#,##0.00", align_right, fill_deduct),
            (c_elden_fark, "#,##0.00", align_right, fill_net),
            (c_net_odenecek, "#,##0.00", align_right, fill_net),
        ]:
            cl, fmt, algn, fill_bg = col_info
            c = ws_p[f"{cl}{r_idx}"]
            c.number_format = fmt
            c.alignment = algn
            if fill_bg:
                c.fill = fill_bg
            elif row_fill:
                c.fill = row_fill
                
        for col_idx in range(1, cur_col_idx):
            cl = get_column_letter(col_idx)
            ws_p[f"{cl}{r_idx}"].border = border_cell
            ws_p[f"{cl}{r_idx}"].font = font_data

    # Puantaj Genel Toplam Satırı
    tot_r_puantaj = start_data_row + len(puantaj_rows)
    ws_p.row_dimensions[tot_r_puantaj].height = 22
    ws_p[f"A{tot_r_puantaj}"] = ""
    ws_p[f"B{tot_r_puantaj}"] = ""
    ws_p[f"C{tot_r_puantaj}"] = "GENEL TOPLAM"
    ws_p[f"C{tot_r_puantaj}"].alignment = align_left
    ws_p[f"C{tot_r_puantaj}"].font = font_total

    ws_p[f"L{tot_r_puantaj}"] = f"=SUM(L4:L{tot_r_puantaj-1})"
    ws_p[f"L{tot_r_puantaj}"].number_format = "#,##0.00"

    for col_let, text, width, is_pazar, d in day_cols:
        ws_p[f"{col_let}{tot_r_puantaj}"] = f"=SUM({col_let}4:{col_let}{tot_r_puantaj-1})"
        ws_p[f"{col_let}{tot_r_puantaj}"].number_format = "0.0"

    for cl in [c_tot_hours, c_work_days, c_ht_days, c_izin_days, c_rap_days, c_sgk_days, c_eksik_days, c_ot_weekday, c_missing_hours, c_ot_pazar]:
        ws_p[f"{cl}{tot_r_puantaj}"] = f"=SUM({cl}4:{cl}{tot_r_puantaj-1})"
        ws_p[f"{cl}{tot_r_puantaj}"].number_format = "0.0" if cl in (c_tot_hours, c_ot_weekday, c_missing_hours, c_ot_pazar) else "0"

    for cl in [c_maas_tl, c_ot_hi_tl, c_ot_hs_tl, c_ot_tot_tl, c_hakedis_tl, c_bordro_net, c_icra_kes, c_elden_fark, c_net_odenecek]:
        ws_p[f"{cl}{tot_r_puantaj}"] = f"=SUM({cl}4:{cl}{tot_r_puantaj-1})"
        ws_p[f"{cl}{tot_r_puantaj}"].number_format = "#,##0.00"

    for col_idx in range(1, cur_col_idx):
        cl = get_column_letter(col_idx)
        c = ws_p[f"{cl}{tot_r_puantaj}"]
        c.font = font_total
        c.fill = fill_total
        c.border = border_total
        if not c.alignment.horizontal:
            c.alignment = align_right

    ws_p.freeze_panes = "D4"

    # -------------------------------------------------------------------------
    # GÜNLÜK ÇALIŞMA SAATİ KORUMASI (Maks. 24 Saat Kuralı):
    # Bir günde 24 saatten fazla çalışma yazılamaz!
    # Excel içinde elle 24 saatten fazla yazılmak istense dahi kesin olarak engeller.
    # -------------------------------------------------------------------------
    first_day_col = day_cols[0][0]
    last_day_col = day_cols[-1][0]
    last_data_r = tot_r_puantaj - 1
    
    dv_max24 = DataValidation(
        type="decimal",
        operator="between",
        formula1="0",
        formula2="24",
        allow_blank=True,
        showInputMessage=True,
        promptTitle="Günlük Çalışma Süresi",
        prompt="Çalışma süresi 0 ile 24 saat arasında olmalıdır.",
        showErrorMessage=True,
        errorTitle="Geçersiz Süre (Maksimum 24 Saat)",
        error="HATA: Bir günde 24 saatten fazla çalışma süresi yazılamaz! Lütfen 0 ile 24.0 saat arasında bir değer giriniz."
    )
    ws_p.add_data_validation(dv_max24)
    dv_max24.add(f"{first_day_col}4:{last_day_col}{last_data_r}")

    # =========================================================================
    # SAYFA 2: Eksik_Basim_Raporu (İstisna & İnsan Kontrolü Raporu)
    # =========================================================================
    ws_miss = wb_new.create_sheet(title="Eksik_Basim_Raporu")
    ws_miss.views.sheetView[0].showGridLines = True

    miss_headers = [
        ("Sıra No", 8), ("Tarih", 12), ("Gün", 12), ("Kart No", 10), ("Sicil No", 10),
        ("Adı Soyadı", 22), ("Bölüm / Lokasyon", 20), ("Tespit Edilen Giriş", 15),
        ("Tespit Edilen Çıkış", 15), ("Tüm Basım Hareketleri", 22),
        ("İnceleme Nedeni (Hata Türü)", 26), ("Puantaja Yazılan Saat", 18),
        ("Fazla Mesai (Saat)", 16), ("İK / Amir Onayı & Notu", 22)
    ]
    ws_miss.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(miss_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_miss[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_miss.column_dimensions[col_let].width = w

    for r_idx, mp in enumerate(engine.exception_records, start=2):
        ws_miss[f"A{r_idx}"] = mp["sno"]
        ws_miss[f"B{r_idx}"] = mp["tarih"]
        ws_miss[f"C{r_idx}"] = mp["gun_adi"]
        ws_miss[f"D{r_idx}"] = mp["kart"]
        ws_miss[f"E{r_idx}"] = mp["sicil"]
        ws_miss[f"F{r_idx}"] = mp["ad_soyad"]
        ws_miss[f"G{r_idx}"] = mp["bolum"] or mp["lokasyon"]
        ws_miss[f"H{r_idx}"] = mp["raw_g_saat"]
        ws_miss[f"I{r_idx}"] = mp["raw_c_saat"]
        ws_miss[f"J{r_idx}"] = mp["all_punches"]
        ws_miss[f"K{r_idx}"] = mp["status_type"]
        ws_miss[f"L{r_idx}"] = mp["final_hours"]
        ws_miss[f"M{r_idx}"] = mp["overtime_hours"]
        ws_miss[f"N{r_idx}"] = "İK / Amir Onayında"
        
        for c_idx in range(1, 15):
            cl = get_column_letter(c_idx)
            ws_miss[f"{cl}{r_idx}"].border = border_cell
            ws_miss[f"{cl}{r_idx}"].font = font_data
            if c_idx == 11:
                ws_miss[f"{cl}{r_idx}"].fill = fill_missing
                ws_miss[f"{cl}{r_idx}"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="C00000")
            elif c_idx in (12, 13):
                ws_miss[f"{cl}{r_idx}"].number_format = "0.0"
                ws_miss[f"{cl}{r_idx}"].alignment = align_center
            if c_idx not in (6, 7, 10):
                ws_miss[f"{cl}{r_idx}"].alignment = align_center

    # Eksik Basım Raporunda L sütunu (Puantaja Yazılan Saat) için de 24 saat koruması:
    if len(engine.exception_records) > 0:
        dv_miss24 = DataValidation(
            type="decimal",
            operator="between",
            formula1="0",
            formula2="24",
            allow_blank=True,
            showErrorMessage=True,
            errorTitle="Geçersiz Süre (Maksimum 24 Saat)",
            error="HATA: Bir günde 24 saatten fazla çalışma süresi yazılamaz! Lütfen 0 ile 24.0 saat arasında bir değer giriniz."
        )
        ws_miss.add_data_validation(dv_miss24)
        dv_miss24.add(f"L2:L{len(engine.exception_records)+1}")

    # =========================================================================
    # SAYFA 3: Bolum_Prim_Mesai_Raporu (Balık Dolum/Kesim & Üretim Prim Denetimi)
    # =========================================================================
    ws_bon = wb_new.create_sheet(title="Bolum_Prim_Mesai_Raporu")
    ws_bon.views.sheetView[0].showGridLines = True

    bon_headers = [
        ("Sıra No", 8), ("Tarih", 12), ("Gün", 12), ("Kart No", 10), ("Sicil No", 10),
        ("Adı Soyadı", 22), ("Bölüm / Görev", 20), ("Giriş Saati", 14),
        ("Çıkış Saati", 14), ("Fiili Çalışma (Saat)", 18), ("Eklenen Prim (Saat)", 18),
        ("Puantaja Yazılan Saat", 18), ("Toplam Fazla Mesai", 18), ("Vardiya Amiri / İK Onayı", 22)
    ]
    ws_bon.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(bon_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_bon[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_bon.column_dimensions[col_let].width = w

    for r_idx, br in enumerate(engine.bonus_records, start=2):
        ws_bon[f"A{r_idx}"] = br["sno"]
        ws_bon[f"B{r_idx}"] = br["tarih"]
        ws_bon[f"C{r_idx}"] = br["gun_adi"]
        ws_bon[f"D{r_idx}"] = br["kart"]
        ws_bon[f"E{r_idx}"] = br["sicil"]
        ws_bon[f"F{r_idx}"] = br["ad_soyad"]
        ws_bon[f"G{r_idx}"] = br["bolum"]
        ws_bon[f"H{r_idx}"] = br["raw_g_saat"]
        ws_bon[f"I{r_idx}"] = br["raw_c_saat"]
        ws_bon[f"J{r_idx}"] = br["fiili_sure"]
        ws_bon[f"K{r_idx}"] = br["bonus_hours"]
        ws_bon[f"L{r_idx}"] = br["final_hours"]
        ws_bon[f"M{r_idx}"] = br["overtime_hours"]
        ws_bon[f"N{r_idx}"] = "İK / Amir Onayında"
        
        for c_idx in range(1, 15):
            cl = get_column_letter(c_idx)
            ws_bon[f"{cl}{r_idx}"].border = border_cell
            ws_bon[f"{cl}{r_idx}"].font = font_data
            if c_idx in (10, 11, 12, 13):
                ws_bon[f"{cl}{r_idx}"].number_format = "0.0"
                ws_bon[f"{cl}{r_idx}"].alignment = align_center
            if c_idx == 11:
                ws_bon[f"{cl}{r_idx}"].fill = fill_bonus
                ws_bon[f"{cl}{r_idx}"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="7030A0")
            if c_idx not in (6, 7):
                ws_bon[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 4: PDKS_Hareket_Kayitlari (Ham Turnike Verileri)
    # =========================================================================
    ws_pdk = wb_new.create_sheet(title="PDKS_Hareket_Kayitlari")
    ws_pdk.views.sheetView[0].showGridLines = True

    pdk_headers = [
        ("Sıra No", 8), ("Sicil No", 10), ("Kart No", 10), ("Adı Soyadı", 22),
        ("Gün", 12), ("Giriş Tarihi", 12), ("Giriş Saati", 12),
        ("Çıkış Tarihi", 12), ("Çıkış Saati", 12), ("Toplam Süre", 12),
        ("Normal Mesai", 12), ("Fazla Mesai", 12), ("Lokasyon", 16)
    ]
    ws_pdk.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(pdk_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_pdk[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_pdk.column_dimensions[col_let].width = w

    for r_idx, rec in enumerate(engine.audit_records, start=2):
        ws_pdk[f"A{r_idx}"] = rec["sno"]
        ws_pdk[f"B{r_idx}"] = rec["sicil"]
        ws_pdk[f"C{r_idx}"] = rec["kart"]
        ws_pdk[f"D{r_idx}"] = rec["ad_soyad"]
        ws_pdk[f"E{r_idx}"] = rec["gun_adi"]
        ws_pdk[f"F{r_idx}"] = rec["tarih"]
        ws_pdk[f"G{r_idx}"] = rec["raw_g_saat"]
        ws_pdk[f"H{r_idx}"] = rec["tarih"]
        ws_pdk[f"I{r_idx}"] = rec["raw_c_saat"]
        ws_pdk[f"J{r_idx}"] = rec["final_hours"]
        ws_pdk[f"K{r_idx}"] = rec["base_hours"]
        ws_pdk[f"L{r_idx}"] = rec["overtime_hours"]
        ws_pdk[f"J{r_idx}"].number_format = "0.0"
        ws_pdk[f"K{r_idx}"].number_format = "0.0"
        ws_pdk[f"L{r_idx}"].number_format = "0.0"
        ws_pdk[f"M{r_idx}"] = rec["lokasyon"]
        
        for c_idx in range(1, 14):
            cl = get_column_letter(c_idx)
            ws_pdk[f"{cl}{r_idx}"].border = border_cell
            ws_pdk[f"{cl}{r_idx}"].font = font_data
            if c_idx not in (4, 13):
                ws_pdk[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 5: Resmi_Bordro_SGK (Yasal SGK Bordro Listesi)
    # =========================================================================
    ws_bor = wb_new.create_sheet(title="Resmi_Bordro_SGK")
    ws_bor.views.sheetView[0].showGridLines = True
    bor_headers = [
        ("Sıra No", 8), ("TC Kimlik No", 14), ("Adı Soyadı", 22), ("Ücret (TL)", 14),
        ("Giriş Tarihi", 12), ("Çıkış Tarihi", 12), ("Normal Kazanç", 14),
        ("Ek Kazanç", 14), ("Yasal Kesinti", 14), ("Toplam Kazanç", 14),
        ("Toplam Kesinti", 14), ("Ödenecek Net", 15), ("SGK Gün", 10),
        ("SGK Matrah (Brüt)", 16), ("Kanun No", 12), ("Meslek Kodu", 14)
    ]
    ws_bor.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(bor_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_bor[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_bor.column_dimensions[col_let].width = w

    for r_idx, b in enumerate(engine.bordro_list, start=2):
        ws_bor[f"A{r_idx}"] = b["sno"]
        ws_bor[f"B{r_idx}"] = b["tc"]
        ws_bor[f"B{r_idx}"].number_format = "@"
        ws_bor[f"C{r_idx}"] = b["ad_soyad"]
        ws_bor[f"D{r_idx}"] = b["ucret"]
        ws_bor[f"E{r_idx}"] = b["giris"]
        ws_bor[f"F{r_idx}"] = b["cikis"]
        ws_bor[f"G{r_idx}"] = b["normal_kazanc"]
        ws_bor[f"H{r_idx}"] = b["ek_kazanc"]
        ws_bor[f"I{r_idx}"] = b["yasal_kesinti"]
        ws_bor[f"J{r_idx}"] = b["toplam_kazanc"]
        ws_bor[f"K{r_idx}"] = b["toplam_kesinti"]
        ws_bor[f"L{r_idx}"] = b["odenecek_net"]
        ws_bor[f"M{r_idx}"] = b["sgk_gun"]
        ws_bor[f"N{r_idx}"] = b["sgk_brut"]
        ws_bor[f"O{r_idx}"] = b["kanun"]
        ws_bor[f"P{r_idx}"] = b["meslek_kodu"]
        
        for c_idx in range(1, 17):
            cl = get_column_letter(c_idx)
            ws_bor[f"{cl}{r_idx}"].border = border_cell
            ws_bor[f"{cl}{r_idx}"].font = font_data
            if c_idx in (4, 7, 8, 9, 10, 11, 12, 14):
                ws_bor[f"{cl}{r_idx}"].number_format = "#,##0.00"
                ws_bor[f"{cl}{r_idx}"].alignment = align_right
            elif c_idx not in (3,):
                ws_bor[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 6: Icra_Takip (Yasal Maaş Haciz & İcra Kesintileri)
    # =========================================================================
    ws_icra = wb_new.create_sheet(title="Icra_Takip")
    ws_icra.views.sheetView[0].showGridLines = True
    icra_headers = [
        ("TC Kimlik No", 14), ("Adı Soyadı", 22), ("Şirket", 16),
        ("İcra Kesintisi (TL)", 16), ("Kalan Borç (TL)", 16),
        ("İcra Dosya No", 24), ("İban No", 28), ("Açıklama", 20)
    ]
    ws_icra.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(icra_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_icra[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_icra.column_dimensions[col_let].width = w

    for r_idx, ic in enumerate(engine.icra_list, start=2):
        tc_found = ""
        norm_ic_name = norm_name_key(ic["isim"])
        for p in puantaj_rows:
            if p["name_key"] == norm_ic_name:
                tc_found = p["tc"]
                break
                
        ws_icra[f"A{r_idx}"] = tc_found
        ws_icra[f"A{r_idx}"].number_format = "@"
        ws_icra[f"B{r_idx}"] = ic["isim"]
        ws_icra[f"C{r_idx}"] = ic["sirket"]
        ws_icra[f"D{r_idx}"] = ic["kesilen"]
        ws_icra[f"D{r_idx}"].number_format = "#,##0.00"
        ws_icra[f"E{r_idx}"] = ic["kalan_borc"]
        ws_icra[f"E{r_idx}"].number_format = "#,##0.00"
        ws_icra[f"F{r_idx}"] = ic["dosya"]
        ws_icra[f"G{r_idx}"] = ic["iban"]
        ws_icra[f"H{r_idx}"] = "Aktif Kesinti"
        
        for c_idx in range(1, 9):
            cl = get_column_letter(c_idx)
            ws_icra[f"{cl}{r_idx}"].border = border_cell
            ws_icra[f"{cl}{r_idx}"].font = font_data
            if c_idx in (4, 5):
                ws_icra[f"{cl}{r_idx}"].alignment = align_right
            elif c_idx not in (2, 6, 7):
                ws_icra[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 7: SGK_Raporlar (İstirahat Rapor Takip Listesi)
    # =========================================================================
    ws_rap = wb_new.create_sheet(title="SGK_Raporlar")
    ws_rap.views.sheetView[0].showGridLines = True
    rap_headers = [
        ("Takip No", 12), ("Sıra No", 8), ("TC Kimlik No", 14), ("Adı Soyadı", 22),
        ("Vaka Türü", 18), ("Poliklinik Tarihi", 15), ("İşbaşı Tarihi", 15)
    ]
    ws_rap.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(rap_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_rap[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_rap.column_dimensions[col_let].width = w

    for r_idx, rp in enumerate(engine.rapor_list, start=2):
        ws_rap[f"A{r_idx}"] = rp["takip_no"]
        ws_rap[f"B{r_idx}"] = rp["sira_no"]
        ws_rap[f"C{r_idx}"] = rp["tc"]
        ws_rap[f"C{r_idx}"].number_format = "@"
        ws_rap[f"D{r_idx}"] = rp["ad_soyad"]
        ws_rap[f"E{r_idx}"] = rp["vaka"]
        ws_rap[f"F{r_idx}"] = rp["pol_tarih"]
        ws_rap[f"G{r_idx}"] = rp["isbasi_tarih"]
        
        for c_idx in range(1, 8):
            cl = get_column_letter(c_idx)
            ws_rap[f"{cl}{r_idx}"].border = border_cell
            ws_rap[f"{cl}{r_idx}"].font = font_data
            if c_idx not in (4, 5):
                ws_rap[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 8: Ucretli_Izinler (Yıllık / Ücretli İzin Cetveli)
    # =========================================================================
    ws_izn = wb_new.create_sheet(title="Ucretli_Izinler")
    ws_izn.views.sheetView[0].showGridLines = True
    izn_headers = [("TC Kimlik No", 14), ("Adı Soyadı", 22), ("İzin Saati", 14), ("İzin Günü", 14)]
    ws_izn.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(izn_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_izn[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_izn.column_dimensions[col_let].width = w

    for r_idx, iz in enumerate(engine.izin_list, start=2):
        ws_izn[f"A{r_idx}"] = iz["tc"]
        ws_izn[f"A{r_idx}"].number_format = "@"
        ws_izn[f"B{r_idx}"] = iz["isim"]
        ws_izn[f"C{r_idx}"] = iz["saat"]
        ws_izn[f"D{r_idx}"] = iz["gun"]
        for c_idx in range(1, 5):
            cl = get_column_letter(c_idx)
            ws_izn[f"{cl}{r_idx}"].border = border_cell
            ws_izn[f"{cl}{r_idx}"].font = font_data
            if c_idx in (3, 4):
                ws_izn[f"{cl}{r_idx}"].alignment = align_right
                ws_izn[f"{cl}{r_idx}"].number_format = "#,##0.0"
            elif c_idx != 2:
                ws_izn[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 9: Daimi_Personel_Listesi (Kadro Durumu)
    # =========================================================================
    ws_dai = wb_new.create_sheet(title="Daimi_Personel_Listesi")
    ws_dai.views.sheetView[0].showGridLines = True
    dai_headers = [
        ("Sıra No", 8), ("TC Kimlik No", 14), ("Adı Soyadı", 22),
        ("SGK Giriş", 12), ("SGK Durum", 14), ("Kadro", 14), ("Maaş (TL)", 15)
    ]
    ws_dai.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(dai_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_dai[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_dai.column_dimensions[col_let].width = w

    for r_idx, da in enumerate(engine.daimi_list, start=2):
        ws_dai[f"A{r_idx}"] = da["sno"]
        ws_dai[f"B{r_idx}"] = da["tc"]
        ws_dai[f"B{r_idx}"].number_format = "@"
        ws_dai[f"C{r_idx}"] = da["ad_soyad"]
        ws_dai[f"D{r_idx}"] = da["sgk_giris"]
        ws_dai[f"E{r_idx}"] = da["sgk_durum"]
        ws_dai[f"F{r_idx}"] = da["durum"]
        ws_dai[f"G{r_idx}"] = da["maas"]
        for c_idx in range(1, 8):
            cl = get_column_letter(c_idx)
            ws_dai[f"{cl}{r_idx}"].border = border_cell
            ws_dai[f"{cl}{r_idx}"].font = font_data
            if c_idx == 7:
                ws_dai[f"{cl}{r_idx}"].alignment = align_right
                ws_dai[f"{cl}{r_idx}"].number_format = "#,##0.00"
            elif c_idx != 3:
                ws_dai[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # SAYFA 10: Elden_Odeme_Farki (Net Maaş ile Resmi Bordro Fark Cetveli)
    # =========================================================================
    ws_eld = wb_new.create_sheet(title="Elden_Odeme_Farki")
    ws_eld.views.sheetView[0].showGridLines = True
    eld_headers = [
        ("Sıra No", 8), ("TC Kimlik No", 14), ("Adı Soyadı", 22),
        ("Anlaşılan Maaş", 15), ("Resmi Bordro Net", 15), ("Elden Fark (TL)", 15)
    ]
    ws_eld.row_dimensions[1].height = 25
    for c_idx, (h_text, w) in enumerate(eld_headers, start=1):
        col_let = get_column_letter(c_idx)
        cell = ws_eld[f"{col_let}1"]
        cell.value = h_text
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_hdr
        cell.border = border_header
        ws_eld.column_dimensions[col_let].width = w

    for r_idx, p in enumerate(puantaj_rows, start=2):
        p_row = r_idx + 2
        ws_eld[f"A{r_idx}"] = p["sno"]
        ws_eld[f"B{r_idx}"] = p["tc"]
        ws_eld[f"B{r_idx}"].number_format = "@"
        ws_eld[f"C{r_idx}"] = p["ad_soyad"]
        ws_eld[f"D{r_idx}"] = f"=Aylik_Puantaj!L{p_row}"
        ws_eld[f"E{r_idx}"] = f"=Aylik_Puantaj!{c_bordro_net}{p_row}"
        ws_eld[f"F{r_idx}"] = f"=Aylik_Puantaj!{c_elden_fark}{p_row}"
        
        for c_idx in range(1, 7):
            cl = get_column_letter(c_idx)
            ws_eld[f"{cl}{r_idx}"].border = border_cell
            ws_eld[f"{cl}{r_idx}"].font = font_data
            if c_idx in (4, 5, 6):
                ws_eld[f"{cl}{r_idx}"].number_format = "#,##0.00"
                ws_eld[f"{cl}{r_idx}"].alignment = align_right
            elif c_idx != 3:
                ws_eld[f"{cl}{r_idx}"].alignment = align_center

    # =========================================================================
    # 3. DOSYAYI KAYDETME
    # =========================================================================
    saved_name = output_filename
    try:
        wb_new.save(output_filename)
        print(f"3. Dosya '{output_filename}' olarak başarıyla kaydedildi.")
    except PermissionError:
        saved_name = output_filename.replace(".xlsx", "_guncel.xlsx")
        wb_new.save(saved_name)
        print(f"3. UYARI: '{output_filename}' açık olduğu için '{saved_name}' olarak kaydedildi.")

    print("==================================================")
    print(f" BAŞARILI! '{saved_name}' dosyası eksiksiz oluşturuldu.")
    print(f" - Özel Bölüm Primleri: {len(engine.bonus_records)} kayıt 'Bolum_Prim_Mesai_Raporu' sayfasına işlendi (Lila renkli)")
    print(f" - İnsan Kontrolü Yönetimi: {len(engine.exception_records)} istisna kaydı 'Eksik_Basim_Raporu' sayfasına işlendi (Turuncu renkli)")
    print(" - Canlı Formüller (SUM, COUNTIF, IF, MIN, VLOOKUP) Aktif")
    print(f" - {len(puantaj_rows)} personelin {engine.days_in_month} günlük çalışma süreleri işlendi")
    print("==================================================")
    return saved_name

def main():
    build_puantaj_workbook(
        pdks_path="pdks.xls",
        puantaj_path="puantaj.xls",
        target_month=9,
        target_year=2026,
        output_filename="puantaj.xlsx"
    )

if __name__ == "__main__":
    main()
