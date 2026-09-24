# -*- coding: utf-8 -*-
"""
===================================================================================================
FİDE KONSERVE — AYIK ÇALIŞMA VE FAZLA MESAİ MUTABAKAT PUSULASI ÜRETECİ (A4 MATBU / PDF)
===================================================================================================
İşçi ve İK karşılıklı ıslak imza onaylı, 4857 Sayılı İş Kanunu ve Personel Özlük Dosyası
standartlarına tam uyumlu tekil ve toplu A4 mutabakat fişi oluşturucu.
"""

import calendar
from typing import Dict, Any, List, Optional
from pdks_engine import clean_display_text, fmt_hours_tr

TR_MONTH_NAMES = [
    "OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN",
    "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"
]

TR_DAY_ABBRS = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz"]


def generate_single_slip_html(
    person: Dict[str, Any],
    summary: Dict[str, Any],
    days_data: List[Dict[str, Any]],
    month: int,
    year: int,
    sundays: List[int]
) -> str:
    """Tek bir personel için A4 formatında Çalışma ve Mesai Mutabakat Pusulası HTML'i üretir."""
    ad_soyad = clean_display_text(person.get("ad_soyad", ""))
    tc_no = person.get("tc") or "-"
    sno = str(person.get("sno") or person.get("sicil") or "-")
    bolum = clean_display_text(person.get("bolum") or "Genel")
    durumu = clean_display_text(person.get("durumu") or "MEVSİMLİK")
    sgk_giris = person.get("sgk_giris") or "-"
    month_name = TR_MONTH_NAMES[month - 1] if 1 <= month <= 12 else str(month)

    tot_work_days = summary.get("total_work_days", 0)
    tot_base = summary.get("total_base_hours", 0.0)
    tot_ot = summary.get("total_ot_hours", 0.0)
    tot_bonus = summary.get("total_bonus_hours", 0.0)
    tot_hours = summary.get("total_hours", 0.0)
    sunday_count = len(sundays)

    tot_break = 0.0
    tot_fiili = 0.0

    # 30 Günlük Tablo Satırları
    rows_html = []
    for d_rec in days_data:
        day_num = d_rec.get("gun", 1)
        is_sun = day_num in sundays
        final_h = d_rec.get("final_hours", 0.0)
        has_work = final_h > 0

        # Tarih ve Gün formatı (Örn: 01.09 Sal)
        try:
            w_idx = calendar.weekday(year, month, day_num)
            day_name = TR_DAY_ABBRS[w_idx]
        except Exception:
            day_name = "Paz" if is_sun else ""

        tarih_display = f"{day_num:02d}.{month:02d} {day_name}"

        raw_g = d_rec.get("raw_g_saat") or "-"
        raw_c = d_rec.get("raw_c_saat") or "-"
        brk_h = d_rec.get("break_hours", 0.0)
        fiili_h = d_rec.get("fiili_sure", 0.0)
        base_h = d_rec.get("base_hours", 0.0)
        ot_h = d_rec.get("overtime_hours", 0.0)
        bonus_h = d_rec.get("bonus_hours", 0.0)

        if has_work:
            tot_break += brk_h
            tot_fiili += fiili_h
            brk_str = fmt_hours_tr(brk_h) if brk_h > 0 else "-"
            fiili_str = fmt_hours_tr(fiili_h)
            base_str = fmt_hours_tr(base_h)
            ot_str = fmt_hours_tr(ot_h) if ot_h > 0 else "-"
            bonus_str = f"+{fmt_hours_tr(bonus_h)}" if bonus_h > 0 else "-"
            final_str = fmt_hours_tr(final_h)

            note_raw = d_rec.get("audit_note") or d_rec.get("status_type") or "Normal"
            if is_sun:
                note_display = f"Pazar Mesaisi ({clean_display_text(note_raw)})"
            elif ot_h > 0:
                note_display = f"Fazla Mesai ({clean_display_text(note_raw)})"
            elif bonus_h > 0:
                note_display = f"Primli ({clean_display_text(note_raw)})"
            else:
                note_display = clean_display_text(note_raw)
        else:
            raw_g = "-"
            raw_c = "-"
            brk_str = "-"
            fiili_str = "-"
            base_str = "-"
            ot_str = "-"
            bonus_str = "-"
            final_str = "-"
            note_display = "Hafta Tatili" if is_sun else "-"

        row_cls = "row-sunday" if is_sun else ""
        ot_cls = "cell-ot" if ot_h > 0 else ""
        bonus_cls = "cell-bonus" if bonus_h > 0 else ""

        rows_html.append(f"""
        <tr class="{row_cls}">
          <td style="text-align: center;">{day_num}</td>
          <td style="text-align: center;" class="cell-mono">{tarih_display}</td>
          <td style="text-align: center;" class="cell-mono">{raw_g}</td>
          <td style="text-align: center;" class="cell-mono">{raw_c}</td>
          <td style="text-align: center;" class="cell-mono">{brk_str}</td>
          <td style="text-align: center; font-weight: 600;" class="cell-mono">{fiili_str}</td>
          <td style="text-align: center;" class="cell-mono">{base_str}</td>
          <td style="text-align: center;" class="cell-mono {ot_cls}">{ot_str}</td>
          <td style="text-align: center;" class="cell-mono {bonus_cls}">{bonus_str}</td>
          <td style="text-align: center; font-weight: 700;" class="cell-mono cell-final">{final_str}</td>
          <td style="font-size: 8.5px; color: #475569;" class="note-col">{note_display}</td>
        </tr>
        """)

    table_rows_str = "\n".join(rows_html)

    bonus_total_str = f"+{fmt_hours_tr(tot_bonus)}" if tot_bonus > 0 else "-"

    # Tek bir A4 sayfası HTML bileşeni
    return f"""
    <div class="slip-page">
      <div>
        <!-- Üst Kurumsal Başlık -->
        <div class="slip-header">
          <div>
            <div class="company-brand">FİDE KONSERVE GIDA SAN. VE TİC. A.Ş.</div>
            <div class="company-sub">Susurluk Fabrikası &bull; Personel Özlük ve Bordro İşleri Servisi</div>
          </div>
          <div class="doc-title-block">
            <div class="doc-main-title">AYLIK ÇALIŞMA VE FAZLA MESAİ MUTABAKAT PUSULASI</div>
            <div class="doc-period">DÖNEM: {month_name} {year}</div>
            <div class="doc-ref">4857 Sayılı İş Kanunu Kapsamında Hazırlanmıştır</div>
          </div>
        </div>

        <!-- Personel Özlük Bilgileri -->
        <div class="person-info-box">
          <div class="info-field"><span class="info-label">Adı Soyadı:</span> <span class="info-val">{ad_soyad}</span></div>
          <div class="info-field"><span class="info-label">T.C. Kimlik:</span> <span class="info-val cell-mono">{tc_no}</span></div>
          <div class="info-field"><span class="info-label">Sicil / Sıra:</span> <span class="info-val cell-mono">{sno}</span></div>
          <div class="info-field"><span class="info-label">Bölümü:</span> <span class="info-val">{bolum}</span></div>
          <div class="info-field"><span class="info-label">SGK Giriş:</span> <span class="info-val cell-mono">{sgk_giris}</span></div>
          <div class="info-field"><span class="info-label">Statü:</span> <span class="info-val">{durumu}</span></div>
        </div>

        <!-- Aylık İcmal Şeridi -->
        <div class="kpi-strip">
          <div class="kpi-strip-item">
            <div class="kpi-strip-label">Çalışılan Gün</div>
            <div class="kpi-strip-val">{tot_work_days} <small style="font-size:9px;font-weight:normal;">gün</small></div>
          </div>
          <div class="kpi-strip-item">
            <div class="kpi-strip-label">Normal Mesai</div>
            <div class="kpi-strip-val">{fmt_hours_tr(tot_base)} <small style="font-size:9px;font-weight:normal;">saat</small></div>
          </div>
          <div class="kpi-strip-item" style="border-color: #fed7aa; background: #fffaf5;">
            <div class="kpi-strip-label" style="color: #c2410c;">Fazla Mesai (FM)</div>
            <div class="kpi-strip-val" style="color: #c2410c;">{fmt_hours_tr(tot_ot)} <small style="font-size:9px;font-weight:normal;">saat</small></div>
          </div>
          <div class="kpi-strip-item" style="border-color: #e9d5ff; background: #faf5ff;">
            <div class="kpi-strip-label" style="color: #7e22ce;">Bölüm Primi</div>
            <div class="kpi-strip-val" style="color: #7e22ce;">{fmt_hours_tr(tot_bonus)} <small style="font-size:9px;font-weight:normal;">saat</small></div>
          </div>
          <div class="kpi-strip-item" style="border-color: #bae6fd; background: #f0f9ff;">
            <div class="kpi-strip-label" style="color: #0369a1;">Toplam Tahakkuk</div>
            <div class="kpi-strip-val" style="color: #0369a1;">{fmt_hours_tr(tot_hours)} <small style="font-size:9px;font-weight:normal;">saat</small></div>
          </div>
        </div>

        <!-- 30 Günlük Detaylı Puantaj Çizelgesi -->
        <div class="table-wrapper">
          <table class="days-table">
            <thead>
              <tr>
                <th style="width: 24px;">Gün</th>
                <th style="width: 65px;">Tarih</th>
                <th style="width: 44px;">İlk Giriş</th>
                <th style="width: 44px;">Son Çıkış</th>
                <th style="width: 42px;">Mola</th>
                <th style="width: 48px;">Fiili Süre</th>
                <th style="width: 46px;">Normal</th>
                <th style="width: 44px;">FM</th>
                <th style="width: 40px;">Prim</th>
                <th style="width: 48px;">Toplam</th>
                <th>Vardiya Açıklaması / İşlem Notu</th>
              </tr>
            </thead>
            <tbody>
              {table_rows_str}
              <tr class="table-total-row">
                <td colspan="4" style="text-align: right; padding-right: 8px;">AYLIK GENEL TOPLAMLAR ({tot_work_days} Gün):</td>
                <td style="text-align: center;" class="cell-mono">{fmt_hours_tr(tot_break)}</td>
                <td style="text-align: center;" class="cell-mono">{fmt_hours_tr(tot_fiili)}</td>
                <td style="text-align: center;" class="cell-mono">{fmt_hours_tr(tot_base)}</td>
                <td style="text-align: center; color: #c2410c;" class="cell-mono">{fmt_hours_tr(tot_ot)}</td>
                <td style="text-align: center; color: #7e22ce;" class="cell-mono">{bonus_total_str}</td>
                <td style="text-align: center; font-size: 11px;" class="cell-mono">{fmt_hours_tr(tot_hours)}</td>
                <td style="font-size: 8.5px; color: #64748b; font-weight: normal;">Pazar Tatili: {sunday_count} Gün</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Alt Yasal Mutabakat ve Çift Islak İmza Bölümü -->
      <div class="slip-footer">
        <div class="legal-statement">
          <strong>4857 Sayılı İş Kanunu ve ilgili yönetmelikler uyarınca;</strong> yukarıda dökümü yapılan {month_name} {year} dönemine ait turnike kart basımlarımı, günlük fiili çalışma, normal çalışma, ara dinlenme (mola), fazla mesai ve prim saatlerimi inceledim. Bilgilerin gerçeğe ve fiili çalışmama uygun olduğunu, fazla çalışma ücretlerimin ve tüm hak edişlerimin bordroma eksiksiz yansıtıldığını, herhangi bir hak ve alacağımın kalmadığını gayrikabili rücu kabul, beyan ve imza ederim.
        </div>

        <div class="signatures-box">
          <div class="sig-card">
            <div class="sig-role">PERSONEL (İŞÇİ) ONAYI</div>
            <div class="sig-meta">Adı Soyadı: <strong>{ad_soyad}</strong></div>
            <div class="sig-meta">T.C. Kimlik No: <strong class="cell-mono">{tc_no}</strong></div>
            <div class="sig-line">
              <span>Tarih: ..... / ..... / {year}</span>
              <span>İmza: ___________________________</span>
            </div>
          </div>

          <div class="sig-card">
            <div class="sig-role">İŞVEREN VEKİLİ (İNSAN KAYNAKLARI / VARDİYA AMİRİ)</div>
            <div class="sig-meta">FİDE KONSERVE GIDA SAN. VE TİC. A.Ş.</div>
            <div class="sig-meta">Yetkili Adı Soyadı: ....................................................</div>
            <div class="sig-line">
              <span>Tarih: ..... / ..... / {year}</span>
              <span>Kaşe & İmza: ___________________________</span>
            </div>
          </div>
        </div>
      </div>
    </div>
    """


def render_slips_document(
    slips_html_list: List[str],
    title: str = "Aylık Personel Puantaj Fişleri",
    period_str: str = "Eylül 2026",
    auto_print: bool = False
) -> str:
    """Tüm personelin sayfalarını içeren, yazdırılabilir tam HTML dökümanı."""
    total_count = len(slips_html_list)
    slips_joined = "\n".join(slips_html_list)
    auto_print_script = "<script>window.onload = function() { window.print(); };</script>" if auto_print else ""

    return f"""<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title} — {period_str} — FİDE KONSERVE</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap');

    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}

    body {{
      background-color: #0f172a;
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      color: #0f172a;
      -webkit-font-smoothing: antialiased;
    }}

    /* Screen Sticky Toolbar */
    .screen-toolbar {{
      position: sticky;
      top: 0;
      z-index: 9999;
      background: #090d16;
      border-bottom: 1px solid rgba(255,255,255,0.1);
      padding: 12px 28px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      box-shadow: 0 4px 20px rgba(0,0,0,0.5);
    }}

    .toolbar-left {{
      display: flex;
      align-items: center;
      gap: 14px;
    }}

    .brand-logo-text {{
      font-size: 15px;
      font-weight: 800;
      color: #ffffff;
      letter-spacing: -0.01em;
      display: flex;
      align-items: center;
      gap: 8px;
    }}

    .toolbar-badge {{
      background: rgba(14, 165, 233, 0.15);
      color: #38bdf8;
      border: 1px solid rgba(14, 165, 233, 0.3);
      padding: 3px 10px;
      border-radius: 999px;
      font-size: 12px;
      font-weight: 600;
    }}

    .toolbar-actions {{
      display: flex;
      align-items: center;
      gap: 12px;
    }}

    .btn-action-print {{
      background: #0284c7;
      color: #ffffff;
      border: none;
      border-radius: 6px;
      padding: 8px 18px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 7px;
      transition: all 0.2s;
    }}
    .btn-action-print:hover {{
      background: #0369a1;
      transform: translateY(-1px);
    }}

    .btn-action-close {{
      background: rgba(255,255,255,0.08);
      color: #cbd5e1;
      border: 1px solid rgba(255,255,255,0.15);
      border-radius: 6px;
      padding: 8px 14px;
      font-size: 13px;
      cursor: pointer;
      transition: all 0.2s;
    }}
    .btn-action-close:hover {{
      background: rgba(255,255,255,0.16);
      color: #ffffff;
    }}

    /* Document Wrapper */
    .document-container {{
      padding: 24px 0 60px 0;
    }}

    /* A4 Single Slip Page */
    .slip-page {{
      background: #ffffff;
      width: 210mm;
      height: 297mm;
      max-height: 297mm;
      margin: 0 auto 30px auto;
      padding: 8mm 12mm 8mm 12mm;
      box-shadow: 0 10px 30px rgba(0,0,0,0.35);
      border-radius: 2px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      box-sizing: border-box;
      page-break-after: always;
      position: relative;
    }}

    /* Slip Header */
    .slip-header {{
      border-bottom: 2px solid #0f172a;
      padding-bottom: 6px;
      margin-bottom: 7px;
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
    }}

    .company-brand {{
      font-size: 16px;
      font-weight: 800;
      color: #0f172a;
      letter-spacing: -0.02em;
    }}

    .company-sub {{
      font-size: 10px;
      color: #475569;
      margin-top: 2px;
    }}

    .doc-title-block {{
      text-align: right;
    }}

    .doc-main-title {{
      font-size: 12.5px;
      font-weight: 800;
      color: #0284c7;
      letter-spacing: -0.01em;
    }}

    .doc-period {{
      font-size: 11px;
      font-weight: 700;
      color: #0f172a;
      margin-top: 1px;
    }}

    .doc-ref {{
      font-size: 8.5px;
      color: #64748b;
    }}

    /* Person Info Grid */
    .person-info-box {{
      border: 1px solid #cbd5e1;
      border-radius: 4px;
      background: #f8fafc;
      padding: 6px 10px;
      margin-bottom: 7px;
      display: grid;
      grid-template-columns: 1.4fr 1fr 1fr;
      gap: 4px 12px;
      font-size: 10px;
    }}

    .info-field {{
      display: flex;
      align-items: center;
      gap: 5px;
    }}

    .info-label {{
      color: #475569;
      font-weight: 600;
      font-size: 9.5px;
    }}

    .info-val {{
      color: #0f172a;
      font-weight: 700;
    }}

    /* KPI Summary Strip */
    .kpi-strip {{
      display: grid;
      grid-template-columns: repeat(5, 1fr);
      gap: 6px;
      margin-bottom: 7px;
    }}

    .kpi-strip-item {{
      border: 1px solid #e2e8f0;
      border-radius: 4px;
      padding: 4px 6px;
      background: #ffffff;
      text-align: center;
    }}

    .kpi-strip-label {{
      font-size: 8.5px;
      color: #64748b;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.02em;
    }}

    .kpi-strip-val {{
      font-size: 12px;
      font-weight: 800;
      font-family: 'JetBrains Mono', monospace;
      color: #0f172a;
      margin-top: 1px;
    }}

    /* 30-Day Table */
    .table-wrapper {{
      width: 100%;
    }}

    .days-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 9px;
    }}

    .days-table th {{
      background: #0f172a;
      color: #ffffff;
      padding: 3px 2px;
      font-weight: 600;
      text-align: center;
      border: 1px solid #0f172a;
      font-size: 8.5px;
    }}

    .days-table td {{
      padding: 1.8px 3px;
      border: 1px solid #cbd5e1;
      color: #1e293b;
      height: 16.5px;
    }}

    .row-sunday {{
      background-color: #f1f5f9 !important;
      font-weight: 600;
    }}

    .row-sunday td {{
      color: #475569;
    }}

    .cell-mono {{
      font-family: 'JetBrains Mono', monospace;
      font-size: 8.5px;
    }}

    .cell-ot {{
      font-weight: 700;
      color: #c2410c;
    }}

    .cell-bonus {{
      font-weight: 700;
      color: #7e22ce;
    }}

    .cell-final {{
      font-weight: 800;
      color: #0f172a;
    }}

    .note-col {{
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 140px;
    }}

    .table-total-row {{
      background: #e2e8f0 !important;
      font-weight: 800;
    }}

    .table-total-row td {{
      border-top: 2px solid #0f172a;
      border-bottom: 2px solid #0f172a;
      padding: 3px 2px;
      font-size: 9.5px;
    }}

    /* Legal disclaimer & signatures */
    .slip-footer {{
      margin-top: 6px;
      padding-top: 5px;
      border-top: 1px solid #cbd5e1;
    }}

    .legal-statement {{
      font-size: 8px;
      line-height: 1.3;
      color: #475569;
      text-align: justify;
      font-style: italic;
      margin-bottom: 7px;
    }}

    .signatures-box {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
    }}

    .sig-card {{
      border: 1px solid #94a3b8;
      border-radius: 4px;
      padding: 6px 8px;
      min-height: 60px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }}

    .sig-role {{
      font-size: 9.5px;
      font-weight: 800;
      color: #0f172a;
      border-bottom: 1px dashed #cbd5e1;
      padding-bottom: 2px;
      margin-bottom: 3px;
    }}

    .sig-meta {{
      font-size: 8.5px;
      color: #475569;
      line-height: 1.25;
    }}

    .sig-line {{
      margin-top: 16px;
      display: flex;
      justify-content: space-between;
      font-size: 8.5px;
      color: #475569;
    }}

    /* Print Rules */
    @media print {{
      @page {{
        size: A4 portrait;
        margin: 6mm 10mm;
      }}

      body {{
        background: #ffffff !important;
        color: #000000 !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }}

      .no-print {{
        display: none !important;
      }}

      .document-container {{
        padding: 0 !important;
      }}

      .slip-page {{
        box-shadow: none !important;
        border: none !important;
        margin: 0 !important;
        padding: 0 0 2mm 0 !important;
        width: 100% !important;
        height: 284mm !important;
        max-height: 284mm !important;
        page-break-after: always !important;
        break-after: page !important;
      }}

      .days-table th {{
        background: #0f172a !important;
        color: #ffffff !important;
      }}

      .row-sunday {{
        background-color: #f1f5f9 !important;
      }}

      .table-total-row {{
        background-color: #e2e8f0 !important;
      }}
    }}
  </style>
</head>
<body>

  <!-- Screen Floating Toolbar (Hidden when printing) -->
  <header class="no-print screen-toolbar">
    <div class="toolbar-left">
      <div class="brand-logo-text">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:20px;height:20px;color:#38bdf8;">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
          <polyline points="14 2 14 8 20 8"></polyline>
          <line x1="16" y1="13" x2="8" y2="13"></line>
          <line x1="16" y1="17" x2="8" y2="17"></line>
          <polyline points="10 9 9 9 8 9"></polyline>
        </svg>
        <span>FİDE KONSERVE &bull; {title}</span>
      </div>
      <span class="toolbar-badge">{period_str}</span>
      <span class="toolbar-badge" style="background: rgba(16, 185, 129, 0.15); color: #34d399; border-color: rgba(16, 185, 129, 0.3);">
        Toplam {total_count} Personel Sayfası
      </span>
    </div>

    <div class="toolbar-actions">
      <button class="btn-action-print" onclick="window.print()">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:16px;height:16px;">
          <polyline points="6 9 6 2 18 2 18 9"></polyline>
          <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path>
          <rect x="6" y="14" width="12" height="8"></rect>
        </svg>
        <span>Yazdır / PDF Olarak Kaydet (Ctrl+P)</span>
      </button>
      <button class="btn-action-close" onclick="window.close()">Pencereyi Kapat</button>
    </div>
  </header>

  <!-- Printable Pages Container -->
  <main class="document-container">
    {slips_joined}
  </main>

  {auto_print_script}
</body>
</html>
"""
