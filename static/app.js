/**
 * ============================================================================
 * FİDE KONSERVE — PDKS & PUANTAJ YÖNETİM PORTALI
 * Modern İstemci Mantığı ve Arayüz Denetleyicisi (Vanilla JavaScript)
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// UYGULAMA DURUMU (APP STATE)
// ----------------------------------------------------------------------------
const AppState = {
  activeTab: 'tab-overview',
  stats: null,
  rules: null,
  
  // Puantaj Matrisi Durumu
  matrix: {
    page: 1,
    pageSize: 50,
    search: '',
    department: 'all',
    status: 'active',
    totalPages: 1,
    totalCount: 0,
    items: [],
    daysInMonth: 30,
    sundays: [6, 13, 20, 27]
  },

  // İstisnalar Masası Durumu
  exceptions: {
    status: 'all',
    search: '',
    items: [],
    pendingCount: 0,
    resolvedCount: 0
  },

  // Bölüm Primleri Durumu
  bonuses: {
    department: 'all',
    search: '',
    items: [],
    totalHours: 0
  },

  // Maliyet & Bütçe Radarı Durumu
  financial: {
    page: 1,
    pageSize: 50,
    search: '',
    department: 'all',
    totalPages: 1,
    totalCount: 0,
    items: [],
    departments: []
  },

  // Yasal Uyum & Risk Radarı Durumu
  compliance: {
    page: 1,
    pageSize: 50,
    search: '',
    department: 'all',
    violationType: 'all',
    severity: 'all',
    totalPages: 1,
    totalCount: 0,
    items: [],
    departments: []
  }
};

// ----------------------------------------------------------------------------
// YARDIMCI BİÇİMLENDİRME FONKSİYONLARI
// ----------------------------------------------------------------------------
function fmtHours(val) {
  if (val === null || val === undefined || isNaN(val)) return '0,0';
  return Number(val).toLocaleString('tr-TR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function fmtInt(val) {
  if (val === null || val === undefined || isNaN(val)) return '0';
  return Number(val).toLocaleString('tr-TR');
}

function fmtCurrency(val) {
  if (val === null || val === undefined || isNaN(val)) return '₺0,00';
  return '₺' + Number(val).toLocaleString('tr-TR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function debounce(func, wait) {
  let timeout;
  return function(...args) {
    clearTimeout(timeout);
    timeout = setTimeout(() => func.apply(this, args), wait);
  };
}

// ----------------------------------------------------------------------------
// TOAST BİLDİRİM SİSTEMİ
// ----------------------------------------------------------------------------
function showToast(message, type = 'info', duration = 3500) {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `
    <span class="toast-text">${message}</span>
  `;

  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(30px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// ----------------------------------------------------------------------------
// MODAL YÖNETİMİ
// ----------------------------------------------------------------------------
function openModal(modalId) {
  const el = document.getElementById(modalId);
  if (el) el.classList.add('active');
}

function closeModal(modalId) {
  const el = document.getElementById(modalId);
  if (el) el.classList.remove('active');
}

// ----------------------------------------------------------------------------
// SEKME (TAB) DEĞİŞTİRME MANTIĞI
// ----------------------------------------------------------------------------
function switchTab(tabId) {
  AppState.activeTab = tabId;

  // Nav butonlarını güncelle
  document.querySelectorAll('.nav-tab').forEach(btn => {
    if (btn.getAttribute('data-tab') === tabId) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });

  // İçerik panellerini güncelle
  document.querySelectorAll('.tab-content').forEach(section => {
    if (section.id === tabId) {
      section.classList.add('active');
    } else {
      section.classList.remove('active');
    }
  });

  // İlgili sekmenin verisini tazele
  if (tabId === 'tab-overview') loadStats();
  if (tabId === 'tab-matrix') loadMatrix();
  if (tabId === 'tab-exceptions') loadExceptions();
  if (tabId === 'tab-bonuses') loadBonuses();
  if (tabId === 'tab-financial') loadFinancialRadar();
  if (tabId === 'tab-compliance') loadComplianceRadar();
  if (tabId === 'tab-reports') loadRules();
}

// ----------------------------------------------------------------------------
// 1. GENEL BAKIŞ & GRAFİKLER (OVERVIEW)
// ----------------------------------------------------------------------------
async function loadStats() {
  try {
    const res = await fetch('/api/stats');
    if (!res.ok) throw new Error('İstatistikler alınamadı');
    const data = await res.json();
    AppState.stats = data;

    // Header & KPI Değerleri
    const periodText = document.getElementById('periodText');
    if (periodText) periodText.innerText = `${data.period} Dönemi`;

    const statusLabel = document.getElementById('statusLabel');
    if (statusLabel) statusLabel.innerText = `PDKS Motoru Aktif (${data.total_personnel} Personel)`;

    document.getElementById('valTotalPersonnel').innerText = fmtInt(data.total_personnel);
    document.getElementById('subTotalPersonnel').innerText = `${data.inactive_personnel} Kart Basmayan`;

    document.getElementById('valActivePersonnel').innerText = fmtInt(data.active_personnel);
    const activePct = ((data.active_personnel / data.total_personnel) * 100).toFixed(1);
    document.getElementById('subActivePersonnel').innerText = `%${activePct.replace('.', ',')} Katılım Oranı`;

    document.getElementById('valTotalHours').innerHTML = `${fmtHours(data.total_hours)} <small>saat</small>`;
    document.getElementById('subBaseHours').innerText = `${fmtHours(data.total_base_hours)}s Normal Mesai`;

    document.getElementById('valOvertimeHours').innerHTML = `${fmtHours(data.total_overtime_hours)} <small>saat</small>`;
    document.getElementById('valBonusHours').innerHTML = `${fmtHours(data.total_bonus_hours)} <small>saat</small>`;
    document.getElementById('subBonusCount').innerText = `${data.total_bonuses_count} Hak Ediş Kaydı`;

    document.getElementById('valExceptions').innerText = fmtInt(data.total_exceptions);
    document.getElementById('subResolvedExceptions').innerText = `${data.resolved_exceptions} Çözüldü, ${data.pending_exceptions} Bekliyor`;

    // Tab rozetleri
    const badgeActive = document.getElementById('badgeTotalActive');
    if (badgeActive) badgeActive.innerText = data.active_personnel;

    const badgeExceptions = document.getElementById('badgePendingExceptions');
    if (badgeExceptions) badgeExceptions.innerText = data.pending_exceptions;

    const badgeBonuses = document.getElementById('badgeBonusesCount');
    if (badgeBonuses) badgeBonuses.innerText = data.total_bonuses_count;

    const badgeFin = document.getElementById('badgeFinancialOtCost');
    if (badgeFin && data.financial_summary) {
      const otM = (data.financial_summary.total_overtime_cost / 1000000).toFixed(2).replace('.', ',');
      badgeFin.innerText = `₺${otM}M`;
    }

    const badgeComp = document.getElementById('badgeComplianceScore');
    if (badgeComp && data.compliance_summary) {
      const score = Math.round(data.compliance_summary.compliance_index);
      badgeComp.innerText = `%${score}`;
    }

    // 30 Günlük Grafik & Departman Dağılımını Çiz
    renderDailyTrendChart(data.daily_trends);
    renderDepartmentBreakdown(data.departments);

    // Matris için departman listesini güncelle
    populateDeptDropdowns(data.departments);

  } catch (err) {
    console.error(err);
    showToast('İstatistikler yüklenirken hata oluştu: ' + err.message, 'error');
  }
}

function renderDailyTrendChart(trends) {
  const container = document.getElementById('dailyTrendChartContainer');
  if (!container || !trends || trends.length === 0) return;

  const width = container.clientWidth || 700;
  const height = 240;
  const padLeft = 40;
  const padBottom = 30;
  const padTop = 20;
  const padRight = 20;

  const chartW = width - padLeft - padRight;
  const chartH = height - padTop - padBottom;

  // Max değerleri bul
  const maxHours = Math.max(...trends.map(t => t.total_hours), 1500);
  const barWidth = Math.max(8, (chartW / trends.length) - 6);

  let svg = `
    <svg class="chart-svg" viewBox="0 0 ${width} ${height}">
      <defs>
        <linearGradient id="cyanGradient" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#06b6d4" stop-opacity="0.9"/>
          <stop offset="100%" stop-color="#0284c7" stop-opacity="0.3"/>
        </linearGradient>
        <linearGradient id="amberGradient" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" stop-color="#f59e0b" stop-opacity="1"/>
          <stop offset="100%" stop-color="#d97706" stop-opacity="0.6"/>
        </linearGradient>
      </defs>
  `;

  // Yatay Kılavuz Çizgileri
  const gridSteps = 4;
  for (let i = 0; i <= gridSteps; i++) {
    const yVal = padTop + (chartH / gridSteps) * i;
    const labelVal = Math.round(maxHours - (maxHours / gridSteps) * i);
    svg += `
      <line x1="${padLeft}" y1="${yVal}" x2="${width - padRight}" y2="${yVal}" stroke="rgba(255,255,255,0.06)" stroke-dasharray="3 3"/>
      <text x="${padLeft - 8}" y="${yVal + 4}" fill="#64748b" font-size="10" font-family="'JetBrains Mono', monospace" text-anchor="end">${labelVal}</text>
    `;
  }

  // Barlar & Çizgiler
  trends.forEach((t, idx) => {
    const x = padLeft + idx * (chartW / trends.length) + ((chartW / trends.length) - barWidth) / 2;
    const barH = (t.total_hours / maxHours) * chartH;
    const y = padTop + chartH - barH;

    const otH = (t.overtime_hours / maxHours) * chartH;
    const otY = padTop + chartH - otH;

    const isSun = t.is_sunday;
    const barFill = isSun ? '#eab308' : 'url(#cyanGradient)';
    const barOpacity = isSun ? '0.45' : '1';

    svg += `
      <g class="chart-bar-group" data-day="${t.day}">
        <rect x="${x}" y="${y}" width="${barWidth}" height="${Math.max(2, barH)}" rx="3" fill="${barFill}" opacity="${barOpacity}">
          <title>${t.date} ${isSun ? '(Pazar)' : ''}&#10;Toplam: ${t.total_hours}s&#10;Fazla Mesai: ${t.overtime_hours}s&#10;Aktif Personel: ${t.active_count}</title>
        </rect>
    `;

    // Fazla mesai vurgusu (eğer varsa)
    if (t.overtime_hours > 0 && !isSun) {
      svg += `
        <rect x="${x}" y="${otY}" width="${barWidth}" height="${Math.max(2, otH)}" rx="2" fill="url(#amberGradient)">
          <title>${t.date} Fazla Mesai: ${t.overtime_hours}s</title>
        </rect>
      `;
    }

    // Gün etiketleri (Her 3 günde bir veya pazarlarda)
    if (t.day % 2 !== 0 || isSun) {
      svg += `
        <text x="${x + barWidth / 2}" y="${height - 8}" fill="${isSun ? '#eab308' : '#94a3b8'}" font-size="9" font-family="'JetBrains Mono', monospace" text-anchor="middle" font-weight="${isSun ? '700' : '500'}">${t.day}</text>
      `;
    }

    svg += `</g>`;
  });

  svg += `</svg>`;
  container.innerHTML = svg;
}

function renderDepartmentBreakdown(departments) {
  const container = document.getElementById('deptBreakdownContainer');
  if (!container || !departments) return;

  const deptList = Object.entries(departments)
    .map(([name, data]) => ({ name, ...data }))
    .sort((a, b) => b.total_hours - a.total_hours);

  const maxDeptHours = Math.max(...deptList.map(d => d.total_hours), 1);

  let html = '';
  deptList.slice(0, 8).forEach(d => {
    const pct = ((d.total_hours / maxDeptHours) * 100).toFixed(0);
    html += `
      <div class="dept-item">
        <div class="dept-header-line">
          <span class="dept-name" title="${d.name}">${d.name}</span>
          <span class="dept-stats">${fmtHours(d.total_hours)}s <small>(${d.active} aktif / ${d.count})</small></span>
        </div>
        <div class="progress-track">
          <div class="progress-bar" style="width: ${pct}%;"></div>
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

function populateDeptDropdowns(departments) {
  const select = document.getElementById('matrixDeptFilter');
  if (!select || !departments) return;

  const currentVal = select.value;
  let options = '<option value="all">Tüm Bölümler (Hepsi)</option>';

  Object.keys(departments).sort().forEach(dept => {
    options += `<option value="${dept}">${dept}</option>`;
  });

  select.innerHTML = options;
  if (currentVal) select.value = currentVal;
}

// ----------------------------------------------------------------------------
// 2. PUANTAJ MATRİSİ (MATRIX)
// ----------------------------------------------------------------------------
async function loadMatrix() {
  const tbody = document.getElementById('matrixTableBody');
  const thead = document.getElementById('matrixTableHeader');
  if (!tbody || !thead) return;

  tbody.innerHTML = `<tr><td colspan="40" style="text-align:center; padding: 40px; color: var(--text-dim);">Puantaj verileri yükleniyor...</td></tr>`;

  try {
    const params = new URLSearchParams({
      page: AppState.matrix.page,
      page_size: AppState.matrix.pageSize,
      status: AppState.matrix.status,
      search: AppState.matrix.search,
      department: AppState.matrix.department
    });

    const res = await fetch(`/api/matrix?${params.toString()}`);
    if (!res.ok) throw new Error('Matris verisi alınamadı');
    const data = await res.json();

    AppState.matrix.items = data.items;
    AppState.matrix.totalPages = data.total_pages;
    AppState.matrix.totalCount = data.total;
    AppState.matrix.daysInMonth = data.days_in_month;
    AppState.matrix.sundays = data.sundays || [];

    renderMatrixTableHeader(data.days_in_month, data.sundays);
    renderMatrixTableBody(data.items, data.days_in_month, data.sundays);
    updateMatrixPagination(data);

  } catch (err) {
    console.error(err);
    tbody.innerHTML = `<tr><td colspan="40" style="text-align:center; color: var(--color-danger); padding: 30px;">Hata: ${err.message}</td></tr>`;
  }
}

function renderMatrixTableHeader(daysInMonth, sundays) {
  const thead = document.getElementById('matrixTableHeader');
  if (!thead) return;

  let headerHtml = `
    <tr>
      <th style="width: 40px;">Sıra</th>
      <th style="width: 100px;">TC Kimlik</th>
      <th style="width: 160px;">Adı Soyadı</th>
      <th style="width: 140px;">Bölüm</th>
  `;

  for (let d = 1; d <= daysInMonth; d++) {
    const isSun = sundays.includes(d);
    headerHtml += `
      <th class="th-day ${isSun ? 'th-sunday' : ''}">${d.toString().padStart(2, '0')}</th>
    `;
  }

  headerHtml += `
      <th style="width: 70px; text-align: right;">Ç. Gün</th>
      <th style="width: 80px; text-align: right;">Normal (s)</th>
      <th style="width: 80px; text-align: right;">FM (s)</th>
      <th style="width: 75px; text-align: right;">Prim (s)</th>
      <th style="width: 85px; text-align: right;">Toplam (s)</th>
    </tr>
  `;

  thead.innerHTML = headerHtml;
}

function renderMatrixTableBody(items, daysInMonth, sundays) {
  const tbody = document.getElementById('matrixTableBody');
  if (!tbody) return;

  if (items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="40" style="text-align:center; padding: 40px; color: var(--text-dim);">Arama kriterlerine uygun personel kaydı bulunamadı.</td></tr>`;
    return;
  }

  let rowsHtml = '';
  items.forEach(p => {
    rowsHtml += `<tr onclick="openPersonnelPunchesModal('${p.name_key}')" title="Kart hareketleri detayını görmek için tıklayın">`;
    rowsHtml += `<td>${p.sno || '-'}</td>`;
    rowsHtml += `<td class="cell-mono">${p.tc || '-'}</td>`;
    rowsHtml += `<td class="cell-bold">${p.ad_soyad}</td>`;
    rowsHtml += `<td title="${p.bolum}">${p.bolum || '-'}</td>`;

    // 1..daysInMonth
    for (let d = 1; d <= daysInMonth; d++) {
      const isSun = sundays.includes(d);
      const detail = p.day_details ? p.day_details[d] : null;
      const h = detail ? detail.hours : (p.daily_hours ? p.daily_hours[d] : 0.0);

      let cellClass = 'td-day';
      if (isSun) cellClass += ' td-sunday';

      if (h === 0 || !h) {
        rowsHtml += `<td class="${cellClass}"><span class="cell-zero">-</span></td>`;
      } else {
        let pillClass = 'cell-pill';
        if (detail && detail.is_bonus) {
          pillClass += ' has-bonus';
        } else if (detail && detail.is_exception) {
          pillClass += ' has-exc';
        } else if (detail && detail.overtime > 0) {
          pillClass += ' has-ot';
        }
        rowsHtml += `<td class="${cellClass}"><span class="${pillClass}">${fmtHours(h)}</span></td>`;
      }
    }

    // Toplamlar
    rowsHtml += `<td class="cell-mono" style="text-align: right; font-weight: 600;">${p.total_work_days}</td>`;
    rowsHtml += `<td class="cell-mono" style="text-align: right;">${fmtHours(p.total_base_hours)}</td>`;
    rowsHtml += `<td class="cell-mono cell-green" style="text-align: right; font-weight: 600;">${fmtHours(p.total_ot_hours)}</td>`;
    rowsHtml += `<td class="cell-mono cell-purple" style="text-align: right; font-weight: 600;">${fmtHours(p.total_bonus_hours)}</td>`;
    rowsHtml += `<td class="cell-mono cell-accent cell-bold" style="text-align: right;">${fmtHours(p.total_hours)}</td>`;

    rowsHtml += `</tr>`;
  });

  tbody.innerHTML = rowsHtml;
}

function updateMatrixPagination(data) {
  const info = document.getElementById('matrixPaginationInfo');
  const display = document.getElementById('matrixPageDisplay');
  const btnPrev = document.getElementById('btnMatrixPrev');
  const btnNext = document.getElementById('btnMatrixNext');

  const startIdx = (data.page - 1) * data.page_size + 1;
  const endIdx = Math.min(data.page * data.page_size, data.total);

  if (info) {
    info.innerText = data.total > 0
      ? `Gösterilen: ${startIdx} - ${endIdx} / ${fmtInt(data.total)} Personel`
      : 'Kayıt bulunamadı';
  }

  if (display) display.innerText = `Sayfa ${data.page} / ${data.total_pages || 1}`;
  if (btnPrev) btnPrev.disabled = data.page <= 1;
  if (btnNext) btnNext.disabled = data.page >= data.total_pages;
}

// ----------------------------------------------------------------------------
// 3. İSTİSNA & AMİR ONAY MASASI (EXCEPTIONS)
// ----------------------------------------------------------------------------
async function loadExceptions() {
  const tbody = document.getElementById('exceptionsTableBody');
  if (!tbody) return;

  tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding: 30px; color: var(--text-dim);">İstisnalar yükleniyor...</td></tr>`;

  try {
    const params = new URLSearchParams({
      status: AppState.exceptions.status,
      search: AppState.exceptions.search
    });

    const res = await fetch(`/api/exceptions?${params.toString()}`);
    if (!res.ok) throw new Error('İstisnalar alınamadı');
    const data = await res.json();

    AppState.exceptions.items = data.items;
    AppState.exceptions.pendingCount = data.pending_count;
    AppState.exceptions.resolvedCount = data.resolved_count;

    // Buton sayaçlarını güncelle
    const allBtn = document.getElementById('segBtnAllExceptions');
    const pendBtn = document.getElementById('segBtnPendingExceptions');
    const resBtn = document.getElementById('segBtnResolvedExceptions');

    if (allBtn) allBtn.innerText = `Tümü (${data.total})`;
    if (pendBtn) pendBtn.innerText = `Onay Bekleyenler (${data.pending_count})`;
    if (resBtn) resBtn.innerText = `Çözülenler (${data.resolved_count})`;

    renderExceptionsTable(data.items);

  } catch (err) {
    console.error(err);
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; color: var(--color-danger); padding: 30px;">Hata: ${err.message}</td></tr>`;
  }
}

function renderExceptionsTable(items) {
  const tbody = document.getElementById('exceptionsTableBody');
  if (!tbody) return;

  if (items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding: 30px; color: var(--text-dim);">Filtreye uygun istisna kaydı bulunmuyor.</td></tr>`;
    return;
  }

  let html = '';
  items.forEach((rec, idx) => {
    const isResolved = rec.is_resolved;
    const resTag = isResolved
      ? `<span class="status-tag tag-resolved">✓ ÇÖZÜLDÜ</span>`
      : `<span class="status-tag tag-pending">ONAY BEKLİYOR</span>`;

    html += `
      <tr>
        <td>${idx + 1}</td>
        <td class="cell-mono">${rec.tarih}</td>
        <td class="cell-bold">${rec.ad_soyad}</td>
        <td>${rec.bolum || '-'}</td>
        <td class="cell-mono">${rec.all_punches || rec.raw_g_saat || '-'}</td>
        <td class="cell-mono">${fmtHours(rec.final_hours)}s</td>
        <td>
          <span style="color: var(--color-warning); font-weight: 600;">${rec.status_type}</span>
          <div style="font-size: 11px; color: var(--text-muted);">${rec.audit_note || ''}</div>
        </td>
        <td>
          ${resTag}
          ${isResolved && rec.resolution ? `<div style="font-size: 10px; color: var(--color-success);">${rec.resolution.resolved_by || 'Amir'}: ${rec.resolution.approved_hours}s</div>` : ''}
        </td>
        <td style="text-align: center;">
          <button class="btn btn-sm ${isResolved ? 'btn-outline' : 'btn-warning'}" onclick='openResolveModal(${JSON.stringify(rec)})'>
            ${isResolved ? 'Düzenle' : 'Onayla'}
          </button>
        </td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

function openResolveModal(rec) {
  const summaryBox = document.getElementById('modalExceptionSummary');
  const nameKeyInput = document.getElementById('modalNameKey');
  const dayInput = document.getElementById('modalDay');
  const hoursInput = document.getElementById('modalApprovedHours');
  const noteInput = document.getElementById('modalNote');
  const resolverInput = document.getElementById('modalResolver');

  if (!summaryBox || !nameKeyInput || !dayInput) return;

  nameKeyInput.value = rec.name_key || rec.ad_soyad;
  dayInput.value = rec.gun;
  hoursInput.value = rec.final_hours > 0 ? rec.final_hours : 7.5;

  if (rec.resolution && rec.resolution.note) {
    noteInput.value = rec.resolution.note;
    resolverInput.value = rec.resolution.resolved_by || 'Vardiya Amiri';
  } else {
    noteInput.value = `Saha amirinden teyit alındı: ${rec.tarih} tarihinde fiilen çalıştı.`;
  }

  summaryBox.innerHTML = `
    <div><strong>Personel:</strong> ${rec.ad_soyad} (Sicil: ${rec.sicil || '-'})</div>
    <div><strong>Tarih:</strong> ${rec.tarih} (${rec.gun_adi || ''})</div>
    <div><strong>Bölüm:</strong> ${rec.bolum || '-'}</div>
    <div><strong>Turnike Basımları:</strong> <span class="cell-mono">${rec.all_punches || rec.raw_g_saat || '-'}</span></div>
    <div><strong>Hata Durumu:</strong> <span style="color:var(--color-warning);">${rec.status_type} (${rec.audit_note || 'Tek Basım'})</span></div>
  `;

  openModal('modalResolveException');
}

// ----------------------------------------------------------------------------
// 4. BÖLÜM PRİMLERİ (BONUSES)
// ----------------------------------------------------------------------------
async function loadBonuses() {
  const tbody = document.getElementById('bonusesTableBody');
  if (!tbody) return;

  tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding: 30px; color: var(--text-dim);">Prim kayıtları yükleniyor...</td></tr>`;

  try {
    const params = new URLSearchParams({
      department: AppState.bonuses.department,
      search: AppState.bonuses.search
    });

    const res = await fetch(`/api/bonuses?${params.toString()}`);
    if (!res.ok) throw new Error('Bölüm primleri alınamadı');
    const data = await res.json();

    AppState.bonuses.items = data.items;
    AppState.bonuses.totalHours = data.total_bonus_hours;

    document.getElementById('valBonusTotalHours').innerHTML = `${fmtHours(data.total_bonus_hours)} <small>saat</small>`;

    // Balık Dolum/Kesim vs Üretim toplamları
    let fishHours = 0;
    let prodHours = 0;
    Object.entries(data.dept_summary || {}).forEach(([dept, s]) => {
      const dUpper = dept.toUpperCase();
      if (dUpper.includes('BALIK')) fishHours += s.total_hours;
      else if (dUpper.includes('URETIM') || dUpper.includes('ÜRETİM')) prodHours += s.total_hours;
    });

    const valFish = document.getElementById('valFishBonus');
    if (valFish) valFish.innerHTML = `${fmtHours(fishHours)} <small>saat</small>`;

    const valProd = document.getElementById('valProdBonus');
    if (valProd) valProd.innerHTML = `${fmtHours(prodHours)} <small>saat</small>`;

    renderBonusesTable(data.items);

  } catch (err) {
    console.error(err);
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; color: var(--color-danger); padding: 30px;">Hata: ${err.message}</td></tr>`;
  }
}

function renderBonusesTable(items) {
  const tbody = document.getElementById('bonusesTableBody');
  if (!tbody) return;

  if (items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding: 30px; color: var(--text-dim);">Prim kaydı bulunamadı.</td></tr>`;
    return;
  }

  let html = '';
  items.forEach((rec, idx) => {
    html += `
      <tr>
        <td>${idx + 1}</td>
        <td class="cell-mono">${rec.tarih}</td>
        <td class="cell-bold">${rec.ad_soyad}</td>
        <td>${rec.bolum}</td>
        <td class="cell-mono">${rec.all_punches || (rec.raw_g_saat + ' - ' + rec.raw_c_saat)}</td>
        <td class="cell-mono">${fmtHours(rec.fiili_sure)}s</td>
        <td class="cell-mono cell-purple" style="font-weight: 700;">+${fmtHours(rec.bonus_hours)}s</td>
        <td class="cell-mono cell-accent" style="font-weight: 700;">${fmtHours(rec.final_hours)}s</td>
        <td style="color: var(--text-muted); font-size: 11px;">${rec.audit_note || 'Bölüm prim kuralı uygulandı'}</td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

// ----------------------------------------------------------------------------
// 5. MALİYET & BÜTÇE RADARI (FINANCIAL & BUDGET RADAR)
// ----------------------------------------------------------------------------
async function loadFinancialRadar() {
  const tbody = document.getElementById('financialPersonnelTableBody');
  if (tbody) {
    tbody.innerHTML = `<tr><td colspan="17" style="text-align:center; padding: 40px; color: var(--text-dim);">Finans ve bordro verileri hesaplanıyor...</td></tr>`;
  }

  try {
    const params = new URLSearchParams({
      page: AppState.financial.page,
      page_size: AppState.financial.pageSize,
      search: AppState.financial.search,
      department: AppState.financial.department
    });

    const res = await fetch(`/api/financial-radar?${params.toString()}`);
    if (!res.ok) throw new Error('Finansal veriler alınamadı');
    const data = await res.json();

    AppState.financial.items = data.items;
    AppState.financial.totalCount = data.total_count;
    AppState.financial.totalPages = data.total_pages;

    // 1. KPI Kartları
    const s = data.summary;
    const valOt = document.getElementById('valFinancialOtCost');
    if (valOt) valOt.innerText = fmtCurrency(s.total_overtime_cost);

    const subOt = document.getElementById('subFinancialOtHours');
    if (subOt) subOt.innerText = `${fmtHours(s.total_overtime_hours)}s Fazla Mesai (4857 SK %50)`;

    const valPayroll = document.getElementById('valFinancialPayrollCost');
    if (valPayroll) valPayroll.innerText = fmtCurrency(s.total_payroll_budget);

    const subBase = document.getElementById('subFinancialBaseCost');
    if (subBase) subBase.innerText = `Maaş Hakedişi (₺${(s.total_base_cost / 1000000).toFixed(2).replace('.', ',')}M) + FM`;

    const valBonus = document.getElementById('valFinancialBonusCost');
    if (valBonus) valBonus.innerText = fmtCurrency(s.total_bonus_cost);

    const subBonus = document.getElementById('subFinancialBonusHours');
    if (subBonus) subBonus.innerText = `${fmtHours(s.total_bonus_hours)}s Prim Karşılığı`;

    const valRate = document.getElementById('valFinancialAvgRate');
    if (valRate) valRate.innerText = fmtCurrency(s.average_overtime_rate_per_hour);

    // Navbar rozeti
    const badgeNav = document.getElementById('badgeFinancialOtCost');
    if (badgeNav) badgeNav.innerText = `₺${(s.total_overtime_cost / 1000000).toFixed(2).replace('.', ',')}M`;

    // 2. Departman Dağılım Çubukları
    renderDeptCostBars(data.department_breakdown);

    // 3. Departman Analiz & Kişi Başı Yük Kartları
    renderDeptInsights(data.department_breakdown);

    // 4. Personel Tablosu
    renderFinancialTable(data.items, (data.page - 1) * data.page_size);

    // 5. Sayfalama
    renderFinancialPagination(data);

    // 6. Departman Filtre Dropdown
    populateFinancialDeptDropdown(data.departments);

  } catch (err) {
    console.error(err);
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="17" style="text-align:center; color: var(--color-danger); padding: 30px;">Hata: ${err.message}</td></tr>`;
    }
  }
}

function renderDeptCostBars(depts) {
  const container = document.getElementById('deptCostBarsContainer');
  if (!container || !depts) return;

  const maxCost = Math.max(...depts.map(d => d.overtime_cost), 1);

  let html = '';
  depts.slice(0, 8).forEach(d => {
    const widthPct = Math.min(100, Math.max(4, (d.overtime_cost / maxCost) * 100)).toFixed(1);
    const isHigh = d.overtime_budget_share_pct > 12.0;

    html += `
      <div class="dept-cost-item">
        <div class="dept-cost-header">
          <div class="dept-cost-name">
            <span>${d.department}</span>
            <small style="color: var(--text-dim); font-weight: normal;">(${d.headcount} kişi)</small>
          </div>
          <div class="dept-cost-meta">
            <span class="dept-cost-amount">${fmtCurrency(d.overtime_cost)}</span>
            <span class="dept-cost-pct">%${d.overtime_budget_share_pct.toFixed(1).replace('.', ',')}</span>
          </div>
        </div>
        <div class="cost-track">
          <div class="cost-fill ${isHigh ? 'fill-high' : ''}" style="width: ${widthPct}%;"></div>
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

function renderDeptInsights(depts) {
  const container = document.getElementById('deptInsightsContainer');
  if (!container || !depts) return;

  let html = '';
  depts.slice(0, 6).forEach(d => {
    let tagHtml = '<span class="dept-badge-tag tag-normal">Dengeli</span>';
    if (d.avg_overtime_hours_per_worker > 30) {
      tagHtml = '<span class="dept-badge-tag tag-critical">🚨 Aşırı Mesai Yükü</span>';
    } else if (d.overtime_cost > 100000) {
      tagHtml = '<span class="dept-badge-tag tag-high">⚡ Yüksek Bütçe Payı</span>';
    }

    html += `
      <div class="dept-insight-card">
        <div class="dept-insight-left">
          <div class="dept-insight-title">${d.department}</div>
          <div class="dept-insight-stats">
            <span>Kadron: <strong>${d.headcount}</strong></span>
            <span>Ort. Mesai: <strong>${fmtHours(d.avg_overtime_hours_per_worker)}s</strong></span>
            <span>Kişi Başı: <strong>${fmtCurrency(d.avg_overtime_cost_per_worker)}</strong></span>
          </div>
        </div>
        <div class="dept-insight-right">
          ${tagHtml}
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

function renderFinancialTable(items, startIdx) {
  const tbody = document.getElementById('financialPersonnelTableBody');
  if (!tbody) return;

  if (items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="17" style="text-align:center; padding: 40px; color: var(--text-dim);">Arama kriterlerine uygun personel kaydı bulunamadı.</td></tr>`;
    return;
  }

  let html = '';
  items.forEach((p, idx) => {
    const hasDiff = p.cash_difference > 0;
    const diffBadge = hasDiff
      ? `<span class="badge badge-success" style="font-family: var(--font-mono); font-weight:700;">+${fmtCurrency(p.cash_difference)}</span>`
      : `<span style="color: var(--text-dim); font-family: var(--font-mono); font-size:11px;">₺0,00</span>`;

    html += `
      <tr>
        <td>${startIdx + idx + 1}</td>
        <td class="cell-mono">${p.tc_no || '-'}</td>
        <td class="cell-bold">${p.ad_soyad}</td>
        <td>${p.bolum}</td>
        <td style="text-align: center;" class="cell-mono">${p.days_worked}</td>
        <td style="text-align: right;" class="cell-mono">${fmtCurrency(p.net_maas)}</td>
        <td style="text-align: right; color: var(--text-muted);" class="cell-mono">${fmtCurrency(p.hourly_base_rate)}</td>
        <td style="text-align: right; color: var(--color-cyan);" class="cell-mono">${fmtCurrency(p.hourly_overtime_rate)}</td>
        <td style="text-align: right;" class="cell-mono">${fmtHours(p.weekday_ot_hours)}</td>
        <td style="text-align: right;" class="cell-mono">${fmtHours(p.sunday_ot_hours)}</td>
        <td style="text-align: right; font-weight: 700;" class="cell-mono">${fmtHours(p.total_ot_hours)}</td>
        <td style="text-align: right; font-weight: 700; color: var(--color-amber);" class="cell-mono">${fmtCurrency(p.overtime_cost)}</td>
        <td style="text-align: right;" class="cell-mono">${p.bonus_hours > 0 ? '+' + fmtHours(p.bonus_hours) : '-'}</td>
        <td style="text-align: right; color: var(--color-purple);" class="cell-mono">${p.bonus_cost > 0 ? fmtCurrency(p.bonus_cost) : '-'}</td>
        <td style="text-align: right; font-weight: 800; color: #ffffff;" class="cell-mono">${fmtCurrency(p.total_net_earned)}</td>
        <td style="text-align: right; color: var(--text-muted);" class="cell-mono">${fmtCurrency(p.bank_net)}</td>
        <td style="text-align: right;">${diffBadge}</td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

function renderFinancialPagination(data) {
  const info = document.getElementById('financialPaginationInfo');
  if (info) {
    const start = data.total_count === 0 ? 0 : (data.page - 1) * data.page_size + 1;
    const end = Math.min(data.page * data.page_size, data.total_count);
    info.innerText = `Gösterilen: ${start} - ${end} / ${data.total_count} Personel`;
  }

  const disp = document.getElementById('financialPageDisplay');
  if (disp) {
    disp.innerText = `Sayfa ${data.page} / ${data.total_pages || 1}`;
  }

  const btnPrev = document.getElementById('btnFinancialPrev');
  if (btnPrev) btnPrev.disabled = data.page <= 1;

  const btnNext = document.getElementById('btnFinancialNext');
  if (btnNext) btnNext.disabled = data.page >= data.total_pages;
}

function populateFinancialDeptDropdown(departments) {
  const select = document.getElementById('financialDeptFilter');
  if (!select || select.dataset.loaded === 'true' || !departments) return;

  let current = select.value;
  let html = '<option value="all">Tüm Bölümler</option>';
  departments.forEach(dept => {
    html += `<option value="${dept}">${dept}</option>`;
  });
  select.innerHTML = html;
  select.value = current;
  select.dataset.loaded = 'true';
}

// ----------------------------------------------------------------------------
// 6. YASAL UYUM & RİSK RADARI (LEGAL COMPLIANCE RADAR - 4857 SAYILI İŞ KANUNU)
// ----------------------------------------------------------------------------
async function loadComplianceRadar() {
  const tbody = document.getElementById('complianceTableBody');
  if (tbody) {
    tbody.innerHTML = `<tr><td colspan="11" style="text-align:center; padding: 40px; color: var(--text-dim);">Yasal uyum ve risk verileri analiz ediliyor...</td></tr>`;
  }

  try {
    const params = new URLSearchParams({
      page: AppState.compliance.page,
      page_size: AppState.compliance.pageSize,
      violation_type: AppState.compliance.violationType,
      severity: AppState.compliance.severity,
      department: AppState.compliance.department,
      search: AppState.compliance.search
    });

    const res = await fetch(`/api/compliance-radar?${params.toString()}`);
    if (!res.ok) throw new Error('Yasal uyum verileri alınamadı');
    const data = await res.json();

    AppState.compliance.items = data.items;
    AppState.compliance.totalCount = data.total_count;
    AppState.compliance.totalPages = data.total_pages;

    // 1. KPI Kartları
    const s = data.summary;
    const valComp = document.getElementById('valComplianceScore');
    if (valComp) valComp.innerText = `%${s.compliance_index.toFixed(1).replace('.', ',')}`;

    const subRisk = document.getElementById('subComplianceRisk');
    if (subRisk) subRisk.innerText = `${s.risk_status} (${s.total_violations} Toplam İhlal)`;

    const valRest = document.getElementById('valRestViolations');
    if (valRest) valRest.innerHTML = `${fmtInt(s.rest_violations_count)} <small>olay</small>`;

    const valConsec = document.getElementById('valConsecutiveViolations');
    if (valConsec) valConsec.innerHTML = `${fmtInt(s.consecutive_work_count)} <small>dönem</small>`;

    const valOtRisk = document.getElementById('valOvertimeRisks');
    if (valOtRisk) valOtRisk.innerHTML = `${fmtInt(s.overtime_limit_count)} <small>kişi</small>`;

    // Segmented Filtre Rozetleri
    const badgeAll = document.getElementById('badgeCompCountAll');
    if (badgeAll) badgeAll.innerText = fmtInt(s.total_violations);

    const badgeRest = document.getElementById('badgeCompCountRest');
    if (badgeRest) badgeRest.innerText = fmtInt(s.rest_violations_count);

    const badgeConsec = document.getElementById('badgeCompCountConsecutive');
    if (badgeConsec) badgeConsec.innerText = fmtInt(s.consecutive_work_count);

    const badgeOt = document.getElementById('badgeCompCountOt');
    if (badgeOt) badgeOt.innerText = fmtInt(s.overtime_limit_count);

    // Navbar Rozeti
    const badgeNav = document.getElementById('badgeComplianceScore');
    if (badgeNav) badgeNav.innerText = `%${Math.round(s.compliance_index)}`;

    // 2. Departman Dağılım Kartları
    renderDeptComplianceList(data.department_compliance);

    // 3. İdari Risk ve Para Cezası Özeti
    renderComplianceRiskSummary(s, data.department_compliance);

    // 4. İhlal Denetim Tablosu
    renderComplianceTable(data.items, (data.page - 1) * data.page_size);

    // 5. Sayfalama
    renderCompliancePagination(data);

    // 6. Departman Filtre Dropdown
    populateComplianceDeptDropdown(data.departments);

  } catch (err) {
    console.error(err);
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="11" style="text-align:center; color: var(--color-danger); padding: 30px;">Hata: ${err.message}</td></tr>`;
    }
  }
}

function renderDeptComplianceList(depts) {
  const container = document.getElementById('deptComplianceListContainer');
  if (!container || !depts) return;

  let html = '';
  depts.slice(0, 8).forEach(d => {
    let riskTag = '<span class="dept-badge-tag tag-normal">Düşük Risk</span>';
    if (d.risk_level === 'YÜKSEK') {
      riskTag = '<span class="dept-badge-tag tag-critical">🚨 Yüksek Risk</span>';
    } else if (d.risk_level === 'ORTA') {
      riskTag = '<span class="dept-badge-tag tag-high">⚠️ Orta Risk</span>';
    }

    html += `
      <div class="dept-compliance-card">
        <div class="dept-comp-header">
          <div class="dept-comp-name">
            <strong>${d.department}</strong>
            <small style="color: var(--text-dim); margin-left: 6px;">(${d.active_headcount} çalışan)</small>
          </div>
          <div class="dept-comp-badge-group">
            <span class="dept-comp-score" style="font-weight: 700; color: ${d.compliance_score < 70 ? 'var(--color-danger)' : d.compliance_score < 85 ? 'var(--color-amber)' : 'var(--color-emerald)'};">Uyum: %${d.compliance_score.toFixed(1).replace('.', ',')}</span>
            ${riskTag}
          </div>
        </div>
        <div class="dept-comp-stats">
          <span class="stat-pill"><small>11s Dinlenme:</small> <strong>${d.rest_count}</strong></span>
          <span class="stat-pill"><small>Hafta Tatili:</small> <strong>${d.consecutive_count}</strong></span>
          <span class="stat-pill"><small>270s FM Riski:</small> <strong>${d.ot_risk_count}</strong></span>
          <span class="stat-pill" style="border-color: rgba(239,68,68,0.3);"><small style="color:var(--color-danger)">Kritik:</small> <strong style="color:var(--color-danger)">${d.critical_count}</strong></span>
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

function renderComplianceRiskSummary(s, depts) {
  const container = document.getElementById('complianceRiskSummaryContainer');
  if (!container || !s) return;

  const topRiskDept = s.highest_risk_department || 'Balık Temizleme';

  const html = `
    <div class="risk-article-item">
      <div class="risk-article-header">
        <div class="risk-article-title">
          <span class="risk-law-ref">4857 SK Md. 68 & Postalar Yön.</span>
          <strong>11 Saat Kesintisiz Dinlenme Kuralı</strong>
        </div>
        <span class="risk-fine-badge badge-critical">${s.rest_violations_count} Vaka Tespit</span>
      </div>
      <p class="risk-article-desc">
        Vardiya değişiminde iki çalışma arası dinlenmenin 11 saatten az olması doğrudan işverenin ağır kusuru sayılır. Olası bir iş kazasında SGK rücu davaları açılır ve kişi başı idari para cezası kesilir.
      </p>
    </div>

    <div class="risk-article-item">
      <div class="risk-article-header">
        <div class="risk-article-title">
          <span class="risk-law-ref">4857 SK Md. 46</span>
          <strong>7+ Gün Aralıksız Çalışma (Hafta Tatili Gaspı)</strong>
        </div>
        <span class="risk-fine-badge badge-warning">${s.consecutive_work_count} Dönem Tespit</span>
      </div>
      <p class="risk-article-desc">
        6 iş günü çalışan personele 7 günlük zaman diliminde kesintisiz 24 saat hafta tatili verilmesi amir hükmüdür. 7+ gün tatilsiz çalışma işçiye kıdem tazminatlı haklı fesih hakkı doğurur.
      </p>
    </div>

    <div class="risk-article-item">
      <div class="risk-article-header">
        <div class="risk-article-title">
          <span class="risk-law-ref">4857 SK Md. 41</span>
          <strong>Yıllık 270 Saat Fazla Mesai Tavanı</strong>
        </div>
        <span class="risk-fine-badge badge-critical">${s.overtime_limit_count} Riskli Personel</span>
      </div>
      <p class="risk-article-desc">
        İşçinin yazılı onayı olsa dahi yılda 270 saat fazla mesai tavanı aşılamaz. Eylül ayında tek başına 35-50+ saat mesaiye ulaşan çalışanlar için acil vardiya rotasyonu şarttır.
      </p>
    </div>

    <div class="risk-summary-footer" style="padding: 12px 14px; background: rgba(239, 68, 68, 0.08); border-radius: 8px; border: 1px solid rgba(239, 68, 68, 0.2); margin-top: 10px;">
      <div style="font-size: 12px; font-weight: 700; color: var(--color-danger); display: flex; align-items: center; gap: 6px;">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:16px;height:16px;">
          <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
        </svg>
        <span>En Yüksek Teftiş Riski: <strong>${topRiskDept}</strong></span>
      </div>
      <div style="font-size: 11px; color: var(--text-muted); margin-top: 4px;">
        Toplam ${s.total_violations} yasal uygunsuzluktan ${s.critical_violations} tanesi KRİTİK seviyededir. Teftişe karşı yukarıdaki Excel Teftiş Raporunu indirip inceleyebilirsiniz.
      </div>
    </div>
  `;

  container.innerHTML = html;
}

function renderComplianceTable(items, startIdx) {
  const tbody = document.getElementById('complianceTableBody');
  if (!tbody) return;

  if (items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="11" style="text-align:center; padding: 40px; color: var(--text-dim);">Arama ve filtre kriterlerine uygun ihlal kaydı bulunamadı.</td></tr>`;
    return;
  }

  let html = '';
  items.forEach((v, idx) => {
    let typeBadge = '';
    if (v.violation_type === 'REST_11H') {
      typeBadge = '<span class="badge badge-warning">11s Dinlenme</span>';
    } else if (v.violation_type === 'CONSECUTIVE_7D') {
      typeBadge = '<span class="badge badge-purple">Hafta Tatili</span>';
    } else if (v.violation_type === 'OVERTIME_270H') {
      typeBadge = '<span class="badge badge-cyan">270s FM Riski</span>';
    }

    const sevBadge = v.severity === 'CRITICAL'
      ? '<span class="badge badge-critical" style="background:rgba(239,68,68,0.15); color:#f87171; border:1px solid rgba(239,68,68,0.3); font-weight:700;">🚨 KRİTİK</span>'
      : '<span class="badge badge-warning" style="background:rgba(245,158,11,0.15); color:#fbbf24; border:1px solid rgba(245,158,11,0.3); font-weight:600;">⚠️ UYARI</span>';

    html += `
      <tr>
        <td>${startIdx + idx + 1}</td>
        <td class="cell-mono">${v.tc || '-'}</td>
        <td class="cell-bold">${v.ad_soyad}</td>
        <td>${v.bolum}</td>
        <td>${typeBadge}</td>
        <td style="text-align: center;" class="cell-mono">${v.date_str}</td>
        <td style="text-align: center; font-weight: 700;" class="cell-mono">${v.metric_value}</td>
        <td style="text-align: center; color: var(--text-muted); font-size: 11px;">${v.legal_limit}</td>
        <td style="text-align: center;">${sevBadge}</td>
        <td class="cell-mono" style="font-size: 11px; color: var(--color-cyan);">${v.legal_article}</td>
        <td style="font-size: 11.5px; color: var(--text-muted);">${v.detail}</td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

function renderCompliancePagination(data) {
  const info = document.getElementById('compliancePaginationInfo');
  if (info) {
    const start = data.total_count === 0 ? 0 : (data.page - 1) * data.page_size + 1;
    const end = Math.min(data.page * data.page_size, data.total_count);
    info.innerText = `Gösterilen: ${start} - ${end} / ${data.total_count} İhlal Kaydı`;
  }

  const disp = document.getElementById('compliancePageDisplay');
  if (disp) {
    disp.innerText = `Sayfa ${data.page} / ${data.total_pages || 1}`;
  }

  const btnPrev = document.getElementById('btnCompliancePrev');
  if (btnPrev) btnPrev.disabled = data.page <= 1;

  const btnNext = document.getElementById('btnComplianceNext');
  if (btnNext) btnNext.disabled = data.page >= data.total_pages;
}

function populateComplianceDeptDropdown(departments) {
  const select = document.getElementById('complianceDeptFilter');
  if (!select || select.dataset.loaded === 'true' || !departments) return;

  let current = select.value;
  let html = '<option value="all">Tüm Bölümler</option>';
  departments.forEach(dept => {
    html += `<option value="${dept}">${dept}</option>`;
  });
  select.innerHTML = html;
  select.value = current;
  select.dataset.loaded = 'true';
}

// ----------------------------------------------------------------------------
// 7. RAPOR MERKEZİ & AYARLAR (RULES & REPORTS)
// ----------------------------------------------------------------------------
async function loadRules() {
  try {
    const res = await fetch('/api/rules');
    if (!res.ok) throw new Error('Kurallar alınamadı');
    const rules = await res.json();
    AppState.rules = rules;

    // Formu doldur
    const dayShift = rules.day_shift || {};
    const bonuses = rules.department_bonuses || {};
    const otPolicy = rules.overtime_policy || {};

    if (document.getElementById('ruleDayStart')) document.getElementById('ruleDayStart').value = dayShift.start_time || '08:00';
    if (document.getElementById('ruleDayTolerance')) document.getElementById('ruleDayTolerance').value = dayShift.tolerance_time || '08:20';
    if (document.getElementById('ruleDayEnd')) document.getElementById('ruleDayEnd').value = dayShift.end_time || '17:00';
    if (document.getElementById('ruleDayHours')) document.getElementById('ruleDayHours').value = dayShift.standard_hours || 7.5;
    if (document.getElementById('ruleLunchBreak')) document.getElementById('ruleLunchBreak').value = dayShift.lunch_break_hours || 1.5;
    if (document.getElementById('ruleOtGrace')) document.getElementById('ruleOtGrace').value = otPolicy.grace_period_minutes || 25;

    // Balık Dolum/Kesim
    const fish = bonuses.BALIK_DOLUM_KESIM || {};
    if (document.getElementById('ruleFishBonus')) document.getElementById('ruleFishBonus').value = fish.bonus_hours || 2.0;
    if (document.getElementById('ruleFishThreshold')) document.getElementById('ruleFishThreshold').value = fish.min_hours_threshold !== undefined ? fish.min_hours_threshold : 0.0;

    // Üretim
    const prod = bonuses.URETIM || {};
    if (document.getElementById('ruleProdBonus')) document.getElementById('ruleProdBonus').value = prod.bonus_hours || 4.0;
    if (document.getElementById('ruleProdThreshold')) document.getElementById('ruleProdThreshold').value = prod.min_hours_threshold || 12.0;

  } catch (err) {
    console.error(err);
    showToast('Kurallar yüklenirken hata: ' + err.message, 'error');
  }
}

async function saveRulesFromForm(e) {
  e.preventDefault();
  const btn = document.getElementById('btnSaveRules');
  if (btn) {
    btn.disabled = true;
    btn.innerText = 'Kaydediliyor ve Yeniden Hesaplanıyor...';
  }

  try {
    const updated = JSON.parse(JSON.stringify(AppState.rules || {}));
    if (!updated.day_shift) updated.day_shift = {};
    if (!updated.department_bonuses) updated.department_bonuses = {};
    if (!updated.overtime_policy) updated.overtime_policy = {};

    updated.day_shift.start_time = document.getElementById('ruleDayStart').value;
    updated.day_shift.tolerance_time = document.getElementById('ruleDayTolerance').value;
    updated.day_shift.end_time = document.getElementById('ruleDayEnd').value;
    updated.day_shift.standard_hours = parseFloat(document.getElementById('ruleDayHours').value) || 7.5;
    updated.day_shift.lunch_break_hours = parseFloat(document.getElementById('ruleLunchBreak').value) || 1.5;
    updated.overtime_policy.grace_period_minutes = parseInt(document.getElementById('ruleOtGrace').value) || 25;

    // Balık
    if (!updated.department_bonuses.BALIK_DOLUM_KESIM) updated.department_bonuses.BALIK_DOLUM_KESIM = {};
    updated.department_bonuses.BALIK_DOLUM_KESIM.bonus_hours = parseFloat(document.getElementById('ruleFishBonus').value) || 2.0;
    updated.department_bonuses.BALIK_DOLUM_KESIM.min_hours_threshold = parseFloat(document.getElementById('ruleFishThreshold').value) || 0.0;

    // Üretim
    if (!updated.department_bonuses.URETIM) updated.department_bonuses.URETIM = {};
    updated.department_bonuses.URETIM.bonus_hours = parseFloat(document.getElementById('ruleProdBonus').value) || 4.0;
    updated.department_bonuses.URETIM.min_hours_threshold = parseFloat(document.getElementById('ruleProdThreshold').value) || 12.0;

    const res = await fetch('/api/rules', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rules: updated })
    });

    if (!res.ok) throw new Error('Kurallar kaydedilemedi');
    const result = await res.json();

    showToast(result.message || 'Kurallar başarıyla güncellendi!', 'success');
    loadStats();

  } catch (err) {
    console.error(err);
    showToast('Hata: ' + err.message, 'error');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerText = 'Kuralları Kaydet ve Sistemi Yeniden Hesapla';
    }
  }
}

async function triggerGenerateAllReports() {
  const btn = document.getElementById('btnGenerateAllReports');
  const headerBtn = document.getElementById('btnHeaderGenerate');
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span>Raporlar Üretiliyor (Lütfen Bekleyin)...</span>';
  }
  if (headerBtn) headerBtn.disabled = true;

  showToast('Resmi çalışma kitapları ve formüller hesaplanıyor...', 'info', 5000);

  try {
    const res = await fetch('/api/generate-reports', { method: 'POST' });
    if (!res.ok) {
      const errData = await res.json();
      throw new Error(errData.detail || 'Rapor üretimi başarısız');
    }
    const data = await res.json();
    showToast(data.message || 'Tüm raporlar başarıyla üretildi!', 'success', 5000);
  } catch (err) {
    console.error(err);
    showToast('Rapor üretme hatası: ' + err.message, 'error', 6000);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = `
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:18px;height:18px;">
          <polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/>
          <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
        </svg>
        <span>Tüm Raporları Sıfırdan Hesapla & Üret</span>
      `;
    }
    if (headerBtn) headerBtn.disabled = false;
  }
}

// ----------------------------------------------------------------------------
// 6. PERSONEL AYLIK HAREKET ÇEKMECESİ (DRILLDOWN MODAL)
// ----------------------------------------------------------------------------
async function openPersonnelPunchesModal(nameKey) {
  const title = document.getElementById('modalPersonName');
  const meta = document.getElementById('modalPersonMeta');
  const tbody = document.getElementById('modalPunchesTableBody');

  if (tbody) tbody.innerHTML = `<tr><td colspan="12" style="text-align:center; padding: 30px;">Kart hareketleri yükleniyor...</td></tr>`;
  openModal('modalPersonnelPunches');

  try {
    const res = await fetch(`/api/personnel/${encodeURIComponent(nameKey)}/punches`);
    if (!res.ok) throw new Error('Personel hareketleri bulunamadı');
    const data = await res.json();

    const p = data.personnel;
    if (title) title.innerText = p.ad_soyad;
    if (meta) meta.innerText = `TC: ${p.tc || '-'} | Sicil: ${p.sicil || '-'} | Bölüm: ${p.bolum || '-'}`;

    let html = '';
    data.days.forEach(d => {
      const isSun = d.status_type === 'TATİL';
      const hasWork = d.final_hours > 0;
      html += `
        <tr style="${isSun ? 'background: rgba(234,179,8,0.05);' : ''}">
          <td>${d.gun}</td>
          <td class="cell-mono">${d.tarih}</td>
          <td class="cell-mono">${d.all_punches || '-'}</td>
          <td class="cell-mono">${d.raw_g_saat || '-'}</td>
          <td class="cell-mono">${d.raw_c_saat || '-'}</td>
          <td class="cell-mono">${d.break_hours ? fmtHours(d.break_hours) : '-'}</td>
          <td class="cell-mono">${hasWork ? fmtHours(d.fiili_sure) : '-'}</td>
          <td class="cell-mono">${hasWork ? fmtHours(d.base_hours) : '-'}</td>
          <td class="cell-mono cell-green">${d.overtime_hours > 0 ? fmtHours(d.overtime_hours) : '-'}</td>
          <td class="cell-mono cell-purple">${d.bonus_hours > 0 ? '+' + fmtHours(d.bonus_hours) : '-'}</td>
          <td class="cell-mono cell-accent cell-bold">${hasWork ? fmtHours(d.final_hours) : '-'}</td>
          <td style="font-size: 11px; color: var(--text-muted);">${d.audit_note || d.status_type || ''}</td>
        </tr>
      `;
    });

    if (tbody) tbody.innerHTML = html;

  } catch (err) {
    console.error(err);
    if (tbody) tbody.innerHTML = `<tr><td colspan="12" style="text-align:center; color: var(--color-danger); padding: 20px;">Hata: ${err.message}</td></tr>`;
  }
}

// ----------------------------------------------------------------------------
// 7. DOSYA YÜKLEME (FILE UPLOAD)
// ----------------------------------------------------------------------------
async function uploadFile(fileType, file) {
  if (!file) return;

  showToast(`${fileType}.xls yükleniyor ve sistem yenileniyor...`, 'info', 4000);

  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch(`/api/upload?file_type=${fileType}`, {
      method: 'POST',
      body: formData
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.detail || 'Yükleme başarısız');
    }

    const data = await res.json();
    showToast(data.message || 'Dosya başarıyla yüklendi!', 'success', 5000);
    loadStats();
    loadMatrix();
  } catch (err) {
    console.error(err);
    showToast('Yükleme hatası: ' + err.message, 'error', 6000);
  }
}

// ----------------------------------------------------------------------------
// EVENT LISTENERS VE BAŞLANGIÇ
// ----------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {

  // 1. Sekme Tıklamaları
  document.querySelectorAll('.nav-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      const tabId = btn.getAttribute('data-tab');
      if (tabId) switchTab(tabId);
    });
  });

  // 2. Matris Filtre Olayları
  const searchInput = document.getElementById('matrixSearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', debounce((e) => {
      AppState.matrix.search = e.target.value;
      AppState.matrix.page = 1;
      loadMatrix();
    }, 300));
  }

  const deptFilter = document.getElementById('matrixDeptFilter');
  if (deptFilter) {
    deptFilter.addEventListener('change', (e) => {
      AppState.matrix.department = e.target.value;
      AppState.matrix.page = 1;
      loadMatrix();
    });
  }

  const statusFilter = document.getElementById('matrixStatusFilter');
  if (statusFilter) {
    statusFilter.addEventListener('change', (e) => {
      AppState.matrix.status = e.target.value;
      AppState.matrix.page = 1;
      loadMatrix();
    });
  }

  const pageSizeSelect = document.getElementById('matrixPageSize');
  if (pageSizeSelect) {
    pageSizeSelect.addEventListener('change', (e) => {
      AppState.matrix.pageSize = parseInt(e.target.value);
      AppState.matrix.page = 1;
      loadMatrix();
    });
  }

  const btnPrev = document.getElementById('btnMatrixPrev');
  if (btnPrev) {
    btnPrev.addEventListener('click', () => {
      if (AppState.matrix.page > 1) {
        AppState.matrix.page--;
        loadMatrix();
      }
    });
  }

  const btnNext = document.getElementById('btnMatrixNext');
  if (btnNext) {
    btnNext.addEventListener('click', () => {
      if (AppState.matrix.page < AppState.matrix.totalPages) {
        AppState.matrix.page++;
        loadMatrix();
      }
    });
  }

  // 3. İstisna Segmentli Filtre Butonları
  document.querySelectorAll('.segmented-control .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.segmented-control .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      AppState.exceptions.status = btn.getAttribute('data-status');
      loadExceptions();
    });
  });

  const excSearchInput = document.getElementById('exceptionsSearchInput');
  if (excSearchInput) {
    excSearchInput.addEventListener('input', debounce((e) => {
      AppState.exceptions.search = e.target.value;
      loadExceptions();
    }, 300));
  }

  // 4. İstisna Çözüm Formu
  const formResolve = document.getElementById('formResolveException');
  if (formResolve) {
    formResolve.addEventListener('submit', async (e) => {
      e.preventDefault();
      const nameKey = document.getElementById('modalNameKey').value;
      const day = parseInt(document.getElementById('modalDay').value);
      const approvedHours = parseFloat(document.getElementById('modalApprovedHours').value);
      const note = document.getElementById('modalNote').value;
      const resolver = document.getElementById('modalResolver').value;

      try {
        const res = await fetch('/api/exceptions/resolve', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name_key: nameKey,
            day: day,
            approved_hours: approvedHours,
            note: note,
            resolved_by: resolver
          })
        });

        if (!res.ok) throw new Error('Onay işlemi kaydedilemedi');
        const data = await res.json();
        showToast(data.message || 'İstisna onaylandı!', 'success');
        closeModal('modalResolveException');
        loadExceptions();
        loadStats();
      } catch (err) {
        showToast('Hata: ' + err.message, 'error');
      }
    });
  }

  // 5. Prim Masası Filtreleri
  const bonusSearch = document.getElementById('bonusesSearchInput');
  if (bonusSearch) {
    bonusSearch.addEventListener('input', debounce((e) => {
      AppState.bonuses.search = e.target.value;
      loadBonuses();
    }, 300));
  }

  const bonusDept = document.getElementById('bonusesDeptFilter');
  if (bonusDept) {
    bonusDept.addEventListener('change', (e) => {
      AppState.bonuses.department = e.target.value;
      loadBonuses();
    });
  }

  // 6. Finans & Bütçe Radarı Filtreleri ve Sayfalama
  const finSearch = document.getElementById('financialSearchInput');
  if (finSearch) {
    finSearch.addEventListener('input', debounce((e) => {
      AppState.financial.search = e.target.value;
      AppState.financial.page = 1;
      loadFinancialRadar();
    }, 300));
  }

  const finDept = document.getElementById('financialDeptFilter');
  if (finDept) {
    finDept.addEventListener('change', (e) => {
      AppState.financial.department = e.target.value;
      AppState.financial.page = 1;
      loadFinancialRadar();
    });
  }

  const btnFinPrev = document.getElementById('btnFinancialPrev');
  if (btnFinPrev) {
    btnFinPrev.addEventListener('click', () => {
      if (AppState.financial.page > 1) {
        AppState.financial.page--;
        loadFinancialRadar();
      }
    });
  }

  const btnFinNext = document.getElementById('btnFinancialNext');
  if (btnFinNext) {
    btnFinNext.addEventListener('click', () => {
      if (AppState.financial.page < AppState.financial.totalPages) {
        AppState.financial.page++;
        loadFinancialRadar();
      }
    });
  }

  // 7. Yasal Uyum & Risk Radarı Filtreleri ve Sayfalama
  const compSegmented = document.getElementById('compTypeSegmented');
  if (compSegmented) {
    compSegmented.querySelectorAll('.seg-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        compSegmented.querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        AppState.compliance.violationType = btn.getAttribute('data-type') || 'all';
        AppState.compliance.page = 1;
        loadComplianceRadar();
      });
    });
  }

  const compSearch = document.getElementById('complianceSearchInput');
  if (compSearch) {
    compSearch.addEventListener('input', debounce((e) => {
      AppState.compliance.search = e.target.value;
      AppState.compliance.page = 1;
      loadComplianceRadar();
    }, 300));
  }

  const compDept = document.getElementById('complianceDeptFilter');
  if (compDept) {
    compDept.addEventListener('change', (e) => {
      AppState.compliance.department = e.target.value;
      AppState.compliance.page = 1;
      loadComplianceRadar();
    });
  }

  const compSeverity = document.getElementById('complianceSeverityFilter');
  if (compSeverity) {
    compSeverity.addEventListener('change', (e) => {
      AppState.compliance.severity = e.target.value;
      AppState.compliance.page = 1;
      loadComplianceRadar();
    });
  }

  const btnCompPrev = document.getElementById('btnCompliancePrev');
  if (btnCompPrev) {
    btnCompPrev.addEventListener('click', () => {
      if (AppState.compliance.page > 1) {
        AppState.compliance.page--;
        loadComplianceRadar();
      }
    });
  }

  const btnCompNext = document.getElementById('btnComplianceNext');
  if (btnCompNext) {
    btnCompNext.addEventListener('click', () => {
      if (AppState.compliance.page < AppState.compliance.totalPages) {
        AppState.compliance.page++;
        loadComplianceRadar();
      }
    });
  }

  // 7. Rapor Üret Butonları
  const btnGenAll = document.getElementById('btnGenerateAllReports');
  if (btnGenAll) btnGenAll.addEventListener('click', triggerGenerateAllReports);

  const btnHeaderGen = document.getElementById('btnHeaderGenerate');
  if (btnHeaderGen) btnHeaderGen.addEventListener('click', triggerGenerateAllReports);

  // 7. Kurallar Formu
  const rulesForm = document.getElementById('rulesForm');
  if (rulesForm) rulesForm.addEventListener('submit', saveRulesFromForm);

  // 8. Dosya Yükleme Olayları
  const inputPdks = document.getElementById('fileInputPdks');
  if (inputPdks) {
    inputPdks.addEventListener('change', (e) => {
      if (e.target.files.length > 0) uploadFile('pdks', e.target.files[0]);
    });
  }

  const inputPuantaj = document.getElementById('fileInputPuantaj');
  if (inputPuantaj) {
    inputPuantaj.addEventListener('change', (e) => {
      if (e.target.files.length > 0) uploadFile('puantaj', e.target.files[0]);
    });
  }

  // Sürükle-bırak olayları
  const setupDropzone = (zoneId, inputId, fileType) => {
    const zone = document.getElementById(zoneId);
    if (!zone) return;

    ['dragenter', 'dragover'].forEach(eventName => {
      zone.addEventListener(eventName, (e) => {
        e.preventDefault();
        zone.classList.add('dragover');
      }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
      zone.addEventListener(eventName, (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
      }, false);
    });

    zone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files.length > 0) {
        uploadFile(fileType, e.dataTransfer.files[0]);
      }
    });
  };

  setupDropzone('dropzonePdks', 'fileInputPdks', 'pdks');
  setupDropzone('dropzonePuantaj', 'fileInputPuantaj', 'puantaj');

  // Başlangıç Yüklemesi
  loadStats();
});
