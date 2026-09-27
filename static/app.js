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

  // Aktif Oturum ve Rol (RBAC)
  currentUser: {
    username: 'admin',
    name: 'Sistem Yöneticisi',
    title: 'Sistem Yöneticisi (Admin)',
    role: 'admin',
    avatar: '🛡️',
    allowed_tabs: ['tab-overview', 'tab-matrix', 'tab-exceptions', 'tab-bonuses', 'tab-finance', 'tab-compliance', 'tab-reports', 'tab-audit'],
    can_resolve_exceptions: true,
    can_view_finance: true,
    can_view_compliance: true,
    can_edit_rules: true,
    can_export_reports: true
  },
  
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
    selectedKeys: new Set(),
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
  },

  // Çoklu Ay & Arşiv Durumu (SQLite Persistence)
  archive: {
    activePeriodKey: '2026-09',
    periods: [],
    cumulativeOvertime: {
      year: 2026,
      items: [],
      filteredItems: [],
      page: 1,
      pageSize: 50,
      totalPages: 1,
      search: '',
      riskFilter: 'all'
    },
    icra: {
      items: [],
      kpis: {}
    }
  },

  // Denetim İzi Masası Durumu (Audit Trail)
  audit: {
    page: 1,
    pageSize: 50,
    search: '',
    actionType: 'ALL',
    username: 'ALL',
    totalPages: 1,
    totalCount: 0,
    items: [],
    summary: {}
  }
};

// ----------------------------------------------------------------------------
// GLOBAL FETCH INTERCEPTOR (RBAC Rol Yetkisi Gönderici)
// ----------------------------------------------------------------------------
const _nativeFetch = window.fetch;
window.fetch = function(url, options = {}) {
  options.headers = options.headers || {};
  const activeUsername = (AppState.currentUser && AppState.currentUser.username) || localStorage.getItem('fide_user_role') || 'admin';
  if (options.headers instanceof Headers) {
    options.headers.set('X-User-Role', activeUsername);
  } else {
    options.headers['X-User-Role'] = activeUsername;
  }
  return _nativeFetch.call(this, url, options);
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
  if (tabId === 'tab-financial') {
    loadFinancialRadar();
    loadIcraRecords();
  }
  if (tabId === 'tab-compliance') {
    loadComplianceRadar();
    loadCumulativeOvertime();
  }
  if (tabId === 'tab-reports') loadRules();
  if (tabId === 'tab-audit') loadAuditLogs();
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
    if (periodText) periodText.innerText = `${data.active_period_name || data.period} Dönemi`;

    const periodStatusBadge = document.getElementById('periodStatusBadge');
    if (periodStatusBadge) {
      periodStatusBadge.innerText = 'Aktif';
      periodStatusBadge.className = 'period-status-pill';
    }

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
    tbody.innerHTML = `<tr><td colspan="11" style="text-align:center; padding: 30px; color: var(--text-dim);">Filtreye uygun istisna kaydı bulunmuyor.</td></tr>`;
    updateFloatingBulkBar();
    return;
  }

  let html = '';
  items.forEach((rec, idx) => {
    const isResolved = rec.is_resolved;
    const itemKey = `${rec.name_key || rec.ad_soyad}:${rec.gun}`;
    const isSelected = AppState.exceptions.selectedKeys.has(itemKey);
    const resTag = isResolved
      ? `<span class="status-tag tag-resolved">✓ ÇÖZÜLDÜ</span>`
      : `<span class="status-tag tag-pending">ONAY BEKLİYOR</span>`;

    const sug = rec.smart_suggestion || {
      suggested_hours: 7.5,
      suggested_g: rec.raw_g_saat || '08:00',
      suggested_c: '17:00',
      confidence: 75,
      source: 'FACTORY_STANDARD',
      badge_class: 'badge-std',
      reason: 'Standart vardiya kuralı (7,5s tam gün)'
    };

    const sourceLabel = sug.source === 'DEPARTMENT_PEERS' 
      ? '👥 Bölüm Mesaisi' 
      : (sug.source === 'PERSONAL_HABIT' ? '🕒 Vardiya Alışkanlığı' : '🏢 Standart Kural');

    html += `
      <tr class="${isSelected ? 'row-selected' : ''}">
        <td style="text-align: center;">
          <input type="checkbox" class="custom-checkbox exc-row-checkbox" data-key="${itemKey}" ${isSelected ? 'checked' : ''} onchange="toggleExceptionSelection('${itemKey}')">
        </td>
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
          <div class="smart-suggestion-card">
            <div class="smart-sug-header">
              <span class="badge ${sug.badge_class || 'badge-peer'}">${sourceLabel}</span>
              <span class="sug-conf-tag">%${sug.confidence} Güven</span>
            </div>
            <div class="smart-sug-body">
              <span class="sug-hours">${fmtHours(sug.suggested_hours)}s</span>
              <span class="sug-detail">(${sug.suggested_g} - ${sug.suggested_c})</span>
            </div>
            <div class="smart-sug-reason" title="${sug.reason}">${sug.reason}</div>
            ${!isResolved ? `
              <button type="button" class="btn-quick-accept" onclick="quickAcceptSuggestion('${itemKey}', ${sug.suggested_hours}, '${encodeURIComponent(sug.reason)}')">
                ⚡ Öneriyi Kabul Et (${fmtHours(sug.suggested_hours)}s)
              </button>
            ` : ''}
          </div>
        </td>
        <td>
          ${resTag}
          ${isResolved && rec.resolution ? `<div style="font-size: 10px; color: var(--color-success);">${rec.resolution.resolved_by || 'Amir'}: ${rec.resolution.approved_hours}s</div>` : ''}
        </td>
        <td style="text-align: center;">
          <button class="btn btn-sm ${isResolved ? 'btn-outline' : 'btn-warning'}" onclick="openResolveModalByKey('${itemKey}')">
            ${isResolved ? 'Düzenle' : 'Onayla'}
          </button>
        </td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
  updateFloatingBulkBar();
}

function toggleExceptionSelection(key) {
  if (AppState.exceptions.selectedKeys.has(key)) {
    AppState.exceptions.selectedKeys.delete(key);
  } else {
    AppState.exceptions.selectedKeys.add(key);
  }
  updateFloatingBulkBar();
}

function toggleSelectAllExceptions(checked) {
  const checkboxes = document.querySelectorAll('.exc-row-checkbox');
  checkboxes.forEach(cb => {
    cb.checked = checked;
    const key = cb.getAttribute('data-key');
    if (key) {
      if (checked) {
        AppState.exceptions.selectedKeys.add(key);
      } else {
        AppState.exceptions.selectedKeys.delete(key);
      }
    }
  });
  updateFloatingBulkBar();
}

function updateFloatingBulkBar() {
  const bar = document.getElementById('floatingBulkBar');
  const countEl = document.getElementById('bulkSelectedCount');
  const checkAll = document.getElementById('checkAllExceptions');
  const count = AppState.exceptions.selectedKeys.size;

  if (countEl) countEl.innerText = count;

  if (bar) {
    bar.style.display = count > 0 ? 'block' : 'none';
  }

  if (checkAll) {
    const allRows = document.querySelectorAll('.exc-row-checkbox');
    if (allRows.length === 0) {
      checkAll.checked = false;
      checkAll.indeterminate = false;
    } else if (count === 0) {
      checkAll.checked = false;
      checkAll.indeterminate = false;
    } else if (count >= allRows.length) {
      checkAll.checked = true;
      checkAll.indeterminate = false;
    } else {
      checkAll.checked = false;
      checkAll.indeterminate = true;
    }
  }
}

function clearBulkSelection() {
  AppState.exceptions.selectedKeys.clear();
  document.querySelectorAll('.exc-row-checkbox').forEach(cb => {
    cb.checked = false;
  });
  const checkAll = document.getElementById('checkAllExceptions');
  if (checkAll) {
    checkAll.checked = false;
    checkAll.indeterminate = false;
  }
  updateFloatingBulkBar();
}

async function quickAcceptSuggestion(key, hours, reasonEncoded) {
  const reason = decodeURIComponent(reasonEncoded || '');
  const parts = key.split(':');
  if (parts.length < 2) return;
  const nameKey = parts[0];
  const day = parseInt(parts[1]);

  try {
    const res = await fetch('/api/exceptions/resolve', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name_key: nameKey,
        day: day,
        approved_hours: hours,
        note: `⚡ Akıllı Öneri Kabul Edildi: ${reason}`,
        resolved_by: 'Vardiya Amiri / Akıllı Onay'
      })
    });
    if (!res.ok) throw new Error('Öneri onaylanamadı');
    const data = await res.json();
    showToast(`⚡ ${hours}s çalışma süresi onaylandı ve puantaja işlendi.`, 'success');
    loadExceptions();
    loadStats();
  } catch (err) {
    showToast('Hata: ' + err.message, 'error');
  }
}

async function bulkApplySmartSuggestions() {
  if (AppState.exceptions.selectedKeys.size === 0) return;
  const items = [];
  for (const key of AppState.exceptions.selectedKeys) {
    const rec = AppState.exceptions.items.find(r => (r.name_key || r.ad_soyad) + ':' + r.gun === key);
    if (rec && rec.smart_suggestion) {
      items.push({
        name_key: rec.name_key || rec.ad_soyad,
        day: rec.gun,
        approved_hours: rec.smart_suggestion.suggested_hours,
        note: `⚡ Akıllı Öneri: ${rec.smart_suggestion.reason}`
      });
    }
  }

  if (items.length === 0) {
    showToast('Seçili kayıtlar için akıllı öneri bulunamadı.', 'warning');
    return;
  }

  try {
    const res = await fetch('/api/exceptions/bulk-resolve', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        items: items,
        action_type: 'smart',
        global_note: '⚡ Akıllı Tahmin Motoru ile Toplu Amir Onayı',
        resolved_by: 'Vardiya Amiri / Akıllı Toplu Onay'
      })
    });
    if (!res.ok) throw new Error('Toplu akıllı onay başarısız oldu');
    const data = await res.json();
    showToast(`⚡ ${data.resolved_count} istisna kaydı kendi akıllı önerileriyle toplu onaylandı!`, 'success');
    clearBulkSelection();
    loadExceptions();
    loadStats();
  } catch (err) {
    showToast('Hata: ' + err.message, 'error');
  }
}

async function bulkApplyStandardHours() {
  if (AppState.exceptions.selectedKeys.size === 0) return;
  const items = [];
  for (const key of AppState.exceptions.selectedKeys) {
    const rec = AppState.exceptions.items.find(r => (r.name_key || r.ad_soyad) + ':' + r.gun === key);
    if (rec) {
      items.push({
        name_key: rec.name_key || rec.ad_soyad,
        day: rec.gun,
        approved_hours: 7.5,
        note: 'Standart tam gün (7,5s) toplu amir onayı'
      });
    }
  }

  try {
    const res = await fetch('/api/exceptions/bulk-resolve', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        items: items,
        action_type: 'standard_7_5',
        global_note: 'Standart tam gün (7,5s) toplu amir onayı',
        resolved_by: 'Vardiya Amiri / Standart Onay'
      })
    });
    if (!res.ok) throw new Error('Toplu standart onay başarısız oldu');
    const data = await res.json();
    showToast(`✅ ${data.resolved_count} istisna kaydına standart 7,5 saat uygulandı ve onaylandı!`, 'success');
    clearBulkSelection();
    loadExceptions();
    loadStats();
  } catch (err) {
    showToast('Hata: ' + err.message, 'error');
  }
}

function openBulkCustomHoursModal() {
  const count = AppState.exceptions.selectedKeys.size;
  if (count === 0) return;
  const countEl = document.getElementById('bulkModalCount');
  if (countEl) countEl.innerText = count;
  openModal('modalBulkCustomHours');
}

async function submitBulkCustomHours(e) {
  e.preventDefault();
  const count = AppState.exceptions.selectedKeys.size;
  if (count === 0) {
    closeModal('modalBulkCustomHours');
    return;
  }
  const hours = parseFloat(document.getElementById('bulkCustomHoursInput').value);
  if (isNaN(hours) || hours < 0 || hours > 24) {
    showToast('Bir günde 24 saatten fazla çalışma yazılamaz! Lütfen 0 ile 24 saat arasında bir değer girin.', 'error');
    return;
  }
  const resolver = document.getElementById('bulkCustomResolver').value || 'Vardiya Amiri / Toplu Onay';
  const note = document.getElementById('bulkCustomNote').value || `${hours}s toplu amir mesai onayı`;

  const items = [];
  for (const key of AppState.exceptions.selectedKeys) {
    const rec = AppState.exceptions.items.find(r => (r.name_key || r.ad_soyad) + ':' + r.gun === key);
    if (rec) {
      items.push({
        name_key: rec.name_key || rec.ad_soyad,
        day: rec.gun,
        approved_hours: hours,
        note: note,
        resolved_by: resolver
      });
    }
  }

  try {
    const res = await fetch('/api/exceptions/bulk-resolve', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        items: items,
        action_type: 'custom',
        global_note: note,
        resolved_by: resolver
      })
    });
    if (!res.ok) throw new Error('Toplu özel saat onayı başarısız oldu');
    const data = await res.json();
    showToast(`✏️ ${data.resolved_count} personele ${hours} saat toplu uygulandı ve onaylandı!`, 'success');
    closeModal('modalBulkCustomHours');
    clearBulkSelection();
    loadExceptions();
    loadStats();
  } catch (err) {
    showToast('Hata: ' + err.message, 'error');
  }
}

function openResolveModalByKey(itemKey) {
  const rec = AppState.exceptions.items.find(r => (r.name_key || r.ad_soyad) + ':' + r.gun === itemKey);
  if (rec) {
    openResolveModal(rec);
  }
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
    ${rec.smart_suggestion ? `
      <div style="margin-top: 8px; padding: 6px 10px; background: rgba(56, 189, 248, 0.1); border: 1px solid rgba(56, 189, 248, 0.25); border-radius: 6px; font-size: 11.5px; color: #bae6fd;">
        ⚡ <strong>Sistem Tahmini:</strong> ${fmtHours(rec.smart_suggestion.suggested_hours)} saat (%${rec.smart_suggestion.confidence} Güven) — ${rec.smart_suggestion.reason}
      </div>
    ` : ''}
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
    const missText = s.total_missing_hours > 0 ? ` - ${fmtHours(s.total_missing_hours)}s Eksik (1.0x)` : '';
    if (subOt) subOt.innerText = `${fmtHours(s.total_overtime_hours)}s FM (1.5x)${missText}`;

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

    const hasMissing = p.missing_hours > 0;
    const missingBadge = hasMissing
      ? `<span style="color: #f87171; font-weight: 700; font-family: var(--font-mono);" title="Eksik Kesintisi (1.0x): -${fmtCurrency(p.missing_deduction_cost || 0)}">-${fmtHours(p.missing_hours)}</span>`
      : `<span style="color: var(--text-dim);">-</span>`;

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
        <td style="text-align: right;" class="cell-mono">${missingBadge}</td>
        <td style="text-align: right;" class="cell-mono">${fmtHours(p.sunday_ot_hours)}</td>
        <td style="text-align: right; font-weight: 700; color: var(--color-amber);" class="cell-mono" title="H.İçi Net Mesai: ${fmtCurrency(p.weekday_ot_cost || 0)} | Pazar: ${fmtCurrency(p.sunday_ot_cost || 0)}">${fmtCurrency(p.overtime_cost)}</td>
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

    const btnPrint = document.getElementById('btnModalPrintSlip');
    if (btnPrint) {
      btnPrint.onclick = () => window.open(`/api/slips/${encodeURIComponent(nameKey)}`, '_blank');
    }

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

  const btnMatrixSlips = document.getElementById('btnMatrixBulkSlips');
  if (btnMatrixSlips) {
    btnMatrixSlips.addEventListener('click', () => {
      const dept = AppState.matrix.department || 'all';
      window.open(`/api/slips/bulk?department=${encodeURIComponent(dept)}`, '_blank');
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
      if (isNaN(approvedHours) || approvedHours < 0 || approvedHours > 24) {
        showToast('Bir günde 24 saatten fazla çalışma yazılamaz! Lütfen 0 ile 24 saat arasında bir değer girin.', 'error');
        return;
      }
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

        if (!res.ok) {
          const errData = await res.json().catch(() => ({}));
          throw new Error(errData.detail || 'Onay işlemi kaydedilemedi');
        }
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

  // 4a. Canlı 24 Saat Giriş Koruması (Elle 24'ten büyük yazılmasını engeller)
  const enforceMax24Input = (inputElId) => {
    const el = document.getElementById(inputElId);
    if (!el) return;
    el.addEventListener('input', (e) => {
      let val = parseFloat(e.target.value);
      if (val > 24) {
        e.target.value = 24;
        showToast('Bir günde 24 saatten fazla çalışma yazılamaz! Değer 24 saate sınırlandı.', 'warning');
      } else if (val < 0) {
        e.target.value = 0;
      }
    });
  };
  enforceMax24Input('modalApprovedHours');
  enforceMax24Input('bulkCustomHoursInput');

  // 4b. Toplu İşlem & Akıllı Amir Onay Masası Dinleyicileri
  const checkAllExceptions = document.getElementById('checkAllExceptions');
  if (checkAllExceptions) {
    checkAllExceptions.addEventListener('change', (e) => {
      toggleSelectAllExceptions(e.target.checked);
    });
  }

  const btnBulkApplySmart = document.getElementById('btnBulkApplySmart');
  if (btnBulkApplySmart) {
    btnBulkApplySmart.addEventListener('click', bulkApplySmartSuggestions);
  }

  const btnBulkApplyStandard = document.getElementById('btnBulkApplyStandard');
  if (btnBulkApplyStandard) {
    btnBulkApplyStandard.addEventListener('click', bulkApplyStandardHours);
  }

  const btnBulkCustomHours = document.getElementById('btnBulkCustomHours');
  if (btnBulkCustomHours) {
    btnBulkCustomHours.addEventListener('click', openBulkCustomHoursModal);
  }

  const btnBulkClearSelection = document.getElementById('btnBulkClearSelection');
  if (btnBulkClearSelection) {
    btnBulkClearSelection.addEventListener('click', clearBulkSelection);
  }

  const formBulkCustom = document.getElementById('formBulkCustomHours');
  if (formBulkCustom) {
    formBulkCustom.addEventListener('submit', submitBulkCustomHours);
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

  // RBAC ve Kullanıcı Profili Dinleyicileri
  const userProfileBtn = document.getElementById('userProfileBtn');
  if (userProfileBtn) {
    userProfileBtn.addEventListener('click', openRoleSwitcherModal);
  }

  const formLogin = document.getElementById('formCustomLogin');
  if (formLogin) {
    formLogin.addEventListener('submit', (e) => {
      e.preventDefault();
      const u = document.getElementById('loginUsername').value.trim();
      const p = document.getElementById('loginPassword').value;
      if (u) switchRole(u, p);
    });
  }

  // Çoklu Ay & Dönem Seçici Dinleyicisi
  const headerPeriodBadge = document.getElementById('headerPeriodBadge');
  if (headerPeriodBadge) {
    headerPeriodBadge.addEventListener('click', openPeriodSwitcher);
  }

  // İcra & Maaş Haczi Dinleyicileri
  const btnOpenNewIcra = document.getElementById('btnOpenNewIcraModal');
  if (btnOpenNewIcra) {
    btnOpenNewIcra.addEventListener('click', openNewIcraModal);
  }

  // 270 Saat Yıllık Fazla Mesai Kütüğü Dinleyicileri
  const otSearchInput = document.getElementById('overtimeSearchInput');
  if (otSearchInput) {
    otSearchInput.addEventListener('input', debounce((e) => {
      AppState.archive.cumulativeOvertime.search = e.target.value.trim();
      AppState.archive.cumulativeOvertime.page = 1;
      applyCumulativeOvertimeFilters();
    }, 300));
  }

  const otStatusFilter = document.getElementById('overtimeStatusFilter');
  if (otStatusFilter) {
    otStatusFilter.addEventListener('change', (e) => {
      AppState.archive.cumulativeOvertime.riskFilter = e.target.value;
      AppState.archive.cumulativeOvertime.page = 1;
      applyCumulativeOvertimeFilters();
    });
  }

  const btnOtPrev = document.getElementById('btnOvertimePrev');
  if (btnOtPrev) {
    btnOtPrev.addEventListener('click', () => {
      if (AppState.archive.cumulativeOvertime.page > 1) {
        AppState.archive.cumulativeOvertime.page--;
        renderCumulativeOvertimeTable();
      }
    });
  }

  const btnOtNext = document.getElementById('btnOvertimeNext');
  if (btnOtNext) {
    btnOtNext.addEventListener('click', () => {
      if (AppState.archive.cumulativeOvertime.page < AppState.archive.cumulativeOvertime.totalPages) {
        AppState.archive.cumulativeOvertime.page++;
        renderCumulativeOvertimeTable();
      }
    });
  }

  // 11. Denetim İzi & Audit Masası Dinleyicileri
  const auditSearch = document.getElementById('auditSearchInput');
  if (auditSearch) {
    auditSearch.addEventListener('input', debounce((e) => {
      AppState.audit.search = e.target.value.trim();
      AppState.audit.page = 1;
      loadAuditLogs();
    }, 300));
  }

  const auditActionFilter = document.getElementById('auditActionFilter');
  if (auditActionFilter) {
    auditActionFilter.addEventListener('change', (e) => {
      AppState.audit.actionType = e.target.value;
      AppState.audit.page = 1;
      loadAuditLogs();
    });
  }

  const auditUserFilter = document.getElementById('auditUserFilter');
  if (auditUserFilter) {
    auditUserFilter.addEventListener('change', (e) => {
      AppState.audit.username = e.target.value;
      AppState.audit.page = 1;
      loadAuditLogs();
    });
  }

  const btnAuditRefresh = document.getElementById('btnAuditRefresh');
  if (btnAuditRefresh) {
    btnAuditRefresh.addEventListener('click', () => {
      AppState.audit.page = 1;
      loadAuditLogs();
    });
  }

  const btnAuditPrev = document.getElementById('btnAuditPrevPage');
  if (btnAuditPrev) {
    btnAuditPrev.addEventListener('click', () => {
      if (AppState.audit.page > 1) {
        AppState.audit.page--;
        loadAuditLogs();
      }
    });
  }

  const btnAuditNext = document.getElementById('btnAuditNextPage');
  if (btnAuditNext) {
    btnAuditNext.addEventListener('click', () => {
      if (AppState.audit.page < AppState.audit.totalPages) {
        AppState.audit.page++;
        loadAuditLogs();
      }
    });
  }

  // Başlangıç Yüklemesi: Önce kullanıcı yetkilerini çek, sonra verileri yükle
  initCurrentUser().then(() => {
    loadStats();
  });
});

// ----------------------------------------------------------------------------
// 9. ROL BAZLI ERİŞİM VE KULLANICI YÖNETİMİ (RBAC İŞLEVLERİ)
// ----------------------------------------------------------------------------
async function initCurrentUser() {
  const savedRole = localStorage.getItem('fide_user_role') || 'admin';
  try {
    const res = await fetch('/api/auth/me', {
      headers: { 'X-User-Role': savedRole }
    });
    if (res.ok) {
      const user = await res.json();
      AppState.currentUser = user;
      applyRolePermissions(user);
    }
  } catch (err) {
    console.error('Kullanıcı bilgisi alınamadı:', err);
  }
}

function applyRolePermissions(user) {
  if (!user) return;

  // 1. Üst bar profil hapı güncelleme
  const avatarEl = document.getElementById('headerUserAvatar');
  const nameEl = document.getElementById('headerUserName');
  const roleEl = document.getElementById('headerUserRole');

  if (avatarEl) avatarEl.innerText = user.avatar || '👤';
  if (nameEl) nameEl.innerText = user.name || user.username;
  if (roleEl) roleEl.innerText = user.role.toUpperCase();

  // 2. Sekme butonlarını yetkiye göre göster / gizle
  const allowed = new Set(user.allowed_tabs || []);
  const allTabs = document.querySelectorAll('.nav-tabs .nav-tab');

  allTabs.forEach(btn => {
    const tabId = btn.getAttribute('data-tab');
    if (allowed.has(tabId)) {
      btn.style.display = 'inline-flex';
    } else {
      btn.style.display = 'none';
    }
  });

  // 3. Eğer mevcut açık sekme yeni role kapalıysa izin verilen ilk sekmeye geç
  if (!allowed.has(AppState.activeTab)) {
    const firstAllowed = user.allowed_tabs && user.allowed_tabs.length > 0 ? user.allowed_tabs[0] : 'tab-overview';
    switchTab(firstAllowed);
  }

  // 4. Yetkiye bağlı eylem butonları
  const btnHeaderGen = document.getElementById('btnHeaderGenerate');
  if (btnHeaderGen) {
    btnHeaderGen.style.display = user.can_edit_rules ? 'inline-flex' : 'none';
  }

  const rulesSubmit = document.querySelector('#rulesForm button[type="submit"]');
  if (rulesSubmit) {
    rulesSubmit.style.display = user.can_edit_rules ? 'inline-flex' : 'none';
  }
}

async function openRoleSwitcherModal() {
  const grid = document.getElementById('roleSwitcherGrid');
  if (!grid) return;

  grid.innerHTML = '<div style="grid-column: 1/-1; text-align: center; padding: 20px; color: var(--text-muted);">Rol profilleri yükleniyor...</div>';
  openModal('modalRoleSwitcher');

  try {
    const res = await fetch('/api/auth/users');
    if (!res.ok) throw new Error('Kullanıcı listesi alınamadı');
    const data = await res.json();

    let html = '';
    const activeUsername = AppState.currentUser ? AppState.currentUser.username : '';

    const roleDescriptions = {
      'hr': 'Tüm puantajı, eksik basımları, yasal uyum risk radarını ve imzalı A4 fişlerini yönetir.',
      'accounting': 'Maaş ve mesai maliyet radarını, net/brüt ve nakit farklarını inceler, resmi Excel ve bordroları alır.',
      'plant_manager': 'Üst düzey fabrika KPI\'larını, maliyet trendlerini ve yasal risk indeksini inceler.',
      'admin': 'Tüm modüllere, rapor üretimine, kurallara ve sistem ayarlarına tam yetkili erişim sağlar.'
    };

    const roleBadges = {
      'hr': ['Puantaj & Fişler', 'İstisna Onayı', 'Yasal Uyum'],
      'accounting': ['Maaş & Finans', 'Resmi Raporlar', 'Bölüm Primleri'],
      'plant_manager': ['Fabrika KPI', 'Maliyet Radarı', 'Uyum İndeksi'],
      'admin': ['Tam Yetki', 'Kural Düzenleme', 'Tüm Sekmeler']
    };

    data.users.forEach(u => {
      const isActive = u.username === activeUsername;
      const desc = roleDescriptions[u.role] || '';
      const tags = roleBadges[u.role] || [];

      html += `
        <div class="role-card ${isActive ? 'active-role' : ''}" onclick="switchRole('${u.username}')">
          <div class="role-card-header">
            <div class="role-card-avatar">${u.avatar || '👤'}</div>
            <div class="role-card-titles">
              <span class="role-card-name">${u.name}</span>
              <span class="role-card-role">${u.title}</span>
            </div>
          </div>
          <div class="role-card-desc">${desc}</div>
          <div class="role-card-scope">
            ${tags.map(t => `<span class="role-scope-tag">${t}</span>`).join('')}
          </div>
        </div>
      `;
    });

    grid.innerHTML = html;
  } catch (err) {
    grid.innerHTML = `<div style="grid-column: 1/-1; text-align: center; color: var(--color-danger); padding: 20px;">Hata: ${err.message}</div>`;
  }
}

async function switchRole(username, password = null) {
  try {
    const res = await fetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username: username, password: password })
    });

    if (!res.ok) {
      const errData = await res.json().catch(() => ({}));
      throw new Error(errData.detail || 'Giriş yapılamadı');
    }

    const data = await res.json();
    AppState.currentUser = data.user;
    localStorage.setItem('fide_user_role', data.user.username);

    applyRolePermissions(data.user);
    closeModal('modalRoleSwitcher');

    showToast(`🔐 Giriş yapıldı: ${data.user.name} (${data.user.title})`, 'success');

    // Aktif sekmenin içeriğini yeni yetkilerle tazele
    loadStats();
    if (AppState.activeTab === 'tab-matrix') loadMatrix();
    if (AppState.activeTab === 'tab-exceptions') loadExceptions();
    if (AppState.activeTab === 'tab-bonuses') loadBonuses();
    if (AppState.activeTab === 'tab-finance') {
      loadFinancialRadar();
      loadIcraRecords();
    }
    if (AppState.activeTab === 'tab-compliance') {
      loadComplianceRadar();
      loadCumulativeOvertime();
    }
    if (AppState.activeTab === 'tab-audit') {
      loadAuditLogs();
    }
  } catch (err) {
    showToast('Hata: ' + err.message, 'error');
  }
}

// ============================================================================
// 10. ÇOKLU AY ARŞİVİ, 270 SAAT MESAİ KÜTÜĞÜ VE İCRA MODÜLÜ (SQLITE PERSISTENCE)
// ============================================================================

// ----------------------------------------------------------------------------
// A) DÖNEM SEÇİCİ & ÇOKLU AY ARŞİVİ
// ----------------------------------------------------------------------------
async function openPeriodSwitcher() {
  const container = document.getElementById('periodListContainer');
  if (!container) return;
  container.innerHTML = '<div style="text-align:center; padding: 20px; color: var(--text-muted);">Dönem kayıtları yükleniyor...</div>';
  openModal('modalPeriodSwitcher');

  try {
    const res = await fetch('/api/archive/periods');
    if (!res.ok) throw new Error('Dönemler alınamadı');
    const data = await res.json();
    AppState.archive.periods = data.periods || [];

    if (AppState.archive.periods.length === 0) {
      container.innerHTML = '<div style="text-align:center; padding: 20px; color: var(--text-muted);">Kayıtlı arşiv dönemi bulunamadı.</div>';
      return;
    }

    let html = '';
    AppState.archive.periods.forEach(p => {
      const isActive = p.status === 'active';
      const statusBadge = isActive
        ? '<span class="period-status-pill">Aktif Dönem</span>'
        : '<span class="period-status-pill status-locked">Arşiv / Kilitli</span>';
      const pName = p.name || p.period_name || `${p.period_key} Dönemi`;
      const pCount = p.active_personnel_count || p.total_personnel || 250;

      html += `
        <div class="period-item-card ${isActive ? 'is-active' : ''}">
          <div class="period-item-info">
            <div class="period-item-icon">${isActive ? '📅' : '🔒'}</div>
            <div class="period-item-meta">
              <span class="period-item-title">
                ${pName}
                ${statusBadge}
              </span>
              <span class="period-item-desc">
                ${p.year} Yılı • ${p.month}. Ay • ${pCount} Personel Kaydı
              </span>
            </div>
          </div>
          <div>
            <button class="btn btn-sm ${isActive ? 'btn-outline' : 'btn-primary'}" onclick="selectArchivePeriod('${p.period_key}')">
              ${isActive ? 'İnceleniyor' : 'Görüntüle'}
            </button>
          </div>
        </div>
      `;
    });

    container.innerHTML = html;
  } catch (err) {
    container.innerHTML = `<div style="text-align:center; color: var(--color-danger); padding: 20px;">Hata: ${err.message}</div>`;
  }
}

async function selectArchivePeriod(periodKey) {
  try {
    const res = await fetch(`/api/archive/periods/${periodKey}`);
    if (!res.ok) throw new Error('Dönem verisi alınamadı');
    const data = await res.json();
    closeModal('modalPeriodSwitcher');
    const p = data.period || {};
    const pName = p.name || p.period_name || periodKey;
    const recCount = (p.records && p.records.length) || p.active_personnel_count || 250;
    showToast(`🗄️ ${pName} dönemi başarıyla yüklendi (${recCount} personel)`, 'success');
  } catch (err) {
    showToast('Hata: ' + err.message, 'error');
  }
}

async function handlePeriodRollover() {
  const activePeriod = (AppState.archive.periods.find(p => p.status === 'active') || {}).period_key || '2026-09';
  const periodParts = activePeriod.split('-');
  const y = parseInt(periodParts[0]);
  const m = parseInt(periodParts[1]);
  const nextM = m === 12 ? 1 : m + 1;
  const nextY = m === 12 ? y + 1 : y;
  const nextPeriodKey = `${nextY}-${String(nextM).padStart(2, '0')}`;

  const confirmMsg = `${activePeriod} dönemi kilitlenecek ve ${nextPeriodKey} dönemine devredilecektir.\n\nİcra kesintileri borçlardan otomatik düşülecek ve kalan bakiyeler sonraki aya devredilecektir.\n\nDevam etmek istiyor musunuz?`;
  if (!confirm(confirmMsg)) return;

  try {
    const res = await fetch('/api/archive/periods/close-and-rollover', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        current_period_key: activePeriod
      })
    });

    if (!res.ok) {
      const errData = await res.json().catch(() => ({}));
      throw new Error(errData.detail || 'Dönem devri yapılamadı');
    }

    const data = await res.json();
    closeModal('modalPeriodSwitcher');
    showToast(`✅ ${data.message}`, 'success', 5000);

    // Verileri güncelle
    loadStats();
    loadIcraRecords();
    loadCumulativeOvertime();
  } catch (err) {
    showToast('Devir Hatası: ' + err.message, 'error');
  }
}

// ----------------------------------------------------------------------------
// B) 270 SAAT YILLIK FAZLA MESAİ KÜTÜĞÜ (MD. 41 SİCİLİ)
// ----------------------------------------------------------------------------
async function loadCumulativeOvertime() {
  try {
    const res = await fetch(`/api/archive/cumulative-overtime?year=${AppState.archive.cumulativeOvertime.year}`);
    if (!res.ok) throw new Error('Yıllık mesai kütüğü alınamadı');
    const data = await res.json();

    AppState.archive.cumulativeOvertime.items = data.personnel || data.records || [];

    // KPI Güncellemesi
    if (data.kpis) {
      const elEx = document.getElementById('valOtCountExceeded');
      if (elEx) elEx.innerHTML = `${data.kpis.exceeded_count || 0} <small>kişi</small>`;

      const elCrit = document.getElementById('valOtCountCritical');
      if (elCrit) elCrit.innerHTML = `${data.kpis.critical_count || 0} <small>kişi</small>`;

      const elWarn = document.getElementById('valOtCountWarning');
      if (elWarn) elWarn.innerHTML = `${data.kpis.warning_count || 0} <small>kişi</small>`;

      const elSafe = document.getElementById('valOtCountSafe');
      if (elSafe) elSafe.innerHTML = `${data.kpis.safe_count || 0} <small>kişi</small>`;
    }

    applyCumulativeOvertimeFilters();
  } catch (err) {
    console.error('Yıllık mesai yükleme hatası:', err);
  }
}

function applyCumulativeOvertimeFilters() {
  const st = AppState.archive.cumulativeOvertime;
  let filtered = [...st.items];

  // Arama filtresi
  if (st.search) {
    const q = st.search.toLowerCase();
    filtered = filtered.filter(item => {
      const name = (item.ad_soyad || item.name || '').toLowerCase();
      const tc = item.tc || item.tc_no || '';
      const dept = (item.bolum || item.department || '').toLowerCase();
      return name.includes(q) || tc.includes(q) || dept.includes(q);
    });
  }

  // Risk filtresi
  if (st.riskFilter !== 'all') {
    filtered = filtered.filter(item => item.risk_status === st.riskFilter);
  }

  st.filteredItems = filtered;
  st.totalPages = Math.ceil(filtered.length / st.pageSize) || 1;
  if (st.page > st.totalPages) st.page = 1;

  renderCumulativeOvertimeTable();
}

function renderCumulativeOvertimeTable() {
  const st = AppState.archive.cumulativeOvertime;
  const tbody = document.getElementById('overtimeRegisterTableBody');
  if (!tbody) return;

  const startIdx = (st.page - 1) * st.pageSize;
  const pageItems = st.filteredItems.slice(startIdx, startIdx + st.pageSize);

  if (pageItems.length === 0) {
    tbody.innerHTML = '<tr><td colspan="20" style="text-align:center; padding: 36px; color: var(--text-muted);">Arama kriterlerine uygun fazla mesai kaydı bulunamadı.</td></tr>';
    return;
  }

  let html = '';
  pageItems.forEach((r, idx) => {
    const rowNum = startIdx + idx + 1;
    const totOt = r.total_ot_hours ?? r.cumulative_total ?? 0;
    const remLimit = r.remaining_limit_hours ?? r.remaining_limit ?? Math.max(0, 270 - totOt);
    const pct = Math.min(100, Math.round((totOt / 270.0) * 100));

    let barClass = 'ot-bar-safe';
    let badgeClass = 'badge-safe';
    let statusText = '🟢 Güvenli';

    if (r.risk_status === 'EXCEEDED') {
      barClass = 'ot-bar-exceeded';
      badgeClass = 'badge-exceeded-ot';
      statusText = '🚨 Kota Aşıldı';
    } else if (r.risk_status === 'CRITICAL') {
      barClass = 'ot-bar-critical';
      badgeClass = 'badge-critical-ot';
      statusText = '⚠️ Kritik Eşik';
    } else if (r.risk_status === 'WARNING') {
      barClass = 'ot-bar-warning';
      badgeClass = 'badge-warning-ot';
      statusText = '🟡 Yaklaşıyor';
    }

    const renderMonth = (val, isCurrent = false) => {
      const cls = isCurrent ? 'ot-month-cell highlight-current' : val > 0 ? 'ot-month-cell has-val' : 'ot-month-cell';
      return `<td class="${cls}">${val > 0 ? fmtHours(val) : '—'}</td>`;
    };

    html += `
      <tr>
        <td style="text-align: center; color: var(--text-dim);">${rowNum}</td>
        <td style="font-family: var(--font-mono); font-size: 11.5px;">${r.tc || r.tc_no || '—'}</td>
        <td style="font-weight: 600; color: #f8fafc;">${r.ad_soyad || r.name}</td>
        <td style="color: var(--text-secondary); font-size: 11.5px;">${r.bolum || r.department || '—'}</td>
        ${renderMonth(r.m01)}
        ${renderMonth(r.m02)}
        ${renderMonth(r.m03)}
        ${renderMonth(r.m04)}
        ${renderMonth(r.m05)}
        ${renderMonth(r.m06)}
        ${renderMonth(r.m07)}
        ${renderMonth(r.m08)}
        ${renderMonth(r.m09, true)}
        ${renderMonth(r.m10)}
        ${renderMonth(r.m11)}
        ${renderMonth(r.m12)}
        <td style="text-align: right; font-weight: 700; color: #38bdf8;">${fmtHours(totOt)}s</td>
        <td style="text-align: right; font-weight: 600; color: ${remLimit <= 0 ? '#f43f5e' : '#a3e635'};">${fmtHours(remLimit)}s</td>
        <td>
          <div class="ot-progress-wrapper">
            <div class="ot-progress-track">
              <div class="ot-progress-bar ${barClass}" style="width: ${pct}%;"></div>
            </div>
            <span class="ot-pct-label">%${pct}</span>
          </div>
        </td>
        <td style="text-align: center;">
          <span class="badge ${badgeClass}">${statusText}</span>
        </td>
      </tr>
    `;
  });

  tbody.innerHTML = html;

  // Pagination info
  const pInfo = document.getElementById('overtimePaginationInfo');
  if (pInfo) {
    pInfo.innerText = `Gösterilen: ${st.filteredItems.length === 0 ? 0 : startIdx + 1} - ${Math.min(startIdx + st.pageSize, st.filteredItems.length)} / ${st.filteredItems.length} Personel`;
  }
  const pDisp = document.getElementById('overtimePageDisplay');
  if (pDisp) {
    pDisp.innerText = `Sayfa ${st.page} / ${st.totalPages}`;
  }
}

// ----------------------------------------------------------------------------
// C) İCRA VE MAAŞ HACZİ TAKİP MASASI
// ----------------------------------------------------------------------------
async function loadIcraRecords() {
  try {
    const res = await fetch(`/api/archive/icra?status=ALL`);
    if (!res.ok) throw new Error('İcra verileri alınamadı');
    const data = await res.json();

    AppState.archive.icra.items = data.files || data.records || [];
    AppState.archive.icra.kpis = data.kpis || {};

    // KPI'lar
    if (data.kpis) {
      const k = data.kpis;
      const elTotal = document.getElementById('valIcraTotalDebt');
      if (elTotal) elTotal.innerText = fmtCurrency(k.total_debt_tl ?? k.total_initial_debt ?? 0);

      const elMonth = document.getElementById('valIcraMonthDeducted');
      if (elMonth) elMonth.innerText = fmtCurrency(k.total_deducted_tl ?? k.current_month_deductions ?? 0);

      const elRem = document.getElementById('valIcraRemainingDebt');
      if (elRem) elRem.innerText = fmtCurrency(k.total_remaining_tl ?? k.total_remaining_debt ?? 0);

      const badgeFiles = document.getElementById('badgeIcraActiveFiles');
      if (badgeFiles) badgeFiles.innerText = `${k.active_files ?? k.active_files_count ?? 0} Aktif Haciz Dosyası`;
    }

    renderIcraTable();
  } catch (err) {
    console.error('İcra verisi yükleme hatası:', err);
  }
}

function renderIcraTable() {
  const tbody = document.getElementById('icraTableBody');
  if (!tbody) return;

  const items = AppState.archive.icra.items;
  if (!items || items.length === 0) {
    tbody.innerHTML = '<tr><td colspan="11" style="text-align:center; padding: 36px; color: var(--text-muted);">Kayıtlı icra veya maaş haczi dosyası bulunamadı.</td></tr>';
    return;
  }

  let html = '';
  items.forEach((r, idx) => {
    const durum = r.durum || r.status || 'AKTİF';
    const isClosed = durum === 'KAPANDI' || (r.kalan_borc !== undefined && r.kalan_borc <= 0);
    const statusBadge = isClosed
      ? '<span class="badge badge-emerald">KAPANDI</span>'
      : '<span class="badge badge-rose">AKTİF HACİZ</span>';

    const totalD = r.toplam_borc ?? r.total_debt ?? 0;
    const dedD = r.kesilen_kumulatif ?? r.deducted_amount ?? 0;
    const remD = r.kalan_borc ?? r.remaining_balance ?? 0;
    const monthDed = Math.round(remD * (r.aylik_kesinti_orani || 0.25) * 100) / 100;
    const nextRem = Math.max(0, round2(remD - monthDed));

    html += `
      <tr>
        <td style="text-align: center; color: var(--text-dim);">${idx + 1}</td>
        <td style="font-family: var(--font-mono); font-weight: 600; color: #f8fafc;">${r.dosya_no || '—'}</td>
        <td style="font-weight: 600; color: #e2e8f0;">${r.ad_soyad || r.name}</td>
        <td style="color: var(--text-secondary); font-size: 11.5px;">${r.bolum || r.department || '—'}</td>
        <td style="font-size: 11.5px; color: var(--text-muted);">${r.sirket || r.icra_dairesi || 'Bergama İcra D.'}</td>
        <td style="text-align: right; font-weight: 600;">${fmtCurrency(totalD)}</td>
        <td style="text-align: right; color: #94a3b8;">${fmtCurrency(dedD)}</td>
        <td style="text-align: right; font-weight: 700; color: #f43f5e;">${fmtCurrency(remD)}</td>
        <td style="text-align: right; font-weight: 600; color: #38bdf8;">${fmtCurrency(monthDed)}</td>
        <td style="text-align: right; font-weight: 700; color: #fbbf24;">${fmtCurrency(nextRem)}</td>
        <td style="text-align: center;">${statusBadge}</td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

function round2(val) {
  return Math.round(val * 100) / 100;
}

async function openNewIcraModal() {
  const sel = document.getElementById('icraPersonnelSelect');
  if (sel) {
    sel.innerHTML = '<option value="">Personel seçiniz...</option>';
    try {
      const res = await fetch('/api/matrix?page=1&page_size=300');
      if (res.ok) {
        const data = await res.json();
        const items = data.items || [];
        items.forEach(p => {
          const opt = document.createElement('option');
          opt.value = p.tc_no;
          opt.dataset.name = p.name;
          opt.dataset.nameKey = p.name_key || p.name;
          opt.dataset.dept = p.department;
          opt.textContent = `${p.name} (TC: ${p.tc_no} - ${p.department})`;
          sel.appendChild(opt);
        });
      }
    } catch (err) {
      console.error('Personel listesi yüklenemedi:', err);
    }
  }
  openModal('modalNewIcra');
}

async function handleNewIcraSubmit(e) {
  e.preventDefault();
  const sel = document.getElementById('icraPersonnelSelect');
  const selectedOpt = sel.options[sel.selectedIndex];
  const tcNo = sel.value;
  const adSoyad = selectedOpt ? selectedOpt.dataset.name : '';
  const nameKey = selectedOpt ? selectedOpt.dataset.nameKey : '';
  const bolum = selectedOpt ? selectedOpt.dataset.dept : 'Genel';
  const dosyaNo = document.getElementById('icraDosyaNo').value.trim();
  const icraDairesi = document.getElementById('icraDaire').value.trim();
  const alacakli = document.getElementById('icraAlacakli').value.trim();
  const totalDebt = parseFloat(document.getElementById('icraTotalDebt').value);

  if (!tcNo || !dosyaNo || !icraDairesi || isNaN(totalDebt)) {
    showToast('Lütfen tüm zorunlu alanları eksiksiz doldurun', 'warning');
    return;
  }

  try {
    const res = await fetch('/api/archive/icra', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tc: tcNo,
        name_key: nameKey || adSoyad,
        ad_soyad: adSoyad,
        bolum: bolum,
        sirket: icraDairesi,
        dosya_no: dosyaNo,
        toplam_borc: totalDebt,
        kesilen_kumulatif: 0.0,
        kalan_borc: totalDebt,
        aylik_kesinti_orani: 0.25,
        durum: "AKTİF",
        aciklama: alacakli
      })
    });

    if (!res.ok) {
      const errData = await res.json().catch(() => ({}));
      throw new Error(errData.detail || 'İcra kaydı oluşturulamadı');
    }

    const data = await res.json();
    closeModal('modalNewIcra');
    document.getElementById('formNewIcra').reset();
    showToast(`✅ ${data.message || 'İcra kaydı başarıyla eklendi.'}`, 'success');
    loadIcraRecords();
  } catch (err) {
    showToast('Hata: ' + err.message, 'error');
  }
}

// ============================================================================
// 11. DEĞİŞİKLİK VE GÜVENLİK DENETİM İZİ MASASI (AUDIT TRAIL)
// ============================================================================
async function loadAuditLogs() {
  const tableBody = document.getElementById('auditTableBody');
  if (!tableBody) return;

  const params = new URLSearchParams({
    page: AppState.audit.page,
    page_size: AppState.audit.pageSize
  });

  if (AppState.audit.search) {
    params.append('search', AppState.audit.search);
  }
  if (AppState.audit.actionType && AppState.audit.actionType !== 'ALL') {
    params.append('action_type', AppState.audit.actionType);
  }
  if (AppState.audit.username && AppState.audit.username !== 'ALL') {
    params.append('username', AppState.audit.username);
  }

  try {
    const res = await fetch(`/api/audit/logs?${params.toString()}`);
    if (!res.ok) throw new Error('Denetim kayıtları alınamadı');
    const data = await res.json();

    AppState.audit.items = data.items || [];
    AppState.audit.totalCount = data.total || 0;
    AppState.audit.totalPages = data.total_pages || 1;
    AppState.audit.summary = data.summary || {};

    renderAuditKPIs(data.summary);
    renderAuditTable(data.items);
    renderAuditPagination(data);

    // Rozet güncelle
    const badgeAudit = document.getElementById('badgeAuditCount');
    if (badgeAudit) {
      badgeAudit.innerText = `${data.total || 0}`;
    }
  } catch (err) {
    console.error('Audit logs fetch error:', err);
    tableBody.innerHTML = `<tr><td colspan="9" style="text-align: center; color: var(--color-danger); padding: 2rem;">Hata: ${err.message}</td></tr>`;
  }
}

function renderAuditKPIs(summary = {}) {
  const valTotal = document.getElementById('valAuditTotal');
  const valHours = document.getElementById('valAuditHours');
  const valToday = document.getElementById('valAuditToday');
  const valUser = document.getElementById('valAuditUser');

  if (valTotal) valTotal.innerText = fmtInt(summary.total_logs || 0);
  if (valHours) valHours.innerText = fmtInt(summary.hours_revisions || 0);
  if (valToday) valToday.innerText = fmtInt(summary.today_logs || 0);
  if (valUser) valUser.innerText = summary.most_active_user || 'Kayıt Yok';
}

function renderAuditTable(items = []) {
  const tableBody = document.getElementById('auditTableBody');
  if (!tableBody) return;

  if (items.length === 0) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="9" style="text-align: center; padding: 2.5rem; color: var(--text-muted);">
          Kriterlere uygun idari değişiklik veya denetim izi kaydı bulunamadı.
        </td>
      </tr>
    `;
    return;
  }

  const actionBadgeMap = {
    'EXCEPTION_RESOLVE': { label: 'Eksik Basım', cls: 'badge-action-exception', icon: '⏱️' },
    'BULK_RESOLVE': { label: 'Toplu Onay', cls: 'badge-action-bulk', icon: '⚡' },
    'RULE_UPDATE': { label: 'Kural Değişimi', cls: 'badge-action-rule', icon: '⚙️' },
    'PERIOD_ROLLOVER': { label: 'Dönem Devir', cls: 'badge-action-rollover', icon: '📅' },
    'ICRA_ADD': { label: 'İcra / Haciz', cls: 'badge-action-icra', icon: '⚖️' },
    'FILE_UPLOAD': { label: 'Excel Yükleme', cls: 'badge-action-upload', icon: '📤' }
  };

  let html = '';
  items.forEach(it => {
    const badgeInfo = actionBadgeMap[it.action_type] || { label: it.action_type, cls: 'badge-action-exception', icon: '📝' };
    const dayStr = it.day ? `Gün ${it.day}` : '-';

    html += `
      <tr>
        <td style="white-space: nowrap; font-family: var(--font-mono); font-size: 11px; color: var(--text-muted);">
          ${it.timestamp || '-'}
        </td>
        <td>
          <div class="audit-user-cell">
            <span class="audit-user-name">${it.user_name || it.username}</span>
            <span class="audit-user-sub">${it.username} (${(it.user_role || '').toUpperCase()})</span>
          </div>
        </td>
        <td>
          <span class="audit-action-badge ${badgeInfo.cls}">
            <span>${badgeInfo.icon}</span>
            <span>${badgeInfo.label}</span>
          </span>
        </td>
        <td>
          <div style="font-weight: 600; color: var(--text-main);">${it.target_name || '-'}</div>
          <div style="font-size: 11px; color: var(--text-muted);">
            ${it.target_tc ? `TC: ${it.target_tc}` : ''} ${it.target_dept ? `| Bölüm: ${it.target_dept}` : ''}
          </div>
        </td>
        <td style="text-align: center; font-weight: 600; font-size: 11px;">
          ${dayStr}
        </td>
        <td>
          <span class="audit-val-old">${it.old_value || '-'}</span>
        </td>
        <td>
          <span class="audit-val-new">${it.new_value || '-'}</span>
        </td>
        <td style="font-size: 12px; color: var(--text-main); max-width: 280px;">
          ${it.reason || '-'}
        </td>
        <td>
          <span class="audit-ip-tag">${it.ip_address || '-'}</span>
        </td>
      </tr>
    `;
  });

  tableBody.innerHTML = html;
}

function renderAuditPagination(data) {
  const info = document.getElementById('auditPaginationInfo');
  const current = document.getElementById('auditCurrentPage');
  const btnPrev = document.getElementById('btnAuditPrevPage');
  const btnNext = document.getElementById('btnAuditNextPage');

  const total = data.total || 0;
  const page = data.page || 1;
  const pageSize = data.page_size || 50;
  const totalPages = data.total_pages || 1;

  const start = total === 0 ? 0 : (page - 1) * pageSize + 1;
  const end = Math.min(page * pageSize, total);

  if (info) info.innerText = `Gösterilen: ${start} - ${end} / ${total} Kayıt`;
  if (current) current.innerText = `Sayfa ${page} / ${totalPages}`;

  if (btnPrev) btnPrev.disabled = (page <= 1);
  if (btnNext) btnNext.disabled = (page >= totalPages);
}


