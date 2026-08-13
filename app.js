(() => {
  'use strict';

  const DB = globalThis.ConstructionDB;
  const Excel = globalThis.ConstructionExcel;
  const main = document.getElementById('appMain');
  const modal = document.getElementById('modal');
  const modalForm = document.getElementById('modalForm');
  const modalTitle = document.getElementById('modalTitle');
  const modalEyebrow = document.getElementById('modalEyebrow');
  const modalBody = document.getElementById('modalBody');
  const modalActions = document.getElementById('modalActions');
  const moreMenu = document.getElementById('moreMenu');
  const excelFileInput = document.getElementById('excelFileInput');
  const backupFileInput = document.getElementById('backupFileInput');

  const state = {
    projects: [],
    vendors: [],
    school: {},
    currentProjectId: null,
    filter: 'active',
    search: '',
    autosaveTimer: null,
    saveToken: 0,
    importMode: 'auto'
  };

  const FIELDS_FOR_IMPORT = [
    ['projectName','공사명'], ['contractNumber','계약번호'], ['vendorName','계약상대자'],
    ['representative','대표자'], ['businessNumber','사업자번호'], ['vendorAddress','업체주소'],
    ['vendorPhone','업체 전화'], ['contractMethod','계약방법'], ['procurementMethod','G2B/S2B'],
    ['estimatedPrice','예정가격'], ['originalContractAmount','최초 계약금액'], ['currentContractAmount','계약금액'],
    ['bidRate','낙찰율'], ['contractDate','계약일'], ['startDate','착공일'], ['completionDueDate','준공기한'],
    ['actualCompletionDate','실제 준공일'], ['completionInspectionDate','검사·검수일'],
    ['paymentDate','지출일'], ['paymentAmount','지출금액'], ['supervisor','공사감독'],
    ['contractSecurityType','계약보증방법'], ['contractSecurityRate','계약보증률'], ['contractSecurityAmount','계약보증금액'],
    ['defectSecurityType','하자보증방법'], ['defectSecurityRate','하자보증률'], ['defectSecurityAmount','하자보증금액'],
    ['defectPeriodYears','하자기간'], ['defectStartDate','하자 시작일'], ['defectEndDate','하자 종료일'],
    ['licenseType','등록면허'], ['fundingSource','재원구분']
  ];

  const MONEY_FIELDS = new Set(['estimatedPrice','originalContractAmount','currentContractAmount','paymentAmount','contractSecurityAmount','defectSecurityAmount']);

  function e(value) {
    return String(value ?? '').replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));
  }

  function normalizeText(value) {
    return String(value ?? '').trim().replace(/\s+/g, '').toLowerCase();
  }

  function meaningful(value) {
    return value !== '' && value !== null && value !== undefined;
  }

  function sameValue(a, b, field) {
    if (!meaningful(a) && !meaningful(b)) return true;
    if (MONEY_FIELDS.has(field) || ['bidRate','contractSecurityRate','defectSecurityRate','defectPeriodYears'].includes(field)) {
      const an = Number(a), bn = Number(b);
      return Number.isFinite(an) && Number.isFinite(bn) ? Math.abs(an - bn) < 0.0001 : normalizeText(a) === normalizeText(b);
    }
    return normalizeText(a) === normalizeText(b);
  }

  function formatMoney(value) {
    if (!meaningful(value)) return '—';
    const n = Number(value);
    return Number.isFinite(n) ? `${n.toLocaleString('ko-KR')}원` : String(value);
  }

  function formatDate(value) {
    if (!value) return '—';
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
    return m ? `${Number(m[1])}.${Number(m[2])}.${Number(m[3])}.` : value;
  }

  function formatDateTime(value) {
    if (!value) return '';
    try { return new Intl.DateTimeFormat('ko-KR', { month:'numeric', day:'numeric', hour:'2-digit', minute:'2-digit' }).format(new Date(value)); }
    catch { return ''; }
  }

  function statusOf(p) {
    if (p.paymentDate) return { key:'done', label:'완료', cls:'done' };
    if (p.actualCompletionDate) return { key:'waiting', label:'준공·지출대기', cls:'wait' };
    if (p.startDate) return { key:'active', label:'공사진행', cls:'' };
    if (p.contractDate) return { key:'contract', label:'계약', cls:'' };
    return { key:'draft', label:'준비', cls:'wait' };
  }

  function missingFor(p) {
    const s = statusOf(p).key;
    const list = [];
    if (!p.projectName) list.push('공사명');
    if (!p.vendorName) list.push('업체');
    if (['contract','active','waiting','done'].includes(s) && !meaningful(p.currentContractAmount)) list.push('계약금액');
    if (['contract','active','waiting','done'].includes(s) && !p.contractDate) list.push('계약일');
    if (['active','waiting','done'].includes(s) && !p.completionDueDate) list.push('준공기한');
    if (s === 'waiting' && !p.paymentDate) list.push('지출일');
    return list;
  }

  function projectSubtitle(p) {
    const bits = [p.fiscalYear ? `${p.fiscalYear}회계연도` : '', p.workType, p.contractNumber ? `계약 ${p.contractNumber}` : ''].filter(Boolean);
    return bits.join(' · ') || '기본정보 입력 중';
  }

  function sourceLabel(p) {
    if (String(p.source || '').includes('edufine') && String(p.source || '').includes('audit')) return '기존 공사이력에서 불러온 뒤 K-에듀파인 자료로 보완한 공사';
    if (p.source === 'manual+edufine') return '웹에서 만든 뒤 K-에듀파인 자료로 보완한 공사';
    if (p.source === 'edufine') return 'K-에듀파인에서 불러온 공사';
    if (p.source === 'audit-excel') return '기존 학교 공사이력에서 불러온 공사';
    return '웹에서 직접 만든 공사';
  }

  function showToast(message, type = '') {
    const region = document.getElementById('toastRegion');
    const div = document.createElement('div');
    div.className = `toast ${type}`.trim();
    div.textContent = message;
    region.appendChild(div);
    setTimeout(() => div.remove(), 3200);
  }

  function openModal({ title, eyebrow = '', body = '', actions = '' }) {
    modalTitle.textContent = title;
    modalEyebrow.textContent = eyebrow;
    modalBody.innerHTML = body;
    modalActions.innerHTML = actions;
    if (!modal.open) modal.showModal();
  }

  function closeModal() { if (modal.open) modal.close(); }

  function openExcelPicker(mode = 'auto') {
    state.importMode = mode;
    excelFileInput.click();
  }

  async function loadState() {
    await DB.openDb();
    const [projects, vendors, school] = await Promise.all([
      DB.getAll('projects'), DB.getAll('vendors'), DB.get('settings', 'school')
    ]);
    state.projects = projects.sort((a,b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
    state.vendors = vendors.sort((a,b) => (a.name || '').localeCompare(b.name || '', 'ko'));
    state.school = school?.value || {};
  }

  function render() {
    if (state.currentProjectId) renderProjectDetail();
    else renderDashboard();
  }

  function renderDashboard() {
    const active = state.projects.filter(p => statusOf(p).key !== 'done').length;
    const waiting = state.projects.filter(p => statusOf(p).key === 'waiting').length;
    const done = state.projects.filter(p => statusOf(p).key === 'done').length;

    let filtered = [...state.projects];
    if (state.filter === 'active') filtered = filtered.filter(p => statusOf(p).key !== 'done');
    if (state.filter === 'done') filtered = filtered.filter(p => statusOf(p).key === 'done');
    if (state.search.trim()) {
      const q = normalizeText(state.search);
      filtered = filtered.filter(p => [p.projectName,p.vendorName,p.contractNumber,p.workType].some(v => normalizeText(v).includes(q)));
    }

    main.innerHTML = `
      <section class="hero">
        <div>
          <p class="eyebrow">공사정보 한 번만</p>
          <h1>${state.school.name ? `${e(state.school.name)} 공사` : '우리학교 공사'}</h1>
          <p>직접 입력하거나 기존 자료를 불러오세요. 이미 있는 정보는 다시 입력하지 않습니다.</p>
          <div class="security-note"><span class="security-dot"></span>브라우저 자동저장 · 입력 데이터 서버 미전송</div>
        </div>
      </section>

      <section class="start-grid" aria-label="공사정보 시작 방법">
        <button class="start-card primary-start" type="button" data-start-new>
          <span class="start-step">1</span>
          <span class="start-copy"><strong>새 공사 직접등록</strong><small>공사명만으로도 시작하고, 확정된 정보만 덧붙입니다.</small></span>
          <span class="start-arrow">›</span>
        </button>
        <button class="start-card" type="button" data-start-history>
          <span class="start-step">2</span>
          <span class="start-copy"><strong>기존 공사이력 불러오기</strong><small>학교 공사 이력 현황.xlsx의 여러 공사를 한꺼번에 가져옵니다.</small></span>
          <span class="start-arrow">›</span>
        </button>
        <button class="start-card" type="button" data-start-edufine>
          <span class="start-step">3</span>
          <span class="start-copy"><strong>에듀파인으로 정보 업데이트</strong><small>자료관리목록.xlsx에서 계약·준공·지출 정보를 찾아 기존 공사를 보완합니다.</small></span>
          <span class="start-arrow">›</span>
        </button>
      </section>

      <section class="summary-grid" aria-label="공사 현황 요약">
        <article class="summary-card"><span>진행 중</span><strong>${active}</strong></article>
        <article class="summary-card"><span>준공·지출대기</span><strong>${waiting}</strong></article>
        <article class="summary-card"><span>완료</span><strong>${done}</strong></article>
      </section>

      <section class="panel">
        <div class="toolbar">
          <div class="toolbar-left">
            <div class="search-wrap"><input id="projectSearch" type="search" value="${e(state.search)}" placeholder="공사명 · 업체명 · 계약번호 검색" aria-label="공사 검색"></div>
          </div>
          <div class="toolbar-right">
            <div class="segmented" aria-label="공사 상태 필터">
              <button type="button" data-filter="active" class="${state.filter==='active'?'active':''}">진행중</button>
              <button type="button" data-filter="done" class="${state.filter==='done'?'active':''}">완료</button>
              <button type="button" data-filter="all" class="${state.filter==='all'?'active':''}">전체</button>
            </div>
          </div>
        </div>
        <div class="project-list">
          ${filtered.length ? filtered.map(projectRowHtml).join('') : emptyStateHtml()}
        </div>
      </section>`;

    document.getElementById('projectSearch')?.addEventListener('input', ev => {
      state.search = ev.target.value;
      renderDashboard();
      const input = document.getElementById('projectSearch');
      input?.focus();
      if (input) input.setSelectionRange(input.value.length, input.value.length);
    });
    main.querySelectorAll('[data-filter]').forEach(btn => btn.addEventListener('click', () => { state.filter = btn.dataset.filter; renderDashboard(); }));
    main.querySelectorAll('[data-project-id]').forEach(row => row.addEventListener('click', () => openProject(row.dataset.projectId)));
    main.querySelector('[data-start-new]')?.addEventListener('click', openNewProjectModal);
    main.querySelector('[data-start-history]')?.addEventListener('click', () => openExcelPicker('history'));
    main.querySelector('[data-start-edufine]')?.addEventListener('click', () => openExcelPicker('edufine'));
  }

  function projectRowHtml(p) {
    const status = statusOf(p);
    const missing = missingFor(p);
    return `
      <div class="project-row" data-project-id="${e(p.id)}" tabindex="0" role="button" aria-label="${e(p.projectName || '이름 없는 공사')} 열기">
        <div class="project-title">
          <strong>${e(p.projectName || '이름 없는 공사')}</strong>
          <small>${e(projectSubtitle(p))}</small>
        </div>
        <div>
          <span class="cell-label">업체</span>
          <span class="cell-value">${e(p.vendorName || '—')}</span>
        </div>
        <div class="project-method">
          <span class="cell-label">계약방법</span>
          <span class="cell-value">${e(p.contractMethod || '—')}</span>
        </div>
        <div>
          <span class="cell-label">계약금액</span>
          <span class="cell-value money">${e(formatMoney(p.currentContractAmount))}</span>
        </div>
        <div>
          <span class="cell-label">상태</span>
          <span class="status-chip ${status.cls}">${e(status.label)}</span>
          ${missing.length ? `<div class="missing">⚠ ${e(missing.slice(0,2).join(', '))}${missing.length>2?' 외':''}</div>` : ''}
        </div>
        <div class="row-arrow">›</div>
      </div>`;
  }

  function emptyStateHtml() {
    if (state.projects.length && (state.search || state.filter !== 'all')) {
      return `<div class="empty-state"><div class="empty-icon">⌕</div><h3>조건에 맞는 공사가 없어요</h3><p>검색어나 상태 필터를 바꿔보세요.</p></div>`;
    }
    return `<div class="empty-state">
      <div class="empty-icon">工</div>
      <h3>아직 등록된 공사가 없어요</h3>
      <p>위의 세 가지 시작 방법 중 지금 가지고 있는 자료에 맞는 방법을 선택하세요.</p>
    </div>`;
  }

  function openProject(id) {
    state.currentProjectId = id;
    window.scrollTo({ top: 0, behavior: 'instant' });
    render();
  }

  function currentProject() { return state.projects.find(p => p.id === state.currentProjectId); }

  function renderProjectDetail() {
    const p = currentProject();
    if (!p) { state.currentProjectId = null; renderDashboard(); return; }
    const status = statusOf(p);
    const missing = missingFor(p);
    const vendorOptions = [`<option value="">직접 입력 / 새 업체</option>`, ...state.vendors.map(v => `<option value="${e(v.id)}" ${p.vendorId===v.id?'selected':''}>${e(v.name)}${v.businessNumber ? ` · ${e(v.businessNumber)}` : ''}</option>`)].join('');

    main.innerHTML = `
      <button class="back-button" id="backToList" type="button">← 공사 목록</button>
      <div class="detail-head">
        <div class="current-project">
          <p class="eyebrow">현재 작업 중인 공사</p>
          <h1>${e(p.projectName || '이름 없는 공사')}</h1>
          <div class="current-meta"><span class="status-chip ${status.cls}">${e(status.label)}</span><span>${e(projectSubtitle(p))}</span></div>
        </div>
        <div class="save-state" id="saveState"><span class="pulse"></span><span>이 기기에 저장됨</span></div>
      </div>

      <div class="detail-layout">
        <section class="panel form-panel" id="projectForm">
          ${sectionHtml('기본정보','공사 자체를 식별하는 핵심 정보', [
            field('projectName','공사명',p.projectName,'text',true),
            field('fiscalYear','회계연도',p.fiscalYear,'number'),
            selectField('workType','공종',p.workType,['','건축공사','전기공사','통신공사','소방공사','기계설비공사','토목공사','기타']),
            selectField('contractMethod','계약방법',p.contractMethod,['','1인수의','2인이상수의','제한경쟁','일반경쟁','조달계약','기타']),
            moneyField('estimatedPrice','예정가격',p.estimatedPrice),
            moneyField('currentContractAmount','계약금액',p.currentContractAmount)
          ])}
          ${sectionHtml('업체정보','한 번 저장한 업체는 다음 공사에서 선택해 재사용할 수 있습니다.', [
            `<div class="field full"><label for="vendorPicker">업체 보관함에서 선택</label><select id="vendorPicker">${vendorOptions}</select><span class="hint">선택하면 대표자·사업자번호·주소·전화가 현재 공사에 자동 반영됩니다.</span></div>`,
            field('vendorName','업체명',p.vendorName),
            field('representative','대표자',p.representative),
            field('businessNumber','사업자등록번호',p.businessNumber),
            field('vendorPhone','대표전화',p.vendorPhone),
            field('vendorAddress','사업장 주소',p.vendorAddress,'text',true),
            field('licenseType','등록면허 / 업종',p.licenseType,'text',true),
            `<div class="field full"><button class="button secondary small" id="saveVendorBtn" type="button">현재 업체정보를 보관함에 반영</button></div>`
          ])}
          ${sectionHtml('계약','계약이 확정되면 생기는 정보를 추가합니다.', [
            field('contractNumber','계약번호',p.contractNumber),
            field('contractDate','계약일',p.contractDate,'date'),
            moneyField('originalContractAmount','최초 계약금액',p.originalContractAmount),
            field('bidRate','낙찰율 (%)',p.bidRate,'number'),
            field('procurementMethod','G2B / S2B / 조달방식',p.procurementMethod),
            field('contractSecurityType','계약보증 방법',p.contractSecurityType),
            field('contractSecurityRate','계약보증률',p.contractSecurityRate,'number'),
            moneyField('contractSecurityAmount','계약보증금액',p.contractSecurityAmount)
          ])}
          ${sectionHtml('착공 · 준공','공사가 진행되면서 새로 생기는 값만 덧붙입니다.', [
            field('plannedStartDate','착공예정일',p.plannedStartDate,'date'),
            field('startDate','착공일',p.startDate,'date'),
            field('completionDueDate','준공기한',p.completionDueDate,'date'),
            field('actualCompletionDate','실제 준공일',p.actualCompletionDate,'date'),
            field('siteManager','현장대리인',p.siteManager),
            field('siteManagerLicense','현장대리인 자격증 종류',p.siteManagerLicense),
            field('completionInspectionDate','검사·검수일',p.completionInspectionDate,'date'),
            field('completionRequestIssueDate','준공검사원 발행일',p.completionRequestIssueDate,'date'),
            field('completionRequestDocDate','준공검사원 문서등록일',p.completionRequestDocDate,'date'),
            field('completionInspectionRecordDate','준공검사조서 작성일',p.completionInspectionRecordDate,'date')
          ])}
          ${sectionHtml('지출 · 하자','준공 후 생기는 정보를 마지막으로 채웁니다.', [
            field('advancePayment','선금 지급 여부',p.advancePayment),
            field('taxInvoiceDate','세금계산서 발행일',p.taxInvoiceDate,'date'),
            field('paymentDate','지출일',p.paymentDate,'date'),
            moneyField('paymentAmount','지출금액',p.paymentAmount),
            field('defectSecurityType','하자보증서 / 각서',p.defectSecurityType),
            field('defectPeriodYears','하자담보기간(년)',p.defectPeriodYears,'number'),
            field('defectStartDate','하자 시작일',p.defectStartDate,'date'),
            field('defectEndDate','하자 종료일',p.defectEndDate,'date'),
            field('defectSecurityRate','하자보증률',p.defectSecurityRate,'number'),
            moneyField('defectSecurityAmount','하자보증금액',p.defectSecurityAmount),
            field('fundingSource','재원구분',p.fundingSource),
            field('ledgerPrint','공사대장 출력',p.ledgerPrint),
            textareaField('notes','비고',p.notes)
          ])}
        </section>

        <aside class="side-stack">
          <section class="side-card">
            <h3>현재 단계</h3>
            <p>입력된 날짜를 기준으로 자동 판단합니다.</p>
            <div class="progress-list">
              ${progressRow('계약', !!p.contractDate)}
              ${progressRow('착공', !!p.startDate)}
              ${progressRow('준공', !!p.actualCompletionDate)}
              ${progressRow('지출', !!p.paymentDate)}
            </div>
          </section>
          <section class="side-card">
            <h3>빠진 정보</h3>
            ${missing.length ? `<p>현재 단계에서 확인하면 좋은 항목입니다.</p><ul class="missing-list">${missing.map(x=>`<li>${e(x)}</li>`).join('')}</ul>` : `<p>✓ 현재 단계의 핵심정보가 입력되어 있습니다.</p>`}
          </section>
          <section class="side-card">
            <h3>공사서류</h3>
            <p>v0.2는 반복입력 절감 흐름을 먼저 안정화합니다.</p>
            <div class="doc-list">
              <div class="doc-item"><span>착공계</span><em>다음 버전</em></div>
              <div class="doc-item"><span>준공계</span><em>다음 버전</em></div>
              <div class="doc-item"><span>준공검사조서</span><em>다음 버전</em></div>
              <div class="doc-item"><span>대금청구서</span><em>다음 버전</em></div>
            </div>
          </section>
          <section class="side-card">
            <h3>자료 출처</h3>
            <p>${e(sourceLabel(p))}</p>
            ${p.sourceUpdatedAt ? `<p style="margin-top:6px">최근 반영 ${e(formatDateTime(p.sourceUpdatedAt))}</p>` : ''}
          </section>
          <section class="side-card danger-zone">
            <h3>공사 삭제</h3>
            <p>이 공사만 삭제되며 다른 공사에는 영향을 주지 않습니다.</p>
            <button class="button danger small" id="deleteProjectBtn" type="button" style="margin-top:10px">이 공사 삭제</button>
          </section>
        </aside>
      </div>`;

    document.getElementById('backToList').addEventListener('click', () => { state.currentProjectId = null; renderDashboard(); });
    document.getElementById('vendorPicker').addEventListener('change', applyVendorToCurrentProject);
    document.getElementById('saveVendorBtn').addEventListener('click', saveCurrentVendor);
    document.getElementById('deleteProjectBtn').addEventListener('click', confirmDeleteProject);
    main.querySelectorAll('[data-field]').forEach(input => {
      input.addEventListener('input', onProjectInput);
      input.addEventListener('change', onProjectInput);
    });
    initDateInputs(main);
    initMoneyInputs(main);
  }

  function sectionHtml(title, description, fields) {
    return `<div class="section"><div class="section-head"><div><h2>${e(title)}</h2><p>${e(description)}</p></div></div><div class="form-grid">${fields.join('')}</div></div>`;
  }

  function field(name, label, value, type = 'text', full = false) {
    if (type === 'date') return dateField(name, label, value, full);
    const step = type === 'number' ? ' step="any" inputmode="decimal"' : '';
    return `<div class="field ${full?'full':''}"><label for="f_${name}">${e(label)}</label><input id="f_${name}" data-field="${e(name)}" type="${e(type)}" value="${e(value)}"${step}></div>`;
  }

  function normalizeDateText(value) {
    const digits = String(value || '').replace(/\D/g, '').slice(0, 8);
    if (digits.length <= 4) return digits;
    if (digits.length <= 6) return `${digits.slice(0,4)}-${digits.slice(4)}`;
    return `${digits.slice(0,4)}-${digits.slice(4,6)}-${digits.slice(6,8)}`;
  }

  function validIsoDate(value) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ''));
    if (!m) return value ? null : '';
    const yy = Number(m[1]), mm = Number(m[2]), dd = Number(m[3]);
    if (yy < 1900 || yy > 2199 || mm < 1 || mm > 12) return null;
    const maxDay = new Date(yy, mm, 0).getDate();
    if (dd < 1 || dd > maxDay) return null;
    return `${m[1]}-${m[2]}-${m[3]}`;
  }

  function dateInputHtml(targetId, value) {
    const textValue = normalizeDateText(value);
    return `<div class="date-input-wrap" data-date-target="${e(targetId)}">
      <input id="${e(targetId)}_text" class="date-text-input" type="text" inputmode="numeric" autocomplete="off" maxlength="10" placeholder="YYYY-MM-DD" aria-label="날짜 YYYY-MM-DD" value="${e(textValue)}">
      <button class="date-picker-button" type="button" aria-label="달력에서 날짜 선택" title="달력에서 선택">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 3v3M17 3v3M4.5 9h15M6 5h12a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2Z"/></svg>
      </button>
      <input id="${e(targetId)}_picker" class="native-date-picker" type="date" tabindex="-1" aria-hidden="true" value="${e(validIsoDate(textValue) || '')}">
    </div>`;
  }

  function dateField(name, label, value, full = false) {
    const id = `f_${name}`;
    return `<div class="field ${full?'full':''}"><label for="${e(id)}_text">${e(label)}</label>${dateInputHtml(id, value)}<input id="${e(id)}" data-field="${e(name)}" type="hidden" value="${e(validIsoDate(normalizeDateText(value)) || '')}"><span class="date-error" hidden>날짜를 YYYY-MM-DD 형식으로 확인해주세요.</span></div>`;
  }

  function modalDateField(id, label, value = '') {
    return `<div class="field"><label for="${e(id)}_text">${e(label)}</label>${dateInputHtml(id, value)}<input id="${e(id)}" type="hidden" value="${e(validIsoDate(normalizeDateText(value)) || '')}"><span class="date-error" hidden>날짜를 YYYY-MM-DD 형식으로 확인해주세요.</span></div>`;
  }

  function initDateInputs(root) {
    root.querySelectorAll('.date-input-wrap').forEach(group => {
      if (group.dataset.dateBound === 'true') return;
      group.dataset.dateBound = 'true';
      const targetId = group.dataset.dateTarget;
      const hidden = root.querySelector(`#${CSS.escape(targetId)}`) || document.getElementById(targetId);
      const text = group.querySelector('.date-text-input');
      const picker = group.querySelector('.native-date-picker');
      const button = group.querySelector('.date-picker-button');
      const error = group.parentElement?.querySelector('.date-error');

      const setValidity = (invalid) => {
        group.classList.toggle('date-invalid', invalid);
        if (error) error.hidden = !invalid;
      };

      const syncHidden = (eventName = 'input') => {
        if (!hidden) return;
        const iso = validIsoDate(text.value);
        const invalid = iso === null && text.value.length === 10;
        setValidity(invalid);
        if (iso === null) return;
        if (picker) picker.value = iso;
        if (hidden.value !== iso) {
          hidden.value = iso;
          hidden.dispatchEvent(new Event(eventName, { bubbles:true }));
        }
      };

      text.addEventListener('input', () => {
        const caretAtEnd = text.selectionStart === text.value.length;
        text.value = normalizeDateText(text.value);
        if (caretAtEnd) text.setSelectionRange(text.value.length, text.value.length);
        if (text.value.length < 10) setValidity(false);
        syncHidden('input');
      });
      text.addEventListener('blur', () => {
        if (!text.value) {
          setValidity(false);
          if (hidden && hidden.value !== '') {
            hidden.value = '';
            hidden.dispatchEvent(new Event('change', { bubbles:true }));
          }
          if (picker) picker.value = '';
          return;
        }
        const iso = validIsoDate(text.value);
        setValidity(iso === null);
        if (iso) syncHidden('change');
      });
      text.addEventListener('paste', ev => {
        const raw = ev.clipboardData?.getData('text') || '';
        const digits = raw.replace(/\D/g, '');
        if (digits.length !== 8) return;
        ev.preventDefault();
        text.value = normalizeDateText(digits);
        syncHidden('input');
      });

      picker.addEventListener('change', () => {
        if (!picker.value) return;
        text.value = picker.value;
        setValidity(false);
        syncHidden('change');
        text.focus();
      });
      button.addEventListener('click', () => {
        if (picker.showPicker) picker.showPicker();
        else picker.click();
      });
    });
  }

  function parseMoneyInput(value) {
    const digits = String(value ?? '').replace(/\D/g, '');
    return digits ? Number(digits) : '';
  }

  function formatMoneyInput(value) {
    const n = parseMoneyInput(value);
    return meaningful(n) ? Number(n).toLocaleString('ko-KR') : '';
  }

  function koreanMoney(value) {
    const n = parseMoneyInput(value);
    if (!meaningful(n)) return '';
    if (n === 0) return '금 영원';
    const digits = ['','일','이','삼','사','오','육','칠','팔','구'];
    const smallUnits = ['','십','백','천'];
    const bigUnits = ['','만','억','조','경'];
    const raw = String(Math.trunc(n));
    const groups = [];
    for (let end = raw.length; end > 0; end -= 4) groups.unshift(raw.slice(Math.max(0, end - 4), end));
    const parts = [];
    groups.forEach((group, groupIndex) => {
      const padded = group.padStart(4, '0');
      let chunk = '';
      for (let i = 0; i < 4; i++) {
        const d = Number(padded[i]);
        if (!d) continue;
        chunk += digits[d] + smallUnits[3 - i];
      }
      if (chunk) {
        const bigIndex = groups.length - 1 - groupIndex;
        parts.push(chunk + (bigUnits[bigIndex] || ''));
      }
    });
    return `금 ${parts.join(' ')}원`;
  }

  function moneyInputHtml(id, value, dataField = '') {
    const numeric = parseMoneyInput(value);
    const formatted = meaningful(numeric) ? Number(numeric).toLocaleString('ko-KR') : '';
    const dataAttr = dataField ? ` data-field="${e(dataField)}"` : '';
    return `<input id="${e(id)}"${dataAttr} data-money="true" type="text" inputmode="numeric" autocomplete="off" value="${e(formatted)}"><span class="money-korean">${meaningful(numeric) ? e(koreanMoney(numeric)) : ''}</span>`;
  }

  function moneyField(name, label, value) {
    return `<div class="field"><label for="f_${name}">${e(label)}</label>${moneyInputHtml(`f_${name}`, value, name)}</div>`;
  }

  function initMoneyInputs(root) {
    root.querySelectorAll('input[data-money="true"]').forEach(input => {
      if (input.dataset.moneyBound === 'true') return;
      input.dataset.moneyBound = 'true';
      const hint = input.parentElement.querySelector('.money-korean');
      const refresh = () => {
        const numeric = parseMoneyInput(input.value);
        input.value = meaningful(numeric) ? Number(numeric).toLocaleString('ko-KR') : '';
        if (hint) hint.textContent = meaningful(numeric) ? koreanMoney(numeric) : '';
      };
      input.addEventListener('input', refresh);
      input.addEventListener('blur', refresh);
      refresh();
    });
  }

  function selectField(name, label, value, options) {
    const extra = value && !options.includes(value) ? [value] : [];
    return `<div class="field"><label for="f_${name}">${e(label)}</label><select id="f_${name}" data-field="${e(name)}">${[...options,...extra].map(o=>`<option value="${e(o)}" ${o===value?'selected':''}>${e(o || '선택')}</option>`).join('')}</select></div>`;
  }

  function textareaField(name, label, value) {
    return `<div class="field full"><label for="f_${name}">${e(label)}</label><textarea id="f_${name}" data-field="${e(name)}">${e(value)}</textarea></div>`;
  }

  function progressRow(label, done) {
    return `<div class="progress-item"><span>${done?'✓':'○'} ${e(label)}</span><span>${done?'확인':'대기'}</span></div>`;
  }

  function onProjectInput(ev) {
    const p = currentProject();
    if (!p) return;
    const fieldName = ev.target.dataset.field;
    let value = ev.target.value;
    if (ev.target.dataset.money) value = parseMoneyInput(value);
    else if (ev.target.type === 'number' && value !== '') value = Number(value);
    p[fieldName] = value;
    if (fieldName === 'contractDate' && value && !p.fiscalYear) p.fiscalYear = value.slice(0,4);
    p.updatedAt = new Date().toISOString();
    scheduleSave(p);
  }

  function scheduleSave(p) {
    const token = ++state.saveToken;
    const saveState = document.getElementById('saveState');
    if (saveState) saveState.innerHTML = '<span class="pulse" style="background:#b5bdc8"></span><span>저장 중…</span>';
    clearTimeout(state.autosaveTimer);
    state.autosaveTimer = setTimeout(async () => {
      try {
        await DB.put('projects', p);
        if (token === state.saveToken) {
          const s = document.getElementById('saveState');
          if (s) s.innerHTML = '<span class="pulse"></span><span>이 기기에 저장됨</span>';
        }
      } catch (err) {
        showToast(`자동저장 실패: ${err.message}`, 'danger');
      }
    }, 350);
  }

  async function applyVendorToCurrentProject(ev) {
    const p = currentProject();
    const vendor = state.vendors.find(v => v.id === ev.target.value);
    if (!p) return;
    if (!vendor) { p.vendorId = ''; scheduleSave(p); return; }
    Object.assign(p, {
      vendorId: vendor.id,
      vendorName: vendor.name,
      representative: vendor.representative,
      businessNumber: vendor.businessNumber,
      vendorAddress: vendor.address,
      vendorPhone: vendor.phone,
      licenseType: vendor.licenseType || p.licenseType,
      workType: p.workType || vendor.workType,
      updatedAt: new Date().toISOString()
    });
    await DB.put('projects', p);
    renderProjectDetail();
    showToast(`${vendor.name} 정보를 불러왔습니다.`);
  }

  async function saveCurrentVendor() {
    const p = currentProject();
    if (!p?.vendorName?.trim()) { showToast('업체명을 먼저 입력해주세요.', 'warn'); return; }
    let vendor = p.vendorId ? state.vendors.find(v => v.id === p.vendorId) : null;
    if (!vendor && p.businessNumber) vendor = state.vendors.find(v => normalizeText(v.businessNumber) === normalizeText(p.businessNumber));
    if (!vendor) vendor = state.vendors.find(v => normalizeText(v.name) === normalizeText(p.vendorName));
    const next = DB.createVendor({
      ...(vendor || {}),
      id: vendor?.id,
      name: p.vendorName,
      representative: p.representative,
      businessNumber: p.businessNumber,
      phone: p.vendorPhone,
      address: p.vendorAddress,
      workType: p.workType,
      licenseType: p.licenseType,
      updatedAt: new Date().toISOString()
    });
    await DB.put('vendors', next);
    p.vendorId = next.id;
    p.updatedAt = new Date().toISOString();
    await DB.put('projects', p);
    await loadState();
    state.currentProjectId = p.id;
    renderProjectDetail();
    showToast('업체 보관함에 반영했습니다.');
  }

  function openNewProjectModal() {
    const options = state.vendors.map(v => `<option value="${e(v.id)}">${e(v.name)}${v.businessNumber ? ` · ${e(v.businessNumber)}` : ''}</option>`).join('');
    openModal({
      eyebrow: '새 공사 직접등록',
      title: '공사명만으로도 시작할 수 있어요',
      body: `<div class="notice">지금 확정된 정보만 입력하세요. 착공일·준공기한·준공일·지출일은 공사를 만든 뒤 필요한 시점에 추가하면 됩니다.</div>
        <div class="modal-grid new-project-grid" style="margin-top:16px">
          <div class="field full"><label>공사명 <span class="required-mark">필수</span></label><input id="newProjectName" autocomplete="off" placeholder="예: 체육관 환경개선공사"></div>
          <div class="field"><label>공종</label><select id="newWorkType"><option value="">나중에 입력</option><option>건축공사</option><option>전기공사</option><option>통신공사</option><option>소방공사</option><option>기계설비공사</option><option>토목공사</option><option>기타</option></select></div>
          <div class="field"><label>업체</label><select id="newVendorId"><option value="">나중에 입력 / 새 업체</option>${options}</select><span class="hint">저장된 업체를 고르면 대표자·사업자번호·주소까지 자동으로 연결됩니다.</span></div>
          <div class="field"><label for="newContractAmount">계약금액</label>${moneyInputHtml('newContractAmount', '')}</div>
          <div class="field"><label>계약방법</label><select id="newContractMethod"><option value="">나중에 입력</option><option>1인수의</option><option>2인이상수의</option><option>제한경쟁</option><option>일반경쟁</option><option>조달계약</option><option>기타</option></select></div>
          ${modalDateField('newContractDate','계약일')}
        </div>
        <div class="new-project-tip"><strong>입력은 여기까지.</strong><span>공사를 만든 뒤 에듀파인 자료를 불러오면 계약번호·예정가격·착공·준공·지출 정보를 추가로 채울 수 있습니다.</span></div>`,
      actions: `<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="createProjectBtn">공사 만들기</button>`
    });
    initDateInputs(modalBody);
    initMoneyInputs(modalBody);
    modalBody.querySelector('#newProjectName')?.focus();
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#createProjectBtn').addEventListener('click', createProjectFromModal);
  }

  async function createProjectFromModal() {
    const projectName = modalBody.querySelector('#newProjectName').value.trim();
    if (!projectName) { showToast('공사명을 입력해주세요.', 'warn'); modalBody.querySelector('#newProjectName').focus(); return; }
    const vendorId = modalBody.querySelector('#newVendorId').value;
    const vendor = state.vendors.find(v => v.id === vendorId);
    const contractDate = modalBody.querySelector('#newContractDate').value;
    const p = DB.createProject({
      projectName,
      fiscalYear: contractDate ? contractDate.slice(0,4) : String(new Date().getFullYear()),
      workType: modalBody.querySelector('#newWorkType').value,
      contractMethod: modalBody.querySelector('#newContractMethod').value,
      currentContractAmount: parseMoneyInput(modalBody.querySelector('#newContractAmount').value),
      contractDate,
      vendorId: vendor?.id || '',
      vendorName: vendor?.name || '', representative: vendor?.representative || '', businessNumber: vendor?.businessNumber || '',
      vendorAddress: vendor?.address || '', vendorPhone: vendor?.phone || '', licenseType: vendor?.licenseType || ''
    });
    await DB.put('projects', p);
    state.projects.unshift(p);
    closeModal();
    openProject(p.id);
    showToast('새 공사를 만들었습니다. 필요한 정보만 이어서 추가하면 됩니다.');
  }

  async function confirmDeleteProject() {
    const p = currentProject();
    if (!p) return;
    openModal({
      eyebrow: '공사 삭제', title: '이 공사를 삭제할까요?',
      body: `<div class="notice danger"><strong>${e(p.projectName || '이름 없는 공사')}</strong><br>이 기기에 저장된 해당 공사정보가 삭제됩니다. 다른 공사와 업체 보관함은 유지됩니다.</div>`,
      actions: `<button class="button secondary" type="button" data-modal-close>취소</button><button class="button danger" type="button" id="deleteConfirmBtn">삭제</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#deleteConfirmBtn').addEventListener('click', async () => {
      await DB.remove('projects', p.id);
      state.projects = state.projects.filter(x => x.id !== p.id);
      state.currentProjectId = null;
      closeModal(); renderDashboard(); showToast('공사를 삭제했습니다.');
    });
  }

  function findProjectMatch(incoming) {
    if (incoming.contractNumber) {
      const byNumber = state.projects.find(p => p.contractNumber && normalizeText(p.contractNumber) === normalizeText(incoming.contractNumber));
      if (byNumber) return byNumber;
    }
    let best = null, bestScore = 0;
    for (const p of state.projects) {
      let score = 0;
      if (incoming.projectName && p.projectName && normalizeText(incoming.projectName) === normalizeText(p.projectName)) score += 4;
      if (incoming.vendorName && p.vendorName && normalizeText(incoming.vendorName) === normalizeText(p.vendorName)) score += 2;
      if (meaningful(incoming.currentContractAmount) && meaningful(p.currentContractAmount) && Number(incoming.currentContractAmount) === Number(p.currentContractAmount)) score += 2;
      if (incoming.contractDate && p.contractDate && incoming.contractDate === p.contractDate) score += 2;
      if (score > bestScore) { best = p; bestScore = score; }
    }
    return bestScore >= 6 ? best : null;
  }

  function analyzeImport(parsed) {
    return parsed.projects.map((incoming, index) => {
      const match = findProjectMatch(incoming);
      const additions = [], conflicts = [];
      if (match) {
        for (const [field, label] of FIELDS_FOR_IMPORT) {
          const iv = incoming[field], cv = match[field];
          if (!meaningful(iv)) continue;
          if (!meaningful(cv)) additions.push({ field, label, value: iv });
          else if (!sameValue(cv, iv, field)) conflicts.push({ field, label, current: cv, incoming: iv });
        }
      }
      return { index, incoming, match, additions, conflicts };
    });
  }

  async function handleExcelFile(file) {
    if (!file) return;
    const mode = state.importMode || 'auto';
    state.importMode = 'auto';
    try {
      showToast('엑셀을 브라우저에서 분석하고 있습니다.');
      const parsed = await Excel.parseImport(file);
      if (!parsed.projects.length) throw new Error('불러올 공사 행을 찾지 못했습니다.');
      if (mode === 'history' && parsed.type === 'edufine') {
        throw new Error('이 파일은 K-에듀파인 자료관리목록으로 보입니다. 「에듀파인으로 정보 업데이트」에서 선택해주세요.');
      }
      if (mode === 'edufine' && parsed.type !== 'edufine') {
        throw new Error('이 파일은 학교 공사 이력 현황으로 보입니다. 「기존 공사이력 불러오기」에서 선택해주세요.');
      }
      const analysis = analyzeImport(parsed);
      openImportPreview(parsed, analysis, mode);
    } catch (err) {
      showToast(err.message || '엑셀을 읽지 못했습니다.', 'danger');
    } finally {
      excelFileInput.value = '';
    }
  }

  function openImportPreview(parsed, analysis, mode = 'auto') {
    const newCount = analysis.filter(x => !x.match).length;
    const updateCount = analysis.filter(x => x.match && x.additions.length && !x.conflicts.length).length;
    const conflictCount = analysis.filter(x => x.conflicts.length).length;
    const unchangedCount = analysis.filter(x => x.match && !x.additions.length && !x.conflicts.length).length;
    const isEdufine = parsed.type === 'edufine';
    const title = isEdufine ? '에듀파인 정보 업데이트 결과' : '기존 공사이력 불러오기 결과';
    const guide = isEdufine
      ? '기존 공사와 계약번호·공사명·업체·금액·계약일을 비교했습니다. 빈 값은 보완하고, 서로 다른 값은 선택한 경우에만 바꿉니다.'
      : '공사이력의 각 행을 기존 공사와 비교했습니다. 처음 보는 공사는 새로 만들고, 기존 공사는 부족한 값만 보완합니다.';
    const previews = analysis.map(a => {
      const label = !a.match ? '새 공사' : a.conflicts.length ? '확인 필요' : a.additions.length ? '기존 공사 보완' : '변경 없음';
      return `<div class="preview-item" data-import-index="${a.index}">
        <div class="preview-top"><div><strong>${e(a.incoming.projectName || '이름 없는 공사')}</strong><p>${e(a.incoming.vendorName || '업체 미확인')} · ${e(formatMoney(a.incoming.currentContractAmount))}${a.incoming.contractNumber ? ` · ${e(a.incoming.contractNumber)}` : ''}</p></div><span class="preview-badge ${a.conflicts.length?'conflict':a.match&&!a.additions.length?'muted-badge':''}">${label}</span></div>
        ${a.match ? `<p style="margin-top:7px">기존: ${e(a.match.projectName)}${a.additions.length ? ` · 새로 채울 값 ${a.additions.length}개` : ''}</p>` : '<p style="margin-top:7px">현재 목록에 같은 공사가 없어 새 공사로 등록됩니다.</p>'}
        ${a.conflicts.length ? `<div class="conflict-box">${a.conflicts.map(c => conflictHtml(a.index,c)).join('')}</div>` : ''}
      </div>`;
    }).join('');
    openModal({
      eyebrow: parsed.label,
      title,
      body: `<div class="notice"><strong>${e(guide)}</strong><br>파일은 서버로 업로드하지 않고 이 브라우저에서만 읽었습니다.${parsed.ignored ? ` 공사가 아닌 계약 ${parsed.ignored}건은 제외했습니다.` : ''}</div>
        <div class="import-summary four"><div class="import-stat"><strong>${newCount}</strong><span>새 공사</span></div><div class="import-stat"><strong>${updateCount}</strong><span>자동 보완</span></div><div class="import-stat"><strong>${conflictCount}</strong><span>확인 필요</span></div><div class="import-stat"><strong>${unchangedCount}</strong><span>변경 없음</span></div></div>
        <div class="import-result-note">총 ${parsed.projects.length}건 · ${newCount + updateCount + conflictCount}건에 반영할 내용이 있습니다.</div>
        <div class="preview-list">${previews}</div>`,
      actions: `<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="applyImportBtn">${newCount + updateCount + conflictCount ? '반영하기' : '변경 없음'}</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    const applyBtn = modalActions.querySelector('#applyImportBtn');
    applyBtn.disabled = !(newCount + updateCount + conflictCount);
    applyBtn.addEventListener('click', () => applyImport(parsed, analysis));
  }

  function conflictHtml(index, c) {
    const current = MONEY_FIELDS.has(c.field) ? formatMoney(c.current) : c.current;
    const incoming = MONEY_FIELDS.has(c.field) ? formatMoney(c.incoming) : c.incoming;
    return `<div class="conflict-row"><label>${e(c.label)}</label>
      <label class="conflict-choice"><input type="radio" name="conf_${index}_${e(c.field)}" value="keep" checked>현재 유지<br><span class="muted">${e(current)}</span></label>
      <label class="conflict-choice"><input type="radio" name="conf_${index}_${e(c.field)}" value="incoming">엑셀 적용<br><span class="muted">${e(incoming)}</span></label></div>`;
  }

  async function applyImport(parsed, analysis) {
    let created = 0, updated = 0, unchanged = 0;
    for (const a of analysis) {
      let p;
      if (!a.match) {
        p = DB.createProject(a.incoming);
        await DB.put('projects', p);
        created++;
      } else {
        p = a.match;
        let changed = false;
        for (const item of a.additions) { p[item.field] = item.value; changed = true; }
        for (const c of a.conflicts) {
          const choice = modalBody.querySelector(`input[name="conf_${a.index}_${c.field}"]:checked`)?.value;
          if (choice === 'incoming') { p[c.field] = c.incoming; changed = true; }
        }
        if (changed) {
          p.sourceUpdatedAt = new Date().toISOString();
          if (parsed.type === 'edufine') {
            if (p.source === 'manual') p.source = 'manual+edufine';
            else if (p.source === 'audit-excel') p.source = 'audit+edufine';
            else if (!String(p.source || '').includes('edufine')) p.source = 'edufine';
          }
          p.updatedAt = new Date().toISOString();
          await DB.put('projects', p);
          updated++;
        } else unchanged++;
      }
      await upsertVendorFromProject(p);
    }
    await loadState();
    closeModal(); renderDashboard();
    showToast(`반영 완료 · 새 공사 ${created}건 · 기존 공사 ${updated}건 보완${unchanged ? ` · 변경 없음 ${unchanged}건` : ''}`);
  }

  async function upsertVendorFromProject(p) {
    if (!p?.vendorName) return;
    let vendor = p.businessNumber ? state.vendors.find(v => normalizeText(v.businessNumber) === normalizeText(p.businessNumber)) : null;
    if (!vendor) vendor = state.vendors.find(v => normalizeText(v.name) === normalizeText(p.vendorName));
    const next = DB.createVendor({
      ...(vendor || {}), id: vendor?.id,
      name: p.vendorName, representative: p.representative, businessNumber: p.businessNumber,
      phone: p.vendorPhone, address: p.vendorAddress, workType: p.workType, licenseType: p.licenseType,
      updatedAt: new Date().toISOString()
    });
    await DB.put('vendors', next);
    if (!p.vendorId) { p.vendorId = next.id; await DB.put('projects', p); }
    if (!state.vendors.some(v => v.id === next.id)) state.vendors.push(next);
  }

  function openExportAuditModal() {
    if (!state.projects.length) { showToast('내보낼 공사가 없습니다.', 'warn'); return; }
    const years = [...new Set(state.projects.map(p=>p.fiscalYear).filter(Boolean))].sort().reverse();
    openModal({
      eyebrow: '학교 공사 이력 현황', title: '학교 양식으로 내보내기',
      body: `<div class="notice">새로 정리한 학교 공사 이력 현황 양식의 열과 서식을 기준으로 생성합니다.</div><div class="field" style="margin-top:16px"><label>범위</label><select id="auditExportYear"><option value="all">전체 연도 (${state.projects.length}건)</option>${years.map(y=>`<option value="${e(y)}">${e(y)}회계연도 (${state.projects.filter(p=>p.fiscalYear===y).length}건)</option>`).join('')}</select></div>`,
      actions: `<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="auditExportConfirm">엑셀 받기</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#auditExportConfirm').addEventListener('click', async () => {
      const year = modalBody.querySelector('#auditExportYear').value;
      const list = year === 'all' ? state.projects : state.projects.filter(p => p.fiscalYear === year);
      try { await Excel.exportAuditWorkbook(list, { year: year === 'all' ? '' : year }); closeModal(); showToast(`${list.length}건의 공사이력 엑셀을 만들었습니다.`); }
      catch (err) { showToast(err.message, 'danger'); }
    });
  }

  async function backupAll() {
    const data = await DB.exportBackup();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type:'application/json;charset=utf-8' });
    const d = new Date();
    const date = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
    Excel.downloadBlob(blob, `공사정보허브_전체백업_${date}.json`);
    showToast(`공사 ${data.projects.length}건 · 업체 ${data.vendors.length}건을 백업했습니다.`);
  }

  async function handleBackupFile(file) {
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (data.app !== 'construction-info-hub') throw new Error('공사정보 허브 백업파일이 아닙니다.');
      openModal({
        eyebrow:'전체 복원', title:'기존 데이터를 백업파일로 바꿀까요?',
        body:`<div class="notice warn">현재 이 브라우저의 공사와 업체정보를 지우고, 백업파일의 내용으로 교체합니다.</div><div class="import-summary"><div class="import-stat"><strong>${data.projects?.length||0}</strong><span>공사</span></div><div class="import-stat"><strong>${data.vendors?.length||0}</strong><span>업체</span></div><div class="import-stat"><strong>${e(data.version||1)}</strong><span>백업 버전</span></div></div>`,
        actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button danger" type="button" id="restoreConfirm">전체 복원</button>`
      });
      modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
      modalActions.querySelector('#restoreConfirm').addEventListener('click', async()=>{
        await DB.importBackup(data); await loadState(); state.currentProjectId=null; closeModal(); renderDashboard(); showToast('백업을 복원했습니다.');
      });
    } catch(err) { showToast(err.message || '백업파일을 읽지 못했습니다.', 'danger'); }
    finally { backupFileInput.value=''; }
  }

  function openSchoolModal() {
    const s = state.school || {};
    openModal({
      eyebrow:'학교 기본정보', title:'공사마다 반복하지 않는 정보',
      body:`<div class="notice">학교정보는 이 브라우저에 한 번 저장해두고 향후 공사서류 생성 시 기본값으로 사용합니다.</div><div class="modal-grid" style="margin-top:16px">
        ${modalField('schoolName','기관명',s.name)}
        ${modalSelect('schoolType','공·사립',s.type,['','공립','사립'])}
        ${modalField('schoolAddress','기관 주소',s.address,true)}
        ${modalField('schoolPhone','대표전화',s.phone)}
        ${modalField('schoolSupervisor','공사감독 기본값',s.supervisor)}
        ${modalField('schoolInspector','검사자 기본값',s.inspector)}
        ${modalField('schoolWitness','준공검사 입회자 기본값',s.witness)}
      </div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="schoolSave">저장</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#schoolSave').addEventListener('click', async()=>{
      const value = { name:v('#schoolName'), type:v('#schoolType'), address:v('#schoolAddress'), phone:v('#schoolPhone'), supervisor:v('#schoolSupervisor'), inspector:v('#schoolInspector'), witness:v('#schoolWitness') };
      await DB.put('settings',{key:'school',value}); state.school=value; closeModal(); render(); showToast('학교 기본정보를 저장했습니다.');
    });
  }

  function modalField(id,label,value,full=false){return `<div class="field ${full?'full':''}"><label>${e(label)}</label><input id="${e(id)}" value="${e(value||'')}"></div>`;}
  function modalSelect(id,label,value,options){return `<div class="field"><label>${e(label)}</label><select id="${e(id)}">${options.map(o=>`<option value="${e(o)}" ${o===value?'selected':''}>${e(o||'선택')}</option>`).join('')}</select></div>`;}
  function v(sel){return modalBody.querySelector(sel)?.value.trim()||'';}

  function openVendorLibrary() {
    const rows = state.vendors.length ? state.vendors.map(vendorRowHtml).join('') : '<div class="notice">아직 저장된 업체가 없습니다. 공사 상세에서 업체정보를 보관함에 반영하거나 여기서 추가할 수 있습니다.</div>';
    openModal({
      eyebrow:'업체 보관함', title:`저장된 업체 ${state.vendors.length}개`,
      body:`<div class="vendor-list">${rows}</div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>닫기</button><button class="button primary" type="button" id="addVendorBtn">+ 업체 추가</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#addVendorBtn').addEventListener('click', ()=>openVendorEdit(null));
    modalBody.querySelectorAll('[data-vendor-edit]').forEach(btn=>btn.addEventListener('click',()=>openVendorEdit(btn.dataset.vendorEdit)));
    modalBody.querySelectorAll('[data-vendor-delete]').forEach(btn=>btn.addEventListener('click',()=>deleteVendor(btn.dataset.vendorDelete)));
  }

  function vendorRowHtml(v) {
    return `<div class="vendor-row"><div><strong>${e(v.name)}</strong><small>${e([v.representative,v.businessNumber,v.phone].filter(Boolean).join(' · ') || '추가정보 없음')}</small></div><div class="actions"><button class="button secondary small" type="button" data-vendor-edit="${e(v.id)}">수정</button><button class="button danger small" type="button" data-vendor-delete="${e(v.id)}">삭제</button></div></div>`;
  }

  function openVendorEdit(id) {
    const vendor = id ? state.vendors.find(x=>x.id===id) : null;
    openModal({
      eyebrow:'업체 보관함', title:vendor?'업체정보 수정':'업체 추가',
      body:`<div class="modal-grid">${modalField('vendorName','업체명',vendor?.name)}${modalField('vendorRep','대표자',vendor?.representative)}${modalField('vendorBiz','사업자등록번호',vendor?.businessNumber)}${modalField('vendorPhone','대표전화',vendor?.phone)}${modalField('vendorAddress','사업장 주소',vendor?.address,true)}${modalField('vendorLicense','등록면허 / 업종',vendor?.licenseType,true)}</div>`,
      actions:`<button class="button secondary" type="button" id="vendorBack">뒤로</button><button class="button primary" type="button" id="vendorSave">저장</button>`
    });
    modalActions.querySelector('#vendorBack').addEventListener('click',openVendorLibrary);
    modalActions.querySelector('#vendorSave').addEventListener('click',async()=>{
      const name=v('#vendorName'); if(!name){showToast('업체명을 입력해주세요.','warn');return;}
      const next=DB.createVendor({...(vendor||{}),id:vendor?.id,name,representative:v('#vendorRep'),businessNumber:Excel.normalizeBusinessNumber(v('#vendorBiz')),phone:v('#vendorPhone'),address:v('#vendorAddress'),licenseType:v('#vendorLicense'),updatedAt:new Date().toISOString()});
      await DB.put('vendors',next); await loadState(); openVendorLibrary(); showToast('업체정보를 저장했습니다.');
    });
  }

  async function deleteVendor(id) {
    const vendor=state.vendors.find(v=>v.id===id); if(!vendor)return;
    if(state.projects.some(p=>p.vendorId===id)){showToast('현재 공사에서 사용 중인 업체입니다. 공사의 업체 연결을 먼저 변경해주세요.','warn');return;}
    await DB.remove('vendors',id); await loadState(); openVendorLibrary(); showToast(`${vendor.name}을 업체 보관함에서 삭제했습니다.`);
  }

  function openHelp() {
    openModal({
      eyebrow:'도움말', title:'v0.2 사용 흐름',
      body:`<div class="notice"><strong>핵심 원칙</strong><br>같은 공사정보는 한 번 입력하고 다시 입력하지 않습니다.</div>
      <div style="display:grid;gap:16px;margin-top:18px;font-size:14px">
        <div><strong>1. 공사를 여러 건 저장</strong><p class="muted">전기·건축·체육관 공사를 동시에 등록해도 각 공사는 독립적으로 자동저장됩니다.</p></div>
        <div><strong>2. 기존 공사이력 재사용</strong><p class="muted">학교 공사 이력 현황.xlsx를 불러오면 여러 공사를 한꺼번에 등록하고 기존 공사의 빈 정보를 보완합니다.</p></div><div><strong>3. 에듀파인으로 업데이트</strong><p class="muted">자료관리목록.xlsx를 다시 내려받아 올리면 계약·준공·지출 단계에서 새로 생긴 값만 기존 공사에 보완합니다. 다른 값은 자동 덮어쓰지 않습니다.</p></div>
        <div><strong>4. 업체 재사용</strong><p class="muted">업체명·대표자·사업자번호·주소·전화·면허를 업체 보관함에 저장해 다음 공사에서 다시 고를 수 있습니다.</p></div>
        <div><strong>5. 인수인계</strong><p class="muted">전체 백업(JSON)은 앱 복원용이고, 공사이력 엑셀은 감사·업무용 결과물입니다.</p></div>
        <div><strong>보안</strong><p class="muted">공사정보와 엑셀 내용은 서버로 전송하지 않습니다. 브라우저 저장소에 남으므로 공용 Windows 계정에서는 PC 접근통제와 정기 백업이 필요합니다.</p></div>
      </div>`,
      actions:`<button class="button primary" type="button" data-modal-close>확인</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click',closeModal);
  }

  function wireGlobal() {
    document.getElementById('goHomeBtn').addEventListener('click',()=>{state.currentProjectId=null;renderDashboard();});
    document.getElementById('newProjectBtn').addEventListener('click',openNewProjectModal);
    document.getElementById('importBtn').addEventListener('click',()=>openExcelPicker('auto'));
    document.getElementById('exportAuditBtn').addEventListener('click',openExportAuditModal);
    document.getElementById('moreBtn').addEventListener('click',(ev)=>{ev.stopPropagation();moreMenu.hidden=!moreMenu.hidden;ev.currentTarget.setAttribute('aria-expanded',String(!moreMenu.hidden));});
    document.addEventListener('click',(ev)=>{if(!moreMenu.hidden&&!moreMenu.contains(ev.target)&&ev.target.id!=='moreBtn'){moreMenu.hidden=true;document.getElementById('moreBtn').setAttribute('aria-expanded','false');}});
    moreMenu.addEventListener('click',ev=>{
      const action=ev.target.dataset.action;if(!action)return;moreMenu.hidden=true;
      if(action==='import')openExcelPicker('auto'); if(action==='export')openExportAuditModal(); if(action==='vendors')openVendorLibrary(); if(action==='school')openSchoolModal(); if(action==='backup')backupAll(); if(action==='restore')backupFileInput.click(); if(action==='help')openHelp();
    });
    excelFileInput.addEventListener('change',()=>handleExcelFile(excelFileInput.files[0]));
    backupFileInput.addEventListener('change',()=>handleBackupFile(backupFileInput.files[0]));
    document.getElementById('modalCloseX').addEventListener('click', closeModal);
    modalForm.addEventListener('submit',ev=>ev.preventDefault());
  }

  async function init() {
    wireGlobal();
    try { await loadState(); renderDashboard(); }
    catch(err){ main.innerHTML=`<div class="empty-state"><h3>브라우저 저장소를 열지 못했습니다.</h3><p>${e(err.message)}</p></div>`; }
  }

  init();
})();
