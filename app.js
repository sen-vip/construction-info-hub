(() => {
  'use strict';

  const DB = globalThis.ConstructionDB;
  const Excel = globalThis.ConstructionExcel;
  const Documents = globalThis.ConstructionDocuments;
  const ReferenceData = globalThis.ConstructionReferenceData;
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
    payouts: [],
    school: {},
    currentProjectId: null,
    filter: 'active',
    search: '',
    autosaveTimer: null,
    saveToken: 0,
    importMode: 'auto',
    detailTab: 'info',
    selectedDocuments: new Set(),
    selectionProjectId: null,
    activeWarrantyInspectionId: null
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

  const MONEY_FIELDS = new Set(['estimatedPrice','originalContractAmount','currentContractAmount','settlementAmount','priorPaymentAmount','deductionAmount','paymentAmount','contractSecurityAmount','defectSecurityAmount']);

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


  function parseIsoDate(value) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ''));
    if (!m) return null;
    return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
  }

  function isoFromDate(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth()+1).padStart(2,'0');
    const d = String(date.getDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
  }

  function addYearsMinusDay(startIso, years) {
    const start = parseIsoDate(startIso);
    const y = Number(years);
    if (!start || !Number.isFinite(y) || y <= 0) return '';
    const result = new Date(start.getFullYear()+y, start.getMonth(), start.getDate());
    result.setDate(result.getDate()-1);
    return isoFromDate(result);
  }

  function warrantyEndDates(p) {
    const items = Array.isArray(p?.warrantyItems) ? p.warrantyItems.filter(x=>x?.endDate) : [];
    if (items.length) return items.map(x=>x.endDate).filter(Boolean).sort();
    return p?.defectEndDate ? [p.defectEndDate] : [];
  }

  function warrantyStatusText(p) {
    const dates = warrantyEndDates(p);
    if (!dates.length) return '';
    const finalDate = dates[dates.length-1];
    const end = parseIsoDate(finalDate);
    const prefix = dates.length > 1 ? `${dates.length}개 공종 · 최종 만료` : '하자종료';
    if (!end) return `${prefix} ${formatDate(finalDate)}`;
    const today = new Date(); today.setHours(0,0,0,0); end.setHours(0,0,0,0);
    const days = Math.ceil((end-today)/86400000);
    if (days >= 0 && days <= 30) return `${prefix} D-${days}`;
    if (days < 0) return `${dates.length > 1 ? `${dates.length}개 공종 · ` : ''}하자기간 종료 ${formatDate(finalDate)}`;
    return `${prefix} ${formatDate(finalDate)}`;
  }

  function statusOf(p) {
    const inspected = !!(p.completionInspectionDate || p.completionInspectionRecordDate);
    if (p.paymentDate) return { key:'done', label:'완료', cls:'done' };
    if (p.actualCompletionDate && inspected) return { key:'payment_wait', label:'지출 대기', cls:'wait' };
    if (p.actualCompletionDate) return { key:'inspection_wait', label:'준공검사 대기', cls:'wait' };
    if (p.startDate) return { key:'active', label:'공사 진행중', cls:'' };
    if (p.contractDate) return { key:'start_wait', label:'착공 대기', cls:'' };
    return { key:'contract_prep', label:'계약 준비', cls:'wait' };
  }

  function missingFor(p) {
    const s = statusOf(p).key;
    const list = [];
    if (!p.projectName) list.push('공사명');
    if (s === 'contract_prep') {
      if (!p.vendorName) list.push('업체');
      if (!meaningful(p.currentContractAmount)) list.push('계약금액');
      if (!p.contractDate) list.push('계약일');
    }
    if (s === 'start_wait') {
      if (!p.startDate) list.push('착공일');
      if (!p.completionDueDate) list.push('준공기한');
    }
    if (s === 'active' && !p.completionDueDate) list.push('준공기한');
    if (s === 'inspection_wait' && !p.completionInspectionDate) list.push('검사·검수일');
    if (s === 'payment_wait') {
      if (!p.taxInvoiceDate) list.push('세금계산서일');
      if (!p.paymentDate) list.push('지출일');
      if (!meaningful(p.paymentAmount)) list.push('지출금액');
    }
    return list;
  }

  function workflowOf(p) {
    const current = statusOf(p).key;
    const steps = [
      { key:'contract', label:'계약', done:!!p.contractDate, active:current==='contract_prep', stateText:'지금 입력', summary:p.contractDate ? `계약일 ${formatDate(p.contractDate)}` : '계약정보 입력' },
      { key:'start', label:'착공', done:!!p.startDate, active:current==='start_wait', stateText:'지금 입력', summary:p.startDate ? `착공일 ${formatDate(p.startDate)}` : '착공정보 대기' },
      { key:'completion', label:'준공', done:!!p.actualCompletionDate && !!(p.completionInspectionDate || p.completionInspectionRecordDate || p.paymentDate), active:['active','inspection_wait'].includes(current), stateText:current==='active'?'공사 진행 중':'지금 입력', summary:p.actualCompletionDate ? `준공일 ${formatDate(p.actualCompletionDate)}` : '공사 진행 중' },
      { key:'payment', label:'지출', done:!!p.paymentDate, active:current==='payment_wait', stateText:'지금 입력', summary:p.paymentDate ? `지출일 ${formatDate(p.paymentDate)}` : '준공 후 입력' },
      { key:'defect', label:'하자', done:!!p.defectStartDate && !!p.defectEndDate, active:!!p.paymentDate && !(p.defectStartDate && p.defectEndDate), stateText:'확인', summary:p.defectEndDate ? `종료 ${formatDate(p.defectEndDate)}` : '해당 시 관리' }
    ];
    return steps.map(step => ({ ...step, future: !step.done && !step.active }));
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

  function openModal({ title, eyebrow = '', body = '', actions = '', wide = false }) {
    modal.classList.toggle('wide-modal', !!wide);
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

  function normalizeProjectData(project) {
    const p = project || {};
    let changed = false;
    if (!p.delayPenaltyRate || ['1.3‰','1.3/1000','0.13%','0.0013'].includes(String(p.delayPenaltyRate).trim())) { p.delayPenaltyRate = '0.5‰'; changed = true; }
    if (!Array.isArray(p.warrantyItems)) {
      p.warrantyItems = [];
      if (meaningful(p.defectPeriodYears) || p.defectStartDate || p.defectEndDate) {
        p.warrantyItems.push({
          id: crypto.randomUUID ? crypto.randomUUID() : `legacy-${Date.now()}-${Math.random().toString(36).slice(2,8)}`,
          category:'기존 입력', subcategory:p.workType || '기존 하자정보',
          years:Number(p.defectPeriodYears || 0) || '', startDate:p.defectStartDate || '', endDate:p.defectEndDate || '',
          recommendedYears:'', recommendationApplied:false, manuallyModified:true, note:'v0.4.1 이전 단일 하자정보에서 자동 이관'
        });
      }
      changed = true;
    }
    if (!p.workCharacteristics || typeof p.workCharacteristics !== 'object') { p.workCharacteristics = {fall:false,electrical:false,confined:false,industrial:false}; changed = true; }
    else p.workCharacteristics = {fall:false,electrical:false,confined:false,industrial:false,...p.workCharacteristics};
    if (!p.safetyChecklists || typeof p.safetyChecklists !== 'object') { p.safetyChecklists = {}; changed = true; }
    if (!Array.isArray(p.selectedDocuments)) { p.selectedDocuments = []; changed = true; }
    return { project:p, changed };
  }

  async function loadState() {
    await DB.openDb();
    const [projects, vendors, payouts, school] = await Promise.all([
      DB.getAll('projects'), DB.getAll('vendors'), DB.getAll('payouts'), DB.get('settings', 'school')
    ]);
    const migrated = projects.map(normalizeProjectData);
    state.projects = migrated.map(x=>x.project).sort((a,b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
    const changed = migrated.filter(x=>x.changed).map(x=>x.project);
    if (changed.length) await Promise.all(changed.map(project => DB.put('projects', project)));
    state.vendors = vendors.sort((a,b) => (a.name || '').localeCompare(b.name || '', 'ko'));
    state.payouts = payouts || [];
    state.school = school?.value || {};
  }

  function render() {
    if (state.currentProjectId) renderProjectDetail();
    else renderDashboard();
  }

  function renderDashboard() {
    const active = state.projects.filter(p => statusOf(p).key !== 'done').length;
    const waiting = state.projects.filter(p => ['inspection_wait','payment_wait'].includes(statusOf(p).key)).length;
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
          ${status.key==='done' && p.defectEndDate ? `<div class="warranty-row-note">${e(warrantyStatusText(p))}</div>` : ''}
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
    state.detailTab = 'info';
    window.scrollTo({ top: 0, behavior: 'instant' });
    render();
  }

  function currentProject() { return state.projects.find(p => p.id === state.currentProjectId); }

  function renderProjectDetail() {
    const p = currentProject();
    if (!p) { state.currentProjectId = null; renderDashboard(); return; }
    const status = statusOf(p);
    const missing = missingFor(p);
    const workflow = workflowOf(p);
    const step = Object.fromEntries(workflow.map(x => [x.key, x]));
    const currentOpen = {
      basic: status.key === 'contract_prep',
      contract: status.key === 'contract_prep',
      start: status.key === 'start_wait',
      completion: ['active','inspection_wait'].includes(status.key),
      payment: status.key === 'payment_wait',
      defect: status.key === 'done' && !(p.defectStartDate && p.defectEndDate)
    };
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

      <nav class="detail-tabs" aria-label="공사 상세 메뉴">
        <button type="button" class="detail-tab ${state.detailTab==='info'?'active':''}" data-detail-tab="info">공사정보</button>
        <button type="button" class="detail-tab ${state.detailTab==='changes'?'active':''}" data-detail-tab="changes">변경계약${p.contractChanges?.length ? ` <span>${p.contractChanges.length}</span>` : ''}</button>
        <button type="button" class="detail-tab ${state.detailTab==='documents'?'active':''}" data-detail-tab="documents">서류 <span>18</span></button>
      </nav>

      <section class="workflow-strip ${state.detailTab==='info'?'':'hidden'}" aria-label="공사 진행 단계">
        ${workflow.map(workflowStepHtml).join('')}
      </section>

      <div class="detail-layout ${state.detailTab==='info'?'':'hidden'}">
        <section class="panel form-panel" id="projectForm">
          ${workflowSectionHtml('basic','공사 기본정보','한 번 입력해 계속 재사용하는 공사·업체 정보', {label:'기본', done:!!p.projectName && !!p.vendorName, active:status.key==='contract_prep'}, [
            field('projectName','공사명',p.projectName,'text',true),
            field('fiscalYear','회계연도',p.fiscalYear,'number'),
            selectField('workType','공종',p.workType,['','건축공사','전기공사','통신공사','소방공사','기계설비공사','토목공사','기타']),
            `<div class="field full"><label for="vendorPicker">업체 보관함에서 선택</label><select id="vendorPicker">${vendorOptions}</select><span class="hint">선택하면 대표자·사업자번호·주소·전화가 현재 공사에 자동 반영됩니다.</span></div>`,
            field('vendorName','업체명',p.vendorName),
            field('representative','대표자',p.representative),
            field('businessNumber','사업자등록번호',p.businessNumber),
            field('vendorPhone','대표전화',p.vendorPhone),
            field('vendorAddress','사업장 주소',p.vendorAddress,'text',true),
            field('licenseType','등록면허 / 업종',p.licenseType,'text',true),
            `<div class="field full"><button class="button secondary small" id="saveVendorBtn" type="button">현재 업체정보를 보관함에 반영</button></div>`
          ], currentOpen.basic)}

          ${workflowSectionHtml('contract','계약','계약이 확정되면 한 번만 입력합니다.',step.contract,[
            selectField('contractMethod','계약방법',p.contractMethod,['','1인수의','2인이상수의','제한경쟁','일반경쟁','조달계약','기타']),
            moneyField('estimatedPrice','예정가격',p.estimatedPrice),
            moneyField('currentContractAmount','현재 계약금액',p.currentContractAmount),
            field('contractNumber','계약번호',p.contractNumber),
            field('contractDate','계약일',p.contractDate,'date'),
            moneyField('originalContractAmount','최초 계약금액',p.originalContractAmount),
            field('bidRate','낙찰율 (%)',p.bidRate,'number'),
            field('procurementMethod','G2B / S2B / 조달방식',p.procurementMethod),
            field('contractSecurityType','계약보증 방법',p.contractSecurityType),
            field('contractSecurityRate','계약보증률 (%)',p.contractSecurityRate,'number'),
            moneyField('contractSecurityAmount','계약보증금액',p.contractSecurityAmount),
            field('delayPenaltyRate','지연배상금률',p.delayPenaltyRate),
            field('priceAdjustmentMethod','물가변동 계약금액 조정방법',p.priceAdjustmentMethod,'text',true),
            `<div class="field full contract-change-block">
              <div class="subsection-head"><div><strong>변경계약 이력</strong><span>최초 계약을 덮어쓰지 않고 변경 내용을 남깁니다.</span></div><button class="button secondary small" id="addContractChangeBtn" type="button">+ 변경계약 추가</button></div>
              ${contractChangeHistoryHtml(p)}
            </div>`
          ], currentOpen.contract)}

          ${workflowSectionHtml('start','착공','착공하면서 새로 생긴 값만 추가합니다.',step.start,[
            field('plannedStartDate','착공예정일',p.plannedStartDate,'date'),
            field('startDate','착공일',p.startDate,'date'),
            field('completionDueDate','준공기한',p.completionDueDate,'date'),
            field('siteManager','현장대리인',p.siteManager),
            field('siteManagerLicense','현장대리인 자격증 종류',p.siteManagerLicense)
          ], currentOpen.start)}

          ${workflowSectionHtml('completion','준공','준공 및 검사 단계에서 필요한 값만 추가합니다.',step.completion,[
            field('actualCompletionDate','실제 준공일',p.actualCompletionDate,'date'),
            field('completionInspectionDate','검사·검수일',p.completionInspectionDate,'date'),
            field('completionRequestIssueDate','준공검사원 발행일',p.completionRequestIssueDate,'date'),
            field('completionRequestDocDate','준공검사원 문서등록일',p.completionRequestDocDate,'date'),
            field('completionInspectionRecordDate','준공검사조서 작성일',p.completionInspectionRecordDate,'date')
          ], currentOpen.completion)}

          ${workflowSectionHtml('payment','지출','준공 후 지급 단계에서 확인하는 정보입니다.',step.payment,[
            moneyField('settlementAmount','준공정산금액',p.settlementAmount),
            moneyField('priorPaymentAmount','기지급액',p.priorPaymentAmount),
            moneyField('deductionAmount','공제금액',p.deductionAmount),
            field('advancePayment','선금 지급 여부',p.advancePayment),
            field('taxInvoiceDate','세금계산서 발행일',p.taxInvoiceDate,'date'),
            field('paymentDate','지출일',p.paymentDate,'date'),
            moneyField('paymentAmount','지출금액',p.paymentAmount),
            field('fundingSource','재원구분',p.fundingSource),
            `<div class="field full ledger-extra-wrap"><details class="ledger-extra-details"><summary>공사대장 추가정보 <span>선택 입력</span></summary><div class="form-grid nested-form-grid">
              ${field('designer','설계자 / 설계사무소',p.designer)}
              ${field('budgetPolicyProject','정책사업',p.budgetPolicyProject)}
              ${field('budgetUnitProject','단위사업',p.budgetUnitProject)}
              ${field('budgetDetailProject','세부사업',p.budgetDetailProject)}
              ${field('budgetDetailItem','세부항목',p.budgetDetailItem)}
              ${field('costStatisticsItem','원가통계목',p.costStatisticsItem)}
            </div></details></div>`
          ], currentOpen.payment)}

          ${workflowSectionHtml('defect','하자','세부공종별 하자담보기간을 추천·관리하고 검사이력을 누적합니다.',step.defect,[
            field('defectSecurityType','하자보증서 / 각서',p.defectSecurityType),
            field('defectSecurityRate','하자보증률 (%)',p.defectSecurityRate,'number'),
            moneyField('defectSecurityAmount','하자보증금액',p.defectSecurityAmount),
            `<div class="field full warranty-management-block">
              <div class="subsection-head"><div><strong>하자담보 공종</strong><span>별표 4 기준 추천 · 복합공종은 각각 따로 관리</span></div><button class="button primary small" id="openWarrantyManagerBtn" type="button">+ 하자공종 관리</button></div>
              ${warrantyItemsHtml(p)}
              ${(p.warrantyItems||[]).length ? `<div class="warranty-summary-note">대표 하자기간 ${e(p.defectPeriodYears||'—')}년 · 최종 만료 ${e(formatDate(p.defectEndDate))}</div>` : ''}
            </div>`,
            `<div class="field full warranty-management-block">
              <div class="subsection-head"><div><strong>하자검사 이력</strong><span>검사할 때마다 새 기록을 추가합니다. 이전 기록은 덮어쓰지 않습니다.</span></div><button class="button secondary small" id="addWarrantyInspectionBtn" type="button">+ 하자검사 추가</button></div>
              ${warrantyInspectionHistoryHtml(p)}
            </div>`,
            textareaField('notes','비고',p.notes)
          ], currentOpen.defect)}
        </section>

        <aside class="side-stack">
          <section class="side-card next-action-card">
            <p class="eyebrow">지금 확인할 것</p>
            <h3>${e(status.label)}</h3>
            ${missing.length ? `<p>현재 단계에서는 아래 정보만 먼저 확인하면 됩니다.</p><ul class="missing-list">${missing.map(x=>`<li>${e(x)}</li>`).join('')}</ul><button class="button primary small" type="button" id="jumpCurrentStage">입력하러 가기</button>` : `<p class="all-good">✓ 현재 단계의 핵심정보가 입력되어 있습니다.</p>`}
          </section>
          <section class="side-card">
            <h3>변경계약</h3>
            <p>${p.contractChanges?.length ? `변경계약 ${p.contractChanges.length}건이 기록되어 있습니다.` : '아직 변경계약 이력이 없습니다.'}</p>
            <button class="button secondary small" id="sideAddContractChangeBtn" type="button" style="margin-top:10px">+ 변경계약 추가</button>
          </section>
          <section class="side-card document-side-card">
            <h3>공사서류</h3>
            <p>저장된 공사정보로 행정기관 내부 양식을 바로 만듭니다.</p>
            <div class="doc-list">
              ${documentQuickItemHtml('constructionLedger', p)}
              ${documentQuickItemHtml('supervisionReport', p)}
              ${documentQuickItemHtml('completionInspectionRecord', p)}
              ${documentQuickItemHtml('safetyGeneral', p)}
              ${documentQuickItemHtml('warrantyInspectionReport', p)}
              ${documentQuickItemHtml('warrantyLedger', p)}
            </div>
            <button class="button secondary small full-button" id="openDocumentsTab" type="button">서류 전체 보기</button>
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
      </div>

      <section class="tab-panel ${state.detailTab==='changes'?'':'hidden'}">
        <div class="panel contract-tab-panel">
          <div class="tab-panel-head"><div><p class="eyebrow">계약 변경 이력</p><h2>변경계약</h2><p>최초 계약정보를 지우지 않고 금액과 준공기한 변경 이력을 계속 남깁니다.</p></div><button class="button primary" id="tabAddContractChangeBtn" type="button">+ 변경계약 추가</button></div>
          <div class="contract-tab-summary"><div><span>최초 계약금액</span><strong>${e(formatMoney(p.originalContractAmount))}</strong></div><div><span>현재 계약금액</span><strong>${e(formatMoney(p.currentContractAmount))}</strong></div><div><span>현재 준공기한</span><strong>${e(formatDate(p.completionDueDate))}</strong></div></div>
          <div class="contract-tab-history">${contractChangeHistoryHtml(p)}</div>
        </div>
      </section>

      <section class="tab-panel ${state.detailTab==='documents'?'':'hidden'}">
        ${documentsTabHtml(p)}
      </section>`;

    document.getElementById('backToList').addEventListener('click', () => { state.currentProjectId = null; renderDashboard(); });
    document.getElementById('vendorPicker')?.addEventListener('change', applyVendorToCurrentProject);
    document.getElementById('saveVendorBtn')?.addEventListener('click', saveCurrentVendor);
    document.getElementById('deleteProjectBtn')?.addEventListener('click', confirmDeleteProject);
    document.getElementById('addContractChangeBtn')?.addEventListener('click', openContractChangeModal);
    document.getElementById('sideAddContractChangeBtn')?.addEventListener('click', openContractChangeModal);
    document.getElementById('tabAddContractChangeBtn')?.addEventListener('click', openContractChangeModal);
    document.getElementById('openWarrantyManagerBtn')?.addEventListener('click', openWarrantyManager);
    document.getElementById('addWarrantyInspectionBtn')?.addEventListener('click', () => openWarrantyInspectionModal());
    document.getElementById('openDocumentsTab')?.addEventListener('click', () => { state.detailTab='documents'; renderProjectDetail(); });
    document.getElementById('jumpCurrentStage')?.addEventListener('click', () => jumpToSection(sectionForStatus(status.key)));
    main.querySelectorAll('[data-detail-tab]').forEach(btn => btn.addEventListener('click', () => { state.detailTab = btn.dataset.detailTab; renderProjectDetail(); }));
    main.querySelectorAll('[data-doc-open]').forEach(btn => btn.addEventListener('click', () => openDocumentPreview(btn.dataset.docOpen)));
    bindDocumentBatchControls(p);
    main.querySelectorAll('[data-jump-section]').forEach(btn => btn.addEventListener('click', () => jumpToSection(btn.dataset.jumpSection)));
    main.querySelectorAll('[data-change-delete]').forEach(btn => btn.addEventListener('click', () => confirmDeleteContractChange(btn.dataset.changeDelete)));
    main.querySelectorAll('[data-warranty-preview]').forEach(btn => btn.addEventListener('click', () => { state.activeWarrantyInspectionId = btn.dataset.warrantyPreview; openDocumentPreview('warrantyInspectionReport'); }));
    main.querySelectorAll('[data-warranty-delete]').forEach(btn => btn.addEventListener('click', () => deleteWarrantyInspection(btn.dataset.warrantyDelete)));
    main.querySelectorAll('[data-field]').forEach(input => {
      input.addEventListener('input', onProjectInput);
      input.addEventListener('change', onProjectInput);
    });
    initDateInputs(main);
    initMoneyInputs(main);
  }


  const DOCUMENT_DEFINITIONS = Documents.definitions;
  const DOCUMENT_FIELD_LABELS = Documents.fieldLabels;
  const DOCUMENT_PRINT_ORDER = Documents.printOrder;
  const DOCUMENT_SETS = Documents.sets;

  function vendorForProject(p) {
    if (!p) return null;
    return (p.vendorId && state.vendors.find(v => v.id === p.vendorId)) ||
      state.vendors.find(v => p.businessNumber && normalizeText(v.businessNumber) === normalizeText(p.businessNumber)) ||
      state.vendors.find(v => p.vendorName && normalizeText(v.name) === normalizeText(p.vendorName)) || null;
  }

  function payoutForProject(p) {
    const vendor = vendorForProject(p);
    return vendor ? (state.payouts.find(x => x.vendorId === vendor.id) || null) : null;
  }

  function activeWarrantyInspection(p) {
    const list = Array.isArray(p?.warrantyInspections) ? p.warrantyInspections : [];
    if (!list.length) return null;
    return list.find(x => x.id === state.activeWarrantyInspectionId) || list[list.length-1] || null;
  }

  function documentValue(field, p) {
    if (field === 'schoolName') return state.school?.name || '';
    if (field === 'schoolAddress') return state.school?.address || '';
    if (field === 'principal') return state.school?.principal || '';
    if (field === 'supervisor') return p?.supervisor || state.school?.supervisor || '';
    if (field === 'inspector') return p?.inspector || state.school?.inspector || '';
    if (field === 'witness') return p?.witness || state.school?.witness || '';
    if (['bankName','accountNumber','accountHolder'].includes(field)) {
      const payout = payoutForProject(p);
      if (field === 'accountHolder') return payout?.accountHolder || p?.vendorName || '';
      return payout?.[field] || '';
    }
    if (field.startsWith('warranty')) {
      const wi = activeWarrantyInspection(p);
      const map = {
        warrantyInspectionDate:'date', warrantyInspector:'inspector', warrantyWitness:'witness',
        warrantyInspectionResult:'result', warrantyIssueDetails:'issueDetails', warrantyActions:'actions', warrantyNotes:'notes'
      };
      return wi?.[map[field]] ?? '';
    }
    if (field === 'contractSecurityAmount' && !meaningful(p?.contractSecurityAmount)) {
      const amount = Number(p?.currentContractAmount), rateRaw = Number(p?.contractSecurityRate);
      const rate = Number.isFinite(rateRaw) && Math.abs(rateRaw) > 1 ? rateRaw / 100 : rateRaw;
      if (Number.isFinite(amount) && Number.isFinite(rate)) return Math.floor((amount * rate) / 10) * 10;
    }
    if (field === 'defectSecurityAmount' && !meaningful(p?.defectSecurityAmount)) {
      const amount = Number(p?.currentContractAmount), rateRaw = Number(p?.defectSecurityRate);
      const rate = Number.isFinite(rateRaw) && Math.abs(rateRaw) > 1 ? rateRaw / 100 : rateRaw;
      if (Number.isFinite(amount) && Number.isFinite(rate)) return Math.floor((amount * rate) / 10) * 10;
    }
    return p?.[field] ?? '';
  }

  function documentMissing(type, p) {
    const def = DOCUMENT_DEFINITIONS[type];
    if (!def) return [];
    return def.required.filter(field => !meaningful(documentValue(field, p)));
  }

  function selectionForProject(p) {
    if (!p) return new Set();
    if (state.selectionProjectId !== p.id) {
      state.selectionProjectId = p.id;
      state.selectedDocuments = new Set();
    }
    return state.selectedDocuments;
  }

  function orderedSelectedTypes(p) {
    const selected = selectionForProject(p);
    return DOCUMENT_PRINT_ORDER.filter(type => selected.has(type));
  }

  function batchMissingFields(types, p) {
    const fields = [];
    types.forEach(type => documentMissing(type, p).forEach(field => {
      if (!fields.includes(field)) fields.push(field);
    }));
    return fields;
  }

  function documentsNeedingField(field, types, p) {
    return types.filter(type => documentMissing(type, p).includes(field)).map(type => DOCUMENT_DEFINITIONS[type]?.label).filter(Boolean);
  }

  function isSafetyDocument(type) { return /^safety/.test(String(type || '')); }
  function safetyRequiresCompletion(type) { return type === 'safetyGeneral'; }

  function safetyChecklistFor(p, type) { return p?.safetyChecklists?.[type] || null; }

  function safetyChecklistComplete(p, type) {
    const data=safetyChecklistFor(p,type);
    const def=ReferenceData?.safetyChecklists?.[type];
    if(!data || !def) return false;
    const resultCount=Object.keys(data.results||{}).filter(k=>data.results[k]).length;
    return !!data.date && !!data.inspector && resultCount===def.items.length;
  }

  function myDocumentTypes(p) {
    return (Array.isArray(p?.selectedDocuments)?p.selectedDocuments:[]).filter(type=>DOCUMENT_DEFINITIONS[type]);
  }

  async function toggleMyDocument(type) {
    const p=currentProject(); if(!p||!DOCUMENT_DEFINITIONS[type])return;
    const set=new Set(myDocumentTypes(p));
    if(set.has(type))set.delete(type);else set.add(type);
    p.selectedDocuments=[...set];p.updatedAt=new Date().toISOString();
    await DB.put('projects',p); await loadState(); state.currentProjectId=p.id; renderProjectDetail();
  }

  function recommendedSafetyTypes(p) { return ReferenceData?.recommendedSafetyTypes?.(p) || ['safetyGeneral']; }

  function safetyRecommendationsHtml(p) {
    const recommended=recommendedSafetyTypes(p);
    const selected=new Set(myDocumentTypes(p));
    const chars=p.workCharacteristics||{};
    return `<section class="safety-recommend-panel"><div class="safety-recommend-head"><div><p class="eyebrow">공종·작업특성 기반</p><h3>추천 안전·보건 서류</h3><p>추천은 참고용입니다. 실제 작업내용을 확인한 뒤 내 서류에 추가하거나 사용하지 않을 수 있습니다.</p></div></div>
      <div class="work-characteristics"><strong>이 공사에 해당하는 작업</strong><label><input type="checkbox" data-work-characteristic="fall" ${chars.fall?'checked':''}> 고소·추락 위험 작업</label><label><input type="checkbox" data-work-characteristic="electrical" ${chars.electrical?'checked':''}> 전기 작업</label><label><input type="checkbox" data-work-characteristic="confined" ${chars.confined?'checked':''}> 밀폐공간 작업</label><label><input type="checkbox" data-work-characteristic="industrial" ${chars.industrial?'checked':''}> 일반 건설작업</label></div>
      <div class="recommended-document-list">${recommended.map(type=>{const def=DOCUMENT_DEFINITIONS[type];if(!def)return'';const source=ReferenceData.safetyChecklists[type]?.owner==='agency'?'기관 점검':'업체 작성·기관 확인';return `<div class="recommended-document"><div><strong>${e(def.label)}</strong><span>${e(source)}</span></div><button class="button ${selected.has(type)?'ghost':'secondary'} small" type="button" data-my-doc-toggle="${e(type)}">${selected.has(type)?'내 서류에서 빼기':'내 서류에 추가'}</button></div>`;}).join('')}</div></section>`;
  }

  function myDocumentsHtml(p) {
    const types=myDocumentTypes(p);
    return `<section class="my-documents-panel"><div class="document-group-title"><strong>내 서류 ${types.length}종</strong><span>이 공사에서 실제 사용하는 서류만 모아봅니다.</span></div>${types.length?`<div class="my-document-chips">${types.map(type=>`<button type="button" data-doc-open="${e(type)}"><span>${e(DOCUMENT_DEFINITIONS[type].label)}</span><em>${isSafetyDocument(type)?(safetyChecklistComplete(p,type)?'작성완료':(safetyRequiresCompletion(type)?'작성 필요':'빈 양식 가능')):(documentMissing(type,p).length?'정보 필요':'생성 가능')}</em></button>`).join('')}</div>`:`<div class="contract-change-empty">아직 내 서류가 없습니다. 아래 추천 또는 전체 서류에서 필요한 문서를 추가하세요.</div>`}</section>`;
  }

  function safetyDocumentCardHtml(type,p){
    const def=DOCUMENT_DEFINITIONS[type];
    const saved=safetyChecklistFor(p,type);
    const complete=safetyChecklistComplete(p,type);
    const mustComplete=safetyRequiresCompletion(type);
    const selected=myDocumentTypes(p).includes(type);
    const batchSelected=selectionForProject(p).has(type);
    const vendorForm=!mustComplete;
    const source=vendorForm?'업체 작성·제출 / 기관 확인':'기관 점검·자체 보관';
    const statusTitle=complete?'✓ 작성 완료':(vendorForm?'빈 양식 출력 가능':'점검결과 입력 필요');
    const statusDesc=saved?.date?`최근 점검 ${e(formatDate(saved.date))}`:(vendorForm?'업체가 직접 점검·작성할 수 있도록 체크하지 않은 원본 양식도 출력할 수 있습니다.':'원본 체크리스트 항목을 웹에서 입력합니다.');
    return `<article class="document-card safety-document-card ${complete||vendorForm?'ready':'needs-info'} ${batchSelected?'selected':''}" data-document-card="${e(type)}"><div class="document-card-top"><label class="document-select"><input type="checkbox" data-doc-select="${e(type)}" ${batchSelected?'checked':''}><span>선택</span></label><div><span class="document-stage">${e(source)}</span><span class="document-version">양식 ${e(def.version)}</span></div></div><h3>${e(def.label)}</h3><p>${e(def.description)}</p><div class="document-requirement ${complete||vendorForm?'ready':''}"><strong>${e(statusTitle)}</strong><span>${statusDesc}</span></div><div class="document-card-actions"><button class="button ${complete?'secondary':'primary'}" type="button" data-safety-edit="${e(type)}">${complete?'작성내용 수정':'체크리스트 작성'}</button>${(complete||vendorForm)?`<button class="button ghost" type="button" data-doc-open="${e(type)}">${complete?'미리보기':'빈 양식 미리보기'}</button>`:''}</div><button class="my-doc-toggle ${selected?'active':''}" type="button" data-my-doc-toggle="${e(type)}">${selected?'✓ 내 서류':'＋ 내 서류'}</button></article>`;
  }

  function openSafetyChecklistModal(type, afterSave=null) {
    const p=currentProject(); const def=ReferenceData?.safetyChecklists?.[type]; if(!p||!def)return;
    const saved=safetyChecklistFor(p,type)||{}; const results=saved.results||{}; const agency=def.owner==='agency';
    const defaultInspector=saved.inspector || (agency?(p.supervisor||state.school?.supervisor||''):'');
    const rows=def.items.map((item,i)=>{const key=String(i+1),v=results[key]||'';return `<div class="safety-edit-row"><div class="safety-edit-number">${i+1}</div><div class="safety-edit-question">${e(item)}</div><div class="safety-edit-options"><label><input type="radio" name="safety_${i}" value="yes" ${v==='yes'?'checked':''}> 예</label><label><input type="radio" name="safety_${i}" value="no" ${v==='no'?'checked':''}> 아니요</label><label><input type="radio" name="safety_${i}" value="na" ${v==='na'?'checked':''}> 해당없음</label></div></div>`;}).join('');
    openModal({eyebrow:`안전·보건 · ${agency?'기관 점검':'업체 작성·기관 확인'}`,title:def.label,wide:true,body:`<div class="notice"><strong>${e(def.subtitle)}</strong><br>원본 「공사서류 원클릭 프로그램(2026.4.)」의 점검항목을 기준으로 작성합니다.</div><div class="modal-grid safety-meta-edit" style="margin-top:16px">${modalDateField('safetyChecklistDate','점검일',saved.date||p.startDate||p.contractDate||'')}${modalField('safetyChecklistInspector',agency?'점검자':'점검자 직접 입력',defaultInspector)}${agency?'':'<div class="field full"><span class="hint">회사 대표자가 아니라 실제 점검한 사람의 이름을 입력합니다.</span></div>'}<div class="field full"><label for="safetyChecklistNotes">비고</label><input id="safetyChecklistNotes" value="${e(saved.notes||'')}"></div></div><div class="safety-edit-list">${rows}</div>${def.footer?`<p class="safety-edit-footer">※ ${e(def.footer)}</p>`:''}`,actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="saveSafetyChecklistBtn">저장${afterSave?'하고 계속':''}</button>`});
    initDateInputs(modalBody); modalActions.querySelector('[data-modal-close]').addEventListener('click',closeModal); modalActions.querySelector('#saveSafetyChecklistBtn').addEventListener('click',()=>saveSafetyChecklist(type,afterSave));
  }

  async function saveSafetyChecklist(type, afterSave=null){
    const p=currentProject(),def=ReferenceData?.safetyChecklists?.[type];if(!p||!def)return;
    const date=modalBody.querySelector('#safetyChecklistDate')?.value||'',inspector=modalBody.querySelector('#safetyChecklistInspector')?.value?.trim()||'';
    if(!date||!inspector){showToast('점검일과 점검자를 입력해주세요.','warn');return;}
    const results={}; let missing=0; def.items.forEach((_,i)=>{const v=modalBody.querySelector(`input[name="safety_${i}"]:checked`)?.value||'';results[String(i+1)]=v;if(!v)missing++;});
    if(missing){showToast(`아직 선택하지 않은 점검항목이 ${missing}개 있습니다.`,'warn');return;}
    p.safetyChecklists={...(p.safetyChecklists||{}),[type]:{date,inspector,notes:modalBody.querySelector('#safetyChecklistNotes')?.value?.trim()||'',results,updatedAt:new Date().toISOString()}};
    if(!p.selectedDocuments?.includes(type))p.selectedDocuments=[...(p.selectedDocuments||[]),type];
    p.updatedAt=new Date().toISOString();await DB.put('projects',p);await loadState();state.currentProjectId=p.id;closeModal();renderProjectDetail();showToast(`${def.label}을 저장했습니다.`);if(typeof afterSave==='function')afterSave();
  }

  function documentQuickItemHtml(type, p) {
    const def = DOCUMENT_DEFINITIONS[type];
    const missing = documentMissing(type, p);
    const safetyIncomplete = isSafetyDocument(type) && safetyRequiresCompletion(type) && !safetyChecklistComplete(p,type);
    const vendorBlank = isSafetyDocument(type) && !safetyRequiresCompletion(type) && !safetyChecklistComplete(p,type);
    const status = safetyIncomplete ? '작성 필요' : (vendorBlank ? '빈 양식 가능' : (missing.length ? `${missing.length}개 정보 필요` : '생성 가능'));
    return `<button class="doc-item doc-item-button ${(missing.length||safetyIncomplete)?'needs-info':'ready'}" type="button" data-doc-open="${e(type)}"><span>${e(def.label)}</span><em>${e(status)}</em></button>`;
  }

  function documentCardHtml(type, p) {
    if (isSafetyDocument(type)) return safetyDocumentCardHtml(type,p);
    const def = DOCUMENT_DEFINITIONS[type];
    const missing = documentMissing(type, p);
    const labels = missing.map(x => DOCUMENT_FIELD_LABELS[x] || x);
    const checked = selectionForProject(p).has(type);
    const inMyDocs = myDocumentTypes(p).includes(type);
    return `<article class="document-card ${missing.length?'needs-info':'ready'} ${checked?'selected':''}" data-document-card="${e(type)}">
      <div class="document-card-top"><label class="document-select"><input type="checkbox" data-doc-select="${e(type)}" ${checked?'checked':''}><span>선택</span></label><div><span class="document-stage">${e(def.stage)}</span><span class="document-version">양식 ${e(def.version)}</span></div></div>
      <h3>${e(def.label)}</h3><p>${e(def.description)}</p>
      ${missing.length ? `<div class="document-requirement"><strong>추가 입력 ${missing.length}개</strong><span>${e(labels.slice(0,3).join(' · '))}${labels.length>3?' 외':''}</span></div>` : `<div class="document-requirement ready"><strong>✓ 바로 생성 가능</strong><span>현재 공사정보를 그대로 사용합니다.</span></div>`}
      <button class="button ${missing.length?'secondary':'primary'}" type="button" data-doc-open="${e(type)}">${missing.length?'부족정보 입력하고 만들기':'미리보기'}</button>
      <button class="my-doc-toggle ${inMyDocs?'active':''}" type="button" data-my-doc-toggle="${e(type)}">${inMyDocs?'✓ 내 서류':'＋ 내 서류'}</button>
    </article>`;
  }

  function recentPrintHistoryHtml(p) {
    const history = Array.isArray(p.printHistory) ? p.printHistory.slice(-3).reverse() : [];
    if (!history.length) return '';
    return `<div class="print-history"><strong>최근 묶음 출력</strong>${history.map(item => `<span>${e(formatDateTime(item.at))} · ${e(item.labels?.join(' · ') || `${item.count || 0}종`)}</span>`).join('')}</div>`;
  }

  function utilityToolCardHtml(p) {
    const u = p.utilityCost || {};
    const total = Number(u.total || 0);
    return `<article class="document-card utility-tool-card ${total?'ready':''}">
      <div class="document-card-top"><div><span class="document-stage">계산도구</span><span class="document-version">원가통계 기준</span></div></div>
      <h3>전력비·수도광열비 계산</h3><p>원클릭 프로그램의 「수도전기료계산식」 기준으로 전력비·수도광열비를 계산하고 대금청구 공제금액에 바로 연결합니다.</p>
      <div class="document-requirement ${total?'ready':''}">${total?`<strong>계산값 ${e(formatMoney(total))}</strong><span>공제금액에 바로 반영할 수 있습니다.</span>`:`<strong>계산 필요</strong><span>직접재료비·직접노무비만 추가 입력하면 됩니다.</span>`}</div>
      <button class="button ${total?'secondary':'primary'}" type="button" data-open-utility-calculator>${total?'계산 다시 보기':'계산하기'}</button>
    </article>`;
  }

  function documentsTabHtml(p) {
    const selected = orderedSelectedTypes(p);
    const missing = batchMissingFields(selected, p);
    const readyCount = selected.filter(type => documentMissing(type,p).length === 0 && (!isSafetyDocument(type) || safetyChecklistComplete(p,type) || !safetyRequiresCompletion(type))).length;
    return `<div class="documents-panel">
      <div class="documents-head"><div><p class="eyebrow">행정기관 작성·관리 우선</p><h2>공사서류</h2><p>공종과 작업특성에 따라 필요한 서류를 추천하고, 실제 사용하는 문서만 ‘내 서류’에 모아 관리합니다.</p></div><div class="documents-head-note"><strong>v0.4.2.2</strong><span>공사대장 정렬 · 점검자 직접 입력</span></div></div>
      ${myDocumentsHtml(p)}
      ${safetyRecommendationsHtml(p)}
      <div class="document-batch-toolbar">
        <div class="document-set-buttons" aria-label="서류 세트 선택">
          <button class="button secondary small" type="button" data-doc-set="agencyManagement">행정기관 관리서류</button>
          <button class="button secondary small" type="button" data-doc-set="safety">안전·보건 5종</button>
          <button class="button secondary small" type="button" data-doc-set="contract">계약서류 4종</button>
          <button class="button secondary small" type="button" data-doc-set="completion">준공서류 4종</button>
          <button class="button ghost small" type="button" data-doc-set="all">전체 18종</button>
          <button class="button ghost small" type="button" data-doc-clear>선택 해제</button>
        </div>
        <div class="document-batch-summary" id="documentBatchSummary">
          <div><strong>${selected.length ? `${selected.length}종 선택` : '서류를 선택하세요'}</strong><span>${selected.length ? (missing.length ? `공통 부족정보 ${missing.length}개 · 바로 출력 ${readyCount}종` : '선택한 서류의 작성상태를 확인한 뒤 출력합니다.') : '개별 선택 또는 서류 세트를 사용할 수 있습니다.'}</span></div>
          <button class="button primary" type="button" id="openBatchPreviewBtn" ${selected.length?'':'disabled'}>${missing.length ? `부족정보 ${missing.length}개 확인` : `선택한 ${selected.length}종 미리보기`}</button>
        </div>
      </div>
      <div class="document-group agency-document-group"><div class="document-group-title"><strong>행정기관 작성·관리</strong><span>공사 등록부터 감사·검사·하자관리까지 같은 공사정보를 계속 재사용합니다.</span></div><div class="document-grid">${documentCardHtml('constructionLedger',p)}${documentCardHtml('supervisionReport',p)}${documentCardHtml('completionInspectionRecord',p)}${documentCardHtml('warrantyInspectionReport',p)}${documentCardHtml('warrantyLedger',p)}${utilityToolCardHtml(p)}</div></div>
      <div class="document-group safety-document-group"><div class="document-group-title"><strong>안전·보건 확인</strong><span>공통 체크리스트는 기관 점검용이며, 재해별 체크리스트는 원본 기준으로 업체 작성·제출 후 기관에서 확인합니다.</span></div><div class="document-grid">${documentCardHtml('safetyGeneral',p)}${documentCardHtml('safetyFall',p)}${documentCardHtml('safetyElectrical',p)}${documentCardHtml('safetyConfined',p)}${documentCardHtml('safetyIndustrial',p)}</div></div>
      <div class="document-group vendor-document-group"><div class="document-group-title"><strong>업체 제출·징구</strong><span>계약·착공·준공·지출 단계에서 업체와 주고받는 서류</span></div>
        <div class="document-subgroup"><h4>계약</h4><div class="document-grid">${documentCardHtml('standardContract',p)}${documentCardHtml('acceptanceTerms',p)}${documentCardHtml('useSealForm',p)}${documentCardHtml('privateContractPledge',p)}</div></div>
        <div class="document-subgroup"><h4>착공</h4><div class="document-grid">${documentCardHtml('startReport',p)}</div></div>
        <div class="document-subgroup"><h4>준공</h4><div class="document-grid">${documentCardHtml('completionReport',p)}${documentCardHtml('completionInspectionRequest',p)}</div></div>
        <div class="document-subgroup"><h4>지출</h4><div class="document-grid">${documentCardHtml('paymentRequest',p)}</div></div>
      </div>
      ${recentPrintHistoryHtml(p)}
      <div class="document-footnote">출력양식과 안전점검 항목은 제공받은 「공사서류 원클릭 프로그램(2026.4.)」을 기준으로 구현했습니다. 하자담보기간 추천은 제공받은 「건설산업기본법 시행령 별표 4」를 참고하며, 추천값은 사용자가 확인한 뒤 적용합니다.</div>
    </div>`;
  }

  function refreshDocumentBatchUi(p) {
    const summary = main.querySelector('#documentBatchSummary');
    if (!summary) return;
    const selected = orderedSelectedTypes(p);
    const missing = batchMissingFields(selected,p);
    const readyCount = selected.filter(type => documentMissing(type,p).length === 0 && (!isSafetyDocument(type) || safetyChecklistComplete(p,type) || !safetyRequiresCompletion(type))).length;
    summary.innerHTML = `<div><strong>${selected.length ? `${selected.length}종 선택` : '서류를 선택하세요'}</strong><span>${selected.length ? (missing.length ? `공통 부족정보 ${missing.length}개 · 바로 출력 ${readyCount}종` : '모든 서류가 바로 출력 가능합니다.') : '개별 선택 또는 단계별 세트를 사용할 수 있습니다.'}</span></div><button class="button primary" type="button" id="openBatchPreviewBtn" ${selected.length?'':'disabled'}>${missing.length ? `부족정보 ${missing.length}개 확인` : `선택한 ${selected.length}종 미리보기`}</button>`;
    summary.querySelector('#openBatchPreviewBtn')?.addEventListener('click',()=>openSelectedDocumentsFlow());
  }

  function bindDocumentBatchControls(p) {
    if (!p || state.detailTab !== 'documents') return;
    main.querySelectorAll('[data-doc-select]').forEach(input => input.addEventListener('change', () => {
      const selected = selectionForProject(p);
      if (input.checked) selected.add(input.dataset.docSelect); else selected.delete(input.dataset.docSelect);
      input.closest('[data-document-card]')?.classList.toggle('selected', input.checked);
      refreshDocumentBatchUi(p);
    }));
    main.querySelectorAll('[data-doc-set]').forEach(btn => btn.addEventListener('click', () => {
      const set = DOCUMENT_SETS[btn.dataset.docSet];
      if (!set) return;
      state.selectedDocuments = new Set(set.types);
      state.selectionProjectId = p.id;
      renderProjectDetail();
    }));
    main.querySelector('[data-doc-clear]')?.addEventListener('click', () => {
      state.selectedDocuments = new Set();
      state.selectionProjectId = p.id;
      renderProjectDetail();
    });
    main.querySelector('#openBatchPreviewBtn')?.addEventListener('click', () => openSelectedDocumentsFlow());
    main.querySelector('[data-open-utility-calculator]')?.addEventListener('click', openUtilityCalculator);
    main.querySelectorAll('[data-my-doc-toggle]').forEach(btn=>btn.addEventListener('click',()=>toggleMyDocument(btn.dataset.myDocToggle)));
    main.querySelectorAll('[data-safety-edit]').forEach(btn=>btn.addEventListener('click',()=>openSafetyChecklistModal(btn.dataset.safetyEdit)));
    main.querySelectorAll('[data-work-characteristic]').forEach(input=>input.addEventListener('change',async()=>{
      const project=currentProject(); if(!project)return;
      project.workCharacteristics={...(project.workCharacteristics||{}),[input.dataset.workCharacteristic]:input.checked};
      project.updatedAt=new Date().toISOString(); await DB.put('projects',project); await loadState(); state.currentProjectId=project.id; renderProjectDetail();
    }));
  }

  function openSelectedDocumentsFlow() {
    const p = currentProject();
    if (!p) return;
    const types = orderedSelectedTypes(p);
    if (!types.length) { showToast('인쇄할 서류를 먼저 선택해주세요.', 'warn'); return; }
    const incompleteSafety = types.find(type => isSafetyDocument(type) && safetyRequiresCompletion(type) && !safetyChecklistComplete(p,type));
    if (incompleteSafety) { showToast(`${DOCUMENT_DEFINITIONS[incompleteSafety].label}을 먼저 작성해주세요.`, 'warn'); openSafetyChecklistModal(incompleteSafety,()=>openSelectedDocumentsFlow()); return; }
    const missing = batchMissingFields(types, p);
    if (missing.length) openBatchMissingModal(types, missing);
    else openBatchPreview(types);
  }

  function batchMissingFieldHtml(field, types, p) {
    const base = documentMissingFieldHtml(field, p);
    const labels = documentsNeedingField(field, types, p);
    return `<div class="batch-missing-field">${base}<div class="batch-used-by">사용 서류 · ${e(labels.join(' · '))}</div></div>`;
  }

  function openBatchMissingModal(types, missing) {
    const p = currentProject();
    openModal({
      eyebrow:`선택한 서류 ${types.length}종 출력 전 점검`,
      title:`${missing.length}개 정보만 한 번 더 확인해주세요`,
      body:`<div class="notice"><strong>같은 정보는 한 번만 입력합니다.</strong><br>여기에서 저장한 값은 공사정보 또는 업체 지급정보에 반영되고 선택한 모든 서류가 함께 사용합니다.</div><div class="batch-doc-labels">${types.map(type => `<span>${e(DOCUMENT_DEFINITIONS[type].label)}</span>`).join('')}</div><div class="modal-grid document-missing-grid batch-missing-grid" style="margin-top:16px">${missing.map(field => batchMissingFieldHtml(field,types,p)).join('')}</div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="saveBatchMissingFields">저장하고 ${types.length}종 미리보기</button>`
    });
    initDateInputs(modalBody);
    initMoneyInputs(modalBody);
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#saveBatchMissingFields').addEventListener('click', () => saveBatchMissingFields(types, missing));
  }

  async function saveFieldsForDocuments(fields) {
    const p = currentProject();
    if (!p) return false;
    let schoolChanged = false;
    const payoutValues = {};
    const warrantyValues = {};
    for (const field of fields) {
      const el = modalBody.querySelector(`#docMissing_${CSS.escape(field)}`);
      let value = el?.value?.trim?.() ?? '';
      if (MONEY_FIELDS.has(field)) value = parseMoneyInput(value);
      if (!meaningful(value)) { showToast(`${DOCUMENT_FIELD_LABELS[field] || field}을(를) 입력해주세요.`, 'warn'); return false; }
      if (['schoolName','schoolAddress','principal'].includes(field)) {
        const key = field === 'schoolName' ? 'name' : field === 'schoolAddress' ? 'address' : 'principal';
        state.school = { ...(state.school||{}), [key]:value };
        schoolChanged = true;
      } else if (['supervisor','inspector','witness'].includes(field)) {
        p[field] = value;
        if (!meaningful(state.school?.[field])) {
          state.school = { ...(state.school||{}), [field]:value };
          schoolChanged = true;
        }
      } else if (['bankName','accountNumber','accountHolder'].includes(field)) {
        payoutValues[field] = value;
      } else if (field.startsWith('warranty')) {
        warrantyValues[field] = value;
      } else p[field] = value;
    }
    p.updatedAt = new Date().toISOString();
    if (!meaningful(p.originalContractAmount) && meaningful(p.currentContractAmount)) p.originalContractAmount = p.currentContractAmount;
    await DB.put('projects',p);
    if (schoolChanged) await DB.put('settings',{key:'school',value:state.school});
    if (Object.keys(payoutValues).length) await savePayoutForProject(p, payoutValues);
    if (Object.keys(warrantyValues).length) {
      const list = Array.isArray(p.warrantyInspections) ? [...p.warrantyInspections] : [];
      let idx = state.activeWarrantyInspectionId ? list.findIndex(x=>x.id===state.activeWarrantyInspectionId) : -1;
      const base = idx >= 0 ? list[idx] : {};
      const map = {warrantyInspectionDate:'date',warrantyInspector:'inspector',warrantyWitness:'witness',warrantyInspectionResult:'result',warrantyIssueDetails:'issueDetails',warrantyActions:'actions',warrantyNotes:'notes'};
      const next = {...base,id:base.id||DB.uuid(),createdAt:base.createdAt||new Date().toISOString(),updatedAt:new Date().toISOString()};
      Object.entries(warrantyValues).forEach(([k,v])=>{ if(map[k]) next[map[k]]=v; });
      if (idx >= 0) list[idx]=next; else { list.push(next); idx=list.length-1; }
      p.warrantyInspections=list; state.activeWarrantyInspectionId=next.id;
      await DB.put('projects',p);
    }
    await loadState();
    state.currentProjectId = p.id;
    return true;
  }

  async function saveBatchMissingFields(types, fields) {
    const ok = await saveFieldsForDocuments(fields);
    if (!ok) return;
    closeModal();
    renderProjectDetail();
    openBatchPreview(types);
  }

  function batchPreviewBody(types, p, index = 0) {
    const type = types[index];
    const def = DOCUMENT_DEFINITIONS[type];
    return `<div class="batch-preview-head"><div><strong>${index+1} / ${types.length}</strong><span>${e(def.label)}</span></div><div class="batch-preview-pages">${types.map((t,i)=>`<button type="button" class="${i===index?'active':''}" data-batch-page="${i}" aria-label="${e(DOCUMENT_DEFINITIONS[t].label)}">${i+1}</button>`).join('')}</div></div><div class="doc-preview-scroll batch-preview-scroll">${documentMarkup(type,p)}</div>`;
  }

  function openBatchPreview(types, startIndex = 0) {
    const p = currentProject();
    if (!p || !types.length) return;
    let index = Math.min(Math.max(0,startIndex), types.length-1);
    const paint = () => {
      modalBody.innerHTML = batchPreviewBody(types,p,index);
      modalTitle.textContent = `선택한 서류 ${types.length}종 미리보기`;
      modalEyebrow.textContent = '연속 미리보기 · 인쇄 순서대로';
      modalActions.innerHTML = `<button class="button ghost" type="button" id="batchPrevBtn" ${index===0?'disabled':''}>← 이전</button><button class="button ghost" type="button" id="batchNextBtn" ${index===types.length-1?'disabled':''}>다음 →</button><span class="modal-action-spacer"></span><button class="button secondary" type="button" data-modal-close>닫기</button><button class="button primary" type="button" id="printBatchBtn">${types.length}종 한번에 인쇄</button>`;
      modalBody.querySelectorAll('[data-batch-page]').forEach(btn => btn.addEventListener('click',()=>{index=Number(btn.dataset.batchPage)||0;paint();}));
      modalActions.querySelector('#batchPrevBtn')?.addEventListener('click',()=>{if(index>0){index--;paint();}});
      modalActions.querySelector('#batchNextBtn')?.addEventListener('click',()=>{if(index<types.length-1){index++;paint();}});
      modalActions.querySelector('[data-modal-close]')?.addEventListener('click',closeModal);
      modalActions.querySelector('#printBatchBtn')?.addEventListener('click',()=>printAdministrativeDocuments(types,p));
    };
    modal.classList.add('wide-modal');
    if (!modal.open) modal.showModal();
    paint();
  }
  function documentMissingFieldHtml(field, p) {
    const id = `docMissing_${field}`;
    const label = DOCUMENT_FIELD_LABELS[field] || field;
    const value = documentValue(field, p);
    if (MONEY_FIELDS.has(field)) return `<div class="field"><label for="${e(id)}">${e(label)}</label>${moneyInputHtml(id, value)}</div>`;
    if (['contractDate','plannedStartDate','startDate','completionDueDate','actualCompletionDate','completionInspectionDate','warrantyInspectionDate','defectStartDate','defectEndDate'].includes(field)) return modalDateField(id, label, value);
    if (['warrantyInspectionResult','warrantyIssueDetails','warrantyActions','warrantyNotes'].includes(field)) return `<div class="field full"><label for="${e(id)}">${e(label)}</label><textarea id="${e(id)}">${e(value || '')}</textarea></div>`;
    if (['bankName','accountNumber','accountHolder'].includes(field)) {
      return `<div class="field"><label for="${e(id)}">${e(label)}</label><input id="${e(id)}" value="${e(value || '')}" autocomplete="off"><span class="hint">업체 지급정보 보관함에 별도로 저장됩니다.</span></div>`;
    }
    return modalField(id, label, value, ['vendorAddress','projectName','schoolAddress','priceAdjustmentMethod'].includes(field));
  }

  function openDocumentPreview(type) {
    const p = currentProject();
    const def = DOCUMENT_DEFINITIONS[type];
    if (!p || !def) return;
    if (isSafetyDocument(type) && safetyRequiresCompletion(type) && !safetyChecklistComplete(p,type)) { openSafetyChecklistModal(type,()=>openDocumentPreview(type)); return; }
    if (type === 'warrantyInspectionReport' && !(p.warrantyInspections || []).length) { openWarrantyInspectionModal(null,true); return; }
    if (type === 'warrantyInspectionReport' && !state.activeWarrantyInspectionId) state.activeWarrantyInspectionId = p.warrantyInspections[p.warrantyInspections.length-1]?.id || null;
    const missing = documentMissing(type, p);
    if (missing.includes('defectPeriodYears') && !(p.warrantyItems || []).length) { showToast('하자담보기간은 별표 4 추천값을 확인한 뒤 적용해주세요.','warn'); openWarrantyManager(); return; }
    if (missing.length) { openDocumentMissingModal(type, missing); return; }
    openModal({
      eyebrow:`${def.stage} 서류 · 양식 ${def.version}`,
      title:`${def.label} 미리보기`,
      wide:true,
      body:`<div class="doc-preview-note">실제 출력될 A4 모습을 확인하세요. 수정이 필요하면 닫고 공사정보에서 고치면 모든 서류에 함께 반영됩니다.</div><div class="doc-preview-scroll">${documentMarkup(type,p)}</div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>닫기</button><button class="button primary" type="button" id="printDocumentBtn">인쇄 / PDF 저장</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#printDocumentBtn').addEventListener('click', () => printAdministrativeDocument(type, p));
  }

  function printPagesInFrame(pages, title, onBeforePrint, options = {}) {
    if (!pages?.length) return;
    const multi = pages.length > 1;
    const landscape = options.orientation === 'landscape';
    const pageWidth = landscape ? '297mm' : '210mm';
    const pageHeight = landscape ? '210mm' : '297mm';
    const frame = document.createElement('iframe');
    frame.className = `document-print-frame${multi?' document-batch-print-frame':''}`;
    frame.setAttribute('aria-hidden','true');
    frame.setAttribute('tabindex','-1');
    frame.style.width = pageWidth;
    frame.style.height = pageHeight;
    document.body.appendChild(frame);

    const cssUrl = new URL('styles.css',window.location.href).href;
    const bodyMarkup = multi
      ? pages.map((page,index)=>`<section class="batch-print-page" data-page="${index+1}">${page}</section>`).join('')
      : pages[0];
    const html = `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${e(title)}</title>
  <link rel="stylesheet" href="${e(cssUrl)}">
  <style>
    @page { size:A4 ${landscape?'landscape':'portrait'}; margin:0; }
    html,body { width:${pageWidth}; ${multi?'':`height:${pageHeight};`} margin:0; padding:0; background:#fff; }
  </style>
</head>
<body class="${multi?'print-batch-documents':'print-only-document'}${landscape?' print-landscape-document':''}">
  ${bodyMarkup}
</body>
</html>`;

    let printed = false;
    const startPrint = () => {
      if (printed || !frame.contentWindow) return;
      printed = true;
      if (typeof onBeforePrint === 'function') onBeforePrint();
      const win = frame.contentWindow;
      const cleanup = () => { if (frame.isConnected) frame.remove(); };
      win.addEventListener('afterprint',cleanup,{once:true});
      window.setTimeout(cleanup,120000);
      window.setTimeout(()=>{
        try { win.focus(); win.print(); }
        catch(err) { cleanup(); showToast('인쇄창을 열지 못했습니다. 다시 시도해주세요.','warn'); }
      }, multi ? 150 : 120);
    };
    const printDoc = frame.contentDocument;
    if (!printDoc) { frame.remove(); showToast('인쇄영역을 만들지 못했습니다. 다시 시도해주세요.','warn'); return; }
    printDoc.open(); printDoc.write(html); printDoc.close();
    const stylesheet = printDoc.querySelector('link[rel="stylesheet"]');
    if (stylesheet) stylesheet.addEventListener('load',startPrint,{once:true});
    window.setTimeout(()=>{ if (printDoc.readyState === 'complete') startPrint(); }, multi ? 600 : 500);
  }

  function printAdministrativeDocument(type, p) {
    const def = DOCUMENT_DEFINITIONS[type];
    if (!def || !p) return;
    const pages = documentPages(type,p);
    printPagesInFrame(pages, `${def.label} - ${p.projectName || '공사서류'}`);
  }

  async function recordPrintHistory(types, p) {
    if (!p || !types?.length) return;
    const history = Array.isArray(p.printHistory) ? [...p.printHistory] : [];
    history.push({
      at:new Date().toISOString(),
      count:types.length,
      types:[...types],
      labels:types.map(type => DOCUMENT_DEFINITIONS[type]?.label).filter(Boolean)
    });
    p.printHistory = history.slice(-20);
    p.updatedAt = new Date().toISOString();
    try { await DB.put('projects',p); }
    catch (err) { console.warn('print history save failed', err); }
  }

  function printAdministrativeDocuments(types, p) {
    const ordered = DOCUMENT_PRINT_ORDER.filter(type => types.includes(type) && DOCUMENT_DEFINITIONS[type]);
    if (!p || !ordered.length) return;
    const incompleteSafety = ordered.find(type => isSafetyDocument(type) && !safetyChecklistComplete(p,type));
    if (incompleteSafety) { showToast(`${DOCUMENT_DEFINITIONS[incompleteSafety].label}을 먼저 작성해주세요.`, 'warn'); return; }
    const missing = batchMissingFields(ordered,p);
    if (missing.length) { openBatchMissingModal(ordered,missing); return; }

    const labels = ordered.map(type => DOCUMENT_DEFINITIONS[type].label);
    const pages = ordered.flatMap(type => documentPages(type,p));
    const title = `${p.projectName || '공사서류'} - ${labels.join(', ')}`;
    printPagesInFrame(pages,title,()=>recordPrintHistory(ordered,p));
  }

  function openDocumentMissingModal(type, missing) {
    const p = currentProject();
    const def = DOCUMENT_DEFINITIONS[type];
    openModal({
      eyebrow:`${def.label} 만들기`, title:`${missing.length}개 정보만 더 입력해주세요`,
      body:`<div class="notice"><strong>이미 저장된 정보는 다시 묻지 않습니다.</strong><br>아래 값은 공사정보에 저장되어 다른 서류에서도 그대로 재사용됩니다. 지급계좌는 업체 지급정보에 별도로 저장됩니다.</div><div class="modal-grid document-missing-grid" style="margin-top:16px">${missing.map(field => documentMissingFieldHtml(field,p)).join('')}</div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="saveMissingDocFields">저장하고 ${e(def.label)} 보기</button>`
    });
    initDateInputs(modalBody);
    initMoneyInputs(modalBody);
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#saveMissingDocFields').addEventListener('click', () => saveDocumentMissingFields(type, missing));
  }

  async function saveDocumentMissingFields(type, fields) {
    const ok = await saveFieldsForDocuments(fields);
    if (!ok) return;
    closeModal();
    renderProjectDetail();
    openDocumentPreview(type);
  }

  async function savePayoutForProject(p, values = {}) {
    let vendor = vendorForProject(p);
    if (!vendor && p?.vendorName) vendor = await upsertVendorFromProject(p);
    if (!vendor) throw new Error('업체정보를 먼저 입력해주세요.');
    const current = state.payouts.find(x => x.vendorId === vendor.id) || null;
    const next = DB.createPayout({
      ...(current || {}),
      vendorId: vendor.id,
      bankName: values.bankName ?? current?.bankName ?? '',
      accountNumber: values.accountNumber ?? current?.accountNumber ?? '',
      accountHolder: values.accountHolder ?? current?.accountHolder ?? p.vendorName ?? '',
      updatedAt: new Date().toISOString()
    });
    await DB.put('payouts', next);
    const idx = state.payouts.findIndex(x => x.vendorId === vendor.id);
    if (idx >= 0) state.payouts[idx] = next; else state.payouts.push(next);
    return next;
  }


  const UTILITY_RATES_2024 = {
    work: {
      '건축': {electric:0.515,water:0.479}, '토목': {electric:0.346,water:0.597},
      '산업설비': {electric:0.163,water:0.422}, '조경': {electric:0.370,water:0.360}
    },
    duration: {
      '6이하': {electric:0.175,water:0.232}, '6초과12이하': {electric:0.212,water:0.410},
      '12초과36이하': {electric:0.505,water:0.509}, '36초과': {electric:0.715,water:0.596}
    },
    size: [
      {max:500000000,electric:0.132,water:0.154,label:'5억 미만'},
      {max:3000000000,electric:0.176,water:0.199,label:'5억~30억 미만'},
      {max:5000000000,electric:0.269,water:0.331,label:'30억~50억 미만'},
      {max:30000000000,electric:0.448,water:0.611,label:'50억~300억 미만'},
      {max:100000000000,electric:0.506,water:0.726,label:'300억~1,000억 미만'}
    ]
  };

  function utilityWorkCategory(workType) {
    const t=String(workType||'');
    if (t.includes('토목')) return '토목';
    if (t.includes('산업설비')) return '산업설비';
    if (t.includes('조경')) return '조경';
    return '건축'; // 전기·통신·소방·전문공사는 원본 안내대로 건축요율 적용
  }

  function utilityDurationCategory(p) {
    const a=parseIsoDate(p?.startDate || p?.plannedStartDate), b=parseIsoDate(p?.completionDueDate);
    if (!a || !b) return '6이하';
    const months=Math.max(0,(b-a)/86400000/30.4375);
    if (months<=6) return '6이하'; if (months<=12) return '6초과12이하'; if (months<=36) return '12초과36이하'; return '36초과';
  }

  function utilitySizeRate(contractAmount) {
    const base=Number(contractAmount||0)/1.1; // 원본 계산식: 부가세 제외 계약금액 기준
    return UTILITY_RATES_2024.size.find(x=>base<x.max) || UTILITY_RATES_2024.size[UTILITY_RATES_2024.size.length-1];
  }

  function roundDown10(value) { return Math.floor(Math.max(0,Number(value)||0)/10)*10; }

  function calculateUtilityCost({directMaterialCost,directLaborCost,facilityUse,workCategory,durationCategory,contractAmount}) {
    const base=Number(directMaterialCost||0)+Number(directLaborCost||0);
    const work=UTILITY_RATES_2024.work[workCategory] || UTILITY_RATES_2024.work['건축'];
    const duration=UTILITY_RATES_2024.duration[durationCategory] || UTILITY_RATES_2024.duration['6이하'];
    const size=utilitySizeRate(contractAmount);
    const average=(kind)=>(Number(work[kind])+Number(duration[kind])+Number(size[kind]))/3/100;
    const allowElectric=facilityUse!=='수도광열비';
    const allowWater=facilityUse!=='전력비';
    const electric=allowElectric?roundDown10(base*average('electric')):0;
    const water=allowWater?roundDown10(base*average('water')):0;
    return {
      electricCost:electric, waterHeatCost:water, total:electric+water, sizeLabel:size.label,
      baseCost:base, contractAmountExVat:Number(contractAmount||0)/1.1,
      workElectricRate:Number(work.electric), durationElectricRate:Number(duration.electric), sizeElectricRate:Number(size.electric),
      workWaterRate:Number(work.water), durationWaterRate:Number(duration.water), sizeWaterRate:Number(size.water),
      allowElectric, allowWater
    };
  }

  function utilityResultHtml(result) {
    if (!result) return '<div class="utility-result-empty">금액을 입력하고 계산해주세요.</div>';
    return `<div class="utility-result-grid"><div><span>전력비</span><strong>${e(formatMoney(result.electricCost))}</strong></div><div><span>수도광열비</span><strong>${e(formatMoney(result.waterHeatCost))}</strong></div><div class="utility-total"><span>공제금액 합계</span><strong>${e(formatMoney(result.total))}</strong></div></div><p class="utility-rate-note">공사규모 구간 · ${e(result.sizeLabel)}</p>`;
  }

  function utilityInputsFromModal(p) {
    return {
      directMaterialCost:parseMoneyInput(modalBody.querySelector('#utilityMaterial')?.value || ''),
      directLaborCost:parseMoneyInput(modalBody.querySelector('#utilityLabor')?.value || ''),
      facilityUse:modalBody.querySelector('#utilityFacilityUse')?.value || '수도광열비·전력비',
      workCategory:modalBody.querySelector('#utilityWorkCategory')?.value || utilityWorkCategory(p.workType),
      durationCategory:modalBody.querySelector('#utilityDurationCategory')?.value || utilityDurationCategory(p),
      contractAmount:Number(p.currentContractAmount||0)
    };
  }

  function openUtilityCalculator() {
    const p=currentProject(); if(!p)return;
    const u=p.utilityCost||{};
    const workCategory=u.workCategory||utilityWorkCategory(p.workType);
    const durationCategory=u.durationCategory||utilityDurationCategory(p);
    const existing=meaningful(u.total)?{electricCost:Number(u.electricCost||0),waterHeatCost:Number(u.waterHeatCost||0),total:Number(u.total||0),sizeLabel:utilitySizeRate(p.currentContractAmount).label}:null;
    openModal({eyebrow:'행정기관 계산도구 · 2024 완성공사 원가통계',title:'전력비·수도광열비 계산',wide:true,
      body:`<div class="notice"><strong>원본 「수도전기료계산식」의 산식을 웹으로 옮겼습니다.</strong><br>전기·통신·소방·전문공사는 건축요율을 적용하고, 공사기간과 계약금액 구간은 현재 공사정보에서 자동 판단합니다.</div><div class="modal-grid utility-input-grid" style="margin-top:16px">
        <div class="field"><label>공사종류 요율</label><select id="utilityWorkCategory">${['건축','토목','산업설비','조경'].map(x=>`<option ${x===workCategory?'selected':''}>${x}</option>`).join('')}</select></div>
        <div class="field"><label>공사기간 요율</label><select id="utilityDurationCategory">${['6이하','6초과12이하','12초과36이하','36초과'].map(x=>`<option ${x===durationCategory?'selected':''}>${x}</option>`).join('')}</select></div>
        <div class="field"><label>적용 항목</label><select id="utilityFacilityUse"><option ${u.facilityUse==='수도광열비'?'selected':''}>수도광열비</option><option ${u.facilityUse==='전력비'?'selected':''}>전력비</option><option ${!u.facilityUse||u.facilityUse==='수도광열비·전력비'?'selected':''}>수도광열비·전력비</option></select></div>
        <div class="field"><label>현재 계약금액</label><input value="${e(formatMoneyInput(p.currentContractAmount))}" disabled><span class="hint">부가세 제외 금액으로 공사규모 요율을 자동 판단합니다.</span></div>
        <div class="field"><label for="utilityMaterial">직접재료비</label>${moneyInputHtml('utilityMaterial',u.directMaterialCost||'')}</div>
        <div class="field"><label for="utilityLabor">직접노무비</label>${moneyInputHtml('utilityLabor',u.directLaborCost||'')}</div>
      </div><div id="utilityResult" class="utility-result-panel">${utilityResultHtml(existing)}</div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>닫기</button><button class="button secondary" type="button" id="calculateUtilityBtn">계산·저장</button><button class="button secondary" type="button" id="printUtilityBtn">계산식 인쇄 / PDF</button><button class="button primary" type="button" id="applyUtilityDeductionBtn">공제금액에 반영</button>`});
    initMoneyInputs(modalBody);
    modalActions.querySelector('[data-modal-close]').addEventListener('click',closeModal);
    modalActions.querySelector('#calculateUtilityBtn').addEventListener('click',()=>saveUtilityCalculation(false));
    modalActions.querySelector('#printUtilityBtn').addEventListener('click',printUtilityCalculationFromModal);
    modalActions.querySelector('#applyUtilityDeductionBtn').addEventListener('click',()=>saveUtilityCalculation(true));
  }

  async function saveUtilityCalculation(applyDeduction=false) {
    const p=currentProject(); if(!p)return;
    const inputs=utilityInputsFromModal(p);
    if (!meaningful(inputs.directMaterialCost) && !meaningful(inputs.directLaborCost)) { showToast('직접재료비 또는 직접노무비를 입력해주세요.','warn'); return; }
    const result=calculateUtilityCost(inputs);
    p.utilityCost={...inputs,...result,calculatedAt:new Date().toISOString()};
    if (applyDeduction) p.deductionAmount=result.total;
    p.updatedAt=new Date().toISOString();
    await DB.put('projects',p); await loadState(); state.currentProjectId=p.id;
    const resultEl=modalBody.querySelector('#utilityResult'); if(resultEl) resultEl.innerHTML=utilityResultHtml(result);
    showToast(applyDeduction ? `공제금액 ${formatMoney(result.total)}에 반영했습니다.` : '전력비·수도광열비 계산값을 저장했습니다.');
    if (applyDeduction) { closeModal(); renderProjectDetail(); }
  }

  function utilityRateTableHtml() {
    const work = UTILITY_RATES_2024.work;
    const duration = UTILITY_RATES_2024.duration;
    const size = UTILITY_RATES_2024.size;
    const pct = value => Number(value || 0).toFixed(3).replace(/0+$/,'').replace(/\.$/,'');
    return `<table class="utility-print-rate-table"><thead>
      <tr><th rowspan="2">구분</th><th colspan="4">공사종류별</th><th colspan="4">공사기간별</th><th colspan="5">공사규모별</th></tr>
      <tr><th>건축</th><th>토목</th><th>산업설비</th><th>조경</th><th>6개월이하</th><th>6~12개월</th><th>12~36개월</th><th>36개월초과</th><th>5억미만</th><th>5~30억</th><th>30~50억</th><th>50~300억</th><th>300~1,000억</th></tr>
    </thead><tbody>
      <tr><th>전력비</th><td>${pct(work['건축'].electric)}</td><td>${pct(work['토목'].electric)}</td><td>${pct(work['산업설비'].electric)}</td><td>${pct(work['조경'].electric)}</td><td>${pct(duration['6이하'].electric)}</td><td>${pct(duration['6초과12이하'].electric)}</td><td>${pct(duration['12초과36이하'].electric)}</td><td>${pct(duration['36초과'].electric)}</td>${size.map(x=>`<td>${pct(x.electric)}</td>`).join('')}</tr>
      <tr><th>수도광열비</th><td>${pct(work['건축'].water)}</td><td>${pct(work['토목'].water)}</td><td>${pct(work['산업설비'].water)}</td><td>${pct(work['조경'].water)}</td><td>${pct(duration['6이하'].water)}</td><td>${pct(duration['6초과12이하'].water)}</td><td>${pct(duration['12초과36이하'].water)}</td><td>${pct(duration['36초과'].water)}</td>${size.map(x=>`<td>${pct(x.water)}</td>`).join('')}</tr>
    </tbody></table>`;
  }

  function utilityCalculationPrintMarkup(p, inputs, result) {
    const won = value => Number(value || 0).toLocaleString('ko-KR');
    const pct = value => `${Number(value || 0).toFixed(3).replace(/0+$/,'').replace(/\.$/,'')}%`;
    const formula = (label, enabled, workRate, durationRate, sizeRate, amount) => enabled
      ? `<div class="utility-print-formula"><strong>${label}</strong><span>= (${won(inputs.directMaterialCost)} + ${won(inputs.directLaborCost)}) × ((${pct(workRate)} + ${pct(durationRate)} + ${pct(sizeRate)}) / 3)</span><b>= ${won(amount)} 원</b></div>`
      : `<div class="utility-print-formula disabled"><strong>${label}</strong><span>적용 제외</span><b>= 0 원</b></div>`;
    return `<article class="paper-a4-landscape utility-cost-sheet document-print-page">
      <div class="utility-print-project"><strong>[공사명]</strong><span>${e(p.projectName || '')}</span></div>
      <h1>□ 2024년도 기준 완성공사 원가통계(경비율)</h1>
      ${utilityRateTableHtml()}
      <div class="utility-print-notes">
        <p>※ 공사종류에서 전기·통신·소방·전문공사는 건축요율 적용</p>
        <p>※ 공사규모 금액은 공사 계약금액 기준(부가세 제외)</p>
        <p>※ 출처: 대한건설협회 「2024년 완성공사원가분석」 기준</p>
      </div>
      <h1>□ 전력비·수도광열비 계산식</h1>
      <table class="utility-print-inputs"><tbody>
        <tr><th>공사종류</th><td>${e(inputs.workCategory)}</td><th>공사기간</th><td>${e(inputs.durationCategory)}</td><th>공사규모</th><td>${e(result.sizeLabel)}</td><th>시설사용</th><td>${e(inputs.facilityUse)}</td></tr>
        <tr><th>계약금액</th><td>${won(inputs.contractAmount)} 원</td><th>직접재료비</th><td>${won(inputs.directMaterialCost)} 원</td><th>직접노무비</th><td>${won(inputs.directLaborCost)} 원</td><th>부가세 제외 계약금액</th><td>${won(result.contractAmountExVat)} 원</td></tr>
      </tbody></table>
      <div class="utility-print-formulas">
        ${formula('1. 전력비', result.allowElectric, result.workElectricRate, result.durationElectricRate, result.sizeElectricRate, result.electricCost)}
        ${formula('2. 수도광열비', result.allowWater, result.workWaterRate, result.durationWaterRate, result.sizeWaterRate, result.waterHeatCost)}
      </div>
      <div class="utility-print-total"><span>합 계 :</span><strong>${won(result.total)} 원</strong></div>
      <p class="utility-print-footnote">공사정보 허브에 저장된 공사정보와 입력한 직접재료비·직접노무비를 기준으로 계산</p>
    </article>`;
  }

  async function printUtilityCalculationFromModal() {
    const p=currentProject(); if(!p)return;
    const inputs=utilityInputsFromModal(p);
    if (!meaningful(inputs.directMaterialCost) && !meaningful(inputs.directLaborCost)) { showToast('직접재료비 또는 직접노무비를 입력해주세요.','warn'); return; }
    const result=calculateUtilityCost(inputs);
    p.utilityCost={...inputs,...result,calculatedAt:new Date().toISOString()};
    p.updatedAt=new Date().toISOString();
    await DB.put('projects',p);
    const resultEl=modalBody.querySelector('#utilityResult'); if(resultEl) resultEl.innerHTML=utilityResultHtml(result);
    const page=utilityCalculationPrintMarkup(p,inputs,result);
    printPagesInFrame([page], `수도광열비 계산식 - ${p.projectName || '공사'}`, null, {orientation:'landscape'});
  }

  function formatKoreanDate(value) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value || '');
    return m ? `${Number(m[1])}년 ${Number(m[2])}월 ${Number(m[3])}일` : e(value || '');
  }

  function documentMoney(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) return e(value || '');
    const words = koreanMoney(n).replace(/^금\s*/, '').replace(/원$/, '').replace(/\s+/g, '');
    return `금 ${n.toLocaleString('ko-KR')} 원(금 ${words} 원정)`;
  }

  function representativeWithSeal(value) {
    const text = String(value || '').trim();
    if (!text) return '';
    return /\(인\)|（인）/.test(text) ? text : `${text}  (인)`;
  }

  function recipientFor(name) {
    const n = String(name || '').trim();
    if (!n) return '';
    if (n.endsWith('교육지원청')) return `${n} 교육장 귀하`;
    if (n.endsWith('교육청')) return `${n} 교육감 귀하`;
    return `${n}장 귀하`;
  }

  function documentFacts(rows) {
    return `<div class="doc-facts">${rows.map(([label,value]) => `<div class="doc-fact-row"><span class="doc-fact-label">${e(label)}</span><span class="doc-fact-value">${value}</span></div>`).join('')}</div>`;
  }

  function documentVendorBlock(p, inspection = false) {
    const business = Excel.normalizeBusinessNumber(p.businessNumber || '');
    const labels = inspection
      ? [['업   체   명',p.vendorName],['사업자등록번호',business],['주        소',p.vendorAddress],['대   표   자',representativeWithSeal(p.representative)]]
      : [['회     사     명',p.vendorName],['사업자등록번호',business],['주             소',p.vendorAddress],['대     표     자',representativeWithSeal(p.representative)]];
    return `<div class="doc-vendor-block ${inspection?'inspection':''}">${labels.map(([label,value]) => `<div><span>${e(label)} :</span><strong>${e(value)}</strong></div>`).join('')}</div>`;
  }

  function moneyNumberText(value) {
    const n = Number(value);
    return Number.isFinite(n) ? `금 ${n.toLocaleString('ko-KR')} 원` : '';
  }

  function moneyWordsText(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) return '';
    const words = koreanMoney(n).replace(/^금\s*/, '').replace(/원$/, '').replace(/\s+/g, '');
    return `(금 ${words} 원정)`;
  }

  function claimAmountFor(p) {
    const contract = Number(p.currentContractAmount);
    const prior = Number(p.priorPaymentAmount || 0);
    const deduction = Number(p.deductionAmount || 0);
    if (![contract, prior, deduction].every(Number.isFinite)) return '';
    return Math.max(0, contract - prior - deduction);
  }

  function percentText(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) return String(value ?? '');
    const pct = Math.abs(n) <= 1 ? n * 100 : n;
    return Number.isInteger(pct) ? String(pct) : String(Number(pct.toFixed(4)));
  }

  function schoolLeaderLabel(name) {
    const n = String(name || '').trim();
    if (n.endsWith('교육지원청')) return '교육장';
    if (n.endsWith('교육청')) return '교육감';
    if (n.endsWith('학교')) return '교장';
    return '대표';
  }

  function documentContext(p) {
    return {
      project: p,
      school: state.school || {},
      payout: payoutForProject(p) || {},
      value: field => documentValue(field, p),
      helpers: {
        e,
        formatKoreanDate,
        documentMoney,
        representativeWithSeal,
        recipientFor,
        documentFacts,
        documentVendorBlock,
        moneyNumberText,
        moneyWordsText,
        claimAmountFor,
        businessNumber: value => Excel.normalizeBusinessNumber(value || ''),
        percentText,
        schoolLeaderLabel
      }
    };
  }

  function documentPages(type, p) {
    return Documents.renderPages(type, documentContext(p));
  }

  function documentMarkup(type, p) {
    return documentPages(type, p).join('<div class="document-page-gap" aria-hidden="true"></div>');
  }

  function sectionForStatus(statusKey) {
    const p = currentProject();
    if (statusKey === 'contract_prep') return (!p?.projectName || !p?.vendorName) ? 'basic' : 'contract';
    if (statusKey === 'start_wait') return 'start';
    if (['active','inspection_wait'].includes(statusKey)) return 'completion';
    if (statusKey === 'payment_wait') return 'payment';
    return 'defect';
  }

  function jumpToSection(key) {
    const target = document.getElementById(`section-${key}`);
    if (!target) return;
    if (target.tagName === 'DETAILS') target.open = true;
    target.scrollIntoView({ behavior:'smooth', block:'start' });
    setTimeout(() => target.querySelector('input:not([type="hidden"]), select, textarea')?.focus({ preventScroll:true }), 380);
  }

  function workflowStepHtml(step) {
    const stateClass = step.done ? 'done' : step.active ? 'active' : 'future';
    const icon = step.done ? '✓' : step.active ? '●' : '○';
    return `<button type="button" class="workflow-step ${stateClass}" data-jump-section="${e(step.key)}"><span class="workflow-icon">${icon}</span><span class="workflow-copy"><strong>${e(step.label)}</strong><small>${e(step.summary)}</small></span></button>`;
  }

  function workflowSectionHtml(key, title, description, step, fields, open = false) {
    const stateClass = step.done ? 'done' : step.active ? 'active' : 'future';
    const stateText = step.done ? '입력 완료' : step.active ? (step.stateText || '지금 입력') : '필요할 때 입력';
    return `<details class="section workflow-section ${stateClass}" id="section-${e(key)}" ${open?'open':''}>
      <summary class="section-head"><div><h2>${e(title)}</h2><p>${e(description)}</p></div><div class="section-state ${stateClass}">${e(stateText)}</div></summary>
      <div class="form-grid">${fields.join('')}</div>
    </details>`;
  }

  function contractChangeHistoryHtml(p) {
    const changes = Array.isArray(p.contractChanges) ? p.contractChanges : [];
    if (!changes.length) return `<div class="contract-change-empty">변경계약이 생기면 여기에서 금액·준공기한 변경 이력을 남길 수 있습니다.</div>`;
    return `<div class="contract-change-list">${changes.map((c, index) => {
      const amountText = meaningful(c.afterAmount) ? `${formatMoney(c.beforeAmount)} → ${formatMoney(c.afterAmount)}` : '금액 변경 없음';
      const dueText = c.afterCompletionDueDate ? `${formatDate(c.beforeCompletionDueDate)} → ${formatDate(c.afterCompletionDueDate)}` : '준공기한 변경 없음';
      return `<div class="contract-change-item"><div class="change-index">${index+1}차</div><div class="change-copy"><strong>${e(formatDate(c.changeDate) || '변경일 미입력')}</strong><span>${e(amountText)}</span><span>${e(dueText)}</span>${c.reason ? `<small>${e(c.reason)}</small>` : ''}</div><button class="button ghost small" type="button" data-change-delete="${e(c.id)}">삭제</button></div>`;
    }).join('')}</div>`;
  }


  function warrantyInspectionHistoryHtml(p) {
    const list = Array.isArray(p?.warrantyInspections) ? [...p.warrantyInspections].reverse() : [];
    if (!list.length) return `<div class="contract-change-empty">아직 하자검사 기록이 없습니다. 검사할 때마다 새 기록을 추가하세요.</div>`;
    return `<div class="warranty-history-list">${list.map((x,idx)=>`<div class="warranty-history-item"><div class="warranty-history-copy"><strong>${e(formatDate(x.date) || '검사일 미입력')}</strong><span>${e(x.hasDefect === 'yes' ? '하자 있음' : x.hasDefect === 'no' ? '이상 없음' : (x.result || '검사결과 미입력'))}</span>${x.result ? `<small>${e(x.result)}</small>` : ''}</div><div class="warranty-history-actions"><button class="button ghost small" type="button" data-warranty-preview="${e(x.id)}">조서 보기</button><button class="button ghost small" type="button" data-warranty-delete="${e(x.id)}">삭제</button></div></div>`).join('')}</div>`;
  }

  function warrantyItemsHtml(p) {
    const items = Array.isArray(p?.warrantyItems) ? p.warrantyItems : [];
    if (!items.length) return `<div class="contract-change-empty">세부공종을 등록하면 별표 4 기준 하자담보기간을 추천하고 공종별 종료일을 따로 관리할 수 있습니다.</div>`;
    return `<div class="warranty-item-list">${items.map(item=>`<div class="warranty-item"><div class="warranty-item-main"><strong>${e(item.subcategory || item.category || '하자공종')}</strong><span>${e(item.category || '')}${item.years ? ` · ${e(item.years)}년` : ''}${item.endDate ? ` · ${e(formatDate(item.endDate))} 종료` : ''}</span>${item.recommendedYears ? `<small>${Number(item.years)===Number(item.recommendedYears)&&!item.manuallyModified?'별표 4 추천값 적용':`추천 ${e(item.recommendedYears)}년 · 사용자 확인값`}</small>` : '<small>기존 입력값</small>'}</div><button class="button ghost small" type="button" data-warranty-item-delete="${e(item.id)}">삭제</button></div>`).join('')}</div>`;
  }

  function warrantyCategoryOptions(selected='') {
    return ReferenceData.warrantyRules.map(x=>`<option value="${e(x.category)}" ${x.category===selected?'selected':''}>${e(x.category)}</option>`).join('');
  }

  function warrantySubcategoryOptions(category, selected='') {
    const group = ReferenceData.warrantyCategory(category) || ReferenceData.warrantyRules[0];
    return (group?.items || []).map(x=>`<option value="${e(x.subcategory)}" ${x.subcategory===selected?'selected':''}>${e(x.subcategory)}</option>`).join('');
  }

  function openWarrantyManager() {
    const p = currentProject(); if (!p) return;
    const firstCategory = ReferenceData.warrantyRules[0]?.category || '';
    const defaultStart = p.actualCompletionDate || p.defectStartDate || '';
    const items = Array.isArray(p.warrantyItems) ? p.warrantyItems : [];
    openModal({
      eyebrow:'하자관리 · 건설산업기본법 시행령 별표 4', title:'하자담보 공종 관리', wide:true,
      body:`<div class="notice"><strong>하자담보기간은 자동 확정하지 않습니다.</strong><br>공사 종류와 세부공종을 고르면 별표 4의 기간을 추천합니다. 적용 전에 확인하거나 직접 수정할 수 있으며, 복합공사는 세부공종별로 여러 건 등록할 수 있습니다.</div>
        <div class="warranty-manager-grid" style="margin-top:16px">
          <div class="field"><label for="warrantyCategory">공사 종류</label><select id="warrantyCategory">${warrantyCategoryOptions(firstCategory)}</select></div>
          <div class="field"><label for="warrantySubcategory">세부공종</label><select id="warrantySubcategory">${warrantySubcategoryOptions(firstCategory)}</select></div>
          <div class="field"><label for="warrantyRecommendedYears">추천기간</label><div class="recommend-readonly"><strong id="warrantyRecommendedYears"></strong><span>별표 4 기준</span></div></div>
          <div class="field"><label for="warrantyYears">적용기간(년)</label><input id="warrantyYears" type="number" min="0" step="1" inputmode="numeric"><span class="hint">추천값을 그대로 쓰거나 직접 변경할 수 있습니다.</span></div>
          ${modalDateField('warrantyItemStart','하자 시작일',defaultStart)}
          <div class="field"><label for="warrantyItemEnd_text">하자 종료일</label>${dateInputHtml('warrantyItemEnd','')}<input id="warrantyItemEnd" type="hidden"><span class="hint">시작일 + 적용기간으로 자동 계산합니다.</span></div>
          <div class="field full"><label for="warrantyItemNote">비고</label><input id="warrantyItemNote" placeholder="필요한 경우만 입력"></div>
        </div>
        <div class="warranty-manager-actions"><button class="button secondary" type="button" id="applyWarrantyRecommendationBtn">추천기간 적용</button><button class="button primary" type="button" id="addWarrantyItemBtn">+ 하자공종 추가</button></div>
        <div class="warranty-manager-existing"><div class="subsection-head"><div><strong>현재 등록된 하자공종 ${items.length}건</strong><span>복합공사는 공종별 책임기간을 각각 관리합니다.</span></div></div>${warrantyItemsHtml(p)}</div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>닫기</button>`
    });
    initDateInputs(modalBody);
    const categoryEl=modalBody.querySelector('#warrantyCategory'), subEl=modalBody.querySelector('#warrantySubcategory'), yearsEl=modalBody.querySelector('#warrantyYears'), recEl=modalBody.querySelector('#warrantyRecommendedYears');
    const syncRecommendation=()=>{
      const rec=ReferenceData.warrantyRecommendation(categoryEl.value,subEl.value);
      recEl.textContent=rec?`${rec.years}년`:'확인 필요';
      yearsEl.dataset.recommended=rec?.years || '';
    };
    const syncEnd=()=>{
      const start=modalBody.querySelector('#warrantyItemStart')?.value || '';
      const years=Number(yearsEl.value || 0);
      const end=addYearsMinusDay(start,years);
      const hidden=modalBody.querySelector('#warrantyItemEnd'), text=modalBody.querySelector('#warrantyItemEnd_text'), picker=modalBody.querySelector('#warrantyItemEnd_picker');
      if(hidden) hidden.value=end; if(text) text.value=end; if(picker) picker.value=end;
    };
    categoryEl.addEventListener('change',()=>{subEl.innerHTML=warrantySubcategoryOptions(categoryEl.value);syncRecommendation();syncEnd();});
    subEl.addEventListener('change',()=>{syncRecommendation();syncEnd();});
    yearsEl.addEventListener('input',syncEnd);
    modalBody.querySelector('#warrantyItemStart_text')?.addEventListener('blur',()=>setTimeout(syncEnd,0));
    modalBody.querySelector('#warrantyItemStart')?.addEventListener('change',syncEnd);
    modalBody.querySelector('#applyWarrantyRecommendationBtn').addEventListener('click',()=>{const rec=ReferenceData.warrantyRecommendation(categoryEl.value,subEl.value); if(rec){yearsEl.value=rec.years;syncEnd();showToast(`${rec.years}년 추천값을 입력했습니다. 저장 전 수정할 수 있습니다.`);}});
    modalBody.querySelector('#addWarrantyItemBtn').addEventListener('click',saveWarrantyItemFromModal);
    modalBody.querySelectorAll('[data-warranty-item-delete]').forEach(btn=>btn.addEventListener('click',()=>deleteWarrantyItem(btn.dataset.warrantyItemDelete,true)));
    modalActions.querySelector('[data-modal-close]').addEventListener('click',closeModal);
    syncRecommendation();
  }

  async function saveWarrantyItemFromModal() {
    const p=currentProject(); if(!p)return;
    const category=modalBody.querySelector('#warrantyCategory')?.value || '';
    const subcategory=modalBody.querySelector('#warrantySubcategory')?.value || '';
    const years=Number(modalBody.querySelector('#warrantyYears')?.value || 0);
    const startDate=modalBody.querySelector('#warrantyItemStart')?.value || '';
    const recommended=ReferenceData.warrantyRecommendation(category,subcategory)?.years || '';
    if(!category || !subcategory || !years || !startDate){showToast('공사 종류·세부공종·적용기간·시작일을 확인해주세요.','warn');return;}
    const endDate=addYearsMinusDay(startDate,years);
    const list=Array.isArray(p.warrantyItems)?[...p.warrantyItems]:[];
    list.push({id:DB.uuid(),category,subcategory,years,startDate,endDate,recommendedYears:recommended,recommendationApplied:Number(years)===Number(recommended),manuallyModified:Number(years)!==Number(recommended),note:modalBody.querySelector('#warrantyItemNote')?.value?.trim()||'',createdAt:new Date().toISOString()});
    p.warrantyItems=list;
    const starts=list.map(x=>x.startDate).filter(Boolean).sort(); const ends=list.map(x=>x.endDate).filter(Boolean).sort();
    p.defectStartDate=starts[0]||p.defectStartDate||''; p.defectEndDate=ends[ends.length-1]||p.defectEndDate||''; p.defectPeriodYears=Math.max(...list.map(x=>Number(x.years)||0),0)||p.defectPeriodYears||'';
    p.updatedAt=new Date().toISOString(); await DB.put('projects',p); await loadState(); state.currentProjectId=p.id;
    closeModal(); renderProjectDetail(); showToast(`${subcategory} 하자기간 ${years}년을 등록했습니다.`);
  }

  async function deleteWarrantyItem(id, reopen=false) {
    const p=currentProject(); if(!p||!id)return;
    p.warrantyItems=(p.warrantyItems||[]).filter(x=>x.id!==id);
    const list=p.warrantyItems; const starts=list.map(x=>x.startDate).filter(Boolean).sort(); const ends=list.map(x=>x.endDate).filter(Boolean).sort();
    p.defectStartDate=starts[0]||''; p.defectEndDate=ends[ends.length-1]||''; p.defectPeriodYears=list.length?Math.max(...list.map(x=>Number(x.years)||0),0):'';
    p.updatedAt=new Date().toISOString(); await DB.put('projects',p); await loadState(); state.currentProjectId=p.id; closeModal(); renderProjectDetail(); if(reopen)openWarrantyManager();
  }

  function openWarrantyInspectionModal(record = null, previewAfterSave = false) {
    const p = currentProject();
    if (!p) return;
    const current = record || {};
    const inspector = current.inspector || '';
    const witness = current.witness || p.witness || state.school?.witness || '';
    openModal({
      eyebrow:'행정기관 하자관리', title:record ? '하자검사 기록 수정' : '하자검사 기록 추가',
      body:`<div class="notice"><strong>검사기록은 누적됩니다.</strong><br>새 검사를 추가해도 이전 검사기록은 사라지지 않습니다.</div><div class="modal-grid" style="margin-top:16px">
        ${modalDateField('warrantyDate','검사일',current.date || '')}
        <div class="field"><label for="warrantyInspectorInput">검사자 <span class="label-optional">선택</span></label><input id="warrantyInspectorInput" value="${e(inspector)}" placeholder="출력 후 직접 기재할 경우 비워두세요"><span class="hint">기본값을 자동 입력하지 않습니다.</span></div>
        ${modalField('warrantyWitnessInput','입회자',witness)}
        <div class="field"><label for="warrantyHasDefect">하자 유무</label><select id="warrantyHasDefect"><option value="">선택</option><option value="no" ${current.hasDefect==='no'?'selected':''}>이상 없음</option><option value="yes" ${current.hasDefect==='yes'?'selected':''}>하자 있음</option></select></div>
        <div class="field full"><label for="warrantyResult">검사결과</label><textarea id="warrantyResult">${e(current.result || '')}</textarea></div>
        <div class="field full"><label for="warrantyIssueDetails">하자발생내용</label><textarea id="warrantyIssueDetails">${e(current.issueDetails || '')}</textarea></div>
        <div class="field full"><label for="warrantyActions">처리사항</label><textarea id="warrantyActions">${e(current.actions || '')}</textarea></div>
        <div class="field full"><label for="warrantyNotes">기타참고사항</label><textarea id="warrantyNotes">${e(current.notes || '')}</textarea></div>
      </div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="saveWarrantyInspectionBtn">저장</button>`
    });
    initDateInputs(modalBody);
    modalActions.querySelector('[data-modal-close]').addEventListener('click',closeModal);
    modalActions.querySelector('#saveWarrantyInspectionBtn').addEventListener('click',()=>saveWarrantyInspection(record?.id || '',previewAfterSave));
  }

  async function saveWarrantyInspection(recordId = '', previewAfterSave = false) {
    const p = currentProject();
    if (!p) return;
    const date = modalBody.querySelector('#warrantyDate')?.value || '';
    const inspector = modalBody.querySelector('#warrantyInspectorInput')?.value?.trim() || '';
    if (!date) { showToast('검사일을 입력해주세요.','warn'); return; }
    const list = Array.isArray(p.warrantyInspections) ? [...p.warrantyInspections] : [];
    const existingIndex = recordId ? list.findIndex(x=>x.id===recordId) : -1;
    const base = existingIndex >= 0 ? list[existingIndex] : {};
    const next = {
      ...base,
      id: base.id || DB.uuid(),
      date,
      inspector,
      witness: modalBody.querySelector('#warrantyWitnessInput')?.value?.trim() || '',
      hasDefect: modalBody.querySelector('#warrantyHasDefect')?.value || '',
      result: modalBody.querySelector('#warrantyResult')?.value?.trim() || '',
      issueDetails: modalBody.querySelector('#warrantyIssueDetails')?.value?.trim() || '',
      actions: modalBody.querySelector('#warrantyActions')?.value?.trim() || '',
      notes: modalBody.querySelector('#warrantyNotes')?.value?.trim() || '',
      createdAt: base.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    if (existingIndex >= 0) list[existingIndex] = next; else list.push(next);
    p.warrantyInspections = list;
    p.updatedAt = new Date().toISOString();
    await DB.put('projects',p);
    state.activeWarrantyInspectionId = next.id;
    await loadState(); state.currentProjectId=p.id;
    closeModal(); renderProjectDetail();
    showToast('하자검사 기록을 저장했습니다.');
    if (previewAfterSave) openDocumentPreview('warrantyInspectionReport');
  }

  async function deleteWarrantyInspection(id) {
    const p = currentProject();
    if (!p || !id) return;
    if (!confirm('이 하자검사 기록을 삭제할까요?')) return;
    p.warrantyInspections = (p.warrantyInspections || []).filter(x=>x.id!==id);
    p.updatedAt = new Date().toISOString();
    if (state.activeWarrantyInspectionId === id) state.activeWarrantyInspectionId = null;
    await DB.put('projects',p); await loadState(); state.currentProjectId=p.id; renderProjectDetail();
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


  function syncDefectDates(p, changedField = '') {
    const years=Number(p.defectPeriodYears);
    if ((!p.defectStartDate || changedField==='actualCompletionDate') && p.actualCompletionDate && Number.isFinite(years) && years>0) {
      const start=parseIsoDate(p.actualCompletionDate);
      if (start) { start.setDate(start.getDate()+1); p.defectStartDate=isoFromDate(start); }
    }
    if (p.defectStartDate && Number.isFinite(years) && years>0 && ['defectStartDate','defectPeriodYears','actualCompletionDate'].includes(changedField)) {
      p.defectEndDate=addYearsMinusDay(p.defectStartDate,years);
    }
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
    if (fieldName === 'currentContractAmount' && ev.type === 'change' && !(p.contractChanges?.length)) p.originalContractAmount = value;
    if (['defectStartDate','defectPeriodYears','actualCompletionDate'].includes(fieldName)) {
      syncDefectDates(p,fieldName);
      const endHidden=document.getElementById('f_defectEndDate');
      const endText=document.getElementById('f_defectEndDate_text');
      const startHidden=document.getElementById('f_defectStartDate');
      const startText=document.getElementById('f_defectStartDate_text');
      if (startHidden && p.defectStartDate) startHidden.value=p.defectStartDate;
      if (startText && p.defectStartDate) startText.value=p.defectStartDate;
      if (endHidden && p.defectEndDate) endHidden.value=p.defectEndDate;
      if (endText && p.defectEndDate) endText.value=p.defectEndDate;
    }
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
      originalContractAmount: parseMoneyInput(modalBody.querySelector('#newContractAmount').value),
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

  function openContractChangeModal() {
    const p = currentProject();
    if (!p) return;
    const beforeAmount = meaningful(p.currentContractAmount) ? p.currentContractAmount : '';
    openModal({
      eyebrow:'변경계약', title:'변경된 내용만 기록하세요',
      body:`<div class="notice">최초 계약정보를 지우지 않고 변경 이력을 남깁니다. 금액 또는 준공기한 중 변경된 항목만 입력해도 됩니다.</div>
        <div class="modal-grid" style="margin-top:16px">
          ${modalDateField('changeDate','변경계약일')}
          <div class="field"><label>변경 전 계약금액</label><input class="readonly" value="${e(formatMoneyInput(beforeAmount))}" readonly><span class="money-korean">${meaningful(beforeAmount)?e(koreanMoney(beforeAmount)):''}</span></div>
          <div class="field"><label for="changeAmount">변경 후 계약금액</label>${moneyInputHtml('changeAmount','')}</div>
          <div class="field"><label>변경 전 준공기한</label><input class="readonly" value="${e(p.completionDueDate || '')}" readonly></div>
          ${modalDateField('changeDueDate','변경 후 준공기한')}
          <div class="field full"><label>변경 사유</label><textarea id="changeReason" placeholder="예: 설계변경에 따른 계약금액 및 공사기간 변경"></textarea></div>
        </div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="saveContractChangeBtn">변경계약 기록</button>`
    });
    initDateInputs(modalBody);
    initMoneyInputs(modalBody);
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#saveContractChangeBtn').addEventListener('click', saveContractChange);
  }

  async function saveContractChange() {
    const p = currentProject();
    if (!p) return;
    const changeDate = v('#changeDate');
    const afterAmount = parseMoneyInput(modalBody.querySelector('#changeAmount')?.value || '');
    const afterCompletionDueDate = v('#changeDueDate');
    const reason = v('#changeReason');
    if (!changeDate) { showToast('변경계약일을 입력해주세요.', 'warn'); return; }
    if (!meaningful(afterAmount) && !afterCompletionDueDate) { showToast('변경된 계약금액 또는 준공기한을 입력해주세요.', 'warn'); return; }
    const beforeAmount = p.currentContractAmount;
    const beforeCompletionDueDate = p.completionDueDate || '';
    if (meaningful(afterAmount) && meaningful(beforeAmount) && Number(afterAmount) === Number(beforeAmount) && (!afterCompletionDueDate || afterCompletionDueDate === beforeCompletionDueDate)) {
      showToast('현재 값과 동일합니다. 변경된 내용을 입력해주세요.', 'warn'); return;
    }
    if (!Array.isArray(p.contractChanges)) p.contractChanges = [];
    if (!meaningful(p.originalContractAmount) && meaningful(beforeAmount)) p.originalContractAmount = beforeAmount;
    p.contractChanges.push({
      id: DB.uuid(),
      changeDate,
      beforeAmount: meaningful(beforeAmount) ? beforeAmount : '',
      afterAmount: meaningful(afterAmount) ? afterAmount : '',
      beforeCompletionDueDate,
      afterCompletionDueDate,
      reason,
      createdAt: new Date().toISOString()
    });
    if (meaningful(afterAmount)) p.currentContractAmount = afterAmount;
    if (afterCompletionDueDate) p.completionDueDate = afterCompletionDueDate;
    p.updatedAt = new Date().toISOString();
    await DB.put('projects', p);
    closeModal();
    renderProjectDetail();
    showToast('변경계약 이력을 기록했습니다. 현재 계약정보에도 반영했습니다.');
  }

  function confirmDeleteContractChange(changeId) {
    const p = currentProject();
    const changes = Array.isArray(p?.contractChanges) ? p.contractChanges : [];
    const change = changes.find(x => x.id === changeId);
    if (!p || !change) return;
    openModal({
      eyebrow:'변경계약 삭제', title:'이 변경이력을 삭제할까요?',
      body:`<div class="notice warn">변경이력만 삭제합니다. 현재 계약금액과 준공기한은 자동으로 되돌리지 않습니다. 필요하면 현재 계약정보를 직접 확인해주세요.</div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button danger" type="button" id="deleteChangeConfirm">이력 삭제</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#deleteChangeConfirm').addEventListener('click', async () => {
      p.contractChanges = changes.filter(x => x.id !== changeId);
      p.updatedAt = new Date().toISOString();
      await DB.put('projects', p);
      closeModal(); renderProjectDetail(); showToast('변경계약 이력을 삭제했습니다.');
    });
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
    return next;
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
    showToast(`공사 ${data.projects.length}건 · 업체 ${data.vendors.length}건 · 지급정보 ${data.payouts?.length || 0}건을 백업했습니다.`);
  }

  async function handleBackupFile(file) {
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (data.app !== 'construction-info-hub') throw new Error('공사정보 허브 백업파일이 아닙니다.');
      openModal({
        eyebrow:'전체 복원', title:'기존 데이터를 백업파일로 바꿀까요?',
        body:`<div class="notice warn">현재 이 브라우저의 공사·업체·지급정보를 지우고, 백업파일의 내용으로 교체합니다.</div><div class="import-summary"><div class="import-stat"><strong>${data.projects?.length||0}</strong><span>공사</span></div><div class="import-stat"><strong>${data.vendors?.length||0}</strong><span>업체</span></div><div class="import-stat"><strong>${data.payouts?.length||0}</strong><span>지급정보</span></div><div class="import-stat"><strong>${e(data.version||1)}</strong><span>백업 버전</span></div></div>`,
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
        ${modalField('schoolPrincipal','기관 대표자(학교장 등)',s.principal)}
        ${modalField('schoolPhone','대표전화',s.phone)}
        ${modalField('schoolSupervisor','공사감독 기본값',s.supervisor)}
        ${modalField('schoolInspector','검사자 기본값',s.inspector)}
        ${modalField('schoolWitness','준공검사 입회자 기본값',s.witness)}
      </div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="schoolSave">저장</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#schoolSave').addEventListener('click', async()=>{
      const value = { name:v('#schoolName'), type:v('#schoolType'), address:v('#schoolAddress'), principal:v('#schoolPrincipal'), phone:v('#schoolPhone'), supervisor:v('#schoolSupervisor'), inspector:v('#schoolInspector'), witness:v('#schoolWitness') };
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
    const hasPayout = state.payouts.some(x => x.vendorId === v.id && (x.bankName || x.accountNumber || x.accountHolder));
    return `<div class="vendor-row"><div><strong>${e(v.name)}</strong><small>${e([v.representative,v.businessNumber,v.phone].filter(Boolean).join(' · ') || '추가정보 없음')}${hasPayout ? ' · 지급정보 저장됨' : ''}</small></div><div class="actions"><button class="button secondary small" type="button" data-vendor-edit="${e(v.id)}">수정</button><button class="button danger small" type="button" data-vendor-delete="${e(v.id)}">삭제</button></div></div>`;
  }

  function openVendorEdit(id) {
    const vendor = id ? state.vendors.find(x=>x.id===id) : null;
    const payout = vendor ? state.payouts.find(x => x.vendorId === vendor.id) : null;
    openModal({
      eyebrow:'업체 보관함', title:vendor?'업체정보 수정':'업체 추가',
      body:`<div class="modal-grid">${modalField('vendorName','업체명',vendor?.name)}${modalField('vendorRep','대표자',vendor?.representative)}${modalField('vendorBiz','사업자등록번호',vendor?.businessNumber)}${modalField('vendorPhone','대표전화',vendor?.phone)}${modalField('vendorAddress','사업장 주소',vendor?.address,true)}${modalField('vendorLicense','등록면허 / 업종',vendor?.licenseType,true)}
        <div class="field full payout-section-head"><strong>지급정보 <span>선택</span></strong><p>대금청구서에만 사용하며 일반 공사정보와 분리해 이 브라우저에 저장합니다. 업체 목록에는 계좌번호를 표시하지 않습니다.</p></div>
        ${modalField('vendorBank','은행명',payout?.bankName)}${modalField('vendorAccountHolder','예금주명',payout?.accountHolder || vendor?.name)}${modalField('vendorAccount','계좌번호',payout?.accountNumber,true)}
      </div>`,
      actions:`<button class="button secondary" type="button" id="vendorBack">뒤로</button><button class="button primary" type="button" id="vendorSave">저장</button>`
    });
    modalActions.querySelector('#vendorBack').addEventListener('click',openVendorLibrary);
    modalActions.querySelector('#vendorSave').addEventListener('click',async()=>{
      const name=v('#vendorName'); if(!name){showToast('업체명을 입력해주세요.','warn');return;}
      const next=DB.createVendor({...(vendor||{}),id:vendor?.id,name,representative:v('#vendorRep'),businessNumber:Excel.normalizeBusinessNumber(v('#vendorBiz')),phone:v('#vendorPhone'),address:v('#vendorAddress'),licenseType:v('#vendorLicense'),updatedAt:new Date().toISOString()});
      await DB.put('vendors',next);
      const payoutValues = { bankName:v('#vendorBank'), accountHolder:v('#vendorAccountHolder') || name, accountNumber:v('#vendorAccount') };
      if (payoutValues.bankName || payoutValues.accountNumber || payoutValues.accountHolder !== name) {
        await DB.put('payouts', DB.createPayout({...(payout||{}),vendorId:next.id,...payoutValues,updatedAt:new Date().toISOString()}));
      } else if (payout) await DB.remove('payouts', next.id);
      await loadState(); openVendorLibrary(); showToast('업체정보를 저장했습니다.');
    });
  }

  async function deleteVendor(id) {
    const vendor=state.vendors.find(v=>v.id===id); if(!vendor)return;
    if(state.projects.some(p=>p.vendorId===id)){showToast('현재 공사에서 사용 중인 업체입니다. 공사의 업체 연결을 먼저 변경해주세요.','warn');return;}
    await Promise.all([DB.remove('vendors',id), DB.remove('payouts',id)]); await loadState(); openVendorLibrary(); showToast(`${vendor.name}을 업체 보관함에서 삭제했습니다.`);
  }

  function openHelp() {
    openModal({
      eyebrow:'도움말', title:'v0.4.2.2 사용 흐름',
      body:`<div class="notice"><strong>핵심 원칙</strong><br>같은 공사정보는 한 번 입력하고 다시 입력하지 않습니다.</div>
      <div style="display:grid;gap:16px;margin-top:18px;font-size:14px">
        <div><strong>1. 공사를 여러 건 저장</strong><p class="muted">전기·건축·체육관 공사를 동시에 등록해도 각 공사는 독립적으로 자동저장됩니다.</p></div>
        <div><strong>2. 기존 공사이력 재사용</strong><p class="muted">학교 공사 이력 현황.xlsx를 불러오면 여러 공사를 한꺼번에 등록하고 기존 공사의 빈 정보를 보완합니다.</p></div><div><strong>3. 에듀파인으로 업데이트</strong><p class="muted">자료관리목록.xlsx를 다시 내려받아 올리면 계약·준공·지출 단계에서 새로 생긴 값만 기존 공사에 보완합니다. 다른 값은 자동 덮어쓰지 않습니다.</p></div>
        <div><strong>4. 업체 재사용</strong><p class="muted">업체명·대표자·사업자번호·주소·전화·면허를 업체 보관함에 저장해 다음 공사에서 다시 고를 수 있습니다.</p></div>
        <div><strong>5. 공사서류 만들기</strong><p class="muted">공사 상세의 「서류」 탭에서 행정기관 작성·관리 서류를 먼저 확인할 수 있습니다. 공사대장·하자검사조서·하자대장과 기존 계약·착공·준공·지출서류를 같은 공사정보로 만들고 묶음 인쇄할 수 있습니다.</p></div>
        <div><strong>6. 계약서류 세트</strong><p class="muted">공사도급표준계약서·승낙사항·사용인감계·수의계약 통합서약서를 한 번에 선택할 수 있습니다. 수의계약 통합서약서는 원본 양식 구조에 맞춰 2페이지로 출력됩니다.</p></div>
        <div><strong>7. 서류 템플릿 분리</strong><p class="muted">서류별 필수정보·출력순서·양식 버전을 별도 정의로 관리해 이후 양식 변경 시 공사 데이터와 다른 서류에 미치는 영향을 줄였습니다.</p></div>
        <div><strong>8. 지급정보 재사용</strong><p class="muted">은행·계좌·예금주는 업체 지급정보로 분리 저장되며 대금청구서에서만 사용합니다. 업체 목록에는 계좌번호를 노출하지 않습니다.</p></div>
        <div><strong>9. 하자담보기간 추천</strong><p class="muted">건설산업기본법 시행령 별표 4를 참고해 공사종류·세부공종별 하자담보기간을 추천합니다. 추천값은 자동 확정하지 않으며 복합공종은 여러 하자공종으로 각각 관리할 수 있습니다.</p></div>
        <div><strong>10. 안전·보건 서류 추천</strong><p class="muted">공종과 고소·전기·밀폐공간·일반 산업재해 작업 특성에 따라 필요한 체크리스트를 추천합니다. 공통 안전·보건 체크리스트는 기관 점검용, 세부 체크리스트는 업체 작성·기관 확인용으로 구분합니다.</p></div>
        <div><strong>11. 내 서류</strong><p class="muted">이 공사에서 실제 사용하는 서류만 내 서류에 모아두고 다음에 다시 열어도 그대로 유지할 수 있습니다.</p></div>
        <div><strong>12. 인수인계</strong><p class="muted">전체 백업(JSON)은 앱 복원용이고, 공사이력 엑셀은 감사·업무용 결과물입니다.</p></div>
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
