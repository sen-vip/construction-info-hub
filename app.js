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
  const modalAlertRegion = document.getElementById('modalAlertRegion');
  const modalActions = document.getElementById('modalActions');
  const moreMenu = document.getElementById('moreMenu');
  const excelFileInput = document.getElementById('excelFileInput');
  const backupFileInput = document.getElementById('backupFileInput');
  const APP_VERSION = '0.5.1.6';
  const REFERENCE_PROGRAM = '서울시교육청 교육시설안전과 「공사서류 원클릭(간소화)프로그램」(2026.5.수정)';

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
    detailTab: 'documents',
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
  const DATE_FIELDS = new Set(['contractDate','plannedStartDate','startDate','completionDueDate','actualCompletionDate','completionInspectionDate','paymentDate','warrantyInspectionDate','defectStartDate','defectEndDate']);
  const PREVIEW_EXTRA_FIELDS = {
    utilityPaymentPledge:['siteManager'],
    constructionLedger:['contractNumber','supervisor','vendorPhone','representative','contractSecurityAmount','contractSecurityType','fundingSource','plannedStartDate','actualCompletionDate','priorPaymentAmount','paymentDate','paymentAmount','settlementAmount','defectStartDate','defectEndDate','completionInspectionDate','inspector','witness','defectSecurityAmount','defectSecurityType','designer','budgetPolicyProject','budgetUnitProject','budgetDetailProject','budgetDetailItem','costStatisticsItem','budgetCalculationDetails'],
    warrantyInspectionReport:['warrantyIssueDetails','warrantyActions','warrantyNotes','warrantyInspectorName','warrantyWitnessName'],
    warrantyLedger:['supervisor','designer','defectSecurityAmount','completionInspectionDate'],
    paymentRequest:['paymentAmount'],
    safetyGeneral:[], safetyFall:[], safetyElectrical:[], safetyConfined:[], safetyIndustrial:[]
  };
  const PREVIEW_FIELD_GROUPS = [
    ['기관·업체', new Set(['schoolName','schoolAddress','principal','vendorName','businessNumber','vendorAddress','vendorPhone','representative'])],
    ['계약·금액', new Set(['contractNumber','contractMethod','workType','estimatedPrice','originalContractAmount','currentContractAmount','contractSecurityAmount','contractSecurityType','settlementAmount','priorPaymentAmount','deductionAmount','paymentAmount','defectSecurityRate','defectSecurityAmount','defectSecurityType','defectPeriodYears','delayPenaltyRate','priceAdjustmentMethod'])],
    ['일정', DATE_FIELDS],
    ['담당자·하자', new Set(['siteManager','supervisor','inspector','witness','warrantyInspectionResult','warrantyIssueDetails','warrantyActions','warrantyNotes','warrantyInspectorName','warrantyWitnessName'])],
    ['예산·기타', new Set(['fundingSource','designer','budgetPolicyProject','budgetUnitProject','budgetDetailProject','budgetDetailItem','costStatisticsItem','budgetCalculationDetails','bankName','accountNumber','accountHolder','settlementReductionReason'])]
  ];

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
    if (String(p.source || '').includes('edufine') && String(p.source || '').includes('audit')) return '공사관리대장에서 불러온 뒤 K-에듀파인 자료로 보완한 공사';
    if (p.source === 'manual+edufine') return '웹에서 만든 뒤 K-에듀파인 자료로 보완한 공사';
    if (p.source === 'edufine') return 'K-에듀파인에서 불러온 공사';
    if (p.source === 'audit-excel') return '공사관리대장에서 불러온 공사';
    return '웹에서 직접 만든 공사';
  }

  function showToast(message, type = '') {
    if (modal?.open && modalAlertRegion) {
      const div = document.createElement('div');
      div.className = `modal-inline-alert ${type}`.trim();
      div.textContent = message;
      modalAlertRegion.replaceChildren(div);
      modalAlertRegion.hidden = false;
      modalAlertRegion.scrollIntoView({ block:'nearest' });
      window.setTimeout(() => {
        if (div.isConnected) div.remove();
        if (!modalAlertRegion.children.length) modalAlertRegion.hidden = true;
      }, 4200);
      return;
    }
    const region = document.getElementById('toastRegion');
    const div = document.createElement('div');
    div.className = `toast ${type}`.trim();
    div.textContent = message;
    region.appendChild(div);
    setTimeout(() => div.remove(), 3200);
  }

  function clearModalAlert() {
    if (!modalAlertRegion) return;
    modalAlertRegion.replaceChildren();
    modalAlertRegion.hidden = true;
  }

  function openModal({ title, eyebrow = '', body = '', actions = '', wide = false }) {
    modal.classList.toggle('wide-modal', !!wide);
    modalTitle.textContent = title;
    modalEyebrow.textContent = eyebrow;
    clearModalAlert();
    modalBody.innerHTML = body;
    modalActions.innerHTML = actions;
    if (!modal.open) modal.showModal();
  }

  function closeModal() {
    clearModalAlert();
    if (modal.open) modal.close();
  }

  function openExcelPicker(mode = 'auto') {
    state.importMode = mode;
    excelFileInput.click();
  }

  function openExcelDropModal(mode) {
    const isEdufine = mode === 'edufine';
    const title = isEdufine ? '에듀파인 자료 불러오기' : '공사관리대장 불러오기';
    const expected = isEdufine ? '자료관리목록.xlsx' : '공사관리대장.xlsx';
    const guide = isEdufine
      ? `<div class="import-path-guide"><strong>다운로드 경로</strong><span>에듀파인 <b>›</b> 학교회계 <b>›</b> 계약관리 <b>›</b> 계약자료관리 <b>›</b> 자료관리</span></div>`
      : `<div class="import-path-guide"><strong>불러올 파일</strong><span>공사관리대장.xlsx</span></div>`;
    openModal({
      eyebrow:'엑셀 불러오기 · 드래그앤드롭',
      title,
      body:`<div class="notice"><strong>${expected}를 이 창에 끌어다 놓을 수 있습니다.</strong><br>파일은 서버로 전송하지 않고 이 브라우저에서만 읽습니다.</div>${guide}
        <div class="excel-drop-zone" id="excelDropZone" tabindex="0" role="button" aria-label="${expected} 드래그앤드롭 영역">
          <div class="excel-drop-icon">↧</div>
          <strong>${expected}를 여기에 놓으세요</strong>
          <span>또는 아래 파일 선택 버튼을 사용하세요.</span>
        </div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="selectExcelFromDropBtn">파일 선택</button>`
    });
    const zone = modalBody.querySelector('#excelDropZone');
    const acceptFile = file => {
      if (!file) return;
      if (!/\.xlsx$/i.test(file.name || '')) {
        showToast('xlsx 파일만 불러올 수 있습니다.','warn');
        return;
      }
      state.importMode = mode;
      handleExcelFile(file);
    };
    const setDrag = active => zone?.classList.toggle('drag-active', !!active);
    zone?.addEventListener('click',()=>openExcelPicker(mode));
    zone?.addEventListener('keydown',ev=>{ if(ev.key==='Enter'||ev.key===' '){ev.preventDefault();openExcelPicker(mode);} });
    zone?.addEventListener('dragenter',ev=>{ev.preventDefault();setDrag(true);});
    zone?.addEventListener('dragover',ev=>{ev.preventDefault();setDrag(true);});
    zone?.addEventListener('dragleave',ev=>{ if(!zone.contains(ev.relatedTarget)) setDrag(false); });
    zone?.addEventListener('drop',ev=>{
      ev.preventDefault(); setDrag(false);
      acceptFile(ev.dataTransfer?.files?.[0] || null);
    });
    modalActions.querySelector('[data-modal-close]')?.addEventListener('click',closeModal);
    modalActions.querySelector('#selectExcelFromDropBtn')?.addEventListener('click',()=>openExcelPicker(mode));
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
      <section class="hero document-output-hero">
        <div>
          <p class="eyebrow">행정실 공사서류 작성·출력</p>
          <h1>출력할 공사를 선택하세요</h1>
          <p>${state.school.name ? `${e(state.school.name)} · ` : ''}공사를 열면 필요한 서류가 바로 보입니다. 공통정보는 한 번만 입력하고 여러 서류에 재사용합니다.</p>
          <div class="security-note"><span class="security-dot"></span>브라우저 자동저장 · 입력 데이터 서버 미전송</div>
        </div>
      </section>

      <section class="start-grid" aria-label="공사 시작 방법">
        <button class="start-card primary-start" type="button" data-start-new>
          <span class="start-step">＋</span>
          <span class="start-copy"><strong>새 공사 시작</strong><small>공사명부터 등록하고 필요한 정보만 이어서 입력합니다.</small></span>
          <span class="start-arrow">›</span>
        </button>
        <button class="start-card" type="button" data-start-history>
          <span class="start-step">↧</span>
          <span class="start-copy"><strong>공사관리대장 불러오기</strong><small>기존 공사 여러 건을 한꺼번에 가져와 바로 서류를 확인합니다.</small></span>
          <span class="start-arrow">›</span>
        </button>
        <button class="start-card" type="button" data-start-edufine>
          <span class="start-step">↻</span>
          <span class="start-copy"><strong>에듀파인 자료로 보완</strong><small>자료관리목록.xlsx의 계약·준공·지출 정보를 기존 공사에 반영합니다.</small></span>
          <span class="start-arrow">›</span>
        </button>
      </section>

      <section class="panel">
        <div class="toolbar">
          <div class="toolbar-left">
            <div class="search-wrap"><input id="projectSearch" type="search" value="${e(state.search)}" placeholder="공사명 · 업체명 · 계약번호 검색" aria-label="공사 검색"></div>
          </div>
          <div class="toolbar-right">
            ${state.projects.length?'<button class="button ghost small reset-projects-toolbar" id="resetProjectsBtn" type="button">공사자료 초기화</button>':''}
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
    main.querySelector('[data-start-history]')?.addEventListener('click', () => openExcelDropModal('history'));
    main.querySelector('[data-start-edufine]')?.addEventListener('click', () => openExcelDropModal('edufine'));
    document.getElementById('resetProjectsBtn')?.addEventListener('click',confirmResetProjects);
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
    state.detailTab = 'documents';
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
      <button class="back-button" id="backToList" type="button"><span class="back-button-icon" aria-hidden="true">←</span><span>공사 목록</span></button>
      <div class="detail-head document-first-head">
        <div class="current-project">
          <p class="eyebrow">현재 공사</p>
          <h1>${e(p.projectName || '이름 없는 공사')}</h1>
          <div class="current-meta"><span>${e(p.vendorName || '업체 미입력')}</span><span class="meta-dot">·</span><strong>${e(formatMoney(p.currentContractAmount))}</strong><span class="status-chip ${status.cls}">${e(status.label)}</span></div>
        </div>
        <div class="detail-head-actions"><div class="save-state" id="saveState"><span class="pulse"></span><span>이 기기에 저장됨</span></div></div>
      </div>

      <nav class="detail-tabs document-workspace-nav" aria-label="공사 서류 작업 메뉴">
        <button type="button" class="detail-tab ${state.detailTab==='documents'?'active':''}" data-detail-tab="documents">
          <svg class="detail-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l3 3v15H6z"></path><path d="M14 3v4h4M9 11h6M9 15h6"></path></svg>
          <span class="detail-tab-label">서류 작성</span>
        </button>
        <button type="button" class="detail-tab ${state.detailTab==='info'?'active':''}" data-detail-tab="info">
          <svg class="detail-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h14v16H5z"></path><path d="M8 8h8M8 12h8M8 16h5"></path></svg>
          <span class="detail-tab-label">공사정보 수정</span>
        </button>
        <button type="button" class="detail-tab ${state.detailTab==='changes'?'active':''}" data-detail-tab="changes">
          <svg class="detail-tab-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M7 7h10v10H7z"></path><path d="M4 10V4h6M20 14v6h-6"></path></svg>
          <span class="detail-tab-label">변경계약</span>${p.contractChanges?.length ? ` <span class="detail-tab-count">${p.contractChanges.length}</span>` : ''}
        </button>
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
            textareaField('settlementReductionReason','준공 감액 사유',p.settlementReductionReason),
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
            field('defectSecurityType','하자보증방법',p.defectSecurityType),
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
            <button class="button danger small" data-delete-project type="button" style="margin-top:10px">이 공사 삭제</button>
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
    main.querySelectorAll('[data-delete-project]').forEach(btn=>btn.addEventListener('click',confirmDeleteProject));
    document.getElementById('addContractChangeBtn')?.addEventListener('click', openContractChangeModal);
    document.getElementById('sideAddContractChangeBtn')?.addEventListener('click', openContractChangeModal);
    document.getElementById('tabAddContractChangeBtn')?.addEventListener('click', openContractChangeModal);
    document.getElementById('openWarrantyManagerBtn')?.addEventListener('click', openWarrantyManager);
    document.getElementById('addWarrantyInspectionBtn')?.addEventListener('click', () => openWarrantyInspectionModal());
    document.getElementById('openDocumentsTab')?.addEventListener('click', () => { state.detailTab='documents'; renderProjectDetail(); });
    document.getElementById('jumpCurrentStage')?.addEventListener('click', () => jumpToSection(sectionForStatus(status.key)));
    main.querySelectorAll('[data-detail-tab]').forEach(btn => btn.addEventListener('click', () => { state.detailTab = btn.dataset.detailTab; renderProjectDetail(); }));
    main.querySelectorAll('[data-doc-open]').forEach(btn => btn.addEventListener('click', () => openDocumentPreview(btn.dataset.docOpen, {mode:btn.dataset.previewMode || ''})));
    main.querySelectorAll('[data-doc-print]').forEach(btn => btn.addEventListener('click', () => printDocumentDirect(btn.dataset.docPrint, {mode:btn.dataset.previewMode || ''})));
    main.querySelectorAll('[data-pledge-edit]').forEach(btn => btn.addEventListener('click', () => openPrivateContractPledgeModal()));
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
        warrantyInspectorPosition:'inspectorPosition', warrantyInspectorName:'inspectorName',
        warrantyWitnessPosition:'witnessPosition', warrantyWitnessName:'witnessName',
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
  function isPrivateContractPledge(type) { return type === 'privateContractPledge'; }
  function privateContractPledgeData(p) { return p?.privateContractPledge || {results:{}}; }
  function privateContractPledgeAnsweredCount(p) { return Object.values(privateContractPledgeData(p).results || {}).filter(Boolean).length; }
  function privateContractPledgeBusinessType(p) {
    const value=privateContractPledgeData(p).businessType || '';
    return ['individual','corporate'].includes(value)?value:'';
  }
  function privateContractPledgeNonConflictItems() {
    return (ReferenceData?.privateContractPledgeItems||[]).filter(item=>!/^conflict[1-8]$/.test(item.key));
  }
  function privateContractConflictDefaults(businessType) {
    const results={};
    if(businessType==='individual') {
      for(let i=1;i<=6;i++)results[`conflict${i}`]='no';
      results.conflict7='na'; results.conflict8='na';
    } else if(businessType==='corporate') {
      for(let i=1;i<=6;i++)results[`conflict${i}`]='na';
      results.conflict7='no'; results.conflict8='no';
    }
    return results;
  }
  function privateContractPledgeRowsHtml(results={}, prefix='pledge_') {
    const items=ReferenceData?.privateContractPledgeItems||[];
    const optionLabel={yes:'예',no:'아니오',na:'해당없음'};
    return items.map((item,i)=>`<div class="safety-edit-row pledge-edit-row"><div class="safety-edit-number">${e(item.number||String(i+1))}</div><div class="safety-edit-question"><strong>${e(item.group||'')}</strong><span>${e(item.text||'')}</span></div><div class="safety-edit-options">${(item.options||['yes','no']).map(opt=>`<label><input type="radio" name="${e(prefix)}${e(item.key)}" value="${e(opt)}" ${results[item.key]===opt?'checked':''}> ${e(optionLabel[opt]||opt)}</label>`).join('')}</div></div>`).join('');
  }
  function privateContractPledgeQuickHtml(businessType='', prefix='') {
    const name=`${prefix}pledgeBusinessType`;
    return `<section class="pledge-quick-panel"><div class="pledge-quick-heading"><strong>빠른 작성</strong><span>사업자 유형을 선택한 뒤 반복되는 선택을 한 번에 채울 수 있습니다. 자동입력 후 각 항목을 다시 확인하세요.</span></div><div class="pledge-business-type"><strong>사업자 유형</strong><label><input type="radio" name="${e(name)}" value="individual" ${businessType==='individual'?'checked':''}> 개인사업자</label><label><input type="radio" name="${e(name)}" value="corporate" ${businessType==='corporate'?'checked':''}> 법인사업자</label></div><div class="pledge-quick-actions"><button class="button secondary small" type="button" data-pledge-fill-yes>각서부분 모두 예</button><button class="button secondary small" type="button" data-pledge-fill-conflict>체결제한 기본값 채우기</button></div><p class="pledge-quick-hint">개인사업자: ①~⑥ 아니오 · ⑦~⑧ 해당없음 / 법인사업자: ①~⑥ 해당없음 · ⑦~⑧ 아니오</p></section>`;
  }
  function applyPledgeResultsToInputs(container, results, prefix='pledge_') {
    (ReferenceData?.privateContractPledgeItems||[]).forEach(item=>{
      const value=results[item.key]||'';
      container.querySelectorAll(`input[name="${prefix}${item.key}"]`).forEach(input=>{input.checked=input.value===value;});
    });
  }

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
    const chars=p.workCharacteristics||{};
    return `<section class="safety-recommend-panel"><div class="work-characteristics"><strong>이 공사에 해당하는 작업</strong><label><input type="checkbox" data-work-characteristic="fall" ${chars.fall?'checked':''}> 고소·추락 위험 작업</label><label><input type="checkbox" data-work-characteristic="electrical" ${chars.electrical?'checked':''}> 전기 작업</label><label><input type="checkbox" data-work-characteristic="confined" ${chars.confined?'checked':''}> 밀폐공간 작업</label><label><input type="checkbox" data-work-characteristic="industrial" ${chars.industrial?'checked':''}> 일반 건설작업</label></div>
      <div class="recommended-document-list">${recommended.map(type=>{const def=DOCUMENT_DEFINITIONS[type];if(!def)return'';const source=ReferenceData.safetyChecklists[type]?.owner==='agency'?'기관 점검용':'업체 작성·기관 확인용';return `<div class="recommended-document"><div><strong>${e(def.label)}</strong><span>${e(source)}</span></div><em class="recommend-badge">추천</em></div>`;}).join('')}</div></section>`;
  }

  function documentChipStatus(type,p) {
    if (isSafetyDocument(type)) return safetyChecklistComplete(p,type)?'작성완료':(safetyRequiresCompletion(type)?'작성 필요':'빈 양식 가능');
    if (type === 'privateContractPledge') {
      const count=privateContractPledgeAnsweredCount(p),total=(ReferenceData?.privateContractPledgeItems||[]).length;
      return count?`선택 ${count}/${total}`:'빈 양식 가능';
    }
    if (type === 'warrantyInspectionReport') return activeWarrantyInspection(p)?'기본값/공란/입력값 선택':'작성 필요';
    return documentMissing(type,p).length?'정보 필요':'생성 가능';
  }

  function myDocumentsHtml(p) {
    const types=myDocumentTypes(p);
    return `<section class="my-documents-panel"><div class="document-group-title"><strong>내 서류 ${types.length}종</strong><span>이 공사에서 실제 사용하는 서류만 모아봅니다.</span></div>${types.length?`<div class="my-document-chips">${types.map(type=>`<button type="button" data-doc-open="${e(type)}"><span>${e(DOCUMENT_DEFINITIONS[type].label)}</span><em>${e(documentChipStatus(type,p))}</em></button>`).join('')}</div>`:`<div class="contract-change-empty">아직 내 서류가 없습니다. 아래 추천 또는 전체 서류에서 필요한 문서를 추가하세요.</div>`}</section>`;
  }

  function safetyDocumentCardHtml(type,p){
    const def=DOCUMENT_DEFINITIONS[type];
    const saved=safetyChecklistFor(p,type);
    const complete=safetyChecklistComplete(p,type);
    const mustComplete=safetyRequiresCompletion(type);
    const vendorForm=!mustComplete;
    const statusTitle=complete?'작성 완료':(vendorForm?'빈 양식 가능':'작성 필요');
    const statusDesc=saved?.date?`최근 점검 ${e(formatDate(saved.date))}`:(vendorForm?'업체에 전달할 빈 양식도 바로 확인할 수 있습니다.':'점검결과를 입력한 뒤 작성본을 미리보기합니다.');
    return `<article class="document-card simple-document-card safety-document-card ${complete||vendorForm?'ready':'needs-info'}" data-document-card="${e(type)}">
      <div class="simple-document-head"><h3>${e(def.label)}</h3><span class="document-status-pill ${complete?'ready':(vendorForm?'neutral':'warn')}">${e(statusTitle)}</span></div>
      <p class="simple-document-status">${statusDesc}</p>
      <div class="document-card-actions two-actions"><button class="button primary" type="button" data-safety-edit="${e(type)}">체크리스트 작성</button><button class="button secondary" type="button" data-doc-open="${e(type)}" data-preview-mode="blank">빈 양식 미리보기</button></div>
    </article>`;
  }

  function openSafetyChecklistModal(type, afterSave=null) {
    const p=currentProject(); const def=ReferenceData?.safetyChecklists?.[type]; if(!p||!def)return;
    const saved=safetyChecklistFor(p,type)||{}; const results=saved.results||{}; const agency=def.owner==='agency';
    const defaultInspector=saved.inspector || (agency?(p.supervisor||state.school?.supervisor||''):'');
    const rows=def.items.map((item,i)=>{const key=String(i+1),v=results[key]||'';return `<div class="safety-edit-row"><div class="safety-edit-number">${i+1}</div><div class="safety-edit-question">${e(item)}</div><div class="safety-edit-options"><label><input type="radio" name="safety_${i}" value="yes" ${v==='yes'?'checked':''}> 예</label><label><input type="radio" name="safety_${i}" value="no" ${v==='no'?'checked':''}> 아니요</label><label><input type="radio" name="safety_${i}" value="na" ${v==='na'?'checked':''}> 해당없음</label></div></div>`;}).join('');
    const signatureNotice = '<div class="safety-signature-edit-note"><strong>출력 시 점검자 서명을 넣을 수 있습니다.</strong><span>작성한 체크리스트는 인쇄 직전 마우스·터치·펜으로 서명하거나 서명 없이 출력할 수 있으며, 서명 이미지는 저장하지 않습니다.</span></div>';
    const related=saved.relatedChecklistResults||{};
    const relatedEditor=agency?`<section class="safety-related-checklist-editor"><div><strong>6번 · 점검 체크리스트 해당여부</strong><span>원클릭 엑셀과 같이 해당 여부를 O / X로 표시합니다.</span></div><div class="safety-related-grid">${[['fall','추락재해 예방 체크리스트'],['electrical','감전재해 예방 체크리스트'],['confined','밀폐공간 질식재해예방 체크리스트'],['industrial','일반 산업재해 예방 체크리스트']].map(([key,label])=>`<div><span>${e(label)}</span><label><input type="radio" name="safety_related_${key}" value="o" ${related[key]==='o'?'checked':''}> O</label><label><input type="radio" name="safety_related_${key}" value="x" ${related[key]==='x'?'checked':''}> X</label></div>`).join('')}</div></section>`:'';
    openModal({eyebrow:`안전·보건 · ${agency?'기관 점검':'업체 작성·기관 확인'}`,title:def.label,wide:true,body:`<div class="notice"><strong>${e(def.subtitle)}</strong><br>${e(REFERENCE_PROGRAM)}의 점검항목을 기준으로 작성합니다.</div>${signatureNotice}<div class="modal-grid safety-meta-edit" style="margin-top:16px">${modalDateField('safetyChecklistDate','점검일',saved.date||p.startDate||p.contractDate||'')}${modalField('safetyChecklistInspector',agency?'점검자':'점검자 직접 입력',defaultInspector)}${agency?'':'<div class="field full"><span class="hint">회사 대표자가 아니라 실제 점검한 사람의 이름을 입력합니다.</span></div>'}<div class="field full"><label for="safetyChecklistNotes">비고</label><input id="safetyChecklistNotes" value="${e(saved.notes||'')}"></div></div><div class="safety-bulk-toolbar"><div><strong>빠른 입력</strong><span>미응답 항목만 ‘예’로 채우며, 이미 선택한 ‘아니요·해당없음’은 유지합니다.</span></div><div class="safety-bulk-actions"><button class="button secondary small" type="button" id="fillUnansweredSafetyYesBtn">미응답 모두 예</button><button class="button ghost small" type="button" id="resetSafetyAnswersBtn">응답 초기화</button></div></div><div class="safety-edit-list">${rows}</div>${relatedEditor}${def.footer?`<p class="safety-edit-footer">※ ${e(def.footer)}</p>`:''}`,actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="saveSafetyChecklistBtn">저장${afterSave?'하고 계속':''}</button>`});
    initDateInputs(modalBody);
    modalActions.querySelector('[data-modal-close]').addEventListener('click',closeModal);
    modalActions.querySelector('#saveSafetyChecklistBtn').addEventListener('click',()=>saveSafetyChecklist(type,afterSave));
    modalBody.querySelector('#fillUnansweredSafetyYesBtn')?.addEventListener('click',()=>{
      let changed=0;
      modalBody.querySelectorAll('.safety-edit-row').forEach(row=>{
        if(row.querySelector('input[type="radio"]:checked')) return;
        const yes=row.querySelector('input[type="radio"][value="yes"]');
        if(yes){yes.checked=true;changed++;}
      });
      showToast(changed?`미응답 ${changed}개 항목을 ‘예’로 채웠습니다.`:'미응답 항목이 없습니다.');
    });
    modalBody.querySelector('#resetSafetyAnswersBtn')?.addEventListener('click',()=>{
      modalBody.querySelectorAll('.safety-edit-row input[type="radio"]:checked').forEach(input=>{input.checked=false;});
      showToast('체크리스트 응답을 화면에서 초기화했습니다. 저장 전까지 기존 저장값은 유지됩니다.');
    });
  }

  async function saveSafetyChecklist(type, afterSave=null){
    const p=currentProject(),def=ReferenceData?.safetyChecklists?.[type];if(!p||!def)return;
    const date=modalBody.querySelector('#safetyChecklistDate')?.value||'',inspector=modalBody.querySelector('#safetyChecklistInspector')?.value?.trim()||'';
    if(!date||!inspector){showToast('점검일과 점검자를 입력해주세요.','warn');return;}
    const results={}; let missing=0; def.items.forEach((_,i)=>{const v=modalBody.querySelector(`input[name="safety_${i}"]:checked`)?.value||'';results[String(i+1)]=v;if(!v)missing++;});
    if(missing){showToast(`아직 선택하지 않은 점검항목이 ${missing}개 있습니다.`,'warn');return;}
    const relatedChecklistResults={};
    if(def.owner==='agency'){
      for(const key of ['fall','electrical','confined','industrial']){
        const value=modalBody.querySelector(`input[name="safety_related_${key}"]:checked`)?.value||'';
        if(!value){showToast('6번 점검 체크리스트의 O / X를 모두 선택해주세요.','warn');return;}
        relatedChecklistResults[key]=value;
      }
    }
    p.safetyChecklists={...(p.safetyChecklists||{}),[type]:{date,inspector,notes:modalBody.querySelector('#safetyChecklistNotes')?.value?.trim()||'',results,relatedChecklistResults,updatedAt:new Date().toISOString()}};
    if(!p.selectedDocuments?.includes(type))p.selectedDocuments=[...(p.selectedDocuments||[]),type];
    p.updatedAt=new Date().toISOString();await DB.put('projects',p);await loadState();state.currentProjectId=p.id;closeModal();renderProjectDetail();showToast(`${def.label}을 저장했습니다.`);if(typeof afterSave==='function')afterSave();
  }


  function openPrivateContractPledgeModal(afterSavePreview = false) {
    const p=currentProject(); if(!p)return;
    const saved=privateContractPledgeData(p); const results={...(saved.results||{})};
    let businessType=privateContractPledgeBusinessType(p);
    const rows=privateContractPledgeRowsHtml(results);
    openModal({eyebrow:'계약서류 · 작성 지원',title:'수의계약 통합서약서 작성',wide:true,body:`<div class="notice"><strong>반복 선택은 빠르게 채우고, 사실관계는 업체가 최종 확인합니다.</strong><br>‘각서부분 모두 예’는 체결제한 ①~⑧을 제외한 서약·동의 항목을 예로 채웁니다. 체결제한은 사업자 유형에 맞는 기본값을 채운 뒤 반드시 항목별로 확인하세요.</div>${privateContractPledgeQuickHtml(businessType)}<div class="safety-edit-list pledge-edit-list">${rows}</div>`,actions:`<button class="button ghost" type="button" id="clearPledgeChoicesBtn">선택 모두 지우기</button><span class="modal-action-spacer"></span><button class="button secondary" type="button" data-modal-close>취소</button><button class="button secondary" type="button" id="savePledgeChoicesBtn">저장</button><button class="button primary" type="button" id="savePreviewPledgeChoicesBtn">저장하고 미리보기</button>`});
    const readBusinessType=()=>modalBody.querySelector('input[name="pledgeBusinessType"]:checked')?.value||'';
    modalActions.querySelector('[data-modal-close]')?.addEventListener('click',closeModal);
    modalActions.querySelector('#clearPledgeChoicesBtn')?.addEventListener('click',()=>modalBody.querySelectorAll('input[name^="pledge_"]').forEach(input=>{input.checked=false;}));
    modalBody.querySelector('[data-pledge-fill-yes]')?.addEventListener('click',()=>{
      privateContractPledgeNonConflictItems().forEach(item=>{const input=modalBody.querySelector(`input[name="pledge_${item.key}"][value="yes"]`);if(input)input.checked=true;});
      showToast('체결제한을 제외한 서약·동의 항목을 ‘예’로 채웠습니다.');
    });
    modalBody.querySelector('[data-pledge-fill-conflict]')?.addEventListener('click',()=>{
      businessType=readBusinessType();
      if(!businessType){showToast('개인사업자 또는 법인사업자를 먼저 선택해주세요.','warn');return;}
      const defaults=privateContractConflictDefaults(businessType);
      Object.entries(defaults).forEach(([key,value])=>{const input=modalBody.querySelector(`input[name="pledge_${key}"][value="${value}"]`);if(input)input.checked=true;});
      showToast(`${businessType==='individual'?'개인사업자':'법인사업자'} 기준으로 체결제한 기본값을 채웠습니다. 항목별 사실관계를 확인해주세요.`);
    });
    modalActions.querySelector('#savePledgeChoicesBtn')?.addEventListener('click',()=>savePrivateContractPledge(false));
    modalActions.querySelector('#savePreviewPledgeChoicesBtn')?.addEventListener('click',()=>savePrivateContractPledge(true));
  }

  async function savePrivateContractPledge(previewAfterSave=false) {
    const p=currentProject(); if(!p)return;
    const items=ReferenceData?.privateContractPledgeItems||[]; const results={};
    items.forEach(item=>{const v=modalBody.querySelector(`input[name="pledge_${item.key}"]:checked`)?.value||''; if(v)results[item.key]=v;});
    const businessType=modalBody.querySelector('input[name="pledgeBusinessType"]:checked')?.value||'';
    p.privateContractPledge={results,businessType,updatedAt:new Date().toISOString()};
    if(!p.selectedDocuments?.includes('privateContractPledge'))p.selectedDocuments=[...(p.selectedDocuments||[]),'privateContractPledge'];
    p.updatedAt=new Date().toISOString(); await DB.put('projects',p); await loadState(); state.currentProjectId=p.id; closeModal(); renderProjectDetail();
    showToast(Object.keys(results).length?`수의계약 통합서약서 선택 ${Object.keys(results).length}개를 저장했습니다.`:'선택을 비웠습니다. 빈 양식으로 출력할 수 있습니다.');
    if(previewAfterSave)openDocumentPreview('privateContractPledge',{mode:Object.keys(results).length?'filled':'blank'});
  }

  function documentQuickItemHtml(type, p) {
    const def = DOCUMENT_DEFINITIONS[type];
    const missing = documentMissing(type, p);
    const safetyIncomplete = isSafetyDocument(type) && safetyRequiresCompletion(type) && !safetyChecklistComplete(p,type);
    const vendorBlank = isSafetyDocument(type) && !safetyRequiresCompletion(type) && !safetyChecklistComplete(p,type);
    const status = safetyIncomplete ? '작성 필요' : (vendorBlank ? '빈 양식 가능' : (missing.length ? `${missing.length}개 정보 필요` : '생성 가능'));
    return `<button class="doc-item doc-item-button ${(missing.length||safetyIncomplete)?'needs-info':'ready'}" type="button" data-doc-open="${e(type)}"><span>${e(def.label)}</span><em>${e(status)}</em></button>`;
  }

  function privateContractPledgeCardHtml(type,p) {
    const def=DOCUMENT_DEFINITIONS[type];
    const missing=documentMissing(type,p);
    const answered=privateContractPledgeAnsweredCount(p);
    const total=(ReferenceData?.privateContractPledgeItems||[]).length;
    const labels=missing.map(x=>DOCUMENT_FIELD_LABELS[x]||x);
    const status = missing.length ? `출력 전 확인 ${missing.length}개 · ${labels.slice(0,2).join(' · ')}${labels.length>2?' 외':''}` : (answered?`선택내용 ${answered}/${total} 저장됨`:'빈 양식부터 확인 가능');
    return `<article class="document-card simple-document-card pledge-document-card ${missing.length?'needs-info':'ready'}" data-document-card="${e(type)}">
      <div class="simple-document-head"><h3>${e(def.label)}</h3><span class="document-status-pill ${missing.length?'warn':'ready'}">${missing.length?'정보 확인':'출력 가능'}</span></div>
      <p class="simple-document-status">${e(status)}</p>
      <div class="document-card-actions pledge-actions"><button class="button secondary" type="button" data-pledge-edit>${answered?'선택내용 수정':'예·아니오 선택'}</button>${answered?`<button class="button primary" type="button" data-doc-open="${e(type)}" data-preview-mode="filled">작성본 미리보기</button>`:''}<button class="button secondary" type="button" data-doc-open="${e(type)}" data-preview-mode="blank">빈 양식 미리보기</button></div>
    </article>`;
  }

  function warrantyInspectionCardHtml(type,p) {
    const def=DOCUMENT_DEFINITIONS[type];
    const record=activeWarrantyInspection(p);
    const missing=documentMissing(type,p);
    const labels=missing.map(x=>DOCUMENT_FIELD_LABELS[x]||x);
    const status = !record ? '검사일·검사결과를 미리보기에서 입력할 수 있습니다.' : (missing.length?`출력 전 확인 ${missing.length}개 · ${labels.slice(0,2).join(' · ')}${labels.length>2?' 외':''}`:'기본값·공란·직접 입력값을 선택해 출력할 수 있습니다.');
    return `<article class="document-card simple-document-card warranty-report-card ${missing.length?'needs-info':'ready'}" data-document-card="${e(type)}">
      <div class="simple-document-head"><h3>${e(def.label)}</h3><span class="document-status-pill ${missing.length?'warn':'ready'}">${record?'미리보기 가능':'검사내용 입력'}</span></div>
      <p class="simple-document-status">${e(status)}</p>
      <button class="button primary full-button" type="button" data-doc-open="${e(type)}" data-preview-mode="default">미리보기</button>
    </article>`;
  }

  function documentCardHtml(type, p) {
    if (isSafetyDocument(type)) return safetyDocumentCardHtml(type,p);
    if (isPrivateContractPledge(type)) return privateContractPledgeCardHtml(type,p);
    if (type === 'warrantyInspectionReport') return warrantyInspectionCardHtml(type,p);
    const def = DOCUMENT_DEFINITIONS[type];
    const missing = documentMissing(type, p);
    const labels = missing.map(x => DOCUMENT_FIELD_LABELS[x] || x);
    const blankSiteManager = type === 'utilityPaymentPledge' && !String(p.siteManager || '').trim();
    const status = missing.length
      ? `출력 전 확인 ${missing.length}개 · ${labels.slice(0,2).join(' · ')}${labels.length>2?' 외':''}`
      : (blankSiteManager?'현장대리인은 공란으로도 출력할 수 있습니다.':'현재 입력정보로 바로 미리보기할 수 있습니다.');
    return `<article class="document-card simple-document-card ${missing.length?'needs-info':'ready'}" data-document-card="${e(type)}">
      <div class="simple-document-head"><h3>${e(def.label)}</h3><span class="document-status-pill ${missing.length?'warn':'ready'}">${missing.length?'정보 확인':'출력 가능'}</span></div>
      <p class="simple-document-status">${e(status)}</p>
      <button class="button ${missing.length?'secondary':'primary'} full-button" type="button" data-doc-open="${e(type)}">미리보기</button>
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
    return `<article class="document-card simple-document-card utility-tool-card ${total?'ready':''}">
      <div class="simple-document-head"><h3>전력비·수도광열비 계산</h3><span class="document-status-pill ${total?'ready':'neutral'}">${total?'계산됨':'계산 도구'}</span></div>
      <p class="simple-document-status">${total?`현재 계산값 ${e(formatMoney(total))}`:'직접재료비·직접노무비를 넣어 계산합니다.'}</p>
      <button class="button ${total?'secondary':'primary'} full-button" type="button" data-open-utility-calculator>${total?'계산 다시 보기':'계산하기'}</button>
    </article>`;
  }

  function batchDocumentSelectorHtml(p) {
    const selected=selectionForProject(p);
    return `<div class="batch-document-selector">${DOCUMENT_PRINT_ORDER.map(type=>{
      const def=DOCUMENT_DEFINITIONS[type];
      if(!def)return'';
      return `<label class="batch-document-check"><input type="checkbox" data-doc-select="${e(type)}" ${selected.has(type)?'checked':''}><span>${e(def.label)}</span></label>`;
    }).join('')}</div>`;
  }

  function oneClickSeparateDocumentsHtml() {
    return `<details class="oneclick-gap-panel" open>
      <summary><span><strong>원클릭 엑셀에서 별도 확인할 서류</strong><small>허브에서 생성하지 않는 조건부·업체 자체 작성 서류입니다.</small></span><em>8종 + 참고 2종</em></summary>
      <div class="oneclick-gap-body">
        <div><h3>착공·현장 관련</h3><ul><li>도시가스배관 등 안전조치 협의서 <span>해당 공사</span></li><li>현장대리인계</li><li>공사예정공정표</li><li>직접시공계획서</li></ul></div>
        <div><h3>노무비 관련</h3><ul><li>노무비 구분관리 및 지급확인 합의서</li><li>노무비 구분관리 제외</li><li>노무비 구분관리 및 지급확인 적용 제외</li><li>공사근로자 노무비 청구서</li></ul></div>
      </div>
      <p class="oneclick-gap-note"><strong>함께 확인할 참고 시트</strong> · 관련 회계예규 적용 · 착공계·공정표·대리인계(간이)</p>
    </details>`;
  }

  function documentsTabHtml(p) {
    const selected = orderedSelectedTypes(p);
    const missing = batchMissingFields(selected, p);
    const readyCount = selected.filter(type => documentMissing(type,p).length === 0 && (!isSafetyDocument(type) || safetyChecklistComplete(p,type) || !safetyRequiresCompletion(type))).length;
    return `<div class="documents-panel document-first-panel">
      <div class="documents-head document-first-title"><div><p class="eyebrow">행정실 공사서류 작성·출력</p><h2>필요한 서류를 선택하세요</h2><p>공사정보를 한 번 입력하면 여러 서류에 자동으로 반영됩니다. 서류를 열어 확인하고 바로 인쇄·PDF 저장하세요.</p></div><button class="text-action-button" type="button" data-detail-tab="info">공사정보 수정 →</button></div>
      <div class="document-purpose-strip"><div><strong>입력정보를 잘못 적었나요?</strong><span>각 서류의 미리보기에서 바로 수정할 수 있고, 수정값은 다른 서류에도 함께 반영됩니다.</span></div><button class="button secondary small" type="button" data-detail-tab="info">공통정보 확인</button></div>
      <div class="reference-program-banner compact-reference"><span class="reference-program-label">기준 자료</span><div class="reference-program-copy"><strong>${e(REFERENCE_PROGRAM)}</strong><span>서식·점검항목의 기준으로 사용합니다.</span></div></div>

      <div class="document-group primary-document-group"><div class="document-group-title"><strong>행정실 작성·출력</strong><span>학교에서 직접 작성하거나 보관하는 서류를 먼저 모았습니다.</span></div><div class="document-grid simplified-document-grid">${documentCardHtml('supervisionReport',p)}${documentCardHtml('completionInspectionRecord',p)}${utilityToolCardHtml(p)}${documentCardHtml('constructionLedger',p)}${documentCardHtml('warrantyLedger',p)}${documentCardHtml('warrantyInspectionReport',p)}</div></div>

      <div class="document-group vendor-document-group"><div class="document-group-title"><strong>업체에 전달할 서식</strong><span>업체가 작성·제출하는 양식도 현재 공사정보를 반영해 바로 준비할 수 있습니다.</span></div>
        <div class="document-subgroup"><h4>계약</h4><div class="document-grid simplified-document-grid">${documentCardHtml('standardContract',p)}${documentCardHtml('acceptanceTerms',p)}${documentCardHtml('useSealForm',p)}${documentCardHtml('privateContractPledge',p)}</div></div>
        <div class="document-subgroup"><h4>착공</h4><div class="document-grid simplified-document-grid">${documentCardHtml('startReport',p)}${documentCardHtml('utilityPaymentPledge',p)}</div></div>
        <div class="document-subgroup"><h4>준공·지출</h4><div class="document-grid simplified-document-grid">${documentCardHtml('completionReport',p)}${documentCardHtml('completionInspectionRequest',p)}${documentCardHtml('defectSecurityDeposit',p)}${documentCardHtml('completionSettlementAgreement',p)}${documentCardHtml('paymentRequest',p)}</div></div>
      </div>

      <details class="document-group safety-document-group safety-collapsible-group">
        <summary class="safety-document-summary"><span class="safety-summary-copy"><strong>안전·보건 체크리스트</strong><small>작성본 또는 업체 전달용 빈 양식을 필요한 방식으로 준비합니다.</small></span><span class="safety-summary-meta"><em>5종</em><span class="safety-summary-toggle" aria-hidden="true"></span></span></summary>
        <div class="safety-document-body"><div class="document-grid simplified-document-grid">${documentCardHtml('safetyGeneral',p)}${documentCardHtml('safetyFall',p)}${documentCardHtml('safetyElectrical',p)}${documentCardHtml('safetyConfined',p)}${documentCardHtml('safetyIndustrial',p)}</div>
          <details class="secondary-feature-details"><summary>공종·작업특성에 맞는 체크리스트 추천 보기</summary>${safetyRecommendationsHtml(p)}</details>
        </div>
      </details>

      <details class="batch-output-details" ${selected.length?'open':''}>
        <summary><span><strong>여러 서류를 한 번에 출력</strong><small>필요할 때만 열어 서류를 선택하거나 세트를 사용합니다.</small></span><em id="batchOutputCount">${selected.length?`${selected.length}종 선택`:'선택 안 함'}</em></summary>
        <div class="batch-output-body">
          <div class="document-set-buttons" aria-label="서류 세트 선택">
            <button class="button secondary small" type="button" data-doc-set="agencyManagement">행정실 서류</button>
            <button class="button secondary small" type="button" data-doc-set="safety">안전·보건 5종</button>
            <button class="button secondary small" type="button" data-doc-set="contract">계약서류 4종</button>
            <button class="button secondary small" type="button" data-doc-set="start">착공서류 2종</button>
            <button class="button secondary small" type="button" data-doc-set="completion">준공서류 6종</button>
            <button class="button ghost small" type="button" data-doc-set="all">전체 21종</button>
            <button class="button ghost small" type="button" data-doc-clear>선택 해제</button>
          </div>
          ${batchDocumentSelectorHtml(p)}
          <div class="document-batch-summary" id="documentBatchSummary">
            <div><strong>${selected.length ? `${selected.length}종 선택` : '출력할 서류를 선택하세요'}</strong><span>${selected.length ? (missing.length ? `출력 전 확인정보 ${missing.length}개 · 바로 출력 ${readyCount}종` : '선택한 서류를 한 번에 미리보기합니다.') : '개별 서류는 위 카드에서 바로 미리보기할 수 있습니다.'}</span></div>
            <button class="button primary" type="button" id="openBatchPreviewBtn" ${selected.length?'':'disabled'}>${selected.length?`선택한 ${selected.length}종 미리보기`:'묶음 미리보기'}</button>
          </div>
        </div>
      </details>

      ${oneClickSeparateDocumentsHtml()}
      ${recentPrintHistoryHtml(p)}
      <div class="document-footnote">출력양식과 안전점검 항목은 ${e(REFERENCE_PROGRAM)} 버전을 기준으로 구현했습니다. 실제 계약·공사 상황에 맞는지 최종 확인 후 사용합니다.</div>
    </div>`;
  }

  function refreshDocumentBatchUi(p) {
    const summary = main.querySelector('#documentBatchSummary');
    if (!summary) return;
    const selected = orderedSelectedTypes(p);
    const missing = batchMissingFields(selected,p);
    const readyCount = selected.filter(type => documentMissing(type,p).length === 0 && (!isSafetyDocument(type) || safetyChecklistComplete(p,type) || !safetyRequiresCompletion(type))).length;
    summary.innerHTML = `<div><strong>${selected.length ? `${selected.length}종 선택` : '출력할 서류를 선택하세요'}</strong><span>${selected.length ? (missing.length ? `출력 전 확인정보 ${missing.length}개 · 바로 출력 ${readyCount}종` : '선택한 서류를 한 번에 미리보기합니다.') : '개별 서류는 위 카드에서 바로 미리보기할 수 있습니다.'}</span></div><button class="button primary" type="button" id="openBatchPreviewBtn" ${selected.length?'':'disabled'}>${selected.length?`선택한 ${selected.length}종 미리보기`:'묶음 미리보기'}</button>`;
    const countLabel=main.querySelector('#batchOutputCount');
    if(countLabel) countLabel.textContent=selected.length?`${selected.length}종 선택`:'선택 안 함';
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
    main.querySelectorAll('[data-safety-edit]').forEach(btn=>btn.addEventListener('click',()=>openSafetyChecklistModal(btn.dataset.safetyEdit,()=>openDocumentPreview(btn.dataset.safetyEdit,{mode:'filled'}))));
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
  function documentMissingFieldHtml(field, p, overrideValue) {
    const id = `docMissing_${field}`;
    const label = DOCUMENT_FIELD_LABELS[field] || field;
    const value = arguments.length >= 3 ? overrideValue : documentValue(field, p);
    if (MONEY_FIELDS.has(field)) return `<div class="field"><label for="${e(id)}">${e(label)}</label>${moneyInputHtml(id, value)}</div>`;
    if (DATE_FIELDS.has(field)) return modalDateField(id, label, value);
    if (['warrantyInspectionResult','warrantyIssueDetails','warrantyActions','warrantyNotes','settlementReductionReason'].includes(field)) return `<div class="field full"><label for="${e(id)}">${e(label)}</label><textarea id="${e(id)}">${e(value || '')}</textarea></div>`;
    if (['bankName','accountNumber','accountHolder'].includes(field)) {
      return `<div class="field"><label for="${e(id)}">${e(label)}</label><input id="${e(id)}" value="${e(value || '')}" autocomplete="off"><span class="hint">업체 지급정보 보관함에 별도로 저장됩니다.</span></div>`;
    }
    return modalField(id, label, value, ['vendorAddress','projectName','schoolAddress','priceAdjustmentMethod'].includes(field));
  }

  function documentPreviewFields(type) {
    const def=DOCUMENT_DEFINITIONS[type];
    if(!def)return [];
    return [...new Set([...(def.required||[]),...(PREVIEW_EXTRA_FIELDS[type]||[])])];
  }

  function previewFieldGroupName(field) {
    const found=PREVIEW_FIELD_GROUPS.find(([,fields])=>fields.has(field));
    return found?.[0] || '기타 정보';
  }

  function documentPreviewEditorHtml(type,p,draftValues,isBlankForm=false) {
    const fields=documentPreviewFields(type);
    if(!fields.length)return '';
    const required=new Set(DOCUMENT_DEFINITIONS[type]?.required||[]);
    const groups=[];
    fields.forEach(field=>{
      const name=previewFieldGroupName(field);
      let group=groups.find(x=>x.name===name);
      if(!group){group={name,fields:[]};groups.push(group);}
      group.fields.push(field);
    });
    const body=groups.map(group=>`<section class="preview-editor-group"><h4>${e(group.name)}</h4><div class="preview-editor-fields">${group.fields.map(field=>{
      const value=Object.prototype.hasOwnProperty.call(draftValues,field)?draftValues[field]:documentValue(field,p);
      const missing=required.has(field)&&!meaningful(value);
      return `<div class="preview-editor-item ${missing?'missing':''}" data-preview-field-wrap="${e(field)}">${documentMissingFieldHtml(field,p,value)}${required.has(field)?'<span class="preview-required-tag">필수</span>':''}</div>`;
    }).join('')}</div></section>`).join('');
    return `<div class="preview-field-editor"><div class="preview-editor-head"><div><strong>이 서류에 사용되는 정보</strong><span>${isBlankForm?'빈 양식 모드에서는 입력값이 출력물에 반영되지 않습니다.':'오른쪽 서류를 보면서 바로 수정하세요.'}</span></div></div>${body}<div class="preview-editor-savebar"><button class="button primary" type="button" id="savePreviewAllFieldsBtn">변경사항 저장</button></div></div>`;
  }

  function warrantyPreviewRenderOptions(p, mode='default') {
    const selectedMode = ['default','blank','filled'].includes(mode) ? mode : 'default';
    return { warrantySignatoryMode:selectedMode };
  }

  function printDocumentDirect(type, previewOptions = {}) {
    const p=currentProject(); const def=DOCUMENT_DEFINITIONS[type];
    if(!p||!def)return;
    const blankSafety=isSafetyDocument(type)&&previewOptions.mode==='blank';
    if(isSafetyDocument(type) && safetyRequiresCompletion(type) && !blankSafety && !safetyChecklistComplete(p,type)) {
      openSafetyChecklistModal(type,()=>printDocumentDirect(type,previewOptions)); return;
    }
    if(type==='warrantyInspectionReport' && !(p.warrantyInspections||[]).length) { openWarrantyInspectionModal(null,true); return; }
    if(type==='warrantyInspectionReport' && !state.activeWarrantyInspectionId) state.activeWarrantyInspectionId=p.warrantyInspections[p.warrantyInspections.length-1]?.id||null;
    let renderOptions={};
    if(blankSafety) renderOptions.blankSafety=true;
    if(type==='warrantyInspectionReport') renderOptions={...renderOptions,...warrantyPreviewRenderOptions(p,previewOptions.mode||'default')};
    if(type==='privateContractPledge') {
      const hasAnswers=privateContractPledgeAnsweredCount(p)>0;
      renderOptions.blankPledge=previewOptions.mode==='blank'||(!hasAnswers&&previewOptions.mode!=='filled');
    }
    const blankForm=!!(renderOptions.blankSafety||renderOptions.blankPledge);
    const missing=blankForm?[]:documentMissing(type,p);
    if(missing.length) { openDocumentMissingModal(type,missing,previewOptions); return; }
    printAdministrativeDocument(type,p,renderOptions);
  }

  function previewOptionsFromRender(type, renderOptions = {}) {
    if (isSafetyDocument(type)) return {mode:renderOptions.blankSafety?'blank':'filled'};
    if (type === 'privateContractPledge') return {mode:renderOptions.blankPledge?'blank':'filled'};
    if (type === 'warrantyInspectionReport') return {mode:renderOptions.warrantySignatoryMode || 'default'};
    return {};
  }

  function openDocumentInfoEditModal(type, previewOptions = {}) {
    const p=currentProject(); const def=DOCUMENT_DEFINITIONS[type];
    if(!p||!def)return;
    const extraFields=type==='warrantyLedger'?['supervisor']:[];
    const fields=[...new Set([...(def.required||[]),...extraFields])];
    openModal({
      eyebrow:`${def.label} · 입력정보 수정`,
      title:'이 서류에 쓰이는 정보를 수정합니다',
      wide:true,
      body:`<div class="notice"><strong>잘못 입력한 값은 여기에서 바로 고칠 수 있습니다.</strong><br>저장하면 현재 공사정보에 반영되어 다른 서류에서도 같은 값을 사용합니다. 값을 비우면 미리보기 상단에 다시 부족정보로 표시됩니다.</div><div class="modal-grid document-missing-grid" style="margin-top:16px">${fields.map(field=>documentMissingFieldHtml(field,p)).join('')}</div>`,
      actions:`<button class="button secondary" type="button" id="backToDocumentPreviewBtn">미리보기로 돌아가기</button><button class="button primary" type="button" id="saveDocumentInfoEditBtn">저장하고 미리보기</button>`
    });
    initDateInputs(modalBody); initMoneyInputs(modalBody);
    modalActions.querySelector('#backToDocumentPreviewBtn')?.addEventListener('click',()=>openDocumentPreview(type,previewOptions));
    modalActions.querySelector('#saveDocumentInfoEditBtn')?.addEventListener('click',()=>saveDocumentInfoEdits(type,fields,previewOptions));
  }

  async function persistDocumentInfoFields(fields) {
    const p=currentProject(); if(!p)return null;
    let schoolChanged=false; const payoutValues={}; const warrantyValues={};
    for(const field of fields){
      const el=modalBody.querySelector(`#docMissing_${CSS.escape(field)}`);
      let value=el?.value?.trim?.() ?? '';
      if(MONEY_FIELDS.has(field)) value=value?parseMoneyInput(value):'';
      if(['schoolName','schoolAddress','principal'].includes(field)){
        const key=field==='schoolName'?'name':field==='schoolAddress'?'address':'principal';
        state.school={...(state.school||{}),[key]:value}; schoolChanged=true;
      } else if(['supervisor','inspector','witness'].includes(field)) {
        p[field]=value;
      } else if(['bankName','accountNumber','accountHolder'].includes(field)) {
        payoutValues[field]=value;
      } else if(field.startsWith('warranty')) {
        warrantyValues[field]=value;
      } else p[field]=value;
    }
    p.updatedAt=new Date().toISOString();
    await DB.put('projects',p);
    if(schoolChanged) await DB.put('settings',{key:'school',value:state.school});
    if(Object.keys(payoutValues).length) await savePayoutForProject(p,payoutValues);
    if(Object.keys(warrantyValues).length){
      const list=Array.isArray(p.warrantyInspections)?[...p.warrantyInspections]:[];
      let i=state.activeWarrantyInspectionId?list.findIndex(x=>x.id===state.activeWarrantyInspectionId):-1;
      const base=i>=0?list[i]:{};
      const map={warrantyInspectionDate:'date',warrantyInspector:'inspector',warrantyWitness:'witness',warrantyInspectorName:'inspectorName',warrantyWitnessName:'witnessName',warrantyInspectionResult:'result',warrantyIssueDetails:'issueDetails',warrantyActions:'actions',warrantyNotes:'notes'};
      const next={...base,id:base.id||DB.uuid(),createdAt:base.createdAt||new Date().toISOString(),updatedAt:new Date().toISOString()};
      Object.entries(warrantyValues).forEach(([k,v])=>{if(map[k])next[map[k]]=v;});
      if(i>=0)list[i]=next;else list.push(next);
      p.warrantyInspections=list;state.activeWarrantyInspectionId=next.id;await DB.put('projects',p);
    }
    await loadState(); state.currentProjectId=p.id;
    return currentProject();
  }

  async function saveDocumentInfoEdits(type, fields, previewOptions = {}) {
    const saved=await persistDocumentInfoFields(fields); if(!saved)return;
    renderProjectDetail();
    showToast('입력정보를 수정했습니다.');
    openDocumentPreview(type,previewOptions);
  }

  function openDocumentPreview(type, previewOptions = {}) {
    let p = currentProject();
    const def = DOCUMENT_DEFINITIONS[type];
    if (!p || !def) return;
    const blankSafety = isSafetyDocument(type) && previewOptions.mode === 'blank';
    if (isSafetyDocument(type) && safetyRequiresCompletion(type) && !blankSafety && !safetyChecklistComplete(p,type)) { openSafetyChecklistModal(type,()=>openDocumentPreview(type,{mode:'filled'})); return; }
    if (type === 'warrantyInspectionReport' && !(p.warrantyInspections || []).length) { openWarrantyInspectionModal(null,true); return; }
    if (type === 'warrantyInspectionReport' && !state.activeWarrantyInspectionId) state.activeWarrantyInspectionId = p.warrantyInspections[p.warrantyInspections.length-1]?.id || null;

    let renderOptions = {};
    if (blankSafety) renderOptions.blankSafety = true;
    if (type === 'privateContractPledge') {
      const hasAnswers=privateContractPledgeAnsweredCount(p)>0;
      renderOptions.blankPledge = previewOptions.mode === 'blank' || (!hasAnswers && previewOptions.mode !== 'filled');
    }
    if (type === 'warrantyInspectionReport') renderOptions = {...renderOptions,...warrantyPreviewRenderOptions(p,previewOptions.mode || 'default')};
    if (type === 'utilityPaymentPledge') renderOptions.utilitySiteManager = String(p.siteManager || '');

    const previewFields=documentPreviewFields(type);
    let draftValues={};
    let pledgeDraftResults={};
    let pledgeBusinessType='';
    const resetDraftValues=()=>{ draftValues=Object.fromEntries(previewFields.map(field=>[field,documentValue(field,p)])); };
    const resetPledgeDraft=()=>{
      pledgeDraftResults={...(privateContractPledgeData(p).results||{})};
      pledgeBusinessType=privateContractPledgeBusinessType(p);
    };
    resetDraftValues();
    if(type==='privateContractPledge')resetPledgeDraft();

    const isBlankForm = () => !!(renderOptions.blankSafety || renderOptions.blankPledge);
    const draftValue = field => Object.prototype.hasOwnProperty.call(draftValues,field) ? draftValues[field] : documentValue(field,p);
    const currentMissing = () => isBlankForm() ? [] : (def.required||[]).filter(field=>!meaningful(draftValue(field)));

    const draftContext = () => {
      const project=JSON.parse(JSON.stringify(p));
      const school={...(state.school||{})};
      const payout={...(payoutForProject(p)||{})};
      previewFields.forEach(field=>{
        const value=draftValue(field);
        if(field==='schoolName') school.name=value;
        else if(field==='schoolAddress') school.address=value;
        else if(field==='principal') school.principal=value;
        else if(['bankName','accountNumber','accountHolder'].includes(field)) payout[field]=value;
        else if(!field.startsWith('warranty')) project[field]=value;
      });
      if(type==='privateContractPledge') project.privateContractPledge={...(project.privateContractPledge||{}),results:{...pledgeDraftResults},businessType:pledgeBusinessType};
      const options={...renderOptions,__previewSchool:school,__previewPayout:payout,__previewValues:{...draftValues}};
      if(type==='utilityPaymentPledge') options.utilitySiteManager=draftValue('siteManager')||'';
      return {project,options};
    };

    const statusHtml=()=>{
      const missing=currentMissing();
      if(isBlankForm()) return `<div class="preview-editor-status blank" id="previewEditorStatus"><strong>빈 양식 미리보기</strong><span>입력정보는 확인·수정할 수 있지만 현재 빈 양식에는 반영되지 않습니다.</span></div>`;
      if(missing.length){const labels=missing.map(field=>DOCUMENT_FIELD_LABELS[field]||field);return `<div class="preview-editor-status warning" id="previewEditorStatus"><strong>출력 전 확인 ${missing.length}개</strong><span>${e(labels.join(' · '))}</span></div>`;}
      return `<div class="preview-editor-status ready" id="previewEditorStatus"><strong>출력 준비 완료</strong><span>왼쪽 값을 수정하면 오른쪽 미리보기에 바로 반영됩니다.</span></div>`;
    };

    const previewModeControls = () => {
      if (type === 'privateContractPledge') {
        return `<div class="document-preview-mode"><strong>출력 방식</strong><label><input type="radio" name="previewMode" value="filled" ${!renderOptions.blankPledge?'checked':''}> 작성내용 반영</label><label><input type="radio" name="previewMode" value="blank" ${renderOptions.blankPledge?'checked':''}> 빈 양식</label></div>`;
      }
      if (type === 'warrantyInspectionReport') {
        const record=activeWarrantyInspection(p);
        const hasSignatories=!!(record&&[record.inspectorName,record.witnessName].some(Boolean));
        const hasDefaults=!!([p?.inspector,state.school?.inspector,p?.witness,state.school?.witness].some(Boolean));
        const mode=renderOptions.warrantySignatoryMode || 'default';
        return `<div class="document-preview-mode warranty-preview-mode"><strong>검사자·입회자 출력</strong><label><input type="radio" name="previewMode" value="default" ${mode==='default'?'checked':''}> 기본값</label><label><input type="radio" name="previewMode" value="blank" ${mode==='blank'?'checked':''}> 공란 출력</label><label><input type="radio" name="previewMode" value="filled" ${mode==='filled'?'checked':''} ${hasSignatories?'':'disabled'}> 입력값 출력</label>${hasDefaults?'':`<span class="preview-mode-hint">학교 기본정보에 검사자·입회자 기본값이 없으면 기본값 모드도 성명은 공란으로 표시됩니다.</span>`}</div>`;
      }
      return '';
    };

    const pledgePreviewEditorHtml=()=>{
      if(type!=='privateContractPledge')return '';
      const count=Object.values(pledgeDraftResults).filter(Boolean).length;
      return `<div class="preview-field-editor pledge-preview-editor">${privateContractPledgeQuickHtml(pledgeBusinessType,'preview')}<div class="pledge-preview-summary"><strong>선택 ${count}/${(ReferenceData?.privateContractPledgeItems||[]).length}</strong><span>오른쪽 서류에 즉시 반영됩니다.</span></div><details class="pledge-preview-details"><summary>항목별 직접 수정</summary><div class="safety-edit-list pledge-edit-list compact">${privateContractPledgeRowsHtml(pledgeDraftResults,'previewPledge_')}</div></details><div class="preview-editor-savebar"><button class="button primary" type="button" id="savePreviewPledgeBtn">서약 선택값 저장</button></div></div>`;
    };

    const readPledgeDraftFromEditor=()=>{
      if(type!=='privateContractPledge')return;
      const selectedType=modalBody.querySelector('input[name="previewpledgeBusinessType"]:checked')?.value||'';
      if(selectedType)pledgeBusinessType=selectedType;
      (ReferenceData?.privateContractPledgeItems||[]).forEach(item=>{
        const value=modalBody.querySelector(`input[name="previewPledge_${item.key}"]:checked`)?.value||'';
        if(value)pledgeDraftResults[item.key]=value; else delete pledgeDraftResults[item.key];
      });
    };

    const readDraftFromEditor=()=>{
      previewFields.forEach(field=>{
        const el=modalBody.querySelector(`#docMissing_${CSS.escape(field)}`);
        if(!el)return;
        let value=el.value?.trim?.() ?? '';
        if(MONEY_FIELDS.has(field)) value=value?parseMoneyInput(value):'';
        draftValues[field]=value;
      });
    };

    const refreshPaper=()=>{
      readDraftFromEditor();
      readPledgeDraftFromEditor();
      const ctx=draftContext();
      const paper=modalBody.querySelector('.document-preview-paper');
      if(paper) paper.innerHTML=documentMarkup(type,ctx.project,ctx.options);
      const status=modalBody.querySelector('#previewEditorStatus');
      if(status){const wrap=document.createElement('div');wrap.innerHTML=statusHtml();status.replaceWith(wrap.firstElementChild);}
      previewFields.forEach(field=>{
        const wrap=modalBody.querySelector(`[data-preview-field-wrap="${CSS.escape(field)}"]`);
        if(wrap) wrap.classList.toggle('missing',(def.required||[]).includes(field)&&!meaningful(draftValue(field))&&!isBlankForm());
      });
    };

    const bindEditor=()=>{
      previewFields.forEach(field=>{
        const el=modalBody.querySelector(`#docMissing_${CSS.escape(field)}`);
        if(!el)return;
        el.addEventListener('input',refreshPaper);
        el.addEventListener('change',refreshPaper);
      });
      if(type==='privateContractPledge'){
        modalBody.querySelectorAll('input[name="previewpledgeBusinessType"]').forEach(input=>input.addEventListener('change',refreshPaper));
        modalBody.querySelectorAll('input[name^="previewPledge_"]').forEach(input=>input.addEventListener('change',()=>{readPledgeDraftFromEditor();renderOptions.blankPledge=false;paintPreview();}));
        modalBody.querySelector('[data-pledge-fill-yes]')?.addEventListener('click',()=>{
          privateContractPledgeNonConflictItems().forEach(item=>{pledgeDraftResults[item.key]='yes';});
          renderOptions.blankPledge=false; paintPreview(); showToast('체결제한을 제외한 서약·동의 항목을 ‘예’로 채웠습니다.');
        });
        modalBody.querySelector('[data-pledge-fill-conflict]')?.addEventListener('click',()=>{
          pledgeBusinessType=modalBody.querySelector('input[name="previewpledgeBusinessType"]:checked')?.value||pledgeBusinessType;
          if(!pledgeBusinessType){showToast('개인사업자 또는 법인사업자를 먼저 선택해주세요.','warn');return;}
          Object.assign(pledgeDraftResults,privateContractConflictDefaults(pledgeBusinessType));
          renderOptions.blankPledge=false; paintPreview();
          showToast(`${pledgeBusinessType==='individual'?'개인사업자':'법인사업자'} 기준으로 체결제한 기본값을 채웠습니다. 항목별 사실관계를 확인해주세요.`);
        });
        modalBody.querySelector('#savePreviewPledgeBtn')?.addEventListener('click',async()=>{
          readPledgeDraftFromEditor();
          p.privateContractPledge={results:{...pledgeDraftResults},businessType:pledgeBusinessType,updatedAt:new Date().toISOString()};
          if(!p.selectedDocuments?.includes('privateContractPledge'))p.selectedDocuments=[...(p.selectedDocuments||[]),'privateContractPledge'];
          p.updatedAt=new Date().toISOString(); await DB.put('projects',p); await loadState(); state.currentProjectId=p.id; p=currentProject()||p; resetPledgeDraft(); renderProjectDetail(); showToast('수의계약 통합서약서 작성내용을 저장했습니다.'); paintPreview();
        });
      }
      modalBody.querySelector('#savePreviewAllFieldsBtn')?.addEventListener('click',async()=>{
        readDraftFromEditor();
        const saved=await persistDocumentInfoFields(previewFields); if(!saved)return;
        p=saved; resetDraftValues(); renderProjectDetail(); showToast('변경사항을 저장했습니다.'); paintPreview();
      });
    };

    const paintPreview = () => {
      p=currentProject() || p;
      const ctx=draftContext();
      modalBody.innerHTML=`<div class="document-preview-workspace"><aside class="document-preview-inspector">${statusHtml()}${previewModeControls()}${pledgePreviewEditorHtml()}${documentPreviewEditorHtml(type,p,draftValues,isBlankForm())}</aside><div class="doc-preview-scroll document-preview-paper">${documentMarkup(type,ctx.project,ctx.options)}</div></div>`;
      initDateInputs(modalBody); initMoneyInputs(modalBody); bindEditor();
      modalBody.querySelectorAll('input[name="previewMode"]').forEach(input=>input.addEventListener('change',()=>{
        if(type==='privateContractPledge')renderOptions.blankPledge=input.value==='blank';
        if(type==='warrantyInspectionReport')renderOptions.warrantySignatoryMode=input.value;
        paintPreview();
      }));
    };

    openModal({eyebrow:`${def.stage} 서류 · 양식 ${def.version}`,title:`${def.label} 미리보기`,wide:true,body:'',actions:`<button class="button secondary" type="button" data-modal-close>닫기</button><button class="button primary" type="button" id="printDocumentBtn">인쇄 / PDF 저장</button>`});
    paintPreview();
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#printDocumentBtn').addEventListener('click', () => {
      readDraftFromEditor();
      readPledgeDraftFromEditor();
      const missing=currentMissing();
      if(missing.length){
        const labels=missing.map(field=>DOCUMENT_FIELD_LABELS[field]||field);
        showToast(`부족한 정보 ${missing.length}개를 먼저 확인해주세요: ${labels.slice(0,3).join(' · ')}${labels.length>3?' 외':''}`,'warn');
        return;
      }
      const ctx=draftContext();
      openChecklistSignatureModal([type],ctx.project,signatures=>{
        const options={...ctx.options,signatures};
        const pages=documentPages(type,ctx.project,options);
        printPagesInFrame(pages,`${def.label} - ${ctx.project.projectName || '공사'}`);
      });
    });
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
    @page { size:${pageWidth} ${pageHeight}; margin:0; }
    html,body { width:${pageWidth}; ${multi?'':`height:${pageHeight};`} margin:0; padding:0; background:#fff; }
    @media print {
      @page { size:${pageWidth} ${pageHeight}; margin:0; }
      html,body { width:${pageWidth} !important; ${multi?'':`height:${pageHeight} !important;`} margin:0 !important; padding:0 !important; }
      ${landscape?'body.print-only-document.print-landscape-document .document-print-page { width:297mm !important; height:210mm !important; min-width:297mm !important; min-height:210mm !important; max-width:297mm !important; max-height:210mm !important; }':''}
    }
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
      const doc = frame.contentDocument;
      if (doc && !doc.getElementById('final-print-orientation')) {
        const finalStyle = doc.createElement('style');
        finalStyle.id = 'final-print-orientation';
        finalStyle.textContent = `@media print { @page { size: ${pageWidth} ${pageHeight}; margin: 0; } html, body { width: ${pageWidth} !important; ${multi?'':`height: ${pageHeight} !important;`} margin: 0 !important; padding: 0 !important; } ${landscape?'body.print-only-document.print-landscape-document, body.print-only-document.print-landscape-document .document-print-page { width: 297mm !important; height: 210mm !important; min-width: 297mm !important; min-height: 210mm !important; max-width: 297mm !important; max-height: 210mm !important; }':''} }`;
        doc.head.appendChild(finalStyle);
      }
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

  function printAdministrativeDocument(type, p, renderOptions = {}) {
    const def = DOCUMENT_DEFINITIONS[type];
    if (!def || !p) return;
    if (isSafetyDocument(type) && renderOptions.blankSafety) {
      const pages=documentPages(type,p,{...renderOptions,signatures:{}});
      printPagesInFrame(pages, `${def.label} - ${p.projectName || '공사서류'}`);
      return;
    }
    openChecklistSignatureModal([type],p,signatures=>{
      const pages = documentPages(type,p,{...renderOptions,signatures});
      printPagesInFrame(pages, `${def.label} - ${p.projectName || '공사서류'}`);
    });
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
    const incompleteSafety = ordered.find(type => isSafetyDocument(type) && safetyRequiresCompletion(type) && !safetyChecklistComplete(p,type));
    if (incompleteSafety) { showToast(`${DOCUMENT_DEFINITIONS[incompleteSafety].label}을 먼저 작성해주세요.`, 'warn'); return; }
    const missing = batchMissingFields(ordered,p);
    if (missing.length) { openBatchMissingModal(ordered,missing); return; }

    const labels = ordered.map(type => DOCUMENT_DEFINITIONS[type].label);
    openChecklistSignatureModal(ordered,p,signatures=>{
      const pages = ordered.flatMap(type => documentPages(type,p,{signatures}));
      const title = `${p.projectName || '공사서류'} - ${labels.join(', ')}`;
      printPagesInFrame(pages,title,()=>recordPrintHistory(ordered,p));
    });
  }

  function openDocumentMissingModal(type, missing, previewOptions = {}) {
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
    modalActions.querySelector('#saveMissingDocFields').addEventListener('click', () => saveDocumentMissingFields(type, missing, previewOptions));
  }

  async function saveDocumentMissingFields(type, fields, previewOptions = {}) {
    const ok = await saveFieldsForDocuments(fields);
    if (!ok) return;
    closeModal();
    renderProjectDetail();
    openDocumentPreview(type, previewOptions);
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
    const rateSum=(kind)=>Number(work[kind])+Number(duration[kind])+Number(size[kind]);
    const averagePercent=(kind)=>rateSum(kind)/3;
    const rawCost=(kind)=>base*(averagePercent(kind)/100);
    const allowElectric=facilityUse!=='수도광열비';
    const allowWater=facilityUse!=='전력비';
    const electricRaw=allowElectric?rawCost('electric'):0;
    const waterRaw=allowWater?rawCost('water'):0;
    const electric=allowElectric?roundDown10(electricRaw):0;
    const water=allowWater?roundDown10(waterRaw):0;
    return {
      electricCost:electric, waterHeatCost:water, total:electric+water, sizeLabel:size.label,
      directMaterialCost:Number(directMaterialCost||0), directLaborCost:Number(directLaborCost||0),
      baseCost:base, contractAmountExVat:Number(contractAmount||0)/1.1,
      workElectricRate:Number(work.electric), durationElectricRate:Number(duration.electric), sizeElectricRate:Number(size.electric),
      workWaterRate:Number(work.water), durationWaterRate:Number(duration.water), sizeWaterRate:Number(size.water),
      electricRateSum:rateSum('electric'), electricAverageRate:averagePercent('electric'), electricRawCost:electricRaw,
      waterRateSum:rateSum('water'), waterAverageRate:averagePercent('water'), waterRawCost:waterRaw,
      allowElectric, allowWater
    };
  }

  function utilityResultHtml(result) {
    if (!result) return '<div class="utility-result-empty">금액을 입력하고 계산해주세요.</div>';
    return `<div class="utility-basis-summary"><div><span>직접재료비 · 직재</span><strong>${e(formatMoney(result.directMaterialCost))}</strong></div><i>+</i><div><span>직접노무비 · 직노</span><strong>${e(formatMoney(result.directLaborCost))}</strong></div><i>=</i><div class="utility-basis-total"><span>계산기준금액 · 직재+직노</span><strong>${e(formatMoney(result.baseCost))}</strong></div></div><div class="utility-result-grid"><div><span>전력비</span><strong>${e(formatMoney(result.electricCost))}</strong></div><div><span>수도광열비</span><strong>${e(formatMoney(result.waterHeatCost))}</strong></div><div class="utility-total"><span>공제금액 합계</span><strong>${e(formatMoney(result.total))}</strong></div></div><p class="utility-rate-note">계산기준금액(직재 + 직노) ${e(formatMoney(result.baseCost))} · 공사규모 ${e(result.sizeLabel)}</p>`;
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
    const existing=meaningful(u.total)?calculateUtilityCost({
      directMaterialCost:u.directMaterialCost, directLaborCost:u.directLaborCost,
      facilityUse:u.facilityUse || '수도광열비·전력비', workCategory, durationCategory,
      contractAmount:Number(p.currentContractAmount||0)
    }):null;
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

  function utilityRateTableHtml(inputs, result) {
    const work = UTILITY_RATES_2024.work;
    const duration = UTILITY_RATES_2024.duration;
    const size = UTILITY_RATES_2024.size;
    const pct = value => Number(value || 0).toFixed(3).replace(/0+$/,'').replace(/\.$/,'');
    const selected = value => value ? ' class="utility-rate-selected"' : '';
    const workDefs = [['건축','건축'],['토목','토목'],['산업설비','산업설비'],['조경','조경']];
    const durationDefs = [['6이하','6개월이하'],['6초과12이하','6~12개월'],['12초과36이하','12~36개월'],['36초과','36개월초과']];
    return `<table class="utility-print-rate-table"><thead>
      <tr><th rowspan="2">구분</th><th colspan="4">공사종류별</th><th colspan="4">공사기간별</th><th colspan="5">공사규모별</th></tr>
      <tr>${workDefs.map(([key,label])=>`<th${selected(inputs?.workCategory===key)}>${label}</th>`).join('')}${durationDefs.map(([key,label])=>`<th${selected(inputs?.durationCategory===key)}>${label}</th>`).join('')}${size.map(x=>`<th${selected(result?.sizeLabel===x.label)}>${x.label}</th>`).join('')}</tr>
    </thead><tbody>
      <tr><th>전력비</th>${workDefs.map(([key])=>`<td${selected(inputs?.workCategory===key)}>${pct(work[key].electric)}</td>`).join('')}${durationDefs.map(([key])=>`<td${selected(inputs?.durationCategory===key)}>${pct(duration[key].electric)}</td>`).join('')}${size.map(x=>`<td${selected(result?.sizeLabel===x.label)}>${pct(x.electric)}</td>`).join('')}</tr>
      <tr><th>수도광열비</th>${workDefs.map(([key])=>`<td${selected(inputs?.workCategory===key)}>${pct(work[key].water)}</td>`).join('')}${durationDefs.map(([key])=>`<td${selected(inputs?.durationCategory===key)}>${pct(duration[key].water)}</td>`).join('')}${size.map(x=>`<td${selected(result?.sizeLabel===x.label)}>${pct(x.water)}</td>`).join('')}</tr>
    </tbody></table>`;
  }

  function utilityCalculationPrintMarkup(p, inputs, result) {
    const won = value => Number(value || 0).toLocaleString('ko-KR');
    const wonDecimal = value => Number(value || 0).toLocaleString('ko-KR',{minimumFractionDigits:0,maximumFractionDigits:2});
    const pct = value => `${Number(value || 0).toFixed(3).replace(/0+$/,'').replace(/\.$/,'')}%`;
    const pctAvg = value => `${Number(value || 0).toFixed(6).replace(/0+$/,'').replace(/\.$/,'')}%`;
    const durationLabel = ({'6이하':'6개월 이하','6초과12이하':'6~12개월','12초과36이하':'12~36개월','36초과':'36개월 초과'})[inputs.durationCategory] || inputs.durationCategory;
    const formula = (label, enabled, workRate, durationRate, sizeRate, rateSum, averageRate, rawAmount, amount) => enabled
      ? `<section class="utility-calc-block">
          <div class="utility-calc-head"><strong>${label}</strong><b>${won(amount)} 원</b></div>
          <div class="utility-rate-flow">
            <div><small>공사종류 · ${e(inputs.workCategory)}</small><strong>${pct(workRate)}</strong></div><i>+</i>
            <div><small>공사기간 · ${e(durationLabel)}</small><strong>${pct(durationRate)}</strong></div><i>+</i>
            <div><small>공사규모 · ${e(result.sizeLabel)}</small><strong>${pct(sizeRate)}</strong></div><i>=</i>
            <div><small>요율 합계</small><strong>${pct(rateSum)}</strong></div><i>÷ 3</i>
            <div class="utility-average-rate"><small>평균요율</small><strong>${pctAvg(averageRate)}</strong></div>
          </div>
          <div class="utility-cost-flow"><span><b>직재 + 직노</b> ${won(result.baseCost)}원</span><i>×</i><span>평균요율 ${pctAvg(averageRate)}</span><i>=</i><span>${wonDecimal(rawAmount)}원</span><i>→</i><strong>10원 미만 절사&nbsp; ${won(amount)}원</strong></div>
        </section>`
      : `<section class="utility-calc-block disabled"><div class="utility-calc-head"><strong>${label}</strong><b>적용 제외</b></div></section>`;
    return `<article class="paper-a4-landscape utility-cost-sheet document-print-page">
      <div class="utility-print-project"><strong>[공사명]</strong><span>${e(p.projectName || '')}</span></div>
      <h1>□ 2024년도 기준 완성공사 원가통계(경비율)</h1>
      ${utilityRateTableHtml(inputs,result)}
      <div class="utility-print-notes">
        <p>※ 파란 표시: 이번 공사에 실제 적용된 공사종류·기간·규모 구간</p>
        <p>※ 전기·통신·소방·전문공사는 건축요율 적용 / 공사규모는 부가세 제외 계약금액 기준</p>
        <p>※ 출처: 대한건설협회 「2024년 완성공사원가분석」 기준</p>
      </div>
      <h1>□ 전력비·수도광열비 계산식</h1>
      <div class="utility-print-basis">
        <div><span>직접재료비 <b>(직재)</b></span><strong>${won(inputs.directMaterialCost)} 원</strong></div><i>+</i>
        <div><span>직접노무비 <b>(직노)</b></span><strong>${won(inputs.directLaborCost)} 원</strong></div><i>=</i>
        <div class="utility-basis-emphasis"><span>계산기준금액 <b>(직재 + 직노)</b></span><strong>${won(result.baseCost)} 원</strong></div>
      </div>
      <div class="utility-print-contract-meta"><span>계약금액 <b>${won(inputs.contractAmount)}원</b></span><span>부가세 제외 <b>${won(result.contractAmountExVat)}원</b></span><span>공사규모 <b>${e(result.sizeLabel)}</b></span><span>시설사용 <b>${e(inputs.facilityUse)}</b></span></div>
      <div class="utility-print-formulas">
        ${formula('1. 전력비', result.allowElectric, result.workElectricRate, result.durationElectricRate, result.sizeElectricRate, result.electricRateSum, result.electricAverageRate, result.electricRawCost, result.electricCost)}
        ${formula('2. 수도광열비', result.allowWater, result.workWaterRate, result.durationWaterRate, result.sizeWaterRate, result.waterRateSum, result.waterAverageRate, result.waterRawCost, result.waterHeatCost)}
      </div>
      <div class="utility-print-total"><span>공제금액 합계</span><strong>${won(result.total)} 원</strong></div>
      <p class="utility-print-footnote">공사서류 작성지원에 저장된 공사정보와 입력한 직접재료비·직접노무비를 기준으로 계산</p>
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

  function documentContext(p, options = {}) {
    const previewValues=options.__previewValues || {};
    return {
      project: p,
      school: options.__previewSchool || state.school || {},
      payout: options.__previewPayout || payoutForProject(p) || {},
      value: field => Object.prototype.hasOwnProperty.call(previewValues,field) ? previewValues[field] : documentValue(field, p),
      signature: type => options.signatures?.[type] || '',
      renderOptions: options,
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

  function documentPages(type, p, options = {}) {
    return Documents.renderPages(type, documentContext(p, options));
  }

  function documentMarkup(type, p, options = {}) {
    return documentPages(type, p, options).join('<div class="document-page-gap" aria-hidden="true"></div>');
  }

  function openChecklistSignatureModal(types, p, onContinue) {
    const targets = [...new Set(types)].filter(type => isSafetyDocument(type) && safetyChecklistComplete(p,type));
    if (!targets.length) { onContinue({}); return; }
    const checklistSummary = targets.map(type => {
      const saved = safetyChecklistFor(p,type) || {};
      return `<li><strong>${e(DOCUMENT_DEFINITIONS[type]?.label || type)}</strong><span>점검자 ${e(saved.inspector || '미입력')}</span></li>`;
    }).join('');
    const applyTargetText = targets.length > 1 ? `작성 완료된 체크리스트 ${targets.length}종에 같은 서명이 적용됩니다.` : '현재 체크리스트의 점검자 서명란에 적용됩니다. 서명 없이 출력할 수도 있습니다.';
    openModal({
      eyebrow:'출력 전 확인', title:'점검자 서명', wide:true,
      body:`<div class="signature-print-notice"><strong>서명은 이번 출력에만 사용됩니다.</strong><span>${e(applyTargetText)} 마우스·터치·펜으로 서명하세요. 브라우저 저장소와 백업파일에는 저장하지 않습니다.</span></div>
        <ul class="signature-target-list">${checklistSummary}</ul>
        <div class="signature-pad-wrap"><canvas id="checklistSignatureCanvas" class="signature-pad" aria-label="점검자 서명 입력 영역"></canvas><span>이 영역에 서명하세요</span></div>`,
      actions:`<button class="button ghost" type="button" id="clearChecklistSignatureBtn">다시 쓰기</button><span class="modal-action-spacer"></span><button class="button secondary" type="button" id="printWithoutChecklistSignatureBtn">서명 없이 출력</button><button class="button primary" type="button" id="applyChecklistSignatureBtn" disabled>서명 적용 후 출력</button>`
    });

    const canvas = modalBody.querySelector('#checklistSignatureCanvas');
    const clearBtn = modalActions.querySelector('#clearChecklistSignatureBtn');
    const skipBtn = modalActions.querySelector('#printWithoutChecklistSignatureBtn');
    const applyBtn = modalActions.querySelector('#applyChecklistSignatureBtn');
    let context = null;
    let drawing = false;
    let hasInk = false;
    let last = null;

    const prepareCanvas = () => {
      if (!canvas) return;
      const rect = canvas.getBoundingClientRect();
      const ratio = Math.max(1, window.devicePixelRatio || 1);
      canvas.width = Math.max(1, Math.round(rect.width * ratio));
      canvas.height = Math.max(1, Math.round(rect.height * ratio));
      context = canvas.getContext('2d');
      context.setTransform(ratio,0,0,ratio,0,0);
      context.lineWidth = 2.2;
      context.lineCap = 'round';
      context.lineJoin = 'round';
      context.strokeStyle = '#111827';
    };
    const point = ev => {
      const rect = canvas.getBoundingClientRect();
      return {x:ev.clientX-rect.left,y:ev.clientY-rect.top};
    };
    const start = ev => {
      if (!context) return;
      ev.preventDefault();
      drawing = true;
      last = point(ev);
      canvas.setPointerCapture?.(ev.pointerId);
    };
    const move = ev => {
      if (!drawing || !context || !last) return;
      ev.preventDefault();
      const next = point(ev);
      context.beginPath();
      context.moveTo(last.x,last.y);
      context.lineTo(next.x,next.y);
      context.stroke();
      last = next;
      hasInk = true;
      applyBtn.disabled = false;
    };
    const end = ev => {
      if (!drawing) return;
      drawing = false;
      last = null;
      canvas.releasePointerCapture?.(ev.pointerId);
    };
    canvas?.addEventListener('pointerdown',start);
    canvas?.addEventListener('pointermove',move);
    canvas?.addEventListener('pointerup',end);
    canvas?.addEventListener('pointercancel',end);
    clearBtn?.addEventListener('click',()=>{
      if (context) context.clearRect(0,0,canvas.width,canvas.height);
      hasInk = false;
      applyBtn.disabled = true;
    });
    skipBtn?.addEventListener('click',()=>{ closeModal(); onContinue({}); });
    applyBtn?.addEventListener('click',()=>{
      if (!hasInk) { showToast('서명란에 서명하거나 서명 없이 출력을 선택해주세요.','warn'); return; }
      const image = canvas.toDataURL('image/png');
      const signatures = Object.fromEntries(targets.map(type => [type,image]));
      closeModal();
      onContinue(signatures);
    });
    window.requestAnimationFrame(prepareCanvas);
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
    const inspectorName = current.inspectorName || String(current.inspector || '').replace(/^행정실장\s*/, '').trim();
    const witnessName = current.witnessName || current.witness || '';
    openModal({
      eyebrow:'행정기관 하자관리', title:record ? '하자검사 기록 수정' : '하자검사 기록 추가',
      body:`<div class="notice"><strong>하자검사조서의 검사자는 행정실장 기준으로 출력합니다.</strong><br>검사자·입회자 성명만 필요할 때 입력하세요. 미리보기에서 ‘기본값 / 공란 출력 / 입력값 출력’ 중 원하는 방식을 선택할 수 있습니다.</div><div class="modal-grid" style="margin-top:16px">
        ${modalDateField('warrantyDate','검사일',current.date || '')}
        <div class="field"><label for="warrantyInspectorNameInput">검사자 성명 <span class="label-optional">선택</span></label><input id="warrantyInspectorNameInput" value="${e(inspectorName)}" placeholder="행정실장 성명"></div>
        <div class="field"><label for="warrantyWitnessNameInput">입회자 성명 <span class="label-optional">선택</span></label><input id="warrantyWitnessNameInput" value="${e(witnessName)}" placeholder="비워두어도 됩니다"></div>
        <div class="field"><label for="warrantyHasDefect">하자 유무</label><select id="warrantyHasDefect"><option value="">선택</option><option value="no" ${current.hasDefect==='no'?'selected':''}>이상 없음</option><option value="yes" ${current.hasDefect==='yes'?'selected':''}>하자 있음</option></select></div>
        <div class="field full"><label for="warrantyResult">검사결과</label><textarea id="warrantyResult">${e(current.result || '')}</textarea></div>
        <div class="field full"><label for="warrantyIssueDetails">하자발생내용</label><textarea id="warrantyIssueDetails">${e(current.issueDetails || '')}</textarea></div>
        <div class="field full"><label for="warrantyActions">처리사항</label><textarea id="warrantyActions">${e(current.actions || '')}</textarea></div>
        <div class="field full"><label for="warrantyNotes">기타참고사항</label><textarea id="warrantyNotes">${e(current.notes || '')}</textarea></div>
      </div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="saveWarrantyInspectionBtn">${previewAfterSave?'저장하고 미리보기':'저장'}</button>`
    });
    initDateInputs(modalBody);
    modalActions.querySelector('[data-modal-close]').addEventListener('click',closeModal);
    modalActions.querySelector('#saveWarrantyInspectionBtn').addEventListener('click',()=>saveWarrantyInspection(record?.id || '',previewAfterSave));
  }

  async function saveWarrantyInspection(recordId = '', previewAfterSave = false) {
    const p = currentProject();
    if (!p) return;
    const date = modalBody.querySelector('#warrantyDate')?.value || '';
    if (!date) { showToast('검사일을 입력해주세요.','warn'); return; }
    const list = Array.isArray(p.warrantyInspections) ? [...p.warrantyInspections] : [];
    const existingIndex = recordId ? list.findIndex(x=>x.id===recordId) : -1;
    const base = existingIndex >= 0 ? list[existingIndex] : {};
    const inspectorPosition = '행정실장';
    const inspectorName = modalBody.querySelector('#warrantyInspectorNameInput')?.value?.trim() || '';
    const witnessPosition = '';
    const witnessName = modalBody.querySelector('#warrantyWitnessNameInput')?.value?.trim() || '';
    const next = {
      ...base,
      id: base.id || DB.uuid(),
      date,
      inspectorPosition,
      inspectorName,
      inspector: [inspectorPosition,inspectorName].filter(Boolean).join(' '),
      witnessPosition,
      witnessName,
      witness: [witnessPosition,witnessName].filter(Boolean).join(' '),
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
    if (previewAfterSave) openDocumentPreview('warrantyInspectionReport',{mode:'default'});
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

  function confirmResetProjects() {
    const count=state.projects.length;
    if(!count){showToast('초기화할 공사자료가 없습니다.');return;}
    openModal({
      eyebrow:'공사자료 초기화', title:`등록된 공사 ${count}건을 모두 삭제할까요?`,
      body:`<div class="notice danger"><strong>공사자료만 전체 삭제합니다.</strong><br>공사 ${count}건과 각 공사에 저장된 체크리스트·서류 입력값·하자검사 이력이 함께 삭제됩니다.<br><br>학교 기본정보와 업체 보관함·지급정보는 유지됩니다. 필요한 경우 먼저 ‘전체 백업’을 받아두세요.</div>`,
      actions:`<button class="button secondary" type="button" data-modal-close>취소</button><button class="button danger" type="button" id="resetProjectsConfirmBtn">공사자료 ${count}건 전체 삭제</button>`
    });
    modalActions.querySelector('[data-modal-close]')?.addEventListener('click',closeModal);
    modalActions.querySelector('#resetProjectsConfirmBtn')?.addEventListener('click',async()=>{
      await DB.clear('projects');
      state.projects=[]; state.currentProjectId=null; state.selectedDocuments=new Set(); state.selectionProjectId=null; state.activeWarrantyInspectionId=null;
      closeModal(); renderDashboard(); showToast('공사자료를 전체 초기화했습니다.');
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
        throw new Error('이 파일은 공사관리대장으로 보입니다. 「공사관리대장 불러오기」에서 선택해주세요.');
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
    const title = isEdufine ? '에듀파인 정보 업데이트 결과' : '공사관리대장 불러오기 결과';
    const guide = isEdufine
      ? '기존 공사와 계약번호·공사명·업체·금액·계약일을 비교했습니다. 빈 값은 보완하고, 서로 다른 값은 선택한 경우에만 바꿉니다.'
      : '공사관리대장의 각 행을 기존 공사와 비교했습니다. 처음 보는 공사는 새로 만들고, 기존 공사는 부족한 값만 보완합니다.';
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
      eyebrow: '공사관리대장', title: '공사관리대장 엑셀로 내보내기',
      body: `<div class="notice">공사관리대장 양식의 열과 서식을 기준으로 생성합니다.</div><div class="field" style="margin-top:16px"><label>범위</label><select id="auditExportYear"><option value="all">전체 연도 (${state.projects.length}건)</option>${years.map(y=>`<option value="${e(y)}">${e(y)}회계연도 (${state.projects.filter(p=>p.fiscalYear===y).length}건)</option>`).join('')}</select></div>`,
      actions: `<button class="button secondary" type="button" data-modal-close>취소</button><button class="button primary" type="button" id="auditExportConfirm">엑셀 받기</button>`
    });
    modalActions.querySelector('[data-modal-close]').addEventListener('click', closeModal);
    modalActions.querySelector('#auditExportConfirm').addEventListener('click', async () => {
      const year = modalBody.querySelector('#auditExportYear').value;
      const list = year === 'all' ? state.projects : state.projects.filter(p => p.fiscalYear === year);
      try { await Excel.exportAuditWorkbook(list, { year: year === 'all' ? '' : year }); closeModal(); showToast(`${list.length}건의 공사관리대장 엑셀을 만들었습니다.`); }
      catch (err) { showToast(err.message, 'danger'); }
    });
  }

  async function backupAll() {
    const data = await DB.exportBackup();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type:'application/json;charset=utf-8' });
    const d = new Date();
    const date = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
    Excel.downloadBlob(blob, `공사서류작성지원_전체백업_${date}.json`);
    showToast(`공사 ${data.projects.length}건 · 업체 ${data.vendors.length}건 · 지급정보 ${data.payouts?.length || 0}건을 백업했습니다.`);
  }

  async function handleBackupFile(file) {
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (data.app !== 'construction-info-hub') throw new Error('공사서류 작성지원 백업파일이 아닙니다.');
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
      eyebrow:'도움말', title:`공사서류 작성지원 v${APP_VERSION}`,
      body:`<div class="notice"><strong>이 도구의 목적</strong><br>행정실에서 공사 관련 서류를 빠르게 작성·확인·출력하기 위한 도구입니다. 공통정보는 한 번만 입력하고 여러 서류에 다시 사용합니다.</div>
      <div style="display:grid;gap:16px;margin-top:18px;font-size:14px">
        <div><strong>1. 출력할 공사를 선택</strong><p class="muted">공사를 열면 ‘서류 작성’ 화면이 먼저 보입니다. 필요한 서류를 선택해 미리보기부터 확인하세요.</p></div>
        <div><strong>2. 공통정보는 한 번만 입력</strong><p class="muted">공사명·업체·대표자·주소·계약금액·날짜 등은 공사정보에 저장되어 여러 서류에 자동 반영됩니다. 공사관리대장.xlsx 또는 에듀파인 자료관리목록.xlsx로 기존 값을 가져올 수도 있습니다.</p></div>
        <div><strong>3. 미리보기에서 바로 수정</strong><p class="muted">잘못 입력한 값이나 부족한 정보는 각 서류 미리보기의 ‘입력정보 수정’에서 고칩니다. 저장한 값은 다른 서류에도 함께 반영됩니다.</p></div>
        <div><strong>4. 체크리스트 작성·빈 양식</strong><p class="muted">안전·보건 체크리스트는 ‘체크리스트 작성’과 ‘빈 양식 미리보기’로 구분합니다. 작성 화면에서는 ‘미응답 모두 예’를 사용할 수 있고, 인쇄 직전 서명을 넣거나 서명 없이 출력할 수 있습니다.</p></div>
        <div><strong>5. 여러 서류 묶음 출력</strong><p class="muted">필요한 경우 서류 화면 아래 ‘여러 서류를 한 번에 출력’을 열어 개별 선택 또는 단계별 세트를 사용합니다.</p></div>
        <div><strong>6. 원클릭 엑셀 별도 확인</strong><p class="muted">현장대리인계·공정표·직접시공계획서와 도시가스·노무비 관련 서류는 서류 화면의 별도 확인 목록에서 안내합니다.</p></div>
        <div><strong>기준 자료</strong><p class="muted">${e(REFERENCE_PROGRAM)} 버전을 기준으로 서식과 점검항목을 구성했습니다.</p></div>
        <div><strong>보안</strong><p class="muted">공사정보와 업로드한 엑셀 내용은 서버로 전송하지 않고 이 브라우저에 저장합니다. 공용 Windows 계정에서는 PC 접근통제와 백업이 필요합니다.</p></div>
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
      if(action==='import')openExcelPicker('auto'); if(action==='export')openExportAuditModal(); if(action==='vendors')openVendorLibrary(); if(action==='school')openSchoolModal(); if(action==='backup')backupAll(); if(action==='restore')backupFileInput.click(); if(action==='reset-projects')confirmResetProjects(); if(action==='help')openHelp();
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
