(() => {
  'use strict';

  const ReferenceData = globalThis.ConstructionReferenceData;

  /*
   * 공사서류 템플릿 레지스트리
   * - 서류 정의/필수값/인쇄순서/출력 마크업을 앱 로직에서 분리한다.
   * - 양식이 개정되면 공사 데이터 구조 대신 이 파일의 해당 템플릿만 교체한다.
   */

  const DEFINITIONS = {
    standardContract: {
      key:'standardContract', label:'공사도급표준계약서', outputTitle:'공 사 도 급 표 준 계 약 서', stage:'계약', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'발주기관과 계약상대자가 체결하는 공사도급 표준계약서',
      required:['schoolName','schoolAddress','principal','projectName','contractNumber','currentContractAmount','contractSecurityAmount','contractDate','plannedStartDate','completionDueDate','vendorName','businessNumber','vendorAddress','vendorPhone','representative','workType','delayPenaltyRate','priceAdjustmentMethod','defectSecurityRate','defectSecurityAmount','defectPeriodYears']
    },
    acceptanceTerms: {
      key:'acceptanceTerms', label:'승낙사항', outputTitle:'승 낙 사 항 (공 사 집 행)', stage:'계약', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'공사 집행 시 계약상대자가 확인하는 승낙사항',
      required:['projectName','contractDate','plannedStartDate','completionDueDate','vendorName','businessNumber','vendorAddress','representative','delayPenaltyRate','defectPeriodYears']
    },
    useSealForm: {
      key:'useSealForm', label:'사용인감계', outputTitle:'사 용 인 감 계', stage:'계약', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'계약 관련 업무에 사용할 사용인감을 신고하는 서류',
      required:['schoolName','projectName','contractDate','vendorName','vendorAddress','representative']
    },
    privateContractPledge: {
      key:'privateContractPledge', label:'수의계약 통합서약서', outputTitle:'수 의 계 약 통 합 서 약 서', stage:'계약', version:'2026.05 수정', pages:2, owner:'vendor',
      description:'수의계약 관련 각서·청렴·안전·개인정보 동의를 통합한 서약서. 예·아니오를 미리 선택하거나 빈 양식으로 출력할 수 있습니다.',
      required:['schoolName','projectName','contractDate','vendorName','businessNumber','vendorAddress','vendorPhone','representative']
    },
    startReport: {
      key:'startReport', label:'착공계', outputTitle:'착 공 신 고 서', stage:'착공', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'착공 신고 시 제출하는 기본 착공계',
      required:['schoolName','projectName','currentContractAmount','contractDate','startDate','completionDueDate','vendorName','businessNumber','vendorAddress','representative']
    },
    utilityPaymentPledge: {
      key:'utilityPaymentPledge', label:'수도·전기료 납부각서', outputTitle:'각 서', stage:'착공', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'학교의 수도·전력을 사용한 경우 계산된 사용료를 납부하겠다는 업체 각서. 현장대리인 이름은 공란으로도 출력할 수 있습니다.',
      required:['schoolName','projectName','currentContractAmount','contractDate','startDate','completionDueDate','vendorName','representative']
    },
    completionReport: {
      key:'completionReport', label:'준공계', outputTitle:'준 공 계', stage:'준공', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'공사 완료 후 제출하는 준공계',
      required:['schoolName','projectName','currentContractAmount','contractDate','startDate','completionDueDate','actualCompletionDate','vendorName','businessNumber','vendorAddress','representative']
    },
    completionInspectionRequest: {
      key:'completionInspectionRequest', label:'준공검사원', outputTitle:'준 공 검 사 원', stage:'준공', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'준공 사실을 확인하고 검사를 요청하는 서류',
      required:['schoolName','projectName','currentContractAmount','contractDate','startDate','completionDueDate','actualCompletionDate','vendorName','businessNumber','vendorAddress','representative']
    },
    supervisionReport: {
      key:'supervisionReport', label:'공사감독조서', outputTitle:'공 사 감 독 조 서', stage:'준공', version:'2026.05 수정', pages:1, owner:'agency',
      description:'공사감독자가 현장 감독 결과를 확인하는 기관용 조서',
      required:['schoolName','projectName','vendorName','representative','currentContractAmount','contractDate','startDate','completionDueDate','actualCompletionDate','supervisor']
    },
    completionInspectionRecord: {
      key:'completionInspectionRecord', label:'준공검사조서', outputTitle:'준 공 검 사 조 서', stage:'준공', version:'2026.05 수정', pages:1, owner:'agency',
      description:'준공검사 결과와 검사자·입회자를 기록하는 기관용 조서',
      required:['schoolName','projectName','vendorName','representative','currentContractAmount','contractDate','startDate','completionDueDate','actualCompletionDate','completionInspectionDate','settlementAmount','inspector','witness']
    },
    defectSecurityDeposit: {
      key:'defectSecurityDeposit', label:'하자보증금납부서', outputTitle:'하 자 보 수 보 증 금 납 부 서', stage:'준공', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'하자보수보증금의 금액·기간·납부방법을 확인하여 제출하는 서류',
      required:['schoolName','projectName','currentContractAmount','contractDate','actualCompletionDate','defectSecurityRate','defectSecurityAmount','defectStartDate','defectEndDate','defectSecurityType','vendorName','businessNumber','vendorAddress','representative']
    },
    completionSettlementAgreement: {
      key:'completionSettlementAgreement', label:'준공정산동의서', outputTitle:'준 공 정 산 동 의 서', stage:'준공', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'계약금액과 준공정산금액의 차액 및 감액 사유에 동의하는 업체 서류',
      required:['schoolName','projectName','currentContractAmount','settlementAmount','settlementReductionReason','actualCompletionDate','vendorName','vendorAddress','representative']
    },
    paymentRequest: {
      key:'paymentRequest', label:'대금청구서', outputTitle:'대 금 청 구 서', stage:'지출', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'준공 후 계약대금을 지정계좌로 청구하는 서류',
      required:['schoolName','projectName','currentContractAmount','settlementAmount','priorPaymentAmount','deductionAmount','completionInspectionDate','vendorName','vendorAddress','representative','bankName','accountNumber','accountHolder']
    },
    constructionLedger: {
      key:'constructionLedger', label:'공사대장', outputTitle:'공  사  대  장', stage:'관리', version:'2026.05 수정', pages:1, owner:'agency',
      description:'개별 공사의 계약·공정·지급·검사·하자 내용을 한 장에 관리하는 기관용 대장',
      required:['projectName','vendorName','currentContractAmount','contractMethod','contractDate','startDate','completionDueDate']
    },
    warrantyInspectionReport: {
      key:'warrantyInspectionReport', label:'하자검사조서', outputTitle:'하  자  검  사  조  서', stage:'하자', version:'2026.05 수정', pages:1, owner:'agency',
      description:'완료공사의 하자(만료)검사 결과를 기록하는 기관용 조서. 미리보기에서 검사자·입회자를 기본값·공란·입력값 중 선택하고 수정할 수 있습니다.',
      required:['projectName','vendorName','representative','workType','currentContractAmount','startDate','actualCompletionDate','defectEndDate','warrantyInspectionDate']
    },
    warrantyLedger: {
      key:'warrantyLedger', label:'하자대장', outputTitle:'하 자 검 사 대 장', stage:'하자', version:'2026.05 수정', pages:1, owner:'agency',
      description:'공사별 하자보증기간과 정기 하자검사 이력을 누적 관리하는 기관용 대장',
      required:['projectName','vendorName','currentContractAmount','contractDate','startDate','completionDueDate','actualCompletionDate','defectStartDate','defectEndDate']
    },
    safetyGeneral: {
      key:'safetyGeneral', label:'안전·보건 체크리스트', outputTitle:'공사(용역) 안전·보건 체크리스트(공통)', stage:'안전·보건', version:'2026.05 수정', pages:1, owner:'agency',
      description:'기관(학교)에서 점검하고 업체에서 확인하여 자체 보관하는 공통 체크리스트', required:['schoolName','projectName','vendorName']
    },
    safetyFall: {
      key:'safetyFall', label:'추락재해 예방 체크리스트', outputTitle:'추락재해 예방 체크리스트', stage:'안전·보건', version:'2026.05 수정', pages:2, owner:'vendor',
      description:'추락위험 작업 시 공사업체가 작성·제출하고 기관이 확인하는 체크리스트', required:['projectName','vendorName']
    },
    safetyElectrical: {
      key:'safetyElectrical', label:'감전재해 예방 체크리스트', outputTitle:'전기공사 감전재해 예방 체크리스트', stage:'안전·보건', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'감전위험 작업 시 공사(용역)업체가 작성·제출하고 기관이 확인하는 체크리스트', required:['projectName','vendorName']
    },
    safetyConfined: {
      key:'safetyConfined', label:'밀폐공간 질식재해예방 체크리스트', outputTitle:'밀폐공간 질식재해예방 체크리스트', stage:'안전·보건', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'밀폐공간 작업 시 공사(용역)업체가 작성·제출하고 기관이 확인하는 체크리스트', required:['projectName','vendorName']
    },
    safetyIndustrial: {
      key:'safetyIndustrial', label:'일반 산업재해 예방 체크리스트', outputTitle:'일반 산업재해 예방 체크리스트', stage:'안전·보건', version:'2026.05 수정', pages:1, owner:'vendor',
      description:'일반 산업재해 우려 작업 시 공사(용역)업체가 작성·제출하고 기관이 확인하는 체크리스트', required:['projectName','vendorName']
    }
  };

  const FIELD_LABELS = {
    schoolName:'기관명', schoolAddress:'기관 주소', principal:'기관 대표자(학교장 등)',
    projectName:'공사명', workType:'공종', contractNumber:'계약번호', currentContractAmount:'계약금액', contractDate:'계약일', plannedStartDate:'착공예정일', startDate:'착공일', completionDueDate:'준공기한', actualCompletionDate:'실제 준공일',
    completionInspectionDate:'준공검사일', settlementAmount:'준공정산금액', settlementReductionReason:'준공 감액 사유', priorPaymentAmount:'기지급액', deductionAmount:'공제금액',
    vendorName:'업체명', businessNumber:'사업자등록번호', vendorAddress:'사업장 주소', vendorPhone:'업체 전화번호', representative:'대표자',
    contractSecurityAmount:'계약보증금액', delayPenaltyRate:'지연배상금률', priceAdjustmentMethod:'물가변동 계약금액 조정방법',
    defectSecurityType:'하자보증방법', defectSecurityRate:'하자보증률', defectSecurityAmount:'하자보증금액', defectPeriodYears:'하자담보기간(년)', defectStartDate:'하자 시작일', defectEndDate:'하자 종료일',
    siteManager:'현장대리인', supervisor:'공사감독', inspector:'검사자', witness:'준공검사 입회자', bankName:'은행명', accountNumber:'계좌번호', accountHolder:'예금주명',
    warrantyInspectionDate:'하자검사일', warrantyInspector:'하자검사자', warrantyInspectorPosition:'하자검사자 직위', warrantyInspectorName:'하자검사자 성명', warrantyWitness:'하자검사 입회자', warrantyWitnessPosition:'하자검사 입회자 직위', warrantyWitnessName:'하자검사 입회자 성명', warrantyInspectionResult:'검사결과', warrantyIssueDetails:'하자발생내용', warrantyActions:'처리사항', warrantyNotes:'기타참고사항'
  
  };

  const PRINT_ORDER = [
    'standardContract','acceptanceTerms','useSealForm','privateContractPledge',
    'startReport','utilityPaymentPledge','completionReport','completionInspectionRequest','supervisionReport','completionInspectionRecord','defectSecurityDeposit','completionSettlementAgreement','paymentRequest',
    'constructionLedger','safetyGeneral','safetyFall','safetyElectrical','safetyConfined','safetyIndustrial','warrantyInspectionReport','warrantyLedger'
  ];

  const SETS = {
    contract: { label:'계약서류 4종', types:['standardContract','acceptanceTerms','useSealForm','privateContractPledge'] },
    start: { label:'착공서류', types:['startReport','utilityPaymentPledge'] },
    completion: { label:'준공서류', types:['completionReport','completionInspectionRequest','supervisionReport','completionInspectionRecord','defectSecurityDeposit','completionSettlementAgreement'] },
    payment: { label:'지출서류', types:['paymentRequest'] },
    agencyManagement: { label:'행정기관 관리서류', types:['constructionLedger','supervisionReport','completionInspectionRecord','safetyGeneral','warrantyInspectionReport','warrantyLedger'] },
    safety: { label:'안전·보건 서류', types:['safetyGeneral','safetyFall','safetyElectrical','safetyConfined','safetyIndustrial'] },
    all: { label:'전체 21종', types:[...PRINT_ORDER] }
  };

  function common(ctx) {
    const p = ctx.project || {};
    const school = ctx.school || {};
    const payout = ctx.payout || {};
    const h = ctx.helpers || {};
    return { p, school, payout, h };
  }

  function sealChoice(text) {
    const value = String(text || '');
    const isLetter = /각서|지급각서/.test(value);
    const isBond = /보증서|증권|보증증권|이행보증/.test(value);
    return `[${isLetter?'✓':' '}] 각서   [${isBond?'✓':' '}] 보증서`;
  }

  function choiceBox(status, target) { return status === target ? '[✓]' : '[ ]'; }
  function plainChoice(status = '') { return `${choiceBox(status,'yes')} 예   ${choiceBox(status,'no')} 아니오`; }
  function plainChoiceWithNone(status = '') { return `${choiceBox(status,'yes')} 예   ${choiceBox(status,'no')} 아니오\n${choiceBox(status,'na')} 해당없음`; }

  function renderStandardContract(ctx) {
    const { p, school, h } = common(ctx);
    const contractSecurityAmount = ctx.value ? ctx.value('contractSecurityAmount') : p.contractSecurityAmount;
    const defectSecurityAmount = ctx.value ? ctx.value('defectSecurityAmount') : p.defectSecurityAmount;
    const defectRate = h.percentText(p.defectSecurityRate);
    const warrantyItems = Array.isArray(p.warrantyItems) && p.warrantyItems.length ? p.warrantyItems : [{subcategory:p.workType,years:p.defectPeriodYears}];
    const warrantyRows = warrantyItems.slice(0,4).map((item,index)=>`<tr><td>${h.e(item.subcategory || item.category || p.workType || '')}</td><td>${index===0?h.e(h.moneyNumberText(p.currentContractAmount)):''}</td><td>${index===0?`(${h.e(defectRate)}%)&nbsp;&nbsp; ${h.e(h.moneyNumberText(defectSecurityAmount))}`:''}</td><td>${item.years?h.e(String(item.years))+'년':''}</td></tr>`).join('');
    return [`<article class="paper-a4 admin-document standard-contract document-print-page">
      <div class="standard-contract-title-row"><h1>공 사 도 급 표 준 계 약 서</h1><table><tbody><tr><th>계약번호 제</th><td>${h.e(p.contractNumber)}</td><td>호</td></tr><tr><th>공고번호 제</th><td></td><td>호</td></tr></tbody></table></div>
      <table class="standard-contract-table party-table"><colgroup><col style="width:10mm"><col style="width:27mm"><col style="width:31mm"><col style="width:50mm"><col style="width:31mm"><col style="width:20mm"><col style="width:21mm"></colgroup><tbody>
        <tr><th class="vertical-head" rowspan="4">계<br>약<br>자</th><th>발 주 처</th><td colspan="5">${h.e(school.name)}</td></tr>
        <tr><th rowspan="3">계약상대자</th><td>상호 또는 법인명칭 :</td><td>${h.e(p.vendorName)}</td><td>사업자등록번호 :</td><td colspan="2">${h.e(h.businessNumber(p.businessNumber))}</td></tr>
        <tr><td>주소 :</td><td>${h.e(p.vendorAddress)}</td><td>전화번호 :</td><td colspan="2">${h.e(p.vendorPhone)}</td></tr>
        <tr><td>대표자 :</td><td colspan="4">${h.e(p.representative)}</td></tr>
        <tr><th class="vertical-head" rowspan="3"></th><th rowspan="3">연대보증인</th><td>상호 또는 법인명칭 :</td><td></td><td>사업자등록번호 :</td><td colspan="2"></td></tr>
        <tr><td>주소 :</td><td></td><td>전화번호 :</td><td colspan="2"></td></tr>
        <tr><td>대표자 :</td><td colspan="4"></td></tr>
      </tbody></table>
      <table class="standard-contract-table contract-detail-table"><tbody>
        <tr><th class="vertical-head" rowspan="9">계<br>약<br>내<br>용</th><th>공 사 명</th><td colspan="5">${h.e(p.projectName)}</td></tr>
        <tr><th>계 약 금 액</th><td colspan="5">${h.e(h.documentMoney(p.currentContractAmount))}</td></tr>
        <tr><th>총공사부기금액</th><td colspan="5">${h.e(h.documentMoney(p.currentContractAmount))}</td></tr>
        <tr><th>계약보증금</th><td colspan="5">${h.e(h.documentMoney(contractSecurityAmount))}</td></tr>
        <tr><th>현 장</th><td colspan="5">${h.e(school.name)}</td></tr>
        <tr><th>지체상금률</th><td colspan="5">${h.e(p.delayPenaltyRate)}</td></tr>
        <tr><th class="small-label">물가변동계약금액<br>조정방법</th><td colspan="5">${h.e(p.priceAdjustmentMethod)}</td></tr>
        <tr><th>착공연월일</th><td colspan="2">${h.e(h.formatKoreanDate(p.plannedStartDate))}</td><th>준공연월일</th><td colspan="2">${h.e(h.formatKoreanDate(p.completionDueDate))}</td></tr>
        <tr><th>기타사항</th><td colspan="5"></td></tr>
      </tbody></table>
      <div class="contract-defect-caption">하자담보책임(복합공종의 경우 공종별 구분 기재)</div>
      <table class="standard-contract-table contract-defect-table"><thead><tr><th>공 종</th><th>공종별 계약금액</th><th>하자보수보증금율(%) 및 금액</th><th>하자담보책임기간</th></tr></thead><tbody>${warrantyRows}${Array.from({length:Math.max(0,2-warrantyItems.length)},()=>'<tr><td>&nbsp;</td><td></td><td></td><td></td></tr>').join('')}</tbody></table>
      <p class="contract-declaration">계약담당자와 계약상대자는 상호 대등한 입장에서 붙임의 계약문서에 의하여 위 공사에 대한 도급계약을 체결하고 신의에 따라 성실히 계약상의 의무를 이행할 것을 확약하며, 연대보증인은 계약자와 연대하여 계약상의 의무를 이행할 것을 확약한다. 이 계약의 증거로서 계약서를 작성하여 당사자가 기명날인한 후 각각 1통씩 보관한다.</p>
      <div class="contract-attachments"><strong>붙임서류</strong><ol><li>공사입찰유의서 1부</li><li>공사계약일반조건 1부</li><li>공사계약특수조건 1부</li><li>설계서 1부</li><li>산출내역서 1부</li></ol></div>
      <p class="contract-date">${h.e(h.formatKoreanDate(p.contractDate))}</p>
      <div class="contract-signatures"><div><strong>계약자</strong><p><span>기관명</span><b>:</b>${h.e(school.name)}</p><p><span>주소</span><b>:</b>${h.e(school.address)}</p><p><span>${h.e(h.schoolLeaderLabel(school.name))}</span><b>:</b>${h.e(h.representativeWithSeal(school.principal))}</p></div><div><strong>계약상대자</strong><p><span>상호</span><b>:</b>${h.e(p.vendorName)}</p><p><span>주소</span><b>:</b>${h.e(p.vendorAddress)}</p><p><span>대표</span><b>:</b>${h.e(h.representativeWithSeal(p.representative))}</p></div></div>
    </article>`];
  }

  function renderAcceptanceTerms(ctx) {
    const { p, h } = common(ctx);
    const warrantyItems = Array.isArray(p.warrantyItems) ? p.warrantyItems : [];
    const years = String(warrantyItems[0]?.years || p.defectPeriodYears || '');
    const warrantySentence = warrantyItems.length > 1 ? '계약상대자는 공사 준공일로부터 각 세부공종별 하자담보책임기간 동안 해당 공종의 하자에 대하여 담보 책임을 진다.' : `계약상대자는 공사 준공일로부터 ${years}년간 그 공사의 공종별 하자에 대하여 담보 책임을 진다.`;
    return [`<article class="paper-a4 admin-document acceptance-terms document-print-page">
      <h1 class="doc-title wide-spacing acceptance-title">승 낙 사 항 (공 사 집 행)</h1>
      <ol class="acceptance-list">
        <li>계약사항에 의하여 ${h.e(h.formatKoreanDate(p.plannedStartDate))} 착공하고 ${h.e(h.formatKoreanDate(p.completionDueDate))}까지 준공하여야 한다.</li>
        <li>설계의 변경에 의하여 계약금액에 증감이 생긴 때에는 명세서상의 단가로 증감하고 그 단가에 의하기 어려운 때에는 설계변경 당시의 단가에 의한다.</li>
        <li>기한 내에 공사를 준공하지 못한 때에는 그 지연일수 1일에 대하여 계약금액의 ${h.e(p.delayPenaltyRate)}에 해당하는 지연배상금을 납부하여야 하며, 납부하여야 할 금액은 계약대가에서 상계할 수 있다.</li>
        <li>${h.e(warrantySentence)}</li>
        <li>기타 이 계약서에 명시되지 아니한 사항은 지방자치단체를 당사자로 하는 계약에 관한 법률 등의 규정을 준용한다.</li>
      </ol>
      <p class="acceptance-date">${h.e(h.formatKoreanDate(p.contractDate))}</p>
      <div class="acceptance-vendor"><p><span>계약명</span><b>:</b>${h.e(p.projectName)}</p><p><span>상호</span><b>:</b>${h.e(p.vendorName)}</p><p><span>사업자등록번호</span><b>:</b>${h.e(h.businessNumber(p.businessNumber))}</p><p><span>주소</span><b>:</b>${h.e(p.vendorAddress)}</p><p><span>대표자</span><b>:</b>${h.e(h.representativeWithSeal(p.representative))}</p></div>
    </article>`];
  }

  function renderUseSeal(ctx) {
    const { p, school, h } = common(ctx);
    return [`<article class="paper-a4 admin-document use-seal-form document-print-page">
      <h1 class="doc-title wide-spacing use-seal-title">사 용 인 감 계</h1>
      <div class="seal-boxes"><div><strong>증 명 인 감</strong><span></span></div><div><strong>사 용 인 감</strong><span></span></div></div>
      <p class="seal-statement">위 인감은 본인이 사용하는 인감으로서 아래 공사의 계약관련 등에 사용하겠으며, 위 인감사용으로 인한 법률상의 모든 책임은 본인(폐사)이 질 것을 확약하고 이에 사용인감계를 제출합니다.</p>
      <p class="seal-project">□ 공&nbsp;&nbsp; 사&nbsp;&nbsp; 명 : ${h.e(p.projectName)}</p>
      <p class="seal-date">${h.e(h.formatKoreanDate(p.contractDate))}</p>
      <div class="seal-vendor"><p><span>회 사 명 :</span>${h.e(p.vendorName)}</p><p><span>대 표 자 :</span>${h.e(h.representativeWithSeal(p.representative))}</p><p><span>주&nbsp;&nbsp;&nbsp;&nbsp;소 :</span>${h.e(p.vendorAddress)}</p></div>
      <p class="doc-recipient seal-recipient">${h.e(h.recipientFor(school.name))}</p>
    </article>`];
  }

  function pledgeRows(p, h, answers = {}) {
    const q = [
      '발주기관의 소속 고위공직자, 배우자, 고위공직자의 직계존속·비속 또는 생계를 같이하는 배우자의 직계존속·비속에 해당하는가?',
      '계약 업무를 법령상·사실상 담당하는 공직자, 배우자, 공직자의 직계존속·비속 또는 생계를 같이하는 배우자의 직계존속·비속에 해당하는가?',
      '발주기관(산하기관)의 감독기관 소속 고위공직자, 배우자, 고위공직자의 직계존속·비속 또는 생계를 같이하는 배우자의 직계존속·비속에 해당하는가?',
      '발주기관(자회사)의 모회사 소속 고위공직자, 배우자, 고위공직자의 직계존속·비속 또는 생계를 같이하는 배우자의 직계존속·비속에 해당하는가?',
      '상임위원회 위원인 국회의원, 배우자, 국회의원의 직계존속·비속 또는 생계를 같이하는 배우자의 직계존속·비속에 해당하는가?',
      '공공기관을 감사 또는 조사하는 지방의회의 의원, 배우자, 의원의 직계존속·비속 또는 생계를 같이하는 배우자의 직계존속·비속에 해당하는가?',
      '앞의 어느 하나에 해당하는 사람이 대표자인 법인 또는 단체에 해당하는가?',
      '앞의 어느 하나에 해당하는 사람과 특수한 관계의 사업자에 해당하는가?'
    ];
    const a = key => answers?.[key] || '';
    return `
      <tr><td>1</td><th>계약일반조건</th><td>상기 본인(법인)은 「지방자치단체 입찰 및 계약 집행기준」 제9장 계약 일반조건을 준수합니다.</td><td>${plainChoice(a('generalConditions'))}</td></tr>
      <tr><td>2</td><th>수의계약 각서</th><td>귀 기관과 수의계약을 체결함에 있어서 [붙임 1] 수의계약 배제사유 중 어느 사유에도 해당되지 않으며 차후에 이러한 사실이 발견된 경우 계약의 해제·해지 및 부정당업자 제재 처분을 받아도 하등의 이유를 제기하지 않겠습니다.<br><b>[붙임 1] 수의계약 배제사유 1부</b></td><td>${plainChoiceWithNone(a('privateContractExclusion'))}</td></tr>
      <tr class="pledge-conflict"><td rowspan="9">3</td><th rowspan="9">수의계약<br>체결 제한<br>여부 확인서</th><td>① ${h.e(q[0])}</td><td>${plainChoiceWithNone(a('conflict1'))}</td></tr>
      ${q.slice(1).map((text,i)=>`<tr class="pledge-conflict"><td>${'②③④⑤⑥⑦⑧'[i]} ${h.e(text)}</td><td>${plainChoiceWithNone(a(`conflict${i+2}`))}</td></tr>`).join('')}
      <tr class="pledge-conflict-confirm"><td colspan="2">「공직자의 이해충돌 방지법」 제12조에 따른 수의계약 체결 제한에 대하여 위와 같이 확인합니다. 만약 위 사항이 사실과 다른 경우에는 어떠한 처벌이나 불이익도 감수할 것을 서약합니다.</td></tr>
      <tr><td>4</td><th>계약보증금</th><td>계약서의 의무를 이행하지 못하여 계약보증금을 귀 기관에 귀속시켜야 할 사유가 발생하면 「지방자치단체를 당사자로 하는 계약에 관한 법률」 제15조제3항에 따라 즉시 해당하는 금액을 현금으로 납부하겠습니다.</td><td>${plainChoice(a('contractSecurity'))}<br>${h.e(sealChoice(p.contractSecurityType))}</td></tr>
      <tr><td>5</td><th>청렴계약</th><td>임직원과 대리인은 발주기관에서 시행하는 공사 등의 입찰·낙찰, 계약체결, 감독, 검사 또는 계약이행 과정에 참여하면서 금품·향응 등을 제공 또는 약속하거나 수수하지 않고, 불공정한 행위와 알선·청탁을 하지 않겠습니다.</td><td>${plainChoice(a('integrity'))}</td></tr>
      <tr><td>6</td><th>조세포탈 여부 확인 서약서</th><td>「지방자치단체를 당사자로 하는 계약에 관한 법률」 제31조의5에 따른 조세포탈 등을 한 자가 아님을 서약하며, 해당 사실이 발견된 때에는 계약 해제·해지 및 관련 제재 처분을 감수하겠습니다.</td><td>${plainChoice(a('taxEvasion'))}</td></tr>
      <tr><td>7</td><th>하자보수</th><td>「지방자치단체를 당사자로 하는 계약에 관한 법률 시행령」 제71조에 따라 하자보수보증금을 귀 학교(기관)에 귀속시켜야 할 사유가 발생하면 즉시 해당하는 금액을 현금으로 납부하겠습니다.</td><td>${plainChoice(a('defectRepair'))}<br>${h.e(sealChoice(p.defectSecurityType))}</td></tr>
      <tr><td>8</td><th>[공사] 전기·수도 사용료 납부 확인</th><td>우리 업체는 학교(기관)의 전기 및 수도를 사용한 경우 관련 기준에 의한 계산식으로 전기료 및 수도료를 학교회계(교육비특별회계)에 세입조치하고 이의를 제기하지 않겠습니다.</td><td>${plainChoiceWithNone(a('utility'))}</td></tr>
      <tr><td>9</td><th>안전 및 보건 확보 의무 준수</th><td>「산업안전보건법」 및 「중대재해 처벌 등에 관한 법률」 등 관련 법규에 따라 종사자의 안전·보건상 유해요인 또는 위험을 방지하기 위한 의무사항을 이행하겠습니다.<br>① 안전·보건 관계법령상 의무사항 이행 ② 유해·위험요인 신고 시 신속한 개선 ③ 작업 전 안전대책 수립 ④ 중대산업재해 발생 시 선보고 후 사고처리</td><td>${plainChoiceWithNone(a('safetyHealth'))}</td></tr>
      <tr><td>10</td><th>개인정보이용·수집 동의</th><td>「개인정보 보호법」 제15조, 제22조에 따라 개인정보를 수집 및 이용하는 것에 동의합니다.<br><span class="pledge-mini">항목: 대표자명, 주소, 생년월일, (휴대)전화번호, 계좌번호, 이메일 / 목적: 계약업무 진행 / 보유·이용기간: 계약체결일로부터 5년</span></td><td>${plainChoice(a('privacy'))}</td></tr>
      <tr><td>11</td><th>기타 사항</th><td>법령, 예규 등 각종 규정은 개정될 수 있으며, 최신 규정을 따름</td><td></td></tr>`;
  }

  function renderPrivateContractPledge(ctx) {
    const { p, school, h } = common(ctx);
    const blank = !!ctx.renderOptions?.blankPledge;
    const answers = blank ? {} : (p.privateContractPledge?.results || {});
    const page1 = `<article class="paper-a4 admin-document private-contract-pledge pledge-page-one document-print-page">
      <h1 class="pledge-title">수의계약 통합서약서</h1>
      <table class="pledge-summary"><tbody><tr><th>계약명</th><td colspan="3">${h.e(p.projectName)}</td></tr><tr><th>발주기관</th><td colspan="3">${h.e(school.name)}</td></tr><tr><th>업체명</th><td>${h.e(p.vendorName)}</td><th>대표자</th><td>${h.e(p.representative)}</td></tr><tr><th>사업자등록번호</th><td>${h.e(h.businessNumber(p.businessNumber))}</td><th>연락처</th><td>${h.e(p.vendorPhone)}</td></tr><tr><th>주소</th><td colspan="3">${h.e(p.vendorAddress)}</td></tr></tbody></table>
      <table class="pledge-table"><thead><tr><th>순</th><th>구분</th><th>이행 내용</th><th>세부내용</th></tr></thead><tbody>${pledgeRows(p,h,answers)}</tbody></table>
      <div class="pledge-sign"><p>${h.e(h.formatKoreanDate(p.contractDate))}</p><p><span>업 체 명 :</span>${h.e(p.vendorName)}<span>대 표 자 :</span>${h.e(h.representativeWithSeal(p.representative))}</p></div>
      <p class="pledge-recipient">${h.e(h.recipientFor(school.name))}</p>
    </article>`;
    const reasons = [
      '① 견적서 제출 마감일 현재 부도·파산·해산·영업정지 등이 확정된 경우. 다만 법원의 회생절차개시결정이 있는 경우에는 수의계약 체결 가능',
      '② 입찰참가자격 제한기간 중에 있는 자(「지방계약법」 제31조제5항에 해당되는 경우 예외)',
      '③ 견적서 제출 마감일을 기준으로 「지방계약법」 제31조 또는 다른 법령에 따라 부실이행, 담합행위, 입찰·계약 서류의 허위·위조 제출, 입찰·낙찰·계약이행 관련 뇌물 제공으로 부정당업자 제재 처분을 받고 그 종료일로부터 3개월이 지나지 아니한 자(법 제31조제5항에 해당되는 경우 예외)',
      '④ 공사 또는 기술용역의 경우 기술자 보유현황이 관련 법령에 따른 업종등록 기준에 미달하는 자\n※ 기술자 보유현황 심사는 낙찰자결정기준의 관련 기술인력 평가방법을 준용하며, 이 경우 입찰공고일은 안내공고일, 적격심사서류 제출마감일은 견적서 제출마감일로 봄',
      '⑤ 견적서 제출 마감일 기준 최근 3개월 이내에 해당 지방자치단체의 입찰·계약 및 그 이행과 관련하여 10일 이상 지연배상금 부과, 정당한 이행명령 거부, 불법하도급, 5회 이상 하자보수 또는 물의를 일으키는 등 신용이 떨어져 계약 체결이 곤란하다고 판단되는 자',
      '⑥ 견적서 제출 마감일 기준 최근 3개월 이내에 해당 지방자치단체와의 계약 및 그 이행과 관련하여 정당한 이유 없이 계약에 응하지 아니하거나 포기서를 제출한 사실이 있는 자\n※ 정당한 이유 없이 계약을 체결하지 아니한 경우 입찰참가자격 제한에는 해당하지 않더라도 수의계약 배제사유에는 해당할 수 있음',
      '⑦ 수의계약 체결일 현재 「지방계약법」 제33조에 해당하는 자\n1. 지방자치단체의 장 또는 지방의회의원의 배우자인 사업자(법인은 대표자)\n2. 지방자치단체의 장 또는 지방의회의원(배우자 포함)의 직계 존·비속인 사업자\n3. 지방자치단체의 장 또는 지방의회의원이 자본금 총액의 50% 이상을 소유한 자\n4. 지방자치단체의 장 또는 지방의회의원 가족(배우자, 직계 존·비속)의 합산금액이 자본금 총액의 50% 이상을 소유한 사업자\n5. 지방자치단체의 장 또는 지방의회의원 소유업체의 계열회사 등',
      '⑧ 발주기관이 제한한 자격요건 등을 충족하지 아니한 자',
      '⑨ 그 밖에 계약담당자가 계약이행능력이 없다고 판단되는 명백한 증거가 있는 자',
      '⑩ 「재난 및 안전관리 기본법」에 따라 특별재난지역으로 선포된 지역의 재난복구공사(용역)의 경우, 결격여부 심사일 현재 계약금액 5천만원 이상 해당 업종 관급공사 또는 계약금액 2천만원 이상 관급용역이 3건(일시 정지 중인 계약 제외) 이상인 자. 다만 동시에 여러 건의 수의계약 체결 예정자로 선정된 경우에는 기존 계약을 포함하여 3건까지 체결 가능하며, 2인 이상 견적서 제출 수의계약에 한함'
    ];
    const page2 = `<article class="paper-a4 admin-document private-contract-pledge pledge-page-two document-print-page"><div class="pledge-attachment-label">[붙임 1]</div><h1 class="pledge-reasons-title">수의계약 배제 사유</h1><div class="pledge-reasons pledge-reasons-full">${reasons.map(text=>`<div class="pledge-reason-item">${h.e(text).replace(/\n/g,'<br>')}</div>`).join('')}</div><p class="pledge-reasons-note">※ 관련 법령·예규 및 계약 집행기준이 개정될 수 있으므로 실제 계약 시 최신 기준을 확인합니다.</p></article>`;
    return [page1,page2];
  }

  function renderStartReport(ctx) {
    const { p, school, h } = common(ctx);
    return [`<article class="paper-a4 admin-document start-report document-print-page">
      <h1 class="doc-title wide-spacing">착 공 신 고 서</h1>
      ${h.documentFacts([
        ['1. 공 사 명 :',h.e(p.projectName)],
        ['2. 계 약 금 액 :',h.e(h.documentMoney(p.currentContractAmount))],
        ['3. 계약연월일 :',h.e(h.formatKoreanDate(p.contractDate))],
        ['4. 착공연월일 :',h.e(h.formatKoreanDate(p.startDate))],
        ['5. 준 공 기 한 :',h.e(h.formatKoreanDate(p.completionDueDate))]
      ])}
      <div class="doc-attachments"><span>붙&nbsp;&nbsp;&nbsp;&nbsp;임 :</span><ol><li>현장대리인계(재직증명서, 건설기술경력증수첩 사본)</li><li>공사예정공정표</li><li>공사도급내역서</li></ol></div>
      <p class="doc-statement start-statement">상기와 같이 공사를 착공하였기에 착공계를 제출합니다.</p>
      <p class="doc-date-center">${h.e(h.formatKoreanDate(p.startDate))}</p>
      ${h.documentVendorBlock(p)}
      <p class="doc-recipient">${h.e(h.recipientFor(school.name))}</p>
    </article>`];
  }

  function renderUtilityPaymentPledge(ctx) {
    const { p, school, h } = common(ctx);
    const hasSiteManagerOverride = Object.prototype.hasOwnProperty.call(ctx.renderOptions || {}, 'utilitySiteManager');
    const siteManager = hasSiteManagerOverride ? ctx.renderOptions.utilitySiteManager : (p.siteManager || '');
    return [`<article class="paper-a4 admin-document vendor-form-page utility-payment-pledge document-print-page">
      <h1 class="vendor-form-title spaced-title">각&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;서</h1>
      <p class="vendor-form-intro">당사는 발주처인 ${h.e(school.name || '')}와 계약체결하여 시공 예정인 아래 공사 건에 대하여 귀 기관의 수도·전력을 사용할 경우에는 대한건설협회 완성공사원가분석 경비율에 의한 계산식으로 수도광열비·전력비 요금을 귀 기관에 세입조치하고 이에 대하여 이의를 제기하지 않을 것을 확약합니다.</p>
      <table class="vendor-form-facts"><tbody>
        <tr><th>○ 공 사 명 :</th><td colspan="3">${h.e(p.projectName)}</td></tr>
        <tr><th>○ 계 약 자 :</th><td>${h.e(p.vendorName)}</td><th class="compact-head">대표 :</th><td>${h.e(p.representative)}</td></tr>
        <tr><th>○ 공사금액 :</th><td colspan="3">${h.e(h.documentMoney(p.currentContractAmount))}</td></tr>
        <tr><th>○ 공사기간 :</th><td>${h.e(h.formatKoreanDate(p.startDate))}</td><th class="range-mark">~</th><td>${h.e(h.formatKoreanDate(p.completionDueDate))}</td></tr>
      </tbody></table>
      <p class="vendor-form-date">${h.e(h.formatKoreanDate(p.contractDate))}</p>
      <div class="vendor-form-signatures utility-signatures">
        <div class="signature-heading">계약자</div>
        <div class="utility-signature-line utility-company-line"><span>업&nbsp;&nbsp;체&nbsp;&nbsp;명 :</span><strong>${h.e(p.vendorName)}</strong></div>
        <div class="utility-signature-line"><span>대&nbsp;&nbsp;표&nbsp;&nbsp;자 :</span><strong class="utility-signature-name">${h.e(p.representative || '')}</strong><em class="utility-signature-seal">(인)</em></div>
        <div class="utility-signature-line"><span>현장대리인 :</span><strong class="utility-signature-name">${h.e(siteManager)}</strong><em class="utility-signature-seal">(인)</em></div>
      </div>
      <p class="doc-recipient">${h.e(h.recipientFor(school.name))}</p>
    </article>`];
  }

  function renderCompletionReport(ctx) {
    const { p, school, h } = common(ctx);
    return [`<article class="paper-a4 admin-document completion-report document-print-page">
      <h1 class="doc-title wide-spacing">준 공 계</h1>
      ${h.documentFacts([
        ['1. 공 사 명 :',h.e(p.projectName)],
        ['2. 계약금액 :',h.e(h.documentMoney(p.currentContractAmount))],
        ['3. 계약일자 :',h.e(h.formatKoreanDate(p.contractDate))],
        ['4. 착공일자 :',h.e(h.formatKoreanDate(p.startDate))],
        ['5. 준공기한 :',h.e(h.formatKoreanDate(p.completionDueDate))],
        ['6. 준공일자 :',h.e(h.formatKoreanDate(p.actualCompletionDate))]
      ])}
      <p class="doc-statement completion-statement">상기공사를 준공하였기에 준공계를 제출합니다.</p>
      <p class="doc-date-center">${h.e(h.formatKoreanDate(p.actualCompletionDate))}</p>
      ${h.documentVendorBlock(p)}
      <p class="doc-recipient">${h.e(h.recipientFor(school.name))}</p>
    </article>`];
  }

  function renderInspectionRequest(ctx) {
    const { p, school, h } = common(ctx);
    return [`<article class="paper-a4 admin-document inspection-request document-print-page">
      <h1 class="doc-title wide-spacing">준 공 검 사 원</h1>
      ${h.documentFacts([
        ['1. 공 사 명 :',h.e(p.projectName)],
        ['2. 계약금액 :',h.e(h.documentMoney(p.currentContractAmount))],
        ['3. 계약일자 :',h.e(h.formatKoreanDate(p.contractDate))],
        ['4. 착공일자 :',h.e(h.formatKoreanDate(p.startDate))],
        ['5. 준공기한 :',h.e(h.formatKoreanDate(p.completionDueDate))],
        ['6. 준공일자 :',h.e(h.formatKoreanDate(p.actualCompletionDate))]
      ])}
      <div class="doc-pledge"><p>위 공사의 도급시행에 있어서 공사전반에 걸쳐 공사설계도서, 품질관리기준 및 기타</p><p>약정대로 어김없이 준공되었음을 확인하오며, 만약 공사시공, 감독 및 검사에 관하여</p><p>하자가 발견될 시는 하자담보기간 전후를 막론하고 실액변상 또는 재시공할 것을</p><p>서약하고 이에 준공검사원을 제출합니다.</p></div>
      <p class="doc-date-center">${h.e(h.formatKoreanDate(p.actualCompletionDate))}</p>
      ${h.documentVendorBlock(p,true)}
      <p class="doc-recipient">${h.e(h.recipientFor(school.name))}</p>
    </article>`];
  }

  function renderSupervision(ctx) {
    const { p, school, h } = common(ctx);
    const supervisor = ctx.value('supervisor');
    return [`<article class="paper-a4 admin-document official-record supervision-record document-print-page">
      <h1 class="doc-title wide-spacing">공 사 감 독 조 서</h1>
      <table class="official-table"><tbody>
        <tr><th>공 사 명</th><td colspan="3">${h.e(p.projectName)}</td></tr>
        <tr><th>도 급 자</th><td>${h.e(p.vendorName)}</td><th class="mini-head">대표</th><td>${h.e(p.representative)}</td></tr>
        <tr><th>계 약 금 액</th><td colspan="3">${h.e(h.documentMoney(p.currentContractAmount))}</td></tr>
        <tr><th>계 약 일</th><td colspan="3">${h.e(h.formatKoreanDate(p.contractDate))}</td></tr>
        <tr><th>착 공 일</th><td colspan="3">${h.e(h.formatKoreanDate(p.startDate))}</td></tr>
        <tr><th>준 공 기 한</th><td colspan="3">${h.e(h.formatKoreanDate(p.completionDueDate))}</td></tr>
        <tr><th>실제준공일</th><td colspan="3">${h.e(h.formatKoreanDate(p.actualCompletionDate))}</td></tr>
        <tr class="memo-row"><th>비 고</th><td colspan="3"></td></tr>
      </tbody></table>
      <div class="record-statement supervision-copy"><p>위 공사의 감독자로&nbsp;&nbsp;&nbsp; ${h.e(h.formatKoreanDate(p.startDate))}&nbsp;&nbsp;&nbsp; ~ &nbsp;&nbsp;&nbsp;${h.e(h.formatKoreanDate(p.actualCompletionDate))}&nbsp;&nbsp;&nbsp;까지</p><p>실지 현장 감독한 결과 공사 전반에 걸쳐 공사설계도서, 제시방서 및 품질</p><p>관리 기준 및 기타 약정대로 어김없이 준공되었음을 인정함.</p></div>
      <p class="record-date">${h.e(h.formatKoreanDate(p.actualCompletionDate))}</p>
      <div class="record-signature-row"><span>공사감독원</span><span>${h.e(school.name)}</span><span>${h.e(h.representativeWithSeal(supervisor))}</span></div>
    </article>`];
  }

  function renderInspectionRecord(ctx) {
    const { p, school, h } = common(ctx);
    const inspector = ctx.value('inspector');
    const witness = ctx.value('witness');
    return [`<article class="paper-a4 admin-document official-record completion-inspection-record document-print-page">
      <h1 class="doc-title wide-spacing">준 공 검 사 조 서</h1>
      <table class="official-table inspection-table"><tbody>
        <tr><th>공 사 명</th><td colspan="3">${h.e(p.projectName)}</td></tr>
        <tr><th>도 급 자</th><td>${h.e(p.vendorName)}</td><th class="mini-head">대표</th><td>${h.e(p.representative)}</td></tr>
        <tr><th>계약금액</th><td colspan="3">${h.e(h.documentMoney(p.currentContractAmount))}</td></tr>
        <tr><th>계 약 일</th><td>${h.e(h.formatKoreanDate(p.contractDate))}</td><th>준 공 기 한</th><td>${h.e(h.formatKoreanDate(p.completionDueDate))}</td></tr>
        <tr><th>착 공 일</th><td>${h.e(h.formatKoreanDate(p.startDate))}</td><th>준 공 일</th><td>${h.e(h.formatKoreanDate(p.actualCompletionDate))}</td></tr>
        <tr><th>준공검사일</th><td>${h.e(h.formatKoreanDate(p.completionInspectionDate))}</td><td colspan="2"></td></tr>
        <tr><th>참 고</th><td colspan="3">준공정산금액 :&nbsp;&nbsp; ${h.e(h.moneyNumberText(p.settlementAmount))}</td></tr>
        <tr class="attachment-row"><th>별 첨</th><td>준공정산서 1부</td><td colspan="2">※공사위치, 재료, 물량 표시된 배치도,<br>평면도 및 공사개요 첨부</td></tr>
      </tbody></table>
      <p class="inspection-finished">위와 같이 준공검사를 필하였음</p>
      <p class="record-date inspection-date">${h.e(h.formatKoreanDate(p.completionInspectionDate))}</p>
      <div class="inspection-signatures"><div><span>검사자</span><span>${h.e(school.name)}</span><span>${h.e(h.representativeWithSeal(inspector))}</span></div><div><span>입회자</span><span></span><span>${h.e(h.representativeWithSeal(witness))}</span></div><div><span>입회자</span><span></span><span>(인)</span></div></div>
    </article>`];
  }

  function renderDefectSecurityDeposit(ctx) {
    const { p, school, h } = common(ctx);
    const recordDate = p.defectStartDate || p.actualCompletionDate;
    return [`<article class="paper-a4 admin-document vendor-form-page defect-security-deposit document-print-page">
      <h1 class="vendor-form-title defect-deposit-title">하 자 보 수 보 증 금&nbsp;&nbsp;납 부 서</h1>
      <table class="vendor-form-table defect-deposit-table"><tbody>
        <tr><th>공 사 명</th><td colspan="3">${h.e(p.projectName)}</td></tr>
        <tr><th>계약금액</th><td colspan="3">${h.e(h.documentMoney(p.currentContractAmount))}</td></tr>
        <tr><th>계약일자</th><td>${h.e(h.formatKoreanDate(p.contractDate))}</td><th>준공일자</th><td>${h.e(h.formatKoreanDate(p.actualCompletionDate))}</td></tr>
        <tr><th>하자보증금율</th><td colspan="3">${h.e(h.percentText(p.defectSecurityRate))}%</td></tr>
        <tr><th>하자보증금액</th><td colspan="3">${h.e(h.documentMoney(ctx.value('defectSecurityAmount')))}</td></tr>
        <tr><th>하자담보책임기간</th><td>${h.e(h.formatKoreanDate(p.defectStartDate))}</td><th class="range-mark">~</th><td>${h.e(h.formatKoreanDate(p.defectEndDate))}</td></tr>
        <tr><th>보증금 납부방법</th><td colspan="3">${h.e(p.defectSecurityType || '')}</td></tr>
      </tbody></table>
      <p class="vendor-form-declaration">위의 금액을 하자보수보증금으로 납부합니다.</p>
      <p class="vendor-form-date">${h.e(h.formatKoreanDate(recordDate))}</p>
      <div class="vendor-form-signatures">
        <div><span>업&nbsp;&nbsp;&nbsp;체&nbsp;&nbsp;&nbsp;명 :</span><strong>${h.e(p.vendorName)}</strong></div>
        <div><span>사업자등록번호 :</span><strong>${h.e(h.businessNumber(p.businessNumber))}</strong></div>
        <div><span>주&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;소 :</span><strong>${h.e(p.vendorAddress)}</strong></div>
        <div><span>대&nbsp;&nbsp;&nbsp;표&nbsp;&nbsp;&nbsp;자 :</span><strong>${h.e(h.representativeWithSeal(p.representative))}</strong></div>
      </div>
      <p class="doc-recipient">${h.e(h.recipientFor(school.name))}</p>
    </article>`];
  }

  function renderCompletionSettlementAgreement(ctx) {
    const { p, school, h } = common(ctx);
    const contractAmount = Number(p.currentContractAmount || 0);
    const settlementAmount = Number(p.settlementAmount || 0);
    const reductionAmount = Math.max(0, contractAmount - settlementAmount);
    const recordDate = p.defectStartDate || p.completionInspectionDate || p.actualCompletionDate;
    return [`<article class="paper-a4 admin-document vendor-form-page completion-settlement-agreement document-print-page">
      <h1 class="vendor-form-title settlement-title">준 공 정 산 동 의 서</h1>
      <table class="vendor-form-table settlement-project-table"><tbody>
        <tr><th>공 사 명</th><td>${h.e(p.projectName)}</td></tr>
        <tr><th>공사현장</th><td>${h.e(school.name || '')}</td></tr>
      </tbody></table>
      <h2 class="settlement-section-title">◈ 동의 사항</h2>
      <h3 class="settlement-subtitle">1. 계약금액</h3>
      <table class="vendor-form-table settlement-amount-table"><thead><tr><th>구분</th><th>당초</th><th>준공</th><th>감액</th><th>비고</th></tr></thead><tbody><tr><th>계약금액</th><td>${h.e(h.moneyNumberText(contractAmount))}</td><td>${h.e(h.moneyNumberText(settlementAmount))}</td><td>${h.e(h.moneyNumberText(reductionAmount))}</td><td></td></tr></tbody></table>
      <h3 class="settlement-subtitle">2. 준공감액 사유</h3>
      <div class="settlement-reason">${h.e(p.settlementReductionReason || '')}</div>
      <p class="vendor-form-declaration settlement-declaration">당사는 귀 기관과 계약 체결하여 공사 완료한 본 공사에 대하여 상기 준공감액 내역 및 붙임 내역서와 같이 감액함에 동의하고 이의 없이 계약자로서의 의무를 성실히 이행할 것을 확약합니다.</p>
      <p class="vendor-form-date">${h.e(h.formatKoreanDate(recordDate))}</p>
      <div class="vendor-form-signatures">
        <div><span>업&nbsp;&nbsp;&nbsp;체&nbsp;&nbsp;&nbsp;명 :</span><strong>${h.e(p.vendorName)}</strong></div>
        <div><span>주&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;소 :</span><strong>${h.e(p.vendorAddress)}</strong></div>
        <div><span>대&nbsp;&nbsp;&nbsp;표&nbsp;&nbsp;&nbsp;자 :</span><strong>${h.e(h.representativeWithSeal(p.representative))}</strong></div>
      </div>
      <p class="doc-recipient">${h.e(h.recipientFor(school.name))}</p>
    </article>`];
  }

  function renderPayment(ctx) {
    const { p, school, payout, h } = common(ctx);
    const claimAmount = h.claimAmountFor(p);
    const accountHolder = payout.accountHolder || p.vendorName || '';
    const rows = [
      ['1. 계 약 건 명 :', p.projectName, ''],
      ['2. 계 약 금 액 :', h.moneyNumberText(p.currentContractAmount), h.moneyWordsText(p.currentContractAmount)],
      ['3. 준 공 금 액 :', h.moneyNumberText(p.settlementAmount), h.moneyWordsText(p.settlementAmount)],
      ['4. 기 지 급 액 :', h.moneyNumberText(p.priorPaymentAmount), h.moneyWordsText(p.priorPaymentAmount)],
      ['5. 청 구 금 액 :', h.moneyNumberText(claimAmount), h.moneyWordsText(claimAmount)],
      ['6. 공 제 금 액 :', h.moneyNumberText(p.deductionAmount), h.moneyWordsText(p.deductionAmount)]
    ];
    return [`<article class="paper-a4 admin-document payment-request document-print-page">
      <h1 class="doc-title wide-spacing">대 금 청 구 서</h1>
      <table class="payment-lines-table"><tbody>${rows.map(([label,value,words]) => `<tr><th>${h.e(label)}</th><td class="payment-value">${h.e(value)}</td><td class="payment-words">${h.e(words)}</td></tr>`).join('')}</tbody></table>
      <p class="payment-statement">위와 같이 청구하오니 아래 계좌에 입금하여 주시기 바랍니다.</p>
      <p class="payment-account-title">□ 지정계좌현황</p>
      <table class="payment-account-table"><tbody><tr><th>은 행 명</th><td>${h.e(payout.bankName || '')}</td></tr><tr><th>계 좌 번 호</th><td>${h.e(payout.accountNumber || '')}</td></tr><tr><th>예 금 주 명</th><td>${h.e(accountHolder)}</td></tr></tbody></table>
      <p class="record-date payment-date">${h.e(h.formatKoreanDate(p.completionInspectionDate))}</p>
      <div class="payment-vendor"><div><span>회 사 명 :</span><strong>${h.e(p.vendorName)}</strong></div><div><span>주&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;소 :</span><strong>${h.e(p.vendorAddress)}</strong></div><div><span>대 표 자 :</span><strong>${h.e(h.representativeWithSeal(p.representative))}</strong></div></div>
      <p class="payment-note">※ 계약도장과 청구서의 도장이 다를 경우 인감증명 및 사용인감신고서 제출</p>
      <p class="doc-recipient">${h.e(h.recipientFor(school.name))}</p>
    </article>`];
  }


  function renderConstructionLedger(ctx) {
    const { p, h } = common(ctx);
    const changes = Array.isArray(p.contractChanges) ? p.contractChanges : [];
    const changeRows = changes.slice(0,4);
    const budgetRows = [
      ['정책사업', p.budgetPolicyProject],
      ['단위사업', p.budgetUnitProject],
      ['세부사업', p.budgetDetailProject],
      ['세부항목', p.budgetDetailItem],
      ['원가통계목', p.costStatisticsItem],
      ['산출내역', p.budgetCalculationDetails]
    ];
    const ledgerSplitRows = budgetRows.map(([label,value],i) => {
      const c = i > 0 && i <= 4 ? changeRows[i-1] : null;
      const changeCells = i === 0
        ? '<th>년 월 일</th><th colspan="2">증감액</th><th colspan="2">변경금액</th>'
        : `<td>${c?h.e(h.formatKoreanDate(c.changeDate)):''}</td><td colspan="2">${c?h.e(h.moneyNumberText(Number(c.afterAmount||0)-Number(c.beforeAmount||0))):''}</td><td colspan="2">${c?h.e(h.moneyNumberText(c.afterAmount)):''}</td>`;
      return `<tr class="ledger-split-row">${i===0?'<th rowspan="6" class="ledger-section-label">설계변경<br>증감</th>':''}${changeCells}${i===0?'<th rowspan="6" class="ledger-section-label">지출과목</th>':''}<th colspan="2" class="ledger-budget-label">${h.e(label)}</th><td colspan="3">${h.e(value || '')}</td></tr>`;
    }).join('');
    return [`<article class="paper-a4 admin-document construction-ledger document-print-page">
      <div class="ledger-title-row"><h1>공  사  대  장</h1><div>계약번호 : <strong>${h.e(p.contractNumber || '')}</strong></div></div>
      <table class="ledger-table">
        <colgroup>${Array.from({length:12},()=>'<col>').join('')}</colgroup>
        <tbody>
        <tr><th>공사명</th><td colspan="4">${h.e(p.projectName)}</td><th>도급자</th><td colspan="3">${h.e(p.vendorName)}</td><th>공사감독<br>성명</th><td colspan="2">${h.e(ctx.value('supervisor') || '')}</td></tr>
        <tr><th rowspan="2">계약금액</th><td colspan="4" rowspan="2">${h.e(h.moneyNumberText(p.currentContractAmount))}</td><th>계약방법</th><td colspan="3">${h.e(p.contractMethod || '')}</td><th>전화번호</th><td colspan="2">${h.e(p.vendorPhone || '')}</td></tr>
        <tr><th>계약일</th><td colspan="3">${h.e(h.formatKoreanDate(p.contractDate))}</td><th>대표자</th><td colspan="2">${h.e(p.representative || '')}</td></tr>

        ${ledgerSplitRows}

        <tr><th>계약내용</th><th colspan="2">계약보증금</th><td colspan="3">${h.e(h.moneyNumberText(ctx.value('contractSecurityAmount')))}</td><td colspan="2">${h.e(p.contractSecurityType || '')}</td><th>재원</th><td colspan="3">${h.e(p.fundingSource || '')}</td></tr>

        <tr><th rowspan="4">공정</th><th>구분</th><th colspan="2">계약상 년월일</th><th colspan="2">실제 년월일</th><th rowspan="4">지급내역</th><th>구분</th><th colspan="2">년월일</th><th colspan="2">지급금액</th></tr>
        <tr><th>착공</th><td colspan="2">${h.e(h.formatKoreanDate(p.plannedStartDate || p.startDate))}</td><td colspan="2">${h.e(h.formatKoreanDate(p.startDate))}</td><th>선급금</th><td colspan="2"></td><td colspan="2">${h.e(h.moneyNumberText(p.priorPaymentAmount))}</td></tr>
        <tr><th>준공</th><td colspan="2">${h.e(h.formatKoreanDate(p.completionDueDate))}</td><td colspan="2">${h.e(h.formatKoreanDate(p.actualCompletionDate))}</td><th>준공금</th><td colspan="2">${h.e(h.formatKoreanDate(p.paymentDate))}</td><td colspan="2">${h.e(h.moneyNumberText(p.paymentAmount || p.settlementAmount))}</td></tr>
        <tr><th>준공정산금</th><td colspan="4">${h.e(h.moneyNumberText(p.settlementAmount))}</td><th>잔액</th><td colspan="4">${h.e(h.moneyNumberText(Math.max(0, Number(p.currentContractAmount||0)-Number(p.paymentAmount||0))))}</td></tr>

        <tr><th rowspan="3">하자담보</th><th>시작일</th><td colspan="3">${h.e(h.formatKoreanDate(p.defectStartDate))}</td><th rowspan="3">검사내역</th><th>준공검사일</th><td colspan="2">${h.e(h.formatKoreanDate(p.completionInspectionDate))}</td><th>검사자</th><td colspan="2">${h.e(ctx.value('inspector') || '')}</td></tr>
        <tr><th>종료일</th><td colspan="3">${h.e(h.formatKoreanDate(p.defectEndDate))}</td><th>공정</th><td colspan="2">1</td><th>입회자</th><td colspan="2">${h.e(ctx.value('witness') || '')}</td></tr>
        <tr><th>보증금</th><td colspan="2">${h.e(h.moneyNumberText(ctx.value('defectSecurityAmount')))}</td><td>${h.e(p.defectSecurityType || '')}</td><th>설계자</th><td colspan="5">${h.e(p.designer || '')}</td></tr>
      </tbody></table>
      <div class="ledger-note">※ 공사서류 작성지원에 저장된 공사정보를 기준으로 자동 작성</div>
    </article>`];
  }

  function renderWarrantyInspection(ctx) {
    const { p, school, h } = common(ctx);
    const signatoryMode = ctx.renderOptions?.warrantySignatoryMode || 'default';
    const useEntered = signatoryMode === 'filled';
    const useDefault = signatoryMode === 'default';
    const cleanPersonName = (value, fixedTitle = '') => {
      let text = String(value || '').trim();
      if (!text) return '';
      if (school.name && text.startsWith(String(school.name))) text = text.slice(String(school.name).length).trim();
      text = text.replace(/^직위\s*/,'').replace(/^성명\s*/,'');
      if (fixedTitle) text = text.replace(new RegExp(`^${fixedTitle}\\s*`), '');
      return text.trim();
    };
    // 원클릭 엑셀 양식 기준: 검사자는 학교의 행정실장으로 표시하고,
    // 출력물에는 '직위'·'성명'이라는 별도 표제어를 넣지 않는다.
    const inspectorPosition = (useEntered || useDefault) ? '행정실장' : '';
    const roleText = value => String(value || '').trim();
    const isAdminManager = value => /^행정실장(?:\s|$)/.test(roleText(value));
    const defaultAdminManagerSource = [p.witness, school.witness, p.inspector, school.inspector].find(isAdminManager) || '';
    const inspectorName = useEntered
      ? cleanPersonName(ctx.value('warrantyInspectorName'), '행정실장')
      : (useDefault ? cleanPersonName(defaultAdminManagerSource, '행정실장') : '');
    const defaultWitnessSource = [p.witness, school.witness].find(value => {
      const text = roleText(value);
      return text && !isAdminManager(text) && cleanPersonName(text) !== inspectorName;
    }) || '';
    const witnessName = useEntered
      ? cleanPersonName(ctx.value('warrantyWitnessName'))
      : (useDefault ? cleanPersonName(defaultWitnessSource) : '');
    const date = ctx.value('warrantyInspectionDate');
    const issue = ctx.value('warrantyIssueDetails');
    const actions = ctx.value('warrantyActions');
    const notes = ctx.value('warrantyNotes');
    return [`<article class="paper-a4 admin-document warranty-inspection-report document-print-page">
      <h1 class="doc-title wide-spacing">하 자 검 사 조 서</h1>
      <table class="official-table warranty-inspection-table">
        <colgroup>
          <col class="wi-col-b"><col class="wi-col-c"><col class="wi-col-d"><col class="wi-col-e"><col class="wi-col-f"><col class="wi-col-g">
        </colgroup>
        <tbody>
          <tr><th>공 사 명</th><td colspan="5">${h.e(p.projectName)}</td></tr>
          <tr><th rowspan="2">도 급 자</th><th>회 사 명</th><td colspan="3">${h.e(p.vendorName)}</td><th>공 종</th></tr>
          <tr><th>대 표</th><td colspan="3">${h.e(p.representative)}</td><td class="wi-work-type">${h.e(p.workType || '')}</td></tr>
          <tr><th>도 급 금 액</th><td colspan="3">${h.e(h.moneyNumberText(p.currentContractAmount))}</td><th>준 공 일</th><td>${h.e(h.formatKoreanDate(p.actualCompletionDate))}</td></tr>
          <tr><th>공 사 기 간</th><td class="wi-period-start">${h.e(h.formatKoreanDate(p.startDate))}</td><td class="wi-tilde">~</td><td class="wi-period-end">${h.e(h.formatKoreanDate(p.actualCompletionDate))}</td><th>하자만료일</th><td>${h.e(h.formatKoreanDate(p.defectEndDate))}</td></tr>
          <tr class="warranty-large-row"><th>하자발생내용</th><td colspan="5">${h.e(issue || '')}</td></tr>
          <tr class="warranty-large-row"><th>처 리 사 항</th><td colspan="5">${h.e(actions || '')}</td></tr>
          <tr class="warranty-large-row"><th>기타참고사항</th><td colspan="5">${h.e(notes || '')}</td></tr>
        </tbody>
      </table>
      <p class="warranty-statement">위와 같이 하자(만료)검사를 필하였음.</p>
      <p class="record-date">${h.e(h.formatKoreanDate(date))}</p>
      <div class="warranty-signatures">
        <div><span class="warranty-signature-label">검 사 자 :</span><span class="warranty-signature-content"><span class="warranty-signature-name">${signatoryMode==='blank'?'':`${h.e(inspectorPosition)}${inspectorName?` ${h.e(inspectorName)}`:''}`}</span><span class="warranty-seal">(인)</span></span></div>
        <div><span class="warranty-signature-label">입 회 자 :</span><span class="warranty-signature-content"><span class="warranty-signature-name">${signatoryMode==='blank'?'':`${witnessName?h.e(witnessName):''}`}</span><span class="warranty-seal">(인)</span></span></div>
      </div>
    </article>`];
  }

  function renderWarrantyLedger(ctx) {
    const { p, h } = common(ctx);
    const inspections = Array.isArray(p.warrantyInspections) ? p.warrantyInspections : [];
    const warrantyItems = Array.isArray(p.warrantyItems) ? p.warrantyItems : [];
    const chunks = inspections.length ? Array.from({length:Math.ceil(inspections.length/5)},(_,i)=>inspections.slice(i*5,i*5+5)) : [[]];
    return chunks.map((items,pageIndex)=>`<article class="paper-a4 admin-document warranty-ledger document-print-page">
      <h1>하 자 검 사 대 장</h1>
      <h2>1. 공사내역</h2>
      <table class="warranty-ledger-info"><tbody>
        <tr><th>공사명</th><td colspan="3">${h.e(p.projectName)}</td><th>계약금액</th><td colspan="2">${h.e(h.moneyNumberText(p.currentContractAmount))}</td><th>하자보증금</th><td colspan="2">${h.e(h.moneyNumberText(ctx.value('defectSecurityAmount')))}</td></tr>
        <tr><th>설계자</th><td colspan="3">${h.e(p.designer || '')}</td><th>계약년월일</th><td colspan="2">${h.e(h.formatKoreanDate(p.contractDate))}</td><th>준공기한</th><td colspan="2">${h.e(h.formatKoreanDate(p.completionDueDate))}</td></tr>
        <tr><th>도급자</th><td colspan="3">${h.e(p.vendorName)}</td><th>착공년월일</th><td colspan="2">${h.e(h.formatKoreanDate(p.startDate))}</td><th>준공년월일</th><td colspan="2">${h.e(h.formatKoreanDate(p.actualCompletionDate))}</td></tr>
        <tr><th>감독관</th><td colspan="3">${h.e(ctx.value('supervisor') || '')}</td><th>준공검사일</th><td colspan="2">${h.e(h.formatKoreanDate(p.completionInspectionDate))}</td><th>하자보증기간</th><td colspan="2">${h.e(h.formatKoreanDate(p.defectStartDate))} ~ ${h.e(h.formatKoreanDate(p.defectEndDate))}</td></tr>
      </tbody></table>
      ${warrantyItems.length?`<h2>2. 하자담보 공종</h2><table class="warranty-ledger-items"><thead><tr><th>세부공종</th><th>기간</th><th>시작일</th><th>종료일</th></tr></thead><tbody>${warrantyItems.map(x=>`<tr><td>${h.e(x.subcategory||x.category||'')}</td><td>${x.years?h.e(String(x.years))+'년':''}</td><td>${h.e(h.formatKoreanDate(x.startDate))}</td><td>${h.e(h.formatKoreanDate(x.endDate))}</td></tr>`).join('')}</tbody></table>`:''}
      <h2>${warrantyItems.length?'3':'2'}. 정기 하자검사 내역${chunks.length>1?` (${pageIndex+1}/${chunks.length})`:''}</h2>
      <table class="warranty-ledger-list"><thead><tr><th>하자검사<br>년 월 일</th><th>하자검사 내용 및 조치사항</th><th>하자검사자</th><th>결재</th></tr></thead><tbody>
        ${(items.length?items:[null]).map(x=>`<tr><td>${x?h.e(h.formatKoreanDate(x.date)):''}</td><td>${x?h.e([x.result,x.issueDetails,x.actions].filter(Boolean).join(' / ')):''}</td><td></td><td></td></tr>`).join('')}
        ${Array.from({length:Math.max(0,5-items.length)},()=>`<tr><td></td><td></td><td></td><td></td></tr>`).join('')}
      </tbody></table>
    </article>`);
  }

  function checklistMark(status, target) {
    return status === target ? '✓' : '';
  }

  function renderSafetyChecklist(ctx, type) {
    const { p, school, h } = common(ctx);
    const def = ReferenceData?.safetyChecklists?.[type];
    if (!def) return [];
    const blankSafety = !!ctx.renderOptions?.blankSafety;
    const saved = blankSafety ? {} : (p.safetyChecklists?.[type] || {});
    const results = blankSafety ? {} : (saved.results || {});
    const isAgency = def.owner === 'agency';
    const inspector = blankSafety ? '' : (saved.inspector || (isAgency ? (p.supervisor || school.supervisor || '') : ''));
    const date = blankSafety ? '' : (saved.date || p.startDate || p.contractDate || '');
    const signature = typeof ctx.signature === 'function' ? ctx.signature(type) : '';
    const signatureLabel = isAgency ? '인' : '서명';
    const signatureMarkup = signature
      ? `<span class="safety-signature-print"><img src="${h.e(signature)}" alt="점검자 서명"><em>(${signatureLabel})</em></span>`
      : `<em>(${signatureLabel})</em>`;
    const header = isAgency
      ? `<div class="safety-meta"><div><strong>■ 학교(기관)명:</strong><span>${h.e(school.name || '')}</span></div><div><strong>■ 점 검 자:</strong><span>${h.e(inspector)}</span>${signatureMarkup}</div><div><strong>■ 점검 일자:</strong><span>${h.e(h.formatKoreanDate(date))}</span></div></div>`
      : `<div class="safety-meta"><div><strong>■ 공사 업체명:</strong><span>${h.e(p.vendorName || '')}</span></div><div><strong>■ 점 검 자:</strong><span>${h.e(inspector)}</span>${signatureMarkup}</div><div><strong>■ 점검 일자:</strong><span>${h.e(h.formatKoreanDate(date))}</span></div></div>`;
    const rows = def.items.map((item,i)=>{ const status=results[String(i+1)] || ''; return `<tr><td>${i+1}</td><td>${h.e(item)}</td><td>${checklistMark(status,'yes')}</td><td>${checklistMark(status,'no')}</td><td>${checklistMark(status,'na')}</td></tr>`; }).join('');
    const page1 = `<article class="paper-a4 admin-document safety-checklist ${h.e(type)} document-print-page"><h1>${h.e(type==='safetyGeneral'?'공사(용역) 안전·보건 체크리스트(공통)':def.label)}</h1><p class="safety-subtitle">[ ${h.e(def.subtitle)} ]</p>${header}<table class="safety-check-table"><thead><tr><th>번호</th><th>점 검 내 용</th><th>예</th><th>아니요</th><th>해당없음</th></tr></thead><tbody>${rows}</tbody></table><p class="safety-footer">※ ${h.e(def.footer || '')}</p>${saved.notes?`<div class="safety-written-note"><strong>비고</strong><span>${h.e(saved.notes)}</span></div>`:''}${isAgency?`<div class="safety-vendor-confirm"><h2>공사(용역)업체 확인·서약서</h2><p>학교(기관)로부터 위 점검항목에 대하여 안내(주지)받았으며, 공사(용역) 전 과정에 걸쳐 참여하는 근로자에게 안전보건교육을 실시하고 필요한 개인보호구를 지급 및 착용하며 작업 시 안전수칙을 준수할 것을 서약합니다.</p><div><span>소속(회사) :</span><span>${h.e(p.vendorName||'')}</span></div><div><span>공사업체 책임자 :</span><span>${h.e(saved.contractorResponsible||'')} (서명)</span></div></div>`:''}</article>`;
    if (type !== 'safetyFall' || !(def.notes||[]).length) return [page1];
    const page2 = `<article class="paper-a4 admin-document safety-reference-page document-print-page"><h1>안전보건 점검사항 및 조치사항(예시)</h1><ol>${def.notes.map(x=>`<li>${h.e(x)}</li>`).join('')}</ol><div class="safety-risk-note"><strong>위험공종 안전관리 담당자 안내</strong><p>2m 이상 고소작업, 1.5m 이상 굴착·가설공사, 철골 구조물 공사, 2m 이상 외부도장 공사, 승강기 설치공사는 위험작업 전 공사감독자에게 작업계획 제출·검토·확인이 필요합니다.</p></div></article>`;
    return [page1,page2];
  }

  const RENDERERS = {
    standardContract: renderStandardContract,
    acceptanceTerms: renderAcceptanceTerms,
    useSealForm: renderUseSeal,
    privateContractPledge: renderPrivateContractPledge,
    startReport: renderStartReport,
    utilityPaymentPledge: renderUtilityPaymentPledge,
    completionReport: renderCompletionReport,
    completionInspectionRequest: renderInspectionRequest,
    supervisionReport: renderSupervision,
    completionInspectionRecord: renderInspectionRecord,
    defectSecurityDeposit: renderDefectSecurityDeposit,
    completionSettlementAgreement: renderCompletionSettlementAgreement,
    paymentRequest: renderPayment,
    constructionLedger: renderConstructionLedger,
    safetyGeneral: ctx => renderSafetyChecklist(ctx,'safetyGeneral'),
    safetyFall: ctx => renderSafetyChecklist(ctx,'safetyFall'),
    safetyElectrical: ctx => renderSafetyChecklist(ctx,'safetyElectrical'),
    safetyConfined: ctx => renderSafetyChecklist(ctx,'safetyConfined'),
    safetyIndustrial: ctx => renderSafetyChecklist(ctx,'safetyIndustrial'),
    warrantyInspectionReport: renderWarrantyInspection,
    warrantyLedger: renderWarrantyLedger
  };

  function renderPages(type, ctx) {
    const fn = RENDERERS[type];
    return fn ? fn(ctx) : [];
  }

  globalThis.ConstructionDocuments = {
    definitions: DEFINITIONS,
    fieldLabels: FIELD_LABELS,
    printOrder: PRINT_ORDER,
    sets: SETS,
    renderPages
  };
})();
