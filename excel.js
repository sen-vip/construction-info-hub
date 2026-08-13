(() => {
  'use strict';

  const JSZipLib = globalThis.JSZip;
  if (!JSZipLib) throw new Error('JSZip을 불러오지 못했습니다.');

  const AUDIT_HEADERS = [
    '회계연도','번호','공사명','계약상대자','계약방법','계약금액','예정가격','낙찰율 ',
    'G2B/S2B','계약일','착공일','현장대리인\n자격증종류','준공기한','준공일','준공검사원발행일',
    '준공검사원 문서등록일','준공검사조서작성일','선금지급여부','세금계산서 발행일','지출일',
    '지출금액','하자보증서/각서','등록면허종류','하자기간(~','~까지)','전화번호','대표','사업자번호',
    '재원구분','공사대장출력','비고'
  ];

  function escapeXml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function decodeXml(value) {
    if (value == null) return '';
    return String(value)
      .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
      .replace(/&#([0-9]+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&');
  }

  function colToNumber(col) {
    let n = 0;
    for (const ch of col) n = n * 26 + ch.charCodeAt(0) - 64;
    return n;
  }

  function numberToCol(n) {
    let s = '';
    while (n > 0) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  }

  function splitAddress(address) {
    const m = /^([A-Z]+)(\d+)$/.exec(address || '');
    return m ? { col: m[1], row: Number(m[2]) } : null;
  }

  function parseSharedStrings(xml = '') {
    const out = [];
    const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    let m;
    while ((m = siRe.exec(xml))) {
      const chunks = [];
      const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      let tm;
      while ((tm = tRe.exec(m[1]))) chunks.push(decodeXml(tm[1]));
      out.push(chunks.join(''));
    }
    return out;
  }

  function parseSheetCells(xml = '', sharedStrings = []) {
    const cells = new Map();
    const re = /<c\b([^>]*)\/>|<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    let m;
    while ((m = re.exec(xml))) {
      const attrs = m[1] || m[2] || '';
      const address = /\br="([A-Z]+\d+)"/.exec(attrs)?.[1];
      if (!address) continue;
      const body = m[3] || '';
      const type = /\bt="([^"]+)"/.exec(attrs)?.[1] || '';
      let value = '';
      if (type === 'inlineStr') {
        const parts = [];
        const tr = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
        let tm;
        while ((tm = tr.exec(body))) parts.push(decodeXml(tm[1]));
        value = parts.join('');
      } else {
        const vm = /<v>([\s\S]*?)<\/v>/.exec(body);
        const raw = vm ? decodeXml(vm[1]) : '';
        if (type === 's' && raw !== '') value = sharedStrings[Number(raw)] ?? '';
        else value = raw;
      }
      cells.set(address, value);
    }
    return cells;
  }

  function rowMap(cells, row) {
    const out = {};
    for (const [address, value] of cells.entries()) {
      const p = splitAddress(address);
      if (p && p.row === row) out[p.col] = value;
    }
    return out;
  }

  function maxRow(cells) {
    let max = 0;
    for (const key of cells.keys()) {
      const p = splitAddress(key);
      if (p) max = Math.max(max, p.row);
    }
    return max;
  }

  async function firstSheetPath(zip) {
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const id = /<sheet\b[^>]*r:id="([^"]+)"/.exec(wbXml)?.[1];
    if (!id) return 'xl/worksheets/sheet1.xml';
    const relXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const relRe = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/?\s*>/g;
    let m;
    while ((m = relRe.exec(relXml))) {
      if (m[1] === id) {
        const target = m[2];
        if (target.startsWith('/')) return target.slice(1);
        return target.startsWith('xl/') ? target : `xl/${target.replace(/^\.\//, '')}`;
      }
    }
    return 'xl/worksheets/sheet1.xml';
  }

  async function readWorkbook(fileOrBuffer) {
    const buffer = fileOrBuffer instanceof ArrayBuffer ? fileOrBuffer : await fileOrBuffer.arrayBuffer();
    const zip = await JSZipLib.loadAsync(buffer);
    const path = await firstSheetPath(zip);
    const sheetXml = await zip.file(path).async('string');
    let shared = [];
    if (zip.file('xl/sharedStrings.xml')) {
      shared = parseSharedStrings(await zip.file('xl/sharedStrings.xml').async('string'));
    }
    const cells = parseSheetCells(sheetXml, shared);
    return { zip, path, sheetXml, cells };
  }

  function toNumber(value) {
    if (value === '' || value == null) return '';
    const n = Number(String(value).replace(/,/g, '').trim());
    return Number.isFinite(n) ? n : '';
  }

  function cleanText(value) {
    return String(value ?? '').trim();
  }

  function normalizeBusinessNumber(value) {
    const digits = cleanText(value).replace(/\D/g, '');
    if (digits.length === 10) return `${digits.slice(0,3)}-${digits.slice(3,5)}-${digits.slice(5)}`;
    return cleanText(value);
  }

  function excelSerialToISO(value) {
    if (value === '' || value == null) return '';
    const s = cleanText(value);
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    if (/^\d{4}[./]\d{1,2}[./]\d{1,2}/.test(s)) {
      const [y,m,d] = s.match(/\d+/g).slice(0,3).map(Number);
      return `${String(y).padStart(4,'0')}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    }
    const n = Number(s);
    if (!Number.isFinite(n) || n < 1) return '';
    const epoch = Date.UTC(1899, 11, 30);
    const date = new Date(epoch + Math.round(n) * 86400000);
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth()+1).padStart(2,'0')}-${String(date.getUTCDate()).padStart(2,'0')}`;
  }

  function isoToExcelSerial(value) {
    if (!value) return '';
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
    if (!m) return '';
    const epoch = Date.UTC(1899, 11, 30);
    const time = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return Math.round((time - epoch) / 86400000);
  }

  function inferFiscalYear(dateIso, fallback = '') {
    return dateIso ? dateIso.slice(0, 4) : cleanText(fallback);
  }

  function parseAudit(cells) {
    const projects = [];
    const last = maxRow(cells);
    for (let r = 5; r <= last; r++) {
      const row = rowMap(cells, r);
      const projectName = cleanText(row.C);
      if (!projectName) continue;
      const contractDate = excelSerialToISO(row.J);
      projects.push({
        fiscalYear: cleanText(row.A) || inferFiscalYear(contractDate),
        sequence: cleanText(row.B),
        projectName,
        vendorName: cleanText(row.D),
        contractMethod: cleanText(row.E),
        currentContractAmount: toNumber(row.F),
        estimatedPrice: toNumber(row.G),
        bidRate: toNumber(row.H),
        procurementMethod: cleanText(row.I),
        contractDate,
        startDate: excelSerialToISO(row.K),
        siteManagerLicense: cleanText(row.L),
        completionDueDate: excelSerialToISO(row.M),
        actualCompletionDate: excelSerialToISO(row.N),
        completionRequestIssueDate: excelSerialToISO(row.O),
        completionRequestDocDate: excelSerialToISO(row.P),
        completionInspectionRecordDate: excelSerialToISO(row.Q),
        advancePayment: cleanText(row.R),
        taxInvoiceDate: excelSerialToISO(row.S),
        paymentDate: excelSerialToISO(row.T),
        paymentAmount: toNumber(row.U),
        defectSecurityType: cleanText(row.V),
        licenseType: cleanText(row.W),
        defectStartDate: excelSerialToISO(row.X),
        defectEndDate: excelSerialToISO(row.Y),
        vendorPhone: cleanText(row.Z),
        representative: cleanText(row.AA),
        businessNumber: normalizeBusinessNumber(row.AB),
        fundingSource: cleanText(row.AC),
        ledgerPrint: cleanText(row.AD),
        notes: cleanText(row.AE),
        source: 'audit-excel',
        sourceUpdatedAt: new Date().toISOString()
      });
    }
    return { type: 'audit', label: '감사용 공사이력현황', projects, ignored: 0 };
  }

  function parseEdufine(cells) {
    const projects = [];
    let ignored = 0;
    const last = maxRow(cells);
    for (let r = 5; r <= last; r++) {
      const row = rowMap(cells, r);
      if (!cleanText(row.A) && !cleanText(row.E)) continue;
      const purpose = cleanText(row.B);
      if (purpose && !purpose.includes('공사')) { ignored++; continue; }
      const contractDate = excelSerialToISO(row.F);
      const vendorName = cleanText(row.Q) || cleanText(row.AS) || cleanText(row.AY);
      const representative = cleanText(row.R) || cleanText(row.AT) || cleanText(row.AZ);
      const businessNumber = normalizeBusinessNumber(row.S || row.AU || row.BA);
      const vendorAddress = cleanText(row.U) || cleanText(row.AW) || cleanText(row.BC);
      const vendorPhone = cleanText(row.V) || cleanText(row.AX) || cleanText(row.BD);
      projects.push({
        fiscalYear: inferFiscalYear(contractDate),
        projectName: cleanText(row.E),
        contractNumber: cleanText(row.A),
        contractMethod: cleanText(row.AD) || cleanText(row.C),
        procurementMethod: cleanText(row.K) || cleanText(row.D),
        vendorName,
        representative,
        businessNumber,
        vendorAddress,
        vendorPhone,
        contractDate,
        currentContractAmount: toNumber(row.G),
        originalContractAmount: toNumber(row.AF) || toNumber(row.G),
        estimatedPrice: toNumber(row.AE) || toNumber(row.AC),
        bidRate: toNumber(row.AP) || toNumber(row.AG),
        startDate: excelSerialToISO(row.W),
        completionDueDate: excelSerialToISO(row.X) || excelSerialToISO(row.AB),
        actualCompletionDate: excelSerialToISO(row.Y),
        contractSecurityType: cleanText(row.BE),
        contractSecurityRate: toNumber(row.BG),
        contractSecurityAmount: toNumber(row.BH),
        defectSecurityType: cleanText(row.BI) ? '하자보증' : '',
        defectSecurityRate: toNumber(row.BQ),
        defectSecurityAmount: toNumber(row.BR),
        defectPeriodYears: toNumber(row.BS),
        defectStartDate: excelSerialToISO(row.BT),
        defectEndDate: excelSerialToISO(row.BU),
        supervisor: cleanText(row.FJ),
        completionInspectionDate: excelSerialToISO(row.FR),
        paymentDate: excelSerialToISO(row.FU),
        paymentAmount: toNumber(row.FV),
        source: 'edufine',
        sourceUpdatedAt: new Date().toISOString()
      });
    }
    return { type: 'edufine', label: 'K-에듀파인 자료관리목록', projects, ignored };
  }

  async function parseImport(file) {
    const { cells } = await readWorkbook(file);
    const isAudit = cleanText(cells.get('A4')) === '회계연도' && cleanText(cells.get('C4')) === '공사명';
    const isEdufine = cleanText(cells.get('A3')) === '계약번호' && cleanText(cells.get('E3')) === '계약명';
    if (isAudit) return parseAudit(cells);
    if (isEdufine) return parseEdufine(cells);
    throw new Error('지원하는 엑셀 형식을 찾지 못했습니다. 감사용 공사이력현황 또는 K-에듀파인 자료관리목록인지 확인해주세요.');
  }

  function projectToAuditRow(p, index) {
    return [
      toNumber(p.fiscalYear) || cleanText(p.fiscalYear),
      toNumber(p.sequence) || index + 1,
      cleanText(p.projectName),
      cleanText(p.vendorName),
      cleanText(p.contractMethod),
      toNumber(p.currentContractAmount),
      toNumber(p.estimatedPrice),
      toNumber(p.bidRate),
      cleanText(p.procurementMethod),
      isoToExcelSerial(p.contractDate),
      isoToExcelSerial(p.startDate),
      cleanText(p.siteManagerLicense || p.siteManager),
      isoToExcelSerial(p.completionDueDate),
      isoToExcelSerial(p.actualCompletionDate),
      isoToExcelSerial(p.completionRequestIssueDate),
      isoToExcelSerial(p.completionRequestDocDate),
      isoToExcelSerial(p.completionInspectionRecordDate),
      cleanText(p.advancePayment),
      isoToExcelSerial(p.taxInvoiceDate),
      isoToExcelSerial(p.paymentDate),
      toNumber(p.paymentAmount),
      cleanText(p.defectSecurityType),
      cleanText(p.licenseType),
      isoToExcelSerial(p.defectStartDate),
      isoToExcelSerial(p.defectEndDate),
      cleanText(p.vendorPhone),
      cleanText(p.representative),
      cleanText(p.businessNumber),
      cleanText(p.fundingSource),
      cleanText(p.ledgerPrint),
      cleanText(p.notes)
    ];
  }

  function cellStyleMap(rowXml) {
    const map = {};
    const re = /<c\b([^>]*)\/>|<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    let m;
    while ((m = re.exec(rowXml))) {
      const attrs = m[1] || m[2] || '';
      const address = /\br="([A-Z]+)5"/.exec(attrs)?.[1];
      if (!address) continue;
      map[address] = /\bs="([^"]+)"/.exec(attrs)?.[1] || '';
    }
    return map;
  }

  function cellXml(col, rowNum, value, style) {
    const address = `${col}${rowNum}`;
    const s = style ? ` s="${style}"` : '';
    if (value === '' || value == null) return `<c r="${address}"${s}/>`;
    if (typeof value === 'number' && Number.isFinite(value)) return `<c r="${address}"${s}><v>${value}</v></c>`;
    const text = escapeXml(value);
    return `<c r="${address}"${s} t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`;
  }

  async function exportAuditWorkbook(projects, options = {}) {
    let buffer = options.templateBuffer || null;
    if (!buffer) {
      const response = await fetch('./assets/audit-template.xlsx', { cache: 'no-store' });
      if (!response.ok) throw new Error('감사용 공사이력 템플릿을 불러오지 못했습니다.');
      buffer = await response.arrayBuffer();
    }
    const zip = await JSZipLib.loadAsync(buffer);
    const path = await firstSheetPath(zip);
    let xml = await zip.file(path).async('string');
    const templateRow = /<row\b[^>]*r="5"[^>]*>[\s\S]*?<\/row>/.exec(xml)?.[0];
    if (!templateRow) throw new Error('감사용 템플릿의 데이터 행을 찾지 못했습니다.');
    const styles = cellStyleMap(templateRow);
    const rows = [];
    const ordered = [...projects].sort((a,b) => {
      const ay = Number(a.fiscalYear) || 0, by = Number(b.fiscalYear) || 0;
      if (ay !== by) return ay - by;
      const ad = a.contractDate || '', bd = b.contractDate || '';
      if (ad !== bd) return ad.localeCompare(bd);
      return (a.projectName || '').localeCompare(b.projectName || '', 'ko');
    });
    ordered.forEach((project, idx) => {
      const rowNum = 5 + idx;
      const values = projectToAuditRow(project, idx);
      const cells = values.map((value, ci) => {
        const col = numberToCol(ci + 1);
        return cellXml(col, rowNum, value, styles[col]);
      }).join('');
      rows.push(`<row r="${rowNum}" spans="1:31">${cells}</row>`);
    });
    if (!rows.length) rows.push(templateRow.replace(/r="5"/g, 'r="5"'));

    const existingDataRows = /<row\b[^>]*r="5"[^>]*>[\s\S]*?<\/row>(?:[\s\S]*?)(?=<\/sheetData>)/;
    if (existingDataRows.test(xml)) xml = xml.replace(existingDataRows, rows.join(''));
    else xml = xml.replace('</sheetData>', rows.join('') + '</sheetData>');

    const endRow = Math.max(5, 4 + ordered.length);
    xml = xml.replace(/<dimension\b[^>]*ref="[^"]+"\s*\/?\s*>/, `<dimension ref="A1:AE${endRow}"/>`);
    xml = xml
      .replace(/sqref="V1:V\d+"/, `sqref="V1:V${endRow}"`)
      .replace(/sqref="AD1:AD\d+"/, `sqref="AD1:AD${endRow}"`)
      .replace(/sqref="I3:I\d+"/, `sqref="I3:I${endRow}"`)
      .replace(/sqref="E3:E\d+"/, `sqref="E3:E${endRow}"`);
    zip.file(path, xml);

    const output = await zip.generateAsync({
      type: options.returnUint8Array ? 'uint8array' : 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    if (options.returnUint8Array) return output;
    const suffix = options.year ? `_${options.year}` : '';
    downloadBlob(output, `학교_공사_이력_현황${suffix}.xlsx`);
    return output;
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  globalThis.ConstructionExcel = {
    AUDIT_HEADERS,
    parseImport,
    exportAuditWorkbook,
    excelSerialToISO,
    isoToExcelSerial,
    normalizeBusinessNumber,
    downloadBlob
  };
})();
