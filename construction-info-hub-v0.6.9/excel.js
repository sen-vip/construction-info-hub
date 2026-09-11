(() => {
  'use strict';

  const JSZipLib = globalThis.JSZip;
  if (!JSZipLib) throw new Error('JSZip을 불러오지 못했습니다.');

  const AUDIT_HEADERS = [
    '회계연도','번호','공사명','계약상대자','계약방법','계약금액','예정가격','낙찰율 ',
    'G2B/S2B','계약일','착공일','현장대리인\n자격증종류','준공기한','준공일','준공검사원발행일',
    '준공검사원 문서등록일','준공검사조서작성일(3천이하 생략가능)','선금지급여부','세금계산서 발행일','지출일',
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
    const siRe = /<(?:[A-Za-z_][\w.-]*:)?si\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?si>/g;
    let m;
    while ((m = siRe.exec(xml))) {
      const chunks = [];
      const tRe = /<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g;
      let tm;
      while ((tm = tRe.exec(m[1]))) chunks.push(decodeXml(tm[1]));
      out.push(chunks.join(''));
    }
    return out;
  }

  function parseSheetCells(xml = '', sharedStrings = []) {
    const cells = new Map();
    const re = /<(?:[A-Za-z_][\w.-]*:)?c\b([^>]*)\/>|<(?:[A-Za-z_][\w.-]*:)?c\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?c>/g;
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
        const tr = /<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g;
        let tm;
        while ((tm = tr.exec(body))) parts.push(decodeXml(tm[1]));
        value = parts.join('');
      } else {
        const vm = /<(?:[A-Za-z_][\w.-]*:)?v>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?v>/.exec(body);
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
    const id = /<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*r:id="([^"]+)"/.exec(wbXml)?.[1];
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

  function findAuditHeaderRow(cells) {
    const last = Math.min(maxRow(cells), 12);
    for (let r = 1; r <= last; r++) {
      const row = rowMap(cells, r);
      if (cleanText(row.A) === '회계연도' && cleanText(row.C) === '공사명') return r;
    }
    return 0;
  }

  function parseAudit(cells) {
    const projects = [];
    const last = maxRow(cells);
    const headerRow = findAuditHeaderRow(cells);
    if (!headerRow) return { type: 'audit', label: '공사관리대장', projects, ignored: 0 };
    for (let r = headerRow + 1; r <= last; r++) {
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
    return { type: 'audit', label: '공사관리대장', projects, ignored: 0 };
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
    const auditHeaderRow = findAuditHeaderRow(cells);
    const isAudit = auditHeaderRow > 0;
    const isEdufine = cleanText(cells.get('A3')) === '계약번호' && cleanText(cells.get('E3')) === '계약명';
    if (isAudit) return parseAudit(cells);
    if (isEdufine) return parseEdufine(cells);
    throw new Error('지원하는 엑셀 형식을 찾지 못했습니다. 공사관리대장 또는 K-에듀파인 자료관리목록인지 확인해주세요.');
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

  function cellStyleMap(rowXml, rowNum) {
    const map = {};
    const re = /<(?:[A-Za-z_][\w.-]*:)?c\b([^>]*)\/>|<(?:[A-Za-z_][\w.-]*:)?c\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?c>/g;
    let m;
    while ((m = re.exec(rowXml))) {
      const attrs = m[1] || m[2] || '';
      const address = new RegExp(`\\br="([A-Z]+)${rowNum}"`).exec(attrs)?.[1];
      if (!address) continue;
      map[address] = /\bs="([^"]+)"/.exec(attrs)?.[1] || '';
    }
    return map;
  }

  function rowOpenTagFromTemplate(rowXml, rowNum) {
    const match = /<((?:[A-Za-z_][\w.-]*:)?row)\b([^>]*)>/.exec(rowXml);
    const tag = match?.[1] || 'row';
    const attrs = match?.[2] || '';
    let next = attrs.replace(/\br="\d+"/, `r="${rowNum}"`);
    if (!/\br="/.test(next)) next = ` r="${rowNum}"${next}`;
    return `<${tag}${next}>`;
  }

  function xmlPrefixFromTemplate(rowXml) {
    return /<([A-Za-z_][\w.-]*:)c\b/.exec(rowXml)?.[1] || '';
  }

  function rowTagFromTemplate(rowXml) {
    return /<((?:[A-Za-z_][\w.-]*:)?row)\b/.exec(rowXml)?.[1] || 'row';
  }

  function cellXml(col, rowNum, value, style, prefix = '') {
    const address = `${col}${rowNum}`;
    const s = style ? ` s="${style}"` : '';
    const c = `${prefix}c`, v = `${prefix}v`, is = `${prefix}is`, t = `${prefix}t`;
    if (value === '' || value == null) return `<${c} r="${address}"${s}/>`;
    if (typeof value === 'number' && Number.isFinite(value)) return `<${c} r="${address}"${s}><${v}>${value}</${v}></${c}>`;
    const text = escapeXml(value);
    return `<${c} r="${address}"${s} t="inlineStr"><${is}><${t} xml:space="preserve">${text}</${t}></${is}></${c}>`;
  }

  async function exportAuditWorkbook(projects, options = {}) {
    let buffer = options.templateBuffer || null;
    if (!buffer) {
      const templateUrl = new URL('./assets/공사관리대장.xlsx', document.baseURI || window.location.href);
      const response = await fetch(templateUrl.href, { cache: 'no-store' });
      if (!response.ok) throw new Error('공사관리대장 템플릿을 불러오지 못했습니다.');
      buffer = await response.arrayBuffer();
    }
    const zip = await JSZipLib.loadAsync(buffer);
    const path = await firstSheetPath(zip);
    let xml = await zip.file(path).async('string');

    let shared = [];
    if (zip.file('xl/sharedStrings.xml')) {
      shared = parseSharedStrings(await zip.file('xl/sharedStrings.xml').async('string'));
    }
    const templateCells = parseSheetCells(xml, shared);
    const headerRow = findAuditHeaderRow(templateCells);
    if (!headerRow) throw new Error('공사관리대장 템플릿의 헤더를 찾지 못했습니다.');
    const dataRow = headerRow + 1;
    const templateRowRe = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?row\\b[^>]*r="${dataRow}"[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?row>`);
    const templateRow = templateRowRe.exec(xml)?.[0];
    if (!templateRow) throw new Error('공사관리대장 템플릿의 데이터 행을 찾지 못했습니다.');
    const styles = cellStyleMap(templateRow, dataRow);
    const xmlPrefix = xmlPrefixFromTemplate(templateRow);
    const rowTag = rowTagFromTemplate(templateRow);

    const ordered = [...projects].sort((a,b) => {
      const ay = Number(a.fiscalYear) || 0, by = Number(b.fiscalYear) || 0;
      if (ay !== by) return ay - by;
      const ad = a.contractDate || '', bd = b.contractDate || '';
      if (ad !== bd) return ad.localeCompare(bd);
      return (a.projectName || '').localeCompare(b.projectName || '', 'ko');
    });

    const makeRow = (rowNum, values) => {
      const cells = values.map((value, ci) => {
        const col = numberToCol(ci + 1);
        return cellXml(col, rowNum, value, styles[col], xmlPrefix);
      }).join('');
      return `${rowOpenTagFromTemplate(templateRow, rowNum)}${cells}</${rowTag}>`;
    };

    const rows = ordered.map((project, idx) => makeRow(dataRow + idx, projectToAuditRow(project, idx)));
    if (!rows.length) rows.push(makeRow(dataRow, new Array(31).fill('')));

    const existingDataRows = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?row\\b[^>]*r="${dataRow}"[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?row>(?:[\\s\\S]*?)(?=<\\/(?:[A-Za-z_][\\w.-]*:)?sheetData>)`);
    if (existingDataRows.test(xml)) xml = xml.replace(existingDataRows, rows.join(''));
    else xml = xml.replace(/<\/(?:[A-Za-z_][\w.-]*:)?sheetData>/, m => rows.join('') + m);

    const endRow = Math.max(dataRow, headerRow + ordered.length);
    xml = xml.replace(/<((?:[A-Za-z_][\w.-]*:)?dimension)\b[^>]*ref="[^"]+"\s*\/?\s*>/, (_, tag) => `<${tag} ref="A1:AE${endRow}"/>`);
    for (const col of ['V', 'AD', 'I', 'E']) {
      const re = new RegExp(`sqref="${col}(\\d+):${col}\\d+"`);
      xml = xml.replace(re, (_, startRow) => `sqref="${col}${startRow}:${col}${endRow}"`);
    }
    zip.file(path, xml);

    const output = await zip.generateAsync({
      type: options.returnUint8Array ? 'uint8array' : 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    if (options.returnUint8Array) return output;
    const suffix = options.year ? `_${options.year}` : '';
    downloadBlob(output, `공사관리대장${suffix}.xlsx`);
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
