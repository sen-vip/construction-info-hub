(() => {
  'use strict';

  const DB_NAME = 'construction-info-hub';
  const DB_VERSION = 2;
  let dbPromise = null;

  function uuid() {
    if (globalThis.crypto?.randomUUID) return crypto.randomUUID();
    return 'id-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
  }

  function openDb() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains('projects')) {
          const s = db.createObjectStore('projects', { keyPath: 'id' });
          s.createIndex('updatedAt', 'updatedAt');
          s.createIndex('fiscalYear', 'fiscalYear');
          s.createIndex('contractNumber', 'contractNumber', { unique: false });
        }
        if (!db.objectStoreNames.contains('vendors')) {
          const s = db.createObjectStore('vendors', { keyPath: 'id' });
          s.createIndex('name', 'name', { unique: false });
          s.createIndex('businessNumber', 'businessNumber', { unique: false });
        }
        if (!db.objectStoreNames.contains('settings')) {
          db.createObjectStore('settings', { keyPath: 'key' });
        }
        if (!db.objectStoreNames.contains('payouts')) {
          db.createObjectStore('payouts', { keyPath: 'vendorId' });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
      req.onblocked = () => reject(new Error('브라우저 저장소를 열 수 없습니다. 다른 탭에서 이 앱을 닫은 뒤 다시 시도해주세요.'));
    });
    return dbPromise;
  }

  async function tx(storeName, mode, operation) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tr = db.transaction(storeName, mode);
      const store = tr.objectStore(storeName);
      let result;
      try { result = operation(store); }
      catch (err) { reject(err); return; }
      tr.oncomplete = () => resolve(result?.result ?? result);
      tr.onerror = () => reject(tr.error);
      tr.onabort = () => reject(tr.error || new Error('저장 작업이 취소되었습니다.'));
    });
  }

  async function getAll(storeName) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tr = db.transaction(storeName, 'readonly');
      const req = tr.objectStore(storeName).getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  }

  async function get(storeName, key) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tr = db.transaction(storeName, 'readonly');
      const req = tr.objectStore(storeName).get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  }

  async function put(storeName, value) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tr = db.transaction(storeName, 'readwrite');
      const req = tr.objectStore(storeName).put(value);
      req.onsuccess = () => resolve(value);
      req.onerror = () => reject(req.error);
    });
  }

  async function remove(storeName, key) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tr = db.transaction(storeName, 'readwrite');
      const req = tr.objectStore(storeName).delete(key);
      req.onsuccess = () => resolve(true);
      req.onerror = () => reject(req.error);
    });
  }

  async function clear(storeName) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tr = db.transaction(storeName, 'readwrite');
      const req = tr.objectStore(storeName).clear();
      req.onsuccess = () => resolve(true);
      req.onerror = () => reject(req.error);
    });
  }

  async function replaceAll(storeName, values) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tr = db.transaction(storeName, 'readwrite');
      const store = tr.objectStore(storeName);
      store.clear();
      (values || []).forEach(v => store.put(v));
      tr.oncomplete = () => resolve(true);
      tr.onerror = () => reject(tr.error);
    });
  }

  function createProject(seed = {}) {
    const now = new Date().toISOString();
    return {
      id: uuid(),
      fiscalYear: seed.fiscalYear || String(new Date().getFullYear()),
      sequence: seed.sequence || '',
      projectName: seed.projectName || '',
      workType: seed.workType || '',
      vendorId: seed.vendorId || '',
      vendorName: seed.vendorName || '',
      representative: seed.representative || '',
      businessNumber: seed.businessNumber || '',
      vendorAddress: seed.vendorAddress || '',
      vendorPhone: seed.vendorPhone || '',
      licenseType: seed.licenseType || '',
      contractMethod: seed.contractMethod || '',
      procurementMethod: seed.procurementMethod || '',
      contractNumber: seed.contractNumber || '',
      estimatedPrice: seed.estimatedPrice ?? '',
      originalContractAmount: seed.originalContractAmount ?? '',
      currentContractAmount: seed.currentContractAmount ?? seed.contractAmount ?? '',
      bidRate: seed.bidRate ?? '',
      contractDate: seed.contractDate || '',
      plannedStartDate: seed.plannedStartDate || '',
      startDate: seed.startDate || '',
      completionDueDate: seed.completionDueDate || '',
      actualCompletionDate: seed.actualCompletionDate || '',
      completionInspectionDate: seed.completionInspectionDate || '',
      completionRequestIssueDate: seed.completionRequestIssueDate || '',
      completionRequestDocDate: seed.completionRequestDocDate || '',
      completionInspectionRecordDate: seed.completionInspectionRecordDate || '',
      siteManager: seed.siteManager || '',
      siteManagerLicense: seed.siteManagerLicense || '',
      settlementAmount: seed.settlementAmount ?? '',
      priorPaymentAmount: seed.priorPaymentAmount ?? '',
      deductionAmount: seed.deductionAmount ?? '',
      supervisor: seed.supervisor || '',
      inspector: seed.inspector || '',
      witness: seed.witness || '',
      advancePayment: seed.advancePayment || '',
      taxInvoiceDate: seed.taxInvoiceDate || '',
      paymentDate: seed.paymentDate || '',
      paymentAmount: seed.paymentAmount ?? '',
      contractSecurityType: seed.contractSecurityType || '',
      contractSecurityRate: seed.contractSecurityRate ?? '',
      contractSecurityAmount: seed.contractSecurityAmount ?? '',
      delayPenaltyRate: seed.delayPenaltyRate || '',
      priceAdjustmentMethod: seed.priceAdjustmentMethod || '',
      defectSecurityType: seed.defectSecurityType || '',
      defectSecurityRate: seed.defectSecurityRate ?? '',
      defectSecurityAmount: seed.defectSecurityAmount ?? '',
      defectPeriodYears: seed.defectPeriodYears ?? '',
      defectStartDate: seed.defectStartDate || '',
      defectEndDate: seed.defectEndDate || '',
      fundingSource: seed.fundingSource || '',
      designer: seed.designer || '',
      budgetPolicyProject: seed.budgetPolicyProject || '',
      budgetUnitProject: seed.budgetUnitProject || '',
      budgetDetailProject: seed.budgetDetailProject || '',
      budgetDetailItem: seed.budgetDetailItem || '',
      costStatisticsItem: seed.costStatisticsItem || '',
      ledgerPrint: seed.ledgerPrint || '',
      warrantyInspections: Array.isArray(seed.warrantyInspections) ? seed.warrantyInspections : [],
      utilityCost: seed.utilityCost && typeof seed.utilityCost === 'object' ? { ...seed.utilityCost } : {
        directMaterialCost: '', directLaborCost: '', facilityUse: '수도광열비·전력비', electricCost: '', waterHeatCost: '', total: '', calculatedAt: ''
      },
      notes: seed.notes || '',
      source: seed.source || 'manual',
      sourceUpdatedAt: seed.sourceUpdatedAt || '',
      contractChanges: Array.isArray(seed.contractChanges) ? seed.contractChanges : [],
      printHistory: Array.isArray(seed.printHistory) ? seed.printHistory : [],
      createdAt: seed.createdAt || now,
      updatedAt: seed.updatedAt || now
    };
  }

  function createVendor(seed = {}) {
    const now = new Date().toISOString();
    return {
      id: seed.id || uuid(),
      name: seed.name || '',
      representative: seed.representative || '',
      businessNumber: seed.businessNumber || '',
      phone: seed.phone || '',
      address: seed.address || '',
      workType: seed.workType || '',
      licenseType: seed.licenseType || '',
      createdAt: seed.createdAt || now,
      updatedAt: seed.updatedAt || now
    };
  }


  function createPayout(seed = {}) {
    const now = new Date().toISOString();
    return {
      vendorId: seed.vendorId || '',
      bankName: seed.bankName || '',
      accountNumber: seed.accountNumber || '',
      accountHolder: seed.accountHolder || '',
      createdAt: seed.createdAt || now,
      updatedAt: seed.updatedAt || now
    };
  }

  async function exportBackup() {
    const [projects, vendors, payouts, settings] = await Promise.all([
      getAll('projects'), getAll('vendors'), getAll('payouts'), getAll('settings')
    ]);
    return {
      app: 'construction-info-hub',
      version: 2,
      exportedAt: new Date().toISOString(),
      projects,
      vendors,
      payouts,
      settings
    };
  }

  async function importBackup(data) {
    if (!data || data.app !== 'construction-info-hub' || !Array.isArray(data.projects)) {
      throw new Error('공사정보 허브 백업파일이 아닙니다.');
    }
    await Promise.all([
      replaceAll('projects', data.projects || []),
      replaceAll('vendors', data.vendors || []),
      replaceAll('payouts', data.payouts || []),
      replaceAll('settings', data.settings || [])
    ]);
  }

  globalThis.ConstructionDB = {
    uuid,
    openDb,
    getAll,
    get,
    put,
    remove,
    clear,
    replaceAll,
    createProject,
    createVendor,
    createPayout,
    exportBackup,
    importBackup
  };
})();
