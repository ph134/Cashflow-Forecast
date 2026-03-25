const STORAGE_KEY = 'cashflow-web-app:state:v3';
const DEFAULT_STATE_KEY = 'cashflow-web-app:default-state:v1';
const DEFAULTS_VERSION_KEY = 'cashflow-web-app:defaults-version';
const DEFAULTS_VERSION = '20260324-v4';
const SNAPSHOT_SCHEMA_VERSION = 1;

function createDefaultState() {
  const engineerId = crypto.randomUUID();
  const materialId = crypto.randomUUID();
  const pmId = crypto.randomUUID();
  const testId = crypto.randomUUID();

  return {
    project: {
      salesforceOpportunity: '1023455',
      revision: '1',
      opportunityName: 'Test - CashBonus',
      manualMargin: '50%',
      contractValue: 325000,
      contractCurrency: 'NOK',
      convertToCurrency: 'NOK',
      contractFxRate: 1,
      contractRateIsManual: false,
      projectStartMonth: '2026-06',
      quotedLeadTimeMonths: 6,
      netDays: 0,
    },
    milestones: [
      { id: crypto.randomUUID(), code: 'MS01', label: 'Upon PO Received', percent: 30, invoiceMonth: '2026-06' },
      { id: crypto.randomUUID(), code: 'MS02', label: 'Upon start of production', percent: 30, invoiceMonth: '2026-08' },
      { id: crypto.randomUUID(), code: 'MS03', label: 'Upon delivery of equipment', percent: 30, invoiceMonth: '2026-10' },
      { id: crypto.randomUUID(), code: 'MS04', label: 'Delivery of documentation', percent: 10, invoiceMonth: '2026-12' },
    ],
    costs: [
      { id: engineerId, label: 'Engineer', totalCost: 17308.56, currency: 'NOK', convertToCurrency: 'NOK', conversionRate: 1, rateIsManual: false },
      { id: materialId, label: 'Material', totalCost: 131759.86, currency: 'NOK', convertToCurrency: 'NOK', conversionRate: 1, rateIsManual: false },
      { id: pmId, label: 'PM', totalCost: 10000, currency: 'NOK', convertToCurrency: 'NOK', conversionRate: 1, rateIsManual: false },
      { id: testId, label: 'Test', totalCost: 18.48, currency: 'NOK', convertToCurrency: 'NOK', conversionRate: 1, rateIsManual: false },
    ],
    progress: {
      [engineerId]: [0, 60, 100, 100, 100, 100, 100],
      [materialId]: [0, 30, 40, 60, 80, 80, 90],
      [pmId]: [10, 30, 40, 59, 65, 80, 100],
      [testId]: [0, 0, 0, 0, 0, 0, 100],
    },
  };
}

function getEffectiveDefaultState() {
  const baseDefaults = createDefaultState();
  try {
    const raw = localStorage.getItem(DEFAULT_STATE_KEY);
    if (!raw) return baseDefaults;
    const custom = JSON.parse(raw);
    return {
      project: { ...baseDefaults.project, ...(custom?.project || {}) },
      milestones: Array.isArray(custom?.milestones) && custom.milestones.length ? custom.milestones : baseDefaults.milestones,
      costs: Array.isArray(custom?.costs) && custom.costs.length ? custom.costs : baseDefaults.costs,
      progress: custom?.progress && typeof custom.progress === 'object' ? custom.progress : baseDefaults.progress,
    };
  } catch {
    return baseDefaults;
  }
}

function loadInitialState() {
  try {
    const seenDefaultsVersion = localStorage.getItem(DEFAULTS_VERSION_KEY);
    if (seenDefaultsVersion !== DEFAULTS_VERSION) {
      localStorage.removeItem('cashflow-web-app:state:v1');
      localStorage.removeItem('cashflow-web-app:state:v2');
      localStorage.removeItem('cashflow-web-app:state:v3');
      localStorage.removeItem(STORAGE_KEY);
      localStorage.setItem(DEFAULTS_VERSION_KEY, DEFAULTS_VERSION);
    }
  } catch {
    // Ignore storage errors
  }

  const defaults = getEffectiveDefaultState();
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaults;
    const saved = JSON.parse(raw);
    return {
      project: { ...defaults.project, ...(saved?.project || {}) },
      milestones: Array.isArray(saved?.milestones) && saved.milestones.length ? saved.milestones : defaults.milestones,
      costs: Array.isArray(saved?.costs) && saved.costs.length ? saved.costs : defaults.costs,
      progress: saved?.progress && typeof saved.progress === 'object' ? saved.progress : {},
    };
  } catch {
    return defaults;
  }
}

function persistState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // Ignore storage errors (private mode/quota)
  }
}

const state = loadInitialState();

const projectFields = [
  { key: 'quotedLeadTimeMonths', label: 'Quoted lead time (months)', type: 'number', step: '1', min: '0' },
  { key: 'projectStartMonth', label: 'Project start month', type: 'month' },
];

const CURRENCY_OPTIONS = ['USD', 'EUR', 'GBP', 'JPY', 'CHF', 'CAD', 'AUD', 'NOK', 'SEK', 'DKK', 'BRL', 'INR', 'MXN', 'SGD', 'HKD', 'NZD'];

const dom = {
  projectForm: document.querySelector('#projectForm'),
  healthChecks: document.querySelector('#healthChecks'),
  milestonesBody: document.querySelector('#milestonesBody'),
  costsBody: document.querySelector('#costsBody'),
  progressGrid: document.querySelector('#progressGrid'),
  summaryList: document.querySelector('#summaryList'),
  forecastHead: document.querySelector('#forecastHead'),
  forecastBody: document.querySelector('#forecastBody'),
  chartLegend: document.querySelector('#chartLegend'),
  chartOpportunity: document.querySelector('#chartOpportunity'),
  chartRevision: document.querySelector('#chartRevision'),
  chartOpportunityName: document.querySelector('#chartOpportunityName'),
  chartDate: document.querySelector('#chartDate'),
  chart: document.querySelector('#cashflowChart'),
  addMilestoneBtn: document.querySelector('#addMilestoneBtn'),
  addCostBtn: document.querySelector('#addCostBtn'),
  exportPptBtn: document.querySelector('#exportPptBtn'),
  exportXlsBtn: document.querySelector('#exportXlsBtn'),
  importXlsBtn: document.querySelector('#importXlsBtn'),
  setDefaultsBtn: document.querySelector('#setDefaultsBtn'),
  saveSnapshotBtn: document.querySelector('#saveSnapshotBtn'),
  loadSnapshotBtn: document.querySelector('#loadSnapshotBtn'),
  resetDefaultsBtn: document.querySelector('#resetDefaultsBtn'),
  cashflowEstimate: document.querySelector('#cashflowEstimate'),
  userGuideBtn: document.querySelector('#userGuideBtn'),
  userGuideModal: document.querySelector('#userGuideModal'),
  closeGuideBtn: document.querySelector('#closeGuideBtn'),
};

function cloneStateForSnapshot(source) {
  return {
    project: { ...(source?.project || {}) },
    milestones: Array.isArray(source?.milestones) ? source.milestones.map((milestone) => ({ ...milestone })) : [],
    costs: Array.isArray(source?.costs) ? source.costs.map((cost) => ({ ...cost })) : [],
    progress: source?.progress && typeof source.progress === 'object' ? JSON.parse(JSON.stringify(source.progress)) : {},
  };
}

function normalizeImportedState(rawState) {
  const defaults = createDefaultState();
  const normalizedProject = { ...defaults.project, ...(rawState?.project || {}) };

  const normalizedMilestones = Array.isArray(rawState?.milestones) && rawState.milestones.length
    ? rawState.milestones.map((milestone, index) => ({
      id: String(milestone?.id || crypto.randomUUID()),
      code: String(milestone?.code || `MS${String(index + 1).padStart(2, '0')}`),
      label: String(milestone?.label || 'Milestone'),
      percent: clampPercent(milestone?.percent),
      invoiceMonth: String(milestone?.invoiceMonth || normalizedProject.projectStartMonth),
    }))
    : defaults.milestones;

  const normalizedCosts = Array.isArray(rawState?.costs) && rawState.costs.length
    ? rawState.costs.map((cost, index) => ({
      id: String(cost?.id || crypto.randomUUID()),
      label: String(cost?.label || `Cost Element ${index + 1}`),
      totalCost: clampNumber(cost?.totalCost, 0),
      currency: String(cost?.currency || normalizedProject.contractCurrency || 'USD').toUpperCase(),
      convertToCurrency: String(cost?.convertToCurrency || normalizedProject.convertToCurrency || 'USD').toUpperCase(),
      conversionRate: clampNumber(cost?.conversionRate, 1),
      rateIsManual: Boolean(cost?.rateIsManual),
    }))
    : defaults.costs;

  const normalizedProgress = {};
  const sourceProgress = rawState?.progress && typeof rawState.progress === 'object' ? rawState.progress : {};
  normalizedCosts.forEach((cost) => {
    const rawRow = Array.isArray(sourceProgress[cost.id]) ? sourceProgress[cost.id] : [];
    normalizedProgress[cost.id] = normalizeProgressRow(rawRow.map((value) => clampPercent(value)));
  });

  return {
    project: normalizedProject,
    milestones: normalizedMilestones,
    costs: normalizedCosts,
    progress: normalizedProgress,
  };
}

function applyState(nextState) {
  state.project = { ...nextState.project };
  state.milestones = nextState.milestones.map((milestone) => ({ ...milestone }));
  state.costs = nextState.costs.map((cost) => ({ ...cost }));
  state.progress = JSON.parse(JSON.stringify(nextState.progress || {}));
  ensureProgressShape();
  rerender();
}

function downloadJsonFile(filename, payload) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = objectUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(objectUrl);
}

function saveSnapshot() {
  const safeName = String(state.project.opportunityName || state.project.salesforceOpportunity || 'cashflow')
    .replace(/[^a-z0-9\-_ ]/gi, '')
    .trim()
    .replace(/\s+/g, '-')
    .toLowerCase() || 'cashflow';const STORAGE_KEY = 'cashflow-web-app:state:v3';
const DEFAULT_STATE_KEY = 'cashflow-web-app:default-state:v1';
const DEFAULTS_VERSION_KEY = 'cashflow-web-app:defaults-version';
const DEFAULTS_VERSION = '20260324-v4';
const SNAPSHOT_SCHEMA_VERSION = 1;

function createDefaultState() {
  const engineerId = crypto.randomUUID();
  const materialId = crypto.randomUUID();
  const pmId = crypto.randomUUID();
  const testId = crypto.randomUUID();

  return {
    project: {
      salesforceOpportunity: '1023455',
      revision: '1',
      opportunityName: 'Test - CashBonus',
      manualMargin: '50%',
      contractValue: 325000,
      contractCurrency: 'NOK',
      convertToCurrency: 'NOK',
      contractFxRate: 1,
      contractRateIsManual: false,
      projectStartMonth: '2026-06',
      quotedLeadTimeMonths: 6,
      netDays: 0,
    },
    milestones: [
      { id: crypto.randomUUID(), code: 'MS01', label: 'Upon PO Received', percent: 30, invoiceMonth: '2026-06' },
      { id: crypto.randomUUID(), code: 'MS02', label: 'Upon start of production', percent: 30, invoiceMonth: '2026-08' },
      { id: crypto.randomUUID(), code: 'MS03', label: 'Upon delivery of equipment', percent: 30, invoiceMonth: '2026-10' },
      { id: crypto.randomUUID(), code: 'MS04', label: 'Delivery of documentation', percent: 10, invoiceMonth: '2026-12' },
    ],
    costs: [
      { id: engineerId, label: 'Engineer', totalCost: 17308.56, currency: 'NOK', convertToCurrency: 'NOK', conversionRate: 1, rateIsManual: false },
      { id: materialId, label: 'Material', totalCost: 131759.86, currency: 'NOK', convertToCurrency: 'NOK', conversionRate: 1, rateIsManual: false },
      { id: pmId, label: 'PM', totalCost: 10000, currency: 'NOK', convertToCurrency: 'NOK', conversionRate: 1, rateIsManual: false },
      { id: testId, label: 'Test', totalCost: 18.48, currency: 'NOK', convertToCurrency: 'NOK', conversionRate: 1, rateIsManual: false },
    ],
    progress: {
      [engineerId]: [0, 60, 100, 100, 100, 100, 100],
      [materialId]: [0, 30, 40, 60, 80, 80, 90],
      [pmId]: [10, 30, 40, 59, 65, 80, 100],
      [testId]: [0, 0, 0, 0, 0, 0, 100],
    },
  };
}

function getEffectiveDefaultState() {
  const baseDefaults = createDefaultState();
  try {
    const raw = localStorage.getItem(DEFAULT_STATE_KEY);
    if (!raw) return baseDefaults;
    const custom = JSON.parse(raw);
    return {
      project: { ...baseDefaults.project, ...(custom?.project || {}) },
      milestones: Array.isArray(custom?.milestones) && custom.milestones.length ? custom.milestones : baseDefaults.milestones,
      costs: Array.isArray(custom?.costs) && custom.costs.length ? custom.costs : baseDefaults.costs,
      progress: custom?.progress && typeof custom.progress === 'object' ? custom.progress : baseDefaults.progress,
    };
  } catch {
    return baseDefaults;
  }
}

function loadInitialState() {
  try {
    const seenDefaultsVersion = localStorage.getItem(DEFAULTS_VERSION_KEY);
    if (seenDefaultsVersion !== DEFAULTS_VERSION) {
      localStorage.removeItem('cashflow-web-app:state:v1');
      localStorage.removeItem('cashflow-web-app:state:v2');
      localStorage.removeItem('cashflow-web-app:state:v3');
      localStorage.removeItem(STORAGE_KEY);
      localStorage.setItem(DEFAULTS_VERSION_KEY, DEFAULTS_VERSION);
    }
  } catch {
    // Ignore storage errors
  }

  const defaults = getEffectiveDefaultState();
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaults;
    const saved = JSON.parse(raw);
    return {
      project: { ...defaults.project, ...(saved?.project || {}) },
      milestones: Array.isArray(saved?.milestones) && saved.milestones.length ? saved.milestones : defaults.milestones,
      costs: Array.isArray(saved?.costs) && saved.costs.length ? saved.costs : defaults.costs,
      progress: saved?.progress && typeof saved.progress === 'object' ? saved.progress : {},
    };
  } catch {
    return defaults;
  }
}

function persistState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // Ignore storage errors (private mode/quota)
  }
}

const state = loadInitialState();

const projectFields = [
  { key: 'quotedLeadTimeMonths', label: 'Quoted lead time (months)', type: 'number', step: '1', min: '0' },
  { key: 'projectStartMonth', label: 'Project start month', type: 'month' },
];

const CURRENCY_OPTIONS = ['USD', 'EUR', 'GBP', 'JPY', 'CHF', 'CAD', 'AUD', 'NOK', 'SEK', 'DKK', 'BRL', 'INR', 'MXN', 'SGD', 'HKD', 'NZD'];

const dom = {
  projectForm: document.querySelector('#projectForm'),
  healthChecks: document.querySelector('#healthChecks'),
  milestonesBody: document.querySelector('#milestonesBody'),
  costsBody: document.querySelector('#costsBody'),
  progressGrid: document.querySelector('#progressGrid'),
  summaryList: document.querySelector('#summaryList'),
  forecastHead: document.querySelector('#forecastHead'),
  forecastBody: document.querySelector('#forecastBody'),
  chartLegend: document.querySelector('#chartLegend'),
  chartOpportunity: document.querySelector('#chartOpportunity'),
  chartRevision: document.querySelector('#chartRevision'),
  chartOpportunityName: document.querySelector('#chartOpportunityName'),
  chartDate: document.querySelector('#chartDate'),
  chart: document.querySelector('#cashflowChart'),
  addMilestoneBtn: document.querySelector('#addMilestoneBtn'),
  addCostBtn: document.querySelector('#addCostBtn'),
  exportPptBtn: document.querySelector('#exportPptBtn'),
  exportXlsBtn: document.querySelector('#exportXlsBtn'),
  importXlsBtn: document.querySelector('#importXlsBtn'),
  setDefaultsBtn: document.querySelector('#setDefaultsBtn'),
  saveSnapshotBtn: document.querySelector('#saveSnapshotBtn'),
  loadSnapshotBtn: document.querySelector('#loadSnapshotBtn'),
  resetDefaultsBtn: document.querySelector('#resetDefaultsBtn'),
  cashflowEstimate: document.querySelector('#cashflowEstimate'),
  userGuideBtn: document.querySelector('#userGuideBtn'),
  userGuideModal: document.querySelector('#userGuideModal'),
  closeGuideBtn: document.querySelector('#closeGuideBtn'),
};

function cloneStateForSnapshot(source) {
  return {
    project: { ...(source?.project || {}) },
    milestones: Array.isArray(source?.milestones) ? source.milestones.map((milestone) => ({ ...milestone })) : [],
    costs: Array.isArray(source?.costs) ? source.costs.map((cost) => ({ ...cost })) : [],
    progress: source?.progress && typeof source.progress === 'object' ? JSON.parse(JSON.stringify(source.progress)) : {},
  };
}

function normalizeImportedState(rawState) {
  const defaults = createDefaultState();
  const normalizedProject = { ...defaults.project, ...(rawState?.project || {}) };

  const normalizedMilestones = Array.isArray(rawState?.milestones) && rawState.milestones.length
    ? rawState.milestones.map((milestone, index) => ({
      id: String(milestone?.id || crypto.randomUUID()),
      code: String(milestone?.code || `MS${String(index + 1).padStart(2, '0')}`),
      label: String(milestone?.label || 'Milestone'),
      percent: clampPercent(milestone?.percent),
      invoiceMonth: String(milestone?.invoiceMonth || normalizedProject.projectStartMonth),
    }))
    : defaults.milestones;

  const normalizedCosts = Array.isArray(rawState?.costs) && rawState.costs.length
    ? rawState.costs.map((cost, index) => ({
      id: String(cost?.id || crypto.randomUUID()),
      label: String(cost?.label || `Cost Element ${index + 1}`),
      totalCost: clampNumber(cost?.totalCost, 0),
      currency: String(cost?.currency || normalizedProject.contractCurrency || 'USD').toUpperCase(),
      convertToCurrency: String(cost?.convertToCurrency || normalizedProject.convertToCurrency || 'USD').toUpperCase(),
      conversionRate: clampNumber(cost?.conversionRate, 1),
      rateIsManual: Boolean(cost?.rateIsManual),
    }))
    : defaults.costs;

  const normalizedProgress = {};
  const sourceProgress = rawState?.progress && typeof rawState.progress === 'object' ? rawState.progress : {};
  normalizedCosts.forEach((cost) => {
    const rawRow = Array.isArray(sourceProgress[cost.id]) ? sourceProgress[cost.id] : [];
    normalizedProgress[cost.id] = normalizeProgressRow(rawRow.map((value) => clampPercent(value)));
  });

  return {
    project: normalizedProject,
    milestones: normalizedMilestones,
    costs: normalizedCosts,
    progress: normalizedProgress,
  };
}

function applyState(nextState) {
  state.project = { ...nextState.project };
  state.milestones = nextState.milestones.map((milestone) => ({ ...milestone }));
  state.costs = nextState.costs.map((cost) => ({ ...cost }));
  state.progress = JSON.parse(JSON.stringify(nextState.progress || {}));
  ensureProgressShape();
  rerender();
}

function downloadJsonFile(filename, payload) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = objectUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(objectUrl);
}

function saveSnapshot() {
  const safeName = String(state.project.opportunityName || state.project.salesforceOpportunity || 'cashflow')
    .replace(/[^a-z0-9\-_ ]/gi, '')
    .trim()
    .replace(/\s+/g, '-')
    .toLowerCase() || 'cashflow';
