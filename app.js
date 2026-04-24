const STORAGE_KEY = 'cashflow-web-app:state:v3';
const DEFAULT_STATE_KEY = 'cashflow-web-app:default-state:v1';
const DEFAULTS_VERSION_KEY = 'cashflow-web-app:defaults-version';
const DEFAULTS_VERSION = '20260423-v7';
const SNAPSHOT_SCHEMA_VERSION = 1;

function createDefaultState() {
  const starterCostId = crypto.randomUUID();
  const starterMonth = new Date().toISOString().slice(0, 7);

  return {
    project: {
      salesforceOpportunity: '',
      revision: '',
      opportunityName: '',
      contractValue: 0,
      contractCurrency: 'USD',
      convertToCurrency: 'USD',
      contractFxRate: 1,
      contractRateIsManual: false,
      projectStartMonth: starterMonth,
      quotedLeadTimeMonths: 0,
      netDays: 30,
    },
    milestones: [
      {
        id: crypto.randomUUID(),
        code: 'MS01',
        label: 'Upon Placement of PO',
        percent: 30,
        invoiceMonth: '',
      },
      {
        id: crypto.randomUUID(),
        code: 'MS02',
        label: 'Raw Material Receipt',
        percent: 30,
        invoiceMonth: '',
      },
      {
        id: crypto.randomUUID(),
        code: 'MS03',
        label: 'Delivery',
        percent: 30,
        invoiceMonth: '',
      },
      {
        id: crypto.randomUUID(),
        code: 'MS04',
        label: 'Documentation',
        percent: 10,
        invoiceMonth: '',
      },
    ],
    costs: [
      {
        id: starterCostId,
        label: 'Cost Element 1',
        totalCost: 0,
        currency: 'USD',
        convertToCurrency: 'USD',
        conversionRate: 1,
        rateIsManual: false,
      },
    ],
    progress: {
      [starterCostId]: [0],
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
      localStorage.removeItem(DEFAULT_STATE_KEY);
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
  resetDefaultsTopBtn: document.querySelector('#resetDefaultsTopBtn'),
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

function getExportBaseName() {
  const opp = String(state.project.salesforceOpportunity || '').trim();
  const rev = String(state.project.revision || '').trim();
  if (!opp) return 'cashflow-forecast';
  const parts = [opp];
  if (rev) parts.push('Rev ' + rev);
  return parts.join(' ')
    .replace(/[^a-z0-9\-_ ]/gi, '')
    .trim() || 'cashflow-forecast';
}

function saveSnapshot() {
  const safeName = getExportBaseName();

  const payload = {
    schemaVersion: SNAPSHOT_SCHEMA_VERSION,
    exportedAt: new Date().toISOString(),
    app: 'cashflow-web-app',
    state: cloneStateForSnapshot(state),
  };

  downloadJsonFile(`${safeName}-snapshot.json`, payload);
}

function loadSnapshotFromFile(file) {
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const text = String(event.target?.result || '');
      const parsed = JSON.parse(text);
      const rawState = parsed?.state || parsed;
      const normalized = normalizeImportedState(rawState);
      applyState(normalized);
      window.alert('Snapshot loaded successfully.');
    } catch (error) {
      window.alert(`Snapshot load failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  };
  reader.readAsText(file);
}

function resetToDefaults() {
  if (!window.confirm('Reset all data to default values? This cannot be undone.')) {
    return;
  }
  try {
    localStorage.removeItem('cashflow-web-app:state:v1');
    localStorage.removeItem('cashflow-web-app:state:v2');
    localStorage.removeItem('cashflow-web-app:state:v3');
    localStorage.setItem(DEFAULTS_VERSION_KEY, DEFAULTS_VERSION);
  } catch {
    // Ignore storage errors
  }
  const defaults = getEffectiveDefaultState();
  applyState(defaults);
}

function saveCurrentAsDefault() {
  try {
    const snapshot = {
      project: { ...state.project },
      milestones: state.milestones.map((milestone) => ({ ...milestone })),
      costs: state.costs.map((cost) => ({ ...cost })),
      progress: JSON.parse(JSON.stringify(state.progress || {})),
    };
    localStorage.setItem(DEFAULT_STATE_KEY, JSON.stringify(snapshot));
    window.alert('Current values saved as local defaults.');
  } catch {
    window.alert('Could not save defaults in this browser session.');
  }
}

function getHorizonMonths() {
  const roundedNet = Math.round(Number(state.project.netDays || 0) / 30);
  const baseHorizon = Math.max(1, Number(state.project.quotedLeadTimeMonths || 0) + roundedNet);
  const startMonthValue = state.project.projectStartMonth;

  if (!startMonthValue) {
    return baseHorizon;
  }

  const startDate = parseMonth(startMonthValue);
  let latestMilestoneMonth = 0;

  state.milestones.forEach((milestone) => {
    if (!milestone.invoiceMonth) {
      return;
    }

    const invoiceDate = parseMonth(milestone.invoiceMonth);
    const receivedDate = addDays(invoiceDate, Math.max(0, Math.round(clampNumber(state.project.netDays, 0))));
    const monthOffset = ((receivedDate.getFullYear() - startDate.getFullYear()) * 12)
      + (receivedDate.getMonth() - startDate.getMonth())
      + 1;

    latestMilestoneMonth = Math.max(latestMilestoneMonth, monthOffset);
  });

  return Math.max(baseHorizon, latestMilestoneMonth, 1);
}

function ensureProgressShape() {
  const horizon = getHorizonMonths();

  for (const cost of state.costs) {
    if (!Array.isArray(state.progress[cost.id])) {
      // Initialize new cost with 100% completion for all months
      state.progress[cost.id] = Array.from({ length: horizon }, () => 100);
    } else {
      // Adjust length if horizon changed
      const row = state.progress[cost.id];
      if (row.length < horizon) {
        while (row.length < horizon) {
          row.push(100);
        }
      } else if (row.length > horizon) {
        row.length = horizon;
      }
    }
  }

  const liveIds = new Set(state.costs.map((cost) => cost.id));
  for (const key of Object.keys(state.progress)) {
    if (!liveIds.has(key)) {
      delete state.progress[key];
    }
  }
}

function parseMonth(monthValue) {
  if (!monthValue || typeof monthValue !== 'string') {
    return new Date();
  }
  const [year, month] = monthValue.split('-').map(Number);
  return new Date(year, month - 1, 1);
}
function addMonths(date, offset) {
  return new Date(date.getFullYear(), date.getMonth() + offset, 1);
}

function addDays(date, offsetDays) {
  const next = new Date(date);
  next.setDate(next.getDate() + offsetDays);
  return next;
}

function monthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function formatMonthLabel(date) {
  return new Intl.DateTimeFormat('en-US', { month: 'short', year: '2-digit' }).format(date);
}

function clampNumber(value, fallback = 0) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : fallback;
  }
  const normalized = String(value ?? '').replace(/,/g, '').trim();
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function clampPercent(value) {
  return Math.min(100, Math.max(0, clampNumber(value, 0)));
}

function normalizeProgressRow(values) {
  const normalized = [];
  let running = 0;

  for (const rawValue of values) {
    running = Math.max(running, clampPercent(rawValue));
    normalized.push(running);
  }

  return normalized;
}

function formatCurrency(value, currency = state.project.convertToCurrency) {
  const code = (currency || '').trim().toUpperCase();
  const safeCode = /^[A-Z]{3}$/.test(code) ? code : 'USD';
  try {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: safeCode,
      maximumFractionDigits: 2,
    }).format(value || 0);
  } catch {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      maximumFractionDigits: 2,
    }).format(value || 0);
  }
}

function formatNumber(value) {
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  }).format(value || 0);
}

function formatNumberFixed2(value) {
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value || 0);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function computeModel() {
  ensureProgressShape();

  const horizon = getHorizonMonths();
  const startDate = parseMonth(state.project.projectStartMonth);
  const months = Array.from({ length: horizon }, (_, index) => addMonths(startDate, index));
  const keys = months.map(monthKey);
  const revenueByMonth = Array.from({ length: horizon }, () => 0);

  const milestoneRows = state.milestones.map((milestone) => {
    const revenueCostCurrency = clampNumber(state.project.contractValue) * (clampPercent(milestone.percent) / 100) * clampNumber(state.project.contractFxRate, 1);
    const paymentDate = parseMonth(milestone.invoiceMonth);
    const receivedDate = addDays(paymentDate, Math.max(0, Math.round(clampNumber(state.project.netDays, 0))));
    const receivedMonth = monthKey(receivedDate);
    const targetIndex = keys.indexOf(receivedMonth);

    if (targetIndex >= 0) {
      revenueByMonth[targetIndex] += revenueCostCurrency;
    }

    return {
      ...milestone,
      revenueCostCurrency,
      forecastMonth: receivedMonth,
      targetIndex,
    };
  });

  const cumulativeRevenue = [];
  let runningRevenue = 0;
  for (const value of revenueByMonth) {
    runningRevenue += value;
    cumulativeRevenue.push(runningRevenue);
  }

  const costRows = state.costs.map((cost) => {
    const normalizedProgress = normalizeProgressRow(state.progress[cost.id] || []);
    const rate = clampNumber(cost.conversionRate, 1) || 1;
    const convertedTotal = clampNumber(cost.totalCost) * rate;
    const monthlyCost = [];
    let priorCost = 0;

    normalizedProgress.forEach((progressValue) => {
      const cumulativeCost = convertedTotal * (progressValue / 100);
      const monthCost = Math.max(0, cumulativeCost - priorCost);
      monthlyCost.push(monthCost);
      priorCost = cumulativeCost;
    });

    return {
      ...cost,
      convertedTotal,
      normalizedProgress,
      monthlyCost,
    };
  });

  const monthlyCost = Array.from({ length: horizon }, (_, index) => {
    return costRows.reduce((sum, row) => sum + (row.monthlyCost[index] || 0), 0);
  });

  const cumulativeCost = [];
  let runningCost = 0;
  for (const value of monthlyCost) {
    runningCost += value;
    cumulativeCost.push(runningCost);
  }

  const netCashFlow = revenueByMonth.map((value, index) => value - monthlyCost[index]);
  const cumulativeNet = [];
  let runningNet = 0;
  for (const value of netCashFlow) {
    runningNet += value;
    cumulativeNet.push(runningNet);
  }

  const totals = {
    totalRevenue: revenueByMonth.reduce((sum, value) => sum + value, 0),
    totalCost: monthlyCost.reduce((sum, value) => sum + value, 0),
    finalNet: netCashFlow.reduce((sum, value) => sum + value, 0),
    peakFundingNeed: Math.min(...cumulativeNet, 0),
    milestonePercentTotal: milestoneRows.reduce((sum, row) => sum + clampPercent(row.percent), 0),
  };

  return {
    months,
    milestoneRows,
    costRows,
    revenueByMonth,
    cumulativeRevenue,
    monthlyCost,
    cumulativeCost,
    netCashFlow,
    cumulativeNet,
    totals,
  };
}

function renderProjectForm(model) {
  if (!dom.projectForm) return;
  dom.projectForm.innerHTML = '';

  // Ã¢â€â‚¬Ã¢â€â‚¬ Contract row (table-style, mirrors cost elements) Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
  const contractFxRate = state.project.contractFxRate ?? 1;
  const contractRateIsManual = state.project.contractRateIsManual;
  const sameCurrency = (state.project.contractCurrency || '').toUpperCase() === (state.project.convertToCurrency || '').toUpperCase();
  const fxBadge = `<span class="rate-badge auto" title="Live rate from open.er-api.com">auto avg: ${Number(contractFxRate).toFixed(6).replace(/\.?0+$/, '')}</span>`;
  const currencySelect = (fieldName, selected) => {
    const options = CURRENCY_OPTIONS.map((curr) => `<option value="${curr}" ${curr === selected ? 'selected' : ''}>${curr}</option>`).join('');
    return `<select class="cell-input" data-kind="project" data-field="${fieldName}" style="max-width:110px;">${options}</select>`;
  };

  const opportunityFields = document.createElement('div');
  opportunityFields.className = 'project-meta-grid';
  opportunityFields.innerHTML = `
    <div class="timeline-field">
      <label class="inline-note" style="display:block;margin-bottom:6px;">Opportunity Number</label>
      <input class="cell-input" type="text" data-kind="project" data-field="salesforceOpportunity" value="${state.project.salesforceOpportunity || ''}" placeholder="Opp Number">
    </div>
    <div class="timeline-field revision-field">
      <label class="inline-note" style="display:block;margin-bottom:6px;">Revision</label>
      <input class="cell-input" type="text" data-kind="project" data-field="revision" value="${state.project.revision || ''}" placeholder="e.g. Rev A">
    </div>
    <div class="timeline-field">
      <label class="inline-note" style="display:block;margin-bottom:6px;">Quoted lead time (months)</label>
      <input class="cell-input" type="text" data-kind="project" data-field="quotedLeadTimeMonths" value="${state.project.quotedLeadTimeMonths}">
    </div>
    <div class="timeline-field">
      <label class="inline-note" style="display:block;margin-bottom:6px;">Opportunity Name</label>
      <input class="cell-input" type="text" data-kind="project" data-field="opportunityName" value="${state.project.opportunityName || ''}" placeholder="Opportunity Name">
    </div>
    <div class="timeline-field" style="grid-column: 2;">
    </div>
    <div class="timeline-field">
      <label class="inline-note" style="display:block;margin-bottom:6px;">Project start month</label>
      <input class="cell-input" type="month" data-kind="project" data-field="projectStartMonth" value="${state.project.projectStartMonth}">
    </div>
  `;

  const contractTable = document.createElement('div');
  contractTable.className = 'table-scroll';
  contractTable.innerHTML = `
    <table class="data-table contract-table">
      <thead>
        <tr>
          <th>Contract value</th>
          <th>Contract currency</th>
          <th>Convert to</th>
          <th>FX Rate</th>
          <th>Converted value</th>
          <th>Margin</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td class="number-cell"><input class="cell-input" type="text" data-kind="project" data-field="contractValue" value="${formatNumberFixed2(clampNumber(state.project.contractValue, 0))}"></td>
          <td>${currencySelect('contractCurrency', (state.project.contractCurrency || 'USD').toUpperCase())}</td>
          <td>${currencySelect('convertToCurrency', (state.project.convertToCurrency || 'USD').toUpperCase())}</td>
          <td style="white-space:nowrap;">
            <input class="cell-input" type="text" data-kind="project" data-field="contractFxRate" value="${contractFxRate}" style="max-width:120px;">
            ${fxBadge}
            <button class="ghost-button" style="padding:4px 8px;font-size:0.78rem;" type="button" data-kind="reset-contract-fx-rate" title="Reset to 3-month average">&#x21ba;</button>
          </td>
          <td class="currency-cell ${sameCurrency ? 'muted-cell' : 'converted-value'}">${formatCurrency(clampNumber(state.project.contractValue) * clampNumber(contractFxRate, 1), state.project.convertToCurrency)}</td>
          <td class="number-cell" id="calculatedMarginCell" style="font-weight:600;">—</td>
        </tr>
      </tbody>
    </table>
  `;
  dom.projectForm.appendChild(opportunityFields);
  dom.projectForm.appendChild(contractTable);
}

function renderHealth(model) {
  if (!dom.healthChecks) return;
  const cards = [];
  const milestoneGap = Math.abs(model.totals.milestonePercentTotal - 100);

  // Calculate Risk %
  let riskScore = 0;
  // Risk from milestone allocation gap (0-30%)
  riskScore += Math.min(30, milestoneGap * 0.3);
  // Risk from negative cash flow (0-40%)
  if (model.totals.peakFundingNeed < 0) {
    const fundingRisk = Math.abs(model.totals.peakFundingNeed) / Math.abs(model.totals.totalRevenue || 1);
    riskScore += Math.min(40, fundingRisk * 100);
  }
  // Risk from timeline mismatch (0-30%)
  if (model.milestoneRows.some((row) => row.targetIndex === -1)) {
    riskScore += 30;
  }
  const riskPercent = Math.min(100, Math.round(riskScore));
  const riskLevel = riskPercent <= 20 ? 'Low' : riskPercent <= 50 ? 'Med' : 'High';

  cards.push({
    tone: riskPercent <= 20 ? 'good' : riskPercent <= 50 ? 'warn' : 'bad',
    title: 'Risk',
    text: `${riskLevel} - Based on allocation, funding, and timeline factors`,
  });

  // Check Cost Integrity
  const allCostsTracked = state.costs.length > 0 && state.costs.every((cost) => {
    const progressRow = state.progress[cost.id];
    return Array.isArray(progressRow) && progressRow.length > 0;
  });
  const costAlertText = state.costs.length === 0
    ? 'No cost elements defined.'
    : allCostsTracked
      ? `All ${state.costs.length} cost element(s) are properly tracked in the schedule.`
      : 'Some cost elements are missing progress data.';

  cards.push({
    tone: state.costs.length === 0 ? 'warn' : allCostsTracked ? 'good' : 'bad',
    title: 'Cost Integrity',
    text: costAlertText,
  });

  cards.push({
    tone: milestoneGap < 0.01 ? 'good' : 'bad',
    title: 'Milestone allocation',
    text: milestoneGap < 0.01
      ? 'Milestone percentages sum to 100%.'
      : `Milestone percentages sum to ${formatNumber(model.totals.milestonePercentTotal)}%.`,
  });

  cards.push({
    tone: model.totals.peakFundingNeed < 0 ? 'warn' : 'good',
    title: 'Funding need',
    text: model.totals.peakFundingNeed < 0
      ? `Peak negative cumulative cash is ${formatCurrency(model.totals.peakFundingNeed)}.`
      : 'The model never goes negative on a cumulative basis.',
  });

  const timelineCoverage = model.milestoneRows.some((row) => row.targetIndex === -1)
    ? 'Some milestone dates fall outside the visible timeline.'
    : 'All milestone dates land inside the visible timeline.';

  cards.push({
    tone: model.milestoneRows.some((row) => row.targetIndex === -1) ? 'warn' : 'good',
    title: 'Timeline coverage',
    text: timelineCoverage,
  });

  dom.healthChecks.innerHTML = cards.map((card) => `
    <article class="health-card ${card.tone}">
      <strong>${card.title}</strong>
      <span>${card.text}</span>
    </article>
  `).join('');
}

function renderMilestones(model) {
  // Render Net Days inside the milestone header to keep the section compact.
  if (dom.milestonesBody) {
    let paymentTermsField = document.querySelector('#paymentTermsField');
    if (!paymentTermsField) {
      paymentTermsField = document.createElement('div');
      paymentTermsField.id = 'paymentTermsField';
      paymentTermsField.className = 'field payment-terms-field';
      paymentTermsField.innerHTML = `
        <label for="netDays">Net Days</label>
        <input id="netDays" name="netDays" type="text" class="cell-input" data-kind="project" data-field="netDays">
      `;
      const section = dom.milestonesBody.closest('section');
      if (section) {
        const actions = section.querySelector('.milestone-actions');
        if (actions) {
          actions.prepend(paymentTermsField);
        }
      }
      const netDaysInput = paymentTermsField.querySelector('input');
      if (netDaysInput) {
        netDaysInput.addEventListener('change', (event) => {
          state.project.netDays = event.target.value;
          rerender();
        });
      }
    }
    const netDaysInput = paymentTermsField.querySelector('input');
    if (netDaysInput) {
      netDaysInput.value = state.project.netDays;
    }
  }

  if (!dom.milestonesBody) return;
  dom.milestonesBody.innerHTML = '';

  model.milestoneRows.forEach((milestone) => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td><input class="cell-input mono" data-kind="milestone" data-id="${milestone.id}" data-field="code" value="${milestone.code}"></td>
      <td><input class="cell-input" data-kind="milestone" data-id="${milestone.id}" data-field="label" value="${milestone.label}"></td>
      <td class="percent-cell"><input class="cell-input" type="text" data-kind="milestone" data-id="${milestone.id}" data-field="percent" value="${milestone.percent}"></td>
      <td><input class="cell-input" type="month" data-kind="milestone" data-id="${milestone.id}" data-field="invoiceMonth" value="${milestone.invoiceMonth}"></td>
      <td class="currency-cell converted-value">${formatCurrency(milestone.revenueCostCurrency, state.project.convertToCurrency)}</td>
      <td><button class="remove-button" type="button" data-kind="remove-milestone" data-id="${milestone.id}">Remove</button></td>
    `;
    dom.milestonesBody.appendChild(row);
  });
}

function renderCashflowEstimate(model) {
  if (!dom.cashflowEstimate) return;
  const contractValueConverted = clampNumber(state.project.contractValue) * clampNumber(state.project.contractFxRate, 1);
  const currency = state.project.convertToCurrency || 'USD';
  const horizon = model.months.length;

  // Build month headers with numbers and date labels
  const monthHeaders = model.months.map((month, index) => {
    const label = formatMonthLabel(month);
    return `<th><div>M${index + 1}</div><div class="inline-note">${label}</div></th>`;
  }).join('');

  // Build milestone rows
  const milestoneRows = state.milestones.map((milestone) => {
    const targetIndex = model.milestoneRows.find(mr => mr.id === milestone.id)?.targetIndex ?? -1;
    const cells = Array.from({ length: horizon }, (_, i) => {
      if (i === targetIndex) {
        const revenue = model.milestoneRows.find(mr => mr.id === milestone.id)?.revenueCostCurrency ?? 0;
        return `<td class="number-cell">${formatNumberFixed2(revenue)}</td>`;
      }
      return `<td></td>`;
    }).join('');

    return `
      <tr>
        <td><strong>${milestone.code}</strong></td>
        <td>${milestone.label}</td>
        ${cells}
      </tr>
    `;
  }).join('');

  // Revenue total row
  const revenueTotalCells = model.revenueByMonth.map((value) => {
    return `<td class="number-cell">${formatNumberFixed2(value)}</td>`;
  }).join('');

  // Accrued/cumulative row
  const accruedCells = model.cumulativeRevenue.map((value) => {
    return `<td class="number-cell">${formatNumberFixed2(value)}</td>`;
  }).join('');

  const html = `
    <div class="estimate-header">
      <div class="estimate-info">
        <div><strong>Contract Value:</strong> ${formatCurrency(contractValueConverted)}</div>
      </div>
    </div>
    <table class="estimate-table">
      <thead>
        <tr>
          <th>Code</th>
          <th>Activity</th>
          ${monthHeaders}
        </tr>
      </thead>
      <tbody>
        ${milestoneRows}
        <tr class="total-row">
          <td colspan="2"><strong>Revenue</strong></td>
          ${revenueTotalCells}
        </tr>
        <tr class="accrued-row">
          <td colspan="2"><strong>Accrued</strong></td>
          ${accruedCells}
        </tr>
      </tbody>
    </table>
  `;

  dom.cashflowEstimate.innerHTML = html;
}

function renderCosts() {
  if (!dom.costsBody) return;
  dom.costsBody.innerHTML = '';

  const currencyDropdown = (fieldName, selected) => {
    const options = CURRENCY_OPTIONS.map(curr => `<option value="${curr}" ${curr === selected ? 'selected' : ''}>${curr}</option>`).join('');
    return `<select class="cell-input" data-kind="cost" data-field="${fieldName}" style="max-width:90px;">${options}</select>`;
  };

  state.costs.forEach((cost) => {
    const currency = cost.currency || state.project.contractCurrency || 'USD';
    const convertToCurrency = cost.convertToCurrency || state.project.convertToCurrency || 'USD';
    const conversionRate = cost.conversionRate ?? 1;
    const convertedTotal = clampNumber(cost.totalCost) * (clampNumber(conversionRate, 1) || 1);
    const needsConversion = currency.toUpperCase() !== convertToCurrency.toUpperCase();
    const rateLabel = cost.rateIsManual
      ? '<span class="rate-badge manual" title="Manually set">manual</span>'
      : `<span class="rate-badge auto" title="Live rate from open.er-api.com">auto avg: ${conversionRate.toFixed !== undefined ? Number(conversionRate).toFixed(6).replace(/\.?0+$/, '') : conversionRate}</span>`;
    const row = document.createElement('tr');
    row.innerHTML = `
      <td><input class="cell-input" data-kind="cost" data-id="${cost.id}" data-field="label" value="${cost.label}"></td>
      <td class="number-cell"><input class="cell-input" type="text" data-kind="cost" data-id="${cost.id}" data-field="totalCost" value="${formatNumberFixed2(clampNumber(cost.totalCost, 0))}"></td>
      <td>${currencyDropdown('currency', currency).replace('<select', `<select data-id="${cost.id}"`)}</td>
      <td>${currencyDropdown('convertToCurrency', convertToCurrency).replace('<select', `<select data-id="${cost.id}"`)}</td>
      <td style="white-space:nowrap;">
        <input class="cell-input" type="text" data-kind="cost" data-id="${cost.id}" data-field="conversionRate" value="${conversionRate}" style="max-width:120px;">
        ${rateLabel}
        <button class="ghost-button" style="padding:4px 8px;font-size:0.78rem;" type="button" data-kind="reset-fx-rate" data-id="${cost.id}" title="Reset to 3-month average">&#x21ba;</button>
      </td>
      <td class="currency-cell ${needsConversion ? 'converted-value' : 'muted-cell'}">${formatCurrency(convertedTotal, convertToCurrency)}</td>
      <td><button class="remove-button" type="button" data-kind="remove-cost" data-id="${cost.id}">Remove</button></td>
    `;
    dom.costsBody.appendChild(row);
  });
}

function renderProgressGrid(model) {
  if (!dom.progressGrid) return;
  const headMonths = model.months.map((month, index) => `<th>M${index + 1}<div class="inline-note">${formatMonthLabel(month)}</div></th>`).join('');
  const rows = model.costRows.map((costRow) => {
    const cells = costRow.normalizedProgress.map((value, index) => `
      <td>
        <input
          class="cell-input"
          type="text"
          data-kind="progress"
          data-id="${costRow.id}"
          data-index="${index}"
          value="${formatNumber(value)}"
        >
      </td>
    `).join('');

    return `
      <tr>
        <td>
          <span>${costRow.label}</span>
          <div class="inline-note">${formatCurrency(costRow.convertedTotal, costRow.convertToCurrency || state.project.convertToCurrency)}</div>
        </td>
        ${cells}
      </tr>
    `;
  }).join('');

  dom.progressGrid.innerHTML = `
    <table class="progress-table">
      <thead>
        <tr>
          <th>Cost element</th>
          ${headMonths}
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function getCalculatedMargin(model) {
  const contractValueConverted = clampNumber(state.project.contractValue) * clampNumber(state.project.contractFxRate, 1);
  const totalCost = model.totals.totalCost;
  if (contractValueConverted === 0) return { pct: 0, text: '0.0%' };
  const pct = ((contractValueConverted - totalCost) / contractValueConverted) * 100;
  return { pct, text: `${pct.toFixed(1)}%` };
}

function getSummaryItems(model) {
  const contractValueConverted = clampNumber(state.project.contractValue) * clampNumber(state.project.contractFxRate, 1);
  const margin = getCalculatedMargin(model);
  const horizon = getHorizonMonths();

  return [
    { label: 'Contract Value', value: formatCurrency(contractValueConverted), note: '' },
    { label: 'Margin', value: margin.text, note: '' },
    { label: 'Cost Total', value: formatCurrency(model.totals.totalCost), note: 'All cost elements' },
    { label: 'Timeline', value: `${horizon} months`, note: 'Quoted lead time plus net days' },
    { label: 'Estimate Net Value', value: formatCurrency(model.totals.finalNet), note: 'Total project net position' },
     { label: 'Peak funding need', value: formatCurrency(model.totals.peakFundingNeed), note: 'Lowest point of cumulative net cash' },
  ];
}

function renderSummary(model) {
  if (!dom.summaryList) return;
  const summaryItems = getSummaryItems(model);

  dom.summaryList.innerHTML = summaryItems.map((item) => `
    <article class="summary-card">
      <span>${item.label}</span>
      <strong class="${item.value.startsWith('-') ? 'negative' : 'positive'}">${item.value}</strong>
      <p class="inline-note">${item.note}</p>
    </article>
  `).join('');
}

function renderForecastTable(model) {
  if (!dom.forecastHead || !dom.forecastBody) return;
  dom.forecastHead.innerHTML = `
    <tr>
      <th>Series</th>
      ${model.months.map((month) => `<th>${formatMonthLabel(month)}</th>`).join('')}
    </tr>
  `;

  const rows = [
    { label: 'Forecast Revenue', values: model.revenueByMonth },
    { label: 'Forecast Cost', values: model.monthlyCost },
    { label: 'Sum Cashflow', values: model.cumulativeNet },
  ];

  dom.forecastBody.innerHTML = rows.map((row) => `
    <tr>
      <td><strong>${row.label}</strong></td>
      ${row.values.map((value) => `<td class="${value < 0 ? 'negative' : ''}">${formatNumberFixed2(value)}</td>`).join('')}
    </tr>
  `).join('');
}

function renderChartHeader() {
  if (dom.chartOpportunity) {
    const opportunityValue = String(state.project.salesforceOpportunity || '').trim();
    const revisionValue = String(state.project.revision || '').trim();
    const revisionText = revisionValue ? `,     Revision: ${revisionValue}` : '';
    dom.chartOpportunity.textContent = `Opportunity Number: ${opportunityValue || 'Not provided'}${revisionText}`;
  }

  if (dom.chartRevision) {
    dom.chartRevision.textContent = '';
    dom.chartRevision.style.display = 'none';
  }

  if (dom.chartOpportunityName) {
    const opportunityName = String(state.project.opportunityName || '').trim();
    dom.chartOpportunityName.textContent = `Opportunity Name: ${opportunityName || 'Not provided'}`;
  }

  if (dom.chartDate) {
    const todayLabel = new Date().toLocaleDateString('en-US', {
      month: 'long',
      day: 'numeric',
      year: 'numeric',
    });
    dom.chartDate.textContent = todayLabel;
  }
}

function renderChart(model) {
  if (!dom.chart) return;
  if (dom.chartLegend) {
    dom.chartLegend.innerHTML = '';
  }

  const svg = dom.chart;
  const width = 880;
  const height = 450;
  const padding = { top: 28, right: 32, bottom: 96, left: 98 };
  const plotWidth = width - padding.left - padding.right;
  const plotHeight = height - padding.top - padding.bottom;

  const combinedValues = [...model.cumulativeNet, 0];
  const minValue = Math.min(...combinedValues);
  const maxValue = Math.max(...combinedValues);
  const valueRange = maxValue - minValue || 1;
  const stepX = plotWidth / Math.max(model.months.length, 1);
  const zeroY = padding.top + ((maxValue - 0) / valueRange) * plotHeight;
  const axisLabelX = 18;

  const yForValue = (value) => padding.top + ((maxValue - value) / valueRange) * plotHeight;

  const gridLines = Array.from({ length: 5 }, (_, index) => {
    const ratio = index / 4;
    const value = maxValue - ratio * valueRange;
    const y = padding.top + ratio * plotHeight;
    return `
      <g>
        <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="rgba(23,33,43,0.12)" stroke-dasharray="5 7" />
        <text x="${padding.left - 14}" y="${y + 4}" text-anchor="end" font-size="12" fill="#3f596a">${formatNumber(value)}</text>
      </g>
    `;
  }).join('');

  const cumulativePath = model.cumulativeNet.map((value, index) => {
    const x = padding.left + index * stepX + stepX / 2;
    const y = yForValue(value);
    return `${index === 0 ? 'M' : 'L'} ${x} ${y}`;
  }).join(' ');

  const lineDots = model.cumulativeNet.map((value, index) => {
    const x = padding.left + index * stepX + stepX / 2;
    const y = yForValue(value);
    return `<circle cx="${x}" cy="${y}" r="4.5" fill="#1F8AB8"></circle>`;
  }).join('');

  const xLabels = model.months.map((month, index) => {
    const x = padding.left + index * stepX + stepX / 2;
    return `<text x="${x}" y="${height - 40}" text-anchor="middle" font-size="12" fill="#314f63">${formatMonthLabel(month)}</text>`;
  }).join('');

  svg.innerHTML = `
    <rect x="0" y="0" width="${width}" height="${height}" fill="transparent"></rect>
    ${gridLines}
    <line x1="${padding.left}" y1="${zeroY}" x2="${width - padding.right}" y2="${zeroY}" stroke="rgba(23,33,43,0.24)" />
    <path d="${cumulativePath}" fill="none" stroke="#1F8AB8" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"></path>
    ${lineDots}
    ${xLabels}
    <text x="${axisLabelX}" y="${height - 14}" text-anchor="start" fill="#314f63" font-size="13" font-weight="600">Currency (${state.project.convertToCurrency || 'USD'})</text>
  `;
}

function svgElementToPngDataUrl(svgElement) {
  return new Promise((resolve, reject) => {
    if (!(svgElement instanceof SVGElement)) {
      reject(new Error('Chart not available.'));
      return;
    }

    const serializer = new XMLSerializer();
    let markup = serializer.serializeToString(svgElement);
    if (!markup.includes('xmlns="http://www.w3.org/2000/svg"')) {
      markup = markup.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
    }

    const blob = new Blob([markup], { type: 'image/svg+xml;charset=utf-8' });
    const blobUrl = URL.createObjectURL(blob);
    const image = new Image();

    image.onload = () => {
      const viewBox = svgElement.viewBox.baseVal;
      const width = viewBox && viewBox.width ? Math.round(viewBox.width) : 880;
      const height = viewBox && viewBox.height ? Math.round(viewBox.height) : 420;
      const canvas = document.createElement('canvas');
      canvas.width = width * 2;
      canvas.height = height * 2;
      const context = canvas.getContext('2d');

      if (!context) {
        URL.revokeObjectURL(blobUrl);
        reject(new Error('Canvas export is not available in this browser.'));
        return;
      }

      context.fillStyle = '#eaf6ff';
      context.fillRect(0, 0, canvas.width, canvas.height);
      context.drawImage(image, 0, 0, canvas.width, canvas.height);
      URL.revokeObjectURL(blobUrl);
      resolve(canvas.toDataURL('image/png'));
    };

    image.onerror = () => {
      URL.revokeObjectURL(blobUrl);
      reject(new Error('Unable to render the chart for export.'));
    };

    image.src = blobUrl;
  });
}

function exportToExcel() {
  if (typeof XLSX === 'undefined') {
    window.alert('Excel export is not available because the SheetJS library did not load.');
    return;
  }

  const model = computeModel();
  const horizon = model.months.length;
  const currency = state.project.convertToCurrency || 'USD';
  const monthLabels = model.months.map(formatMonthLabel);
  const wb = XLSX.utils.book_new();

  // Ã¢â€â‚¬Ã¢â€â‚¬ Sheet 1: Inputs Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
  const inputsRows = [
    ['OPPORTUNITY INFORMATION'],
    ['Opportunity Number', state.project.salesforceOpportunity || ''],
    ['Revision', state.project.revision || ''],
    ['Opportunity Name',       state.project.opportunityName || ''],
    ['Contract Value',         clampNumber(state.project.contractValue)],
    ['Contract Currency',      state.project.contractCurrency || ''],
    ['Convert To Currency',    currency],
    ['FX Rate',                clampNumber(state.project.contractFxRate, 1)],
    ['Project Start Month',    state.project.projectStartMonth || ''],
    ['Quoted Lead Time (months)', clampNumber(state.project.quotedLeadTimeMonths)],
    ['Net Days',               clampNumber(state.project.netDays)],
    [],
    ['MILESTONES'],
    ['Code', 'Label', 'Percent (%)', 'Invoice Month', `Revenue (${currency})`],
    ...state.milestones.map((ms) => {
      const rev = clampNumber(state.project.contractValue)
        * (clampPercent(ms.percent) / 100)
        * clampNumber(state.project.contractFxRate, 1);
      return [ms.code, ms.label, clampPercent(ms.percent), ms.invoiceMonth, rev];
    }),
    [],
    ['COST ELEMENTS'],
    ['Label', 'Total Cost', 'Currency', 'Convert To', 'FX Rate', `Converted Cost (${currency})`],
    ...state.costs.map((cost) => [
      cost.label,
      clampNumber(cost.totalCost),
      cost.currency || '',
      cost.convertToCurrency || '',
      clampNumber(cost.conversionRate, 1),
      clampNumber(cost.totalCost) * clampNumber(cost.conversionRate, 1),
    ]),
    [],
    ['PLANNED EXPENDITURES'],
    [`Cost Element (${currency})`, 'Total', ...monthLabels],
    ...model.costRows.map((costRow) => [
      costRow.label,
      costRow.convertedTotal,
      ...costRow.monthlyCost,
    ]),
    [],
    [
      'Total Monthly Cost',
      model.totals.totalCost,
      ...model.monthlyCost,
    ],
  ];

  const wsInputs = XLSX.utils.aoa_to_sheet(inputsRows);
  wsInputs['!cols'] = [{ wch: 30 }, { wch: 24 }, { wch: 14 }, { wch: 16 }, { wch: 24 }, ...Array(horizon).fill({ wch: 14 })];
  
  // Apply consistent font size to all cells (11pt)
  Object.keys(wsInputs).forEach((cell) => {
    if (cell.startsWith('!')) return;
    if (!wsInputs[cell]) wsInputs[cell] = {};
    wsInputs[cell].font = { size: 11 };
  });
  
  XLSX.utils.book_append_sheet(wb, wsInputs, 'Inputs');

  // Ã¢â€â‚¬Ã¢â€â‚¬ Sheet 2: Cashflow Forecast Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
  // Build Cashflow Estimate section (milestone rows + revenue + accrued)
  const estimateAoa = [
    ['CASHFLOW ESTIMATE'],
    ['Code', 'Activity', ...monthLabels],
  ];

  // Add milestone rows with revenue in their target months
  state.milestones.forEach((milestone) => {
    const milestoneRow = model.milestoneRows.find(mr => mr.id === milestone.id);
    const targetIndex = milestoneRow?.targetIndex ?? -1;
    const revenueCostCurrency = milestoneRow?.revenueCostCurrency ?? 0;

    const monthlyRevenue = Array.from({ length: horizon }, (_, i) => {
      return i === targetIndex ? revenueCostCurrency : '';
    });

    estimateAoa.push([milestone.code, milestone.label, ...monthlyRevenue]);
  });

  // Add Revenue total row
  estimateAoa.push(['', 'Revenue', ...model.revenueByMonth]);

  // Add Accrued row with cumulative revenue
  estimateAoa.push(['', 'Accrued', ...model.cumulativeRevenue]);

  // Combine Cashflow Estimate and Forecast data with spacing
  const forecastAoa = [
    [],
    ['FORECAST'],
    [`Series (${currency})`, ...monthLabels],
    ['Forecast Revenue',          ...model.revenueByMonth],
    ['Forecast Cost',             ...model.monthlyCost],
    ['Sum Cashflow',              ...model.cumulativeNet],
  ];

  const combinedAoa = [...estimateAoa, ...forecastAoa];

  const wsForecast = XLSX.utils.aoa_to_sheet(combinedAoa);
  wsForecast['!cols'] = [{ wch: 28 }, ...Array(horizon).fill({ wch: 14 })];
  
  // Apply consistent font size to all cells (11pt)
  Object.keys(wsForecast).forEach((cell) => {
    if (cell.startsWith('!')) return;
    if (!wsForecast[cell]) wsForecast[cell] = {};
    wsForecast[cell].font = { size: 11 };
  });

  XLSX.utils.book_append_sheet(wb, wsForecast, 'Cashflow Forecast');

  const safeName = getExportBaseName();
  XLSX.writeFile(wb, `${safeName}.xlsx`);
}

function importFromExcel(file) {
  if (typeof XLSX === 'undefined') {
    window.alert('Excel import is not available because the SheetJS library did not load.');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets['Inputs'];
      if (!ws) {
        window.alert('Could not find an "Inputs" sheet. Please use a file exported from this tool.');
        return;
      }

      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      // Build a lookup map: label -> value for single-value rows
      const lookup = {};
      rows.forEach((row) => {
        if (row[0] && row[1] !== undefined) lookup[String(row[0]).trim()] = row[1];
      });

      // Ã¢â€â‚¬Ã¢â€â‚¬ Opportunity Information Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
      const newProject = { ...state.project };
      if (lookup['Opportunity Number'] !== undefined) newProject.salesforceOpportunity = String(lookup['Opportunity Number']);
      if (lookup['Revision'] !== undefined) newProject.revision = String(lookup['Revision']);
      if (lookup['Opportunity Name'] !== undefined) newProject.opportunityName = String(lookup['Opportunity Name']);
      if (lookup['Contract Value'] !== undefined) newProject.contractValue = Number(lookup['Contract Value']) || 0;
      if (lookup['Contract Currency']) newProject.contractCurrency = String(lookup['Contract Currency']);
      if (lookup['Convert To Currency']) newProject.convertToCurrency = String(lookup['Convert To Currency']);
      if (lookup['FX Rate'] !== undefined) newProject.contractFxRate = Number(lookup['FX Rate']) || 1;
      if (lookup['Project Start Month']) newProject.projectStartMonth = String(lookup['Project Start Month']);
      if (lookup['Quoted Lead Time (months)'] !== undefined) newProject.quotedLeadTimeMonths = Number(lookup['Quoted Lead Time (months)']) || 0;
      if (lookup['Net Days'] !== undefined) newProject.netDays = Number(lookup['Net Days']) || 0;

      // Ã¢â€â‚¬Ã¢â€â‚¬ Find section start rows Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
      let msHeaderRow = -1, costHeaderRow = -1, plannedHeaderRow = -1;
      rows.forEach((row, i) => {
        const cell = String(row[0] || '').trim();
        if (cell === 'MILESTONES') msHeaderRow = i;
        if (cell === 'COST ELEMENTS') costHeaderRow = i;
        if (cell === 'PLANNED EXPENDITURES') plannedHeaderRow = i;
      });

      // Ã¢â€â‚¬Ã¢â€â‚¬ Milestones Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
      const newMilestones = [];
      if (msHeaderRow >= 0) {
        // Skip the section title row and the header row (Code, Label, ...)
        let i = msHeaderRow + 2;
        while (i < rows.length && rows[i][0] !== '' && rows[i][0] !== undefined) {
          const [code, label, percent, invoiceMonth] = rows[i];
          if (code || label) {
            newMilestones.push({
              id: crypto.randomUUID(),
              code: String(code || ''),
              label: String(label || ''),
              percent: Number(percent) || 0,
              invoiceMonth: String(invoiceMonth || ''),
            });
          }
          i++;
        }
      }

      // Ã¢â€â‚¬Ã¢â€â‚¬ Cost Elements Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
      const newCosts = [];
      if (costHeaderRow >= 0) {
        let i = costHeaderRow + 2;
        while (i < rows.length && rows[i][0] !== '' && rows[i][0] !== undefined) {
          const [label, totalCost, currency, convertToCurrency, fxRate] = rows[i];
          if (label) {
            newCosts.push({
              id: crypto.randomUUID(),
              label: String(label),
              totalCost: Number(totalCost) || 0,
              currency: String(currency || 'USD'),
              convertToCurrency: String(convertToCurrency || 'USD'),
              conversionRate: Number(fxRate) || 1,
              rateIsManual: false,
            });
          }
          i++;
        }
      }

      // Ã¢â€â‚¬Ã¢â€â‚¬ Planned Expenditures (progress %) Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
      // The exported rows are monthly *cost amounts*, not percentages.
      // We derive cumulative progress % from each row relative to total cost.
      const newProgress = {};
      if (plannedHeaderRow >= 0 && newCosts.length) {
        const monthCount = getHorizonMonths();
        let i = plannedHeaderRow + 2; // skip section title + header row
        newCosts.forEach((cost) => {
          if (i >= rows.length) return;
          const dataRow = rows[i];
          // dataRow[0] = label, dataRow[1] = total, dataRow[2..] = monthly costs
          const monthlyCosts = dataRow.slice(2, 2 + monthCount).map((v) => Number(v) || 0);
          const total = Number(dataRow[1]) || 0;
          if (total > 0) {
            let cumulative = 0;
            newProgress[cost.id] = monthlyCosts.map((v) => {
              cumulative += v;
              return Math.round((cumulative / total) * 100 * 100) / 100;
            });
          } else {
            newProgress[cost.id] = Array.from({ length: monthCount }, () => 0);
          }
          i++;
        });
      }

      // Ã¢â€â‚¬Ã¢â€â‚¬ Apply to state Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
      if (!window.confirm('Import will replace all current data. Continue?')) return;

      state.project = newProject;
      state.milestones = newMilestones.length ? newMilestones : state.milestones;
      state.costs = newCosts.length ? newCosts : state.costs;
      state.progress = newProgress;

      persistState();
      rerender();
      // Auto-fetch fresh FX rates for all cost rows and contract rate
      autoFetchContractRate();
      state.costs.forEach((cost) => autoFetchRateForCost(cost.id));
      window.alert('Import successful!');
    } catch (err) {
      window.alert(`Import failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  };
  reader.readAsArrayBuffer(file);
}

async function exportToPowerPoint() {
  if (typeof window.PptxGenJS === 'undefined') {
    window.alert('PowerPoint export is not available because the PPT library did not load.');
    return;
  }

  const button = dom.exportPptBtn;
  if (button) {
    button.disabled = true;
    button.textContent = 'Exporting...';
  }

  try {
    const model = computeModel();
    const summaryItems = getSummaryItems(model);
    const chartImage = await svgElementToPngDataUrl(dom.chart);
    const pptx = new window.PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'GitHub Copilot';
    pptx.company = 'HMH';
    pptx.subject = 'Cashflow forecast';
    pptx.title = 'Cashflow Forecast';

    const slide = pptx.addSlide();
    slide.background = { color: 'F0F5F7' };

    // HMH logo embedded as base64
    const logoDataUrl = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCAIwBdwDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAopK5bx98SvDnw10wX3iHU4bCJsiJDlpJWGOEQct1HTpkZq4QlUkowV2zGrWp0IOpVkkl1Z1NFfGnjX9vK8kZovCugRwx5wLrVGLMw/65oQB/30fpXlGoftZfFG/m3jxL9mTtFBZwqo/NSfzNfRUuH8ZUV2lH1Z8FiuOcpw7cYOU/RafjY/SLNLX5sW37VnxQt5g//CUySY6rJawMD+afyr0zwT+3ZrdnLHF4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xcEcM/WZrM8XH3F8KfV9/RdB1FItLXzp++BRRRSKCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAK97/Ys8aHw18Z7fTZJCttrVvJakdvMUeYh/8AHSv/AAOvAzW14J19vCvjHRNZVmUWF7DcNtOPlVwx/MAj8a2ozdOpFo8jOMIsfl9bDPeUXb16fifrvRUcMiyQoyncrAEEd6kr7U/jbrYKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigApKWkLAAk9KAKupahbaTYXF3dTJb2sEbSySyNtVFAyST2AFfm7+0l8e5/jR4s8izLxeGdNYraQt/wAtm5BmYepHQHoD2y1ej/tiftEL4mvLnwL4du/+JZaybdTuYW/18q/8sgf7qnr6sMduflZa8DGYpS/dQ+Z+/cE8MfVoRzPFx99/Cuy7+r/ISiiivF0P2IKKKKkAooooAKKKKACiiigAooooAKKKKACiiigAooooAlXtU/c1AvarC/epnn1ida+3f2F/Fz6h4Q1rw9NIGbTrhbiEHg+XKCCAPQMjH/gdfES179+xf4h/sj4xx2LH5dTs5oAM/wASjzB+iH8678E+WtFn53xZhVicsq6ax95fLf8AA+/qKKK+rP5zCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA+CP26P+Svad/2B4v/AEfNXztX0T+3R/yV7Tv+wPF/6Pmr52r9gyn/AHKl6H8qcTf8jfEf4gooor2D5cKKKKACiiigArY8Hf8AI3aF/wBf0H/oxax62PB3/I3aF/1/Qf8Aoxaifwv0OjD/AMaHqvzP1li/1afQU+mRf6tPoKfX4W9z+yY/CgooopFBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABX51ftgf8l31r/rjb/wDopa/RWvzq/bA/5LvrX/XG3/8ARS19Vw3/AL6/R/mj804+/wCRXH/GvyZ4lupQc0m2lAxX6cfzuKBml20gOKXdQLUSiiigYUUUUAdv8Dv+SweDP+wnB/6GK/Uhfuj6V+W/wO/5LB4M/wCwnB/6GK/Uhfuj6V+ccTf7xD0P3vw8/wByrf4v0Fooor44/WAooooAKKKKACiiigAooooAKKKKACiiigAooooAK+B/26P+Svad/wBgeL/0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DfJZ6bZXGoXb/dhtYjI5/AV9IfCf9ifWNbmhv/Gsv9k6dlWGnwsGuJOc4ZhlUHT1PUfKea+zdD8N6T4btRaaVpdpplsvSG0gWJPyUAVpdK8fFcQ16y5aK5V+J9RlnAeEws1Uxc/aNdNl/wAEy/DfhvTPCWj22laPZw2Gn267IoIVwqj+p9SeSeTWpRRXykpOTu3qfp8YRpxUYKyQ+iiipNQooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigBMUtFFABSUtFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFACUtFFABSYpaKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigApKWigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAEpaKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACkpaKAEpaKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA//9k=';

    // Shared top header band for chart + summary panels
    slide.addShape('roundRect', {
      x: 0.35,
      y: 0.35,
      w: 12.6,
      h: 0.78,
      rectRadius: 0.06,
      fill: { color: 'F1F2F4' },
      line: { color: 'D0D7DD', pt: 1 },
    });

    slide.addText('Cashflow Chart', {
      x: 0.58,
      y: 0.55,
      w: 5.0,
      h: 0.24,
      fontFace: 'Arial',
      fontSize: 20,
      bold: true,
      color: '118CBE',
    });

    slide.addText('Forecast Summary', {
      x: 8.58,
      y: 0.55,
      w: 3.2,
      h: 0.24,
      fontFace: 'Arial',
      fontSize: 20,
      bold: true,
      color: '118CBE',
    });

    // HMH logo after Forecast Summary in header band
    if (logoDataUrl) {
      slide.addImage({
        data: logoDataUrl,
        x: 12.1,
        y: 0.58,
        h: 0.13,
        w: 0.34,
      });
    }

    // Shared body background (single-slide look)
    slide.addShape('rect', {
      x: 0.35,
      y: 1.08,
      w: 12.6,
      h: 5.55,
      fill: { color: '1493BA' },
      line: { color: '1493BA', pt: 0 },
    });

    // Left chart container
    slide.addShape('roundRect', {
      x: 0.58,
      y: 1.28,
      w: 7.8,
      h: 5.12,
      rectRadius: 0.08,
      fill: { color: '1182A6' },
      line: { color: '1182A6', pt: 0 },
    });

    slide.addImage({
      data: chartImage,
      x: 0.72,
      y: 1.43,
      w: 7.52,
      h: 3.95,
    });

    const pptOppNumber = String(state.project.salesforceOpportunity || '').trim() || 'Not provided';
    const pptRevision = String(state.project.revision || '').trim() || 'Not provided';
    slide.addText(`Opportunity Number: ${pptOppNumber},     Revision: ${pptRevision}`, {
      x: 0.72,
      y: 5.45,
      w: 7.2,
      h: 0.2,
      fontFace: 'Arial',
      fontSize: 11,
      bold: true,
      color: 'F6FCFF',
    });

    slide.addText(`Opportunity Name: ${String(state.project.opportunityName || '').trim() || 'Not provided'}`, {
      x: 0.72,
      y: 5.75,
      w: 7.2,
      h: 0.2,
      fontFace: 'Arial',
      fontSize: 11,
      bold: true,
      color: 'F6FCFF',
    });

    // Right summary container
    slide.addShape('roundRect', {
      x: 8.58,
      y: 1.28,
      w: 4.15,
      h: 5.12,
      rectRadius: 0.08,
      fill: { color: '1182A6' },
      line: { color: '1182A6', pt: 0 },
    });

    slide.addShape('roundRect', {
      x: 8.73,
      y: 1.43,
      w: 3.85,
      h: 4.78,
      rectRadius: 0.06,
      fill: { color: 'EAF6FF' },
      line: { color: 'D0D7DD', pt: 1 },
    });

    summaryItems.forEach((item, index) => {
      const row = Math.floor(index / 2);
      const col = index % 2;
      const x = 8.84 + col * 1.92;
      const y = 1.58 + row * 1.08;
      const isNegative = String(item.value).trim().startsWith('-');

      slide.addShape('roundRect', {
        x,
        y,
        w: 1.76,
        h: 0.84,
        rectRadius: 0.08,
        fill: { color: 'F4FBFF', transparency: 0 },
        line: { color: 'C2D3DF', pt: 1 },
      });
      slide.addText(item.label, {
        x: x + 0.1,
        y: y + 0.08,
        w: 1.55,
        h: 0.16,
        fontFace: 'Arial',
        fontSize: 8,
        color: '5B6773',
        bold: false,
      });
      slide.addText(String(item.value), {
        x: x + 0.1,
        y: y + 0.28,
        w: 1.55,
        h: 0.24,
        fontFace: 'Arial',
        fontSize: 11,
        color: isNegative ? 'E4002B' : '166a94',
        bold: true,
      });
      if (item.note) {
        slide.addText(item.note, {
          x: x + 0.1,
          y: y + 0.54,
          w: 1.55,
          h: 0.16,
          fontFace: 'Arial',
          fontSize: 6,
          color: '5B6773',
        });
      }
    });

    await pptx.writeFile({ fileName: `${getExportBaseName()}.pptx` });
  } catch (error) {
    window.alert(error instanceof Error ? error.message : 'PowerPoint export failed.');
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = 'Export PPT';
    }
  }
}

// Ã¢â€â‚¬Ã¢â€â‚¬ FX auto-fetch Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
// Primary: open.er-api.com (free, no key, live rates)
// Fallback: Frankfurter latest (no key, may have older data)
async function fetchAvgFxRate(from, to) {
  const fromCode = from.trim().toUpperCase();
  const toCode   = to.trim().toUpperCase();
  if (!fromCode || !toCode || fromCode === toCode) return 1;
  if (!/^[A-Z]{3}$/.test(fromCode) || !/^[A-Z]{3}$/.test(toCode)) return null;

  try {
    // Primary: open.er-api.com Ã¢â‚¬â€ live rates, free, no key required
    const erUrl = `https://open.er-api.com/v6/latest/${fromCode}`;
    const erRes = await fetch(erUrl);
    if (erRes.ok) {
      const erData = await erRes.json();
      const rate = Number(erData?.rates?.[toCode]);
      if (Number.isFinite(rate) && rate > 0) {
        return Math.round(rate * 1000000) / 1000000;
      }
    }
  } catch (err) {
    console.warn(`open.er-api fetch error for ${fromCode}/${toCode}:`, err.message);
  }

  try {
    // Fallback: Frankfurter latest
    const latestUrl = `https://api.frankfurter.app/latest?from=${fromCode}&to=${toCode}`;
    const latestRes = await fetch(latestUrl);
    if (latestRes.ok) {
      const latestData = await latestRes.json();
      const latest = Number(latestData?.rates?.[toCode]);
      if (Number.isFinite(latest) && latest > 0) {
        return Math.round(latest * 1000000) / 1000000;
      }
    }
  } catch (err) {
    console.warn(`Frankfurter fetch error for ${fromCode}/${toCode}:`, err.message);
  }

  console.warn(`FX rate unavailable for ${fromCode}/${toCode}`);
  return null;
}

async function autoFetchRateForCost(costId) {
  const cost = state.costs.find((c) => c.id === costId);
  if (!cost || cost.rateIsManual) return;
  const from = (cost.currency || '').trim().toUpperCase();
  const to   = (cost.convertToCurrency || '').trim().toUpperCase();
  if (!from || !to || from === to) {
    cost.conversionRate = 1;
    rerender();
    return;
  }
  const rate = await fetchAvgFxRate(from, to);
  if (rate !== null) {
    cost.conversionRate = rate;
    rerender();
  }
}

function rerender() {
  const model = computeModel();
  renderProjectForm(model);
  renderHealth(model);
  renderMilestones(model);
  renderCashflowEstimate(model);
  renderCosts();
  renderProgressGrid(model);
  renderSummary(model);
  renderForecastTable(model);
  renderChartHeader();
  renderChart(model);

  // Update calculated margin in contract table
  const marginCell = document.getElementById('calculatedMarginCell');
  if (marginCell) {
    const margin = getCalculatedMargin(model);
    marginCell.textContent = margin.text;
    marginCell.style.color = margin.pct < 0 ? '#E4002B' : '';
  }

  persistState();
}

function openUserGuideInNewTab() {
  const guideBody = dom.userGuideModal?.querySelector('.guide-modal-body');
  const guideContent = guideBody
    ? guideBody.innerHTML
    : '<p>User guide content is unavailable.</p>';

  const guideWindow = window.open('', '_blank');
  if (!guideWindow) {
    return;
  }

  guideWindow.document.open();
  guideWindow.document.write(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1.0" />
      <title>User Guide</title>
      <style>
        body {
          margin: 0;
          font-family: "Space Grotesk", sans-serif;
          background: #f7f1e8;
          color: #17212b;
        }
        .guide-wrap {
          max-width: 900px;
          margin: 32px auto;
          padding: 0 20px;
        }
        .guide-card {
          background: #fffdf9;
          border: 1px solid rgba(23, 33, 43, 0.14);
          border-radius: 16px;
          padding: 24px;
          box-shadow: 0 10px 28px rgba(23, 33, 43, 0.08);
        }
        h1 {
          margin: 0 0 16px;
          font-size: 1.6rem;
        }
        p {
          margin: 0 0 10px;
          line-height: 1.5;
        }
      </style>
    </head>
    <body>
      <main class="guide-wrap">
        <section class="guide-card">
          <h1>User Guide</h1>
          ${guideContent}
        </section>
      </main>
    </body>
    </html>
  `);
  guideWindow.document.close();
}

// TEXT fields (label, currency, code, etc.) use 'change' (fires on blur/Enter)
// so rerender() is NOT called on every keystroke Ã¢â‚¬â€ prevents focus loss.
// NUMERIC fields keep 'input' so numbers and charts update live.

const TEXT_FIELDS = new Set(['label', 'code', 'currency', 'convertToCurrency', 'contractCurrency', 'netDays', 'percent', 'conversionRate', 'contractFxRate', 'totalCost']);
const NUMERIC_TEXT_FIELDS = new Set(['contractValue', 'quotedLeadTimeMonths', 'netDays', 'percent', 'conversionRate', 'contractFxRate']);

function handleFieldUpdate(target, callRerender = true) {
  const kind = target.dataset.kind;
  const id = target.dataset.id;
  const field = target.dataset.field;

  if (kind === 'milestone' && id && field) {
    const milestone = state.milestones.find((item) => item.id === id);
    if (!milestone) return;
    milestone[field] = field === 'percent' ? clampNumber(target.value, 0) : target.value;
    if (callRerender) rerender();
  }

  if (kind === 'cost' && id && field) {
    const cost = state.costs.find((item) => item.id === id);
    if (!cost) return;
    const numericFields = ['totalCost', 'conversionRate'];
    cost[field] = numericFields.includes(field) ? clampNumber(target.value, 0) : target.value;
    if (callRerender) rerender();
  }

  if (kind === 'progress' && id) {
    const index = Number(target.dataset.index);
    if (!Array.isArray(state.progress[id])) state.progress[id] = [];
    state.progress[id][index] = clampPercent(target.value);
    if (callRerender) rerender();
  }
}

// Numeric inputs: update state + rerender on every keystroke
document.addEventListener('input', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) && !(target instanceof HTMLTextAreaElement)) return;

  const kind = target.dataset.kind;
  const field = target.dataset.field;

  // Skip text fields here Ã¢â‚¬â€ handled by 'change' below
  if (TEXT_FIELDS.has(field)) return;
  
  // Handle progress updates WITHOUT re-rendering on keystroke
  if (kind === 'progress' && target.dataset.id) {
    const index = Number(target.dataset.index);
    if (!Array.isArray(state.progress[target.dataset.id])) state.progress[target.dataset.id] = [];
    state.progress[target.dataset.id][index] = clampPercent(target.value);
    persistState(); // Save to localStorage immediately, but don't rerender
    return;
  }
  // Also skip project text fields that go through renderProjectForm
  if (kind === 'project' && target.type !== 'number' && target.type !== 'month') return;

  // Mark FX rate as manually overridden when the user edits it directly
  if (kind === 'cost' && field === 'conversionRate' && target.dataset.id) {
    const cost = state.costs.find((c) => c.id === target.dataset.id);
    if (cost) cost.rateIsManual = true;
  }
  if (kind === 'project' && field === 'contractFxRate') {
    state.project.contractRateIsManual = true;
  }

  // Project form numeric/month fields
  if (kind === 'project' && field) {
    const value = target.type === 'number' ? target.valueAsNumber : target.value;
    state.project[field] = Number.isNaN(value) ? 0 : value;
    rerender();
    return;
  }

  handleFieldUpdate(target, true);
});

// Text inputs: update state + rerender only when user leaves the field (blur/Enter)
document.addEventListener('change', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) && !(target instanceof HTMLSelectElement) && !(target instanceof HTMLTextAreaElement)) return;

  const kind = target.dataset.kind;
  const field = target.dataset.field;

  if (!TEXT_FIELDS.has(field) && kind !== 'progress' && !(kind === 'project' && target.type === 'text')) return;

  // Progress inputs: rerender when user leaves the field (blur only)
  // Enter key is handled by the keydown handler which does its own rerender+navigate
  if (kind === 'progress') {
    if (!progressEnterNavigating) rerender();
    progressEnterNavigating = false;
    return;
  }

  // For cost currency fields (including select dropdowns), update state then auto-fetch the new rate
  if (kind === 'cost' && (field === 'currency' || field === 'convertToCurrency') && target.dataset.id) {
    const cost = state.costs.find((c) => c.id === target.dataset.id);
    if (cost) {
      const newValue = target.value.trim().toUpperCase();
      cost[field] = newValue;
      cost.rateIsManual = false; // reset manual flag so new pair gets auto-fetched
      persistState(); // Save state immediately
      
      // Fetch the new rate before rerender to avoid race conditions
      const from = (cost.currency || '').trim().toUpperCase();
      const to = (cost.convertToCurrency || '').trim().toUpperCase();
      if (from && to && from !== to && !cost.rateIsManual) {
        fetchAvgFxRate(from, to).then((rate) => {
          if (rate !== null) {
            cost.conversionRate = rate;
            persistState();
            rerender();
          } else {
            rerender();
          }
        });
      } else {
        rerender();
      }
      return;
    }
  }

  if (kind === 'cost' && field === 'conversionRate' && target.dataset.id) {
    const cost = state.costs.find((c) => c.id === target.dataset.id);
    if (cost) {
      cost.conversionRate = clampNumber(target.value, 0);
      cost.rateIsManual = true;
      rerender();
      return;
    }
  }

  // For project contract currency fields, auto-fetch the new rate
  if (kind === 'project' && (field === 'contractCurrency' || field === 'convertToCurrency')) {
    state.project[field] = target.value.trim().toUpperCase();
    state.project.contractRateIsManual = false;

    if (field === 'convertToCurrency') {
      state.costs.forEach((cost) => {
        cost.convertToCurrency = state.project.convertToCurrency;
        cost.rateIsManual = false;
      });
    }

    rerender();
    autoFetchContractRate();

    if (field === 'convertToCurrency') {
      state.costs.forEach((cost) => autoFetchRateForCost(cost.id));
    }
    return;
  }

  if (kind === 'project' && field) {
    state.project[field] = NUMERIC_TEXT_FIELDS.has(field) ? clampNumber(target.value, 0) : target.value;
    if (field === 'contractFxRate') {
      state.project.contractRateIsManual = true;
    }
    rerender();
    return;
  }

  handleFieldUpdate(target, true);
});

document.addEventListener('click', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  if (target.id === 'userGuideBtn') {
    openUserGuideInNewTab();
    return;
  }

  if (target.id === 'closeGuideBtn' || target.id === 'userGuideModal') {
    if (dom.userGuideModal) {
      dom.userGuideModal.hidden = true;
    }
    return;
  }

  if (target.dataset.kind === 'remove-milestone') {
    state.milestones = state.milestones.filter((item) => item.id !== target.dataset.id);
    rerender();
  }

  if (target.dataset.kind === 'remove-cost') {
    state.costs = state.costs.filter((item) => item.id !== target.dataset.id);
    rerender();
  }

  if (target.dataset.kind === 'reset-fx-rate' && target.dataset.id) {
    const cost = state.costs.find((c) => c.id === target.dataset.id);
    if (cost) {
      cost.rateIsManual = false;
      autoFetchRateForCost(cost.id);
    }
  }

  if (target.dataset.kind === 'reset-contract-fx-rate') {
    state.project.contractRateIsManual = false;
    autoFetchContractRate();
  }
});

document.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape') return;
  if (!dom.userGuideModal || dom.userGuideModal.hidden) return;
  dom.userGuideModal.hidden = true;
});

// Flag to prevent change event from double-rerending when Enter navigates progress cells
let progressEnterNavigating = false;

// General Enter-to-next-field navigation for all inputs
document.addEventListener('keydown', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) && !(target instanceof HTMLSelectElement)) return;
  if (event.key !== 'Enter') return;

  // Progress inputs have their own specialized handler below
  if (target.dataset.kind === 'progress') return;

  event.preventDefault();

  // Gather all visible, enabled inputs/selects in the page in DOM order
  const fieldSelector = 'input.cell-input:not([disabled]):not([type="hidden"]), select.cell-input:not([disabled])';
  const allFields = Array.from(document.querySelectorAll(fieldSelector)).filter(el => el.offsetParent !== null);
  const currentIndex = allFields.indexOf(target);
  if (currentIndex === -1) return;

  // Identify the next field by its data attributes before blur/rerender destroys the DOM
  const nextEl = allFields[currentIndex + 1];
  let nextSelector = null;
  if (nextEl) {
    const kind = nextEl.dataset.kind;
    const field = nextEl.dataset.field;
    const id = nextEl.dataset.id;
    if (kind && field) {
      nextSelector = id
        ? `[data-kind="${kind}"][data-field="${field}"][data-id="${id}"]`
        : `[data-kind="${kind}"][data-field="${field}"]`;
    }
  }

  // Trigger blur so that 'change' event fires, state updates, and rerender runs
  target.blur();

  // After rerender, find the next field in the fresh DOM and focus it
  if (nextSelector) {
    requestAnimationFrame(() => {
      const freshNext = document.querySelector(nextSelector);
      if (freshNext) {
        freshNext.focus();
        if (freshNext instanceof HTMLInputElement) freshNext.select();
      }
    });
  }
});

// Keyboard navigation: Enter moves to next progress field to the right, then next row
document.addEventListener('keydown', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  if (event.key !== 'Enter') return;
  if (target.dataset.kind !== 'progress') return;

  event.preventDefault();

  const currentId = target.dataset.id;
  const currentColIndex = Number(target.dataset.index);

  // Save current cell value to state immediately
  if (!Array.isArray(state.progress[currentId])) state.progress[currentId] = [];
  state.progress[currentId][currentColIndex] = clampPercent(target.value);

  // Build ordered list of all row IDs
  const allInputs = Array.from(document.querySelectorAll('input[data-kind="progress"]'));
  const allRowIds = [];
  allInputs.forEach((el) => {
    if (!allRowIds.includes(el.dataset.id)) allRowIds.push(el.dataset.id);
  });
  const rowInputs = allInputs.filter((el) => el.dataset.id === currentId);
  const colCount = rowInputs.length;
  const isLastCol = currentColIndex >= colCount - 1;

  let nextId, nextCol;
  if (!isLastCol) {
    nextId = currentId;
    nextCol = currentColIndex + 1;
  } else {
    const rowPos = allRowIds.indexOf(currentId);
    nextId = allRowIds[rowPos + 1] ?? allRowIds[0];
    nextCol = 0;
  }

  // Suppress the change event's rerender (we handle it here)
  progressEnterNavigating = true;

  // Rerender synchronously then immediately focus the next cell
  rerender();
  const newInput = document.querySelector(
    `input[data-kind="progress"][data-id="${nextId}"][data-index="${nextCol}"]`
  );
  if (newInput) { newInput.focus(); newInput.select(); }
});

dom.addMilestoneBtn.addEventListener('click', () => {
  state.milestones.push({
    id: crypto.randomUUID(),
    code: `MS${String(state.milestones.length + 1).padStart(2, '0')}`,
    label: 'New milestone',
    percent: 0,
    invoiceMonth: state.project.projectStartMonth,
  });
  rerender();
});

dom.addCostBtn.addEventListener('click', () => {
  const id = crypto.randomUUID();
  state.costs.push({
    id,
    label: `Cost Element ${state.costs.length + 1}`,
    totalCost: 0,
    currency: state.project.contractCurrency || 'USD',
    convertToCurrency: state.project.convertToCurrency || 'USD',
    conversionRate: 1,
    rateIsManual: false,
  });
  state.progress[id] = Array.from({ length: getHorizonMonths() }, () => 100);
  rerender();
  autoFetchRateForCost(id);
});

if (dom.exportPptBtn) {
  dom.exportPptBtn.addEventListener('click', () => {
    exportToPowerPoint();
  });
}

if (dom.exportXlsBtn) {
  dom.exportXlsBtn.addEventListener('click', () => {
    exportToExcel();
  });
}

if (dom.importXlsBtn) {
  dom.importXlsBtn.addEventListener('change', (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    importFromExcel(file);
    // Reset so the same file can be re-imported if needed
    event.target.value = '';
  });
}

if (dom.saveSnapshotBtn) {
  dom.saveSnapshotBtn.addEventListener('click', () => {
    saveSnapshot();
  });
}

if (dom.loadSnapshotBtn) {
  dom.loadSnapshotBtn.addEventListener('change', (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    loadSnapshotFromFile(file);
    event.target.value = '';
  });
}

if (dom.resetDefaultsBtn) {
  dom.resetDefaultsBtn.addEventListener('click', () => {
    resetToDefaults();
  });
}

if (dom.resetDefaultsTopBtn) {
  dom.resetDefaultsTopBtn.addEventListener('click', () => {
    resetToDefaults();
  });
}

if (dom.setDefaultsBtn) {
  dom.setDefaultsBtn.addEventListener('click', () => {
    saveCurrentAsDefault();
  });
}

ensureProgressShape();
rerender();
// Auto-fetch contract FX rate at startup if currencies differ
async function autoFetchContractRate() {
  if (state.project.contractRateIsManual) return;
  const from = (state.project.contractCurrency || '').trim().toUpperCase();
  const to   = (state.project.convertToCurrency || '').trim().toUpperCase();
  if (!from || !to || from === to) {
    state.project.contractFxRate = 1;
    rerender();
    return;
  }
  const rate = await fetchAvgFxRate(from, to);
  if (rate !== null) {
    state.project.contractFxRate = rate;
    rerender();
  }
}
const contractFrom = (state.project.contractCurrency || '').trim().toUpperCase();
const contractTo   = (state.project.convertToCurrency || '').trim().toUpperCase();
if (contractFrom && contractTo && contractFrom !== contractTo && !state.project.contractRateIsManual) {
  autoFetchContractRate();
}
// Auto-fetch rates for any cost rows where currencies already differ at startup
state.costs.forEach((cost) => {
  const from = (cost.currency || '').trim().toUpperCase();
  const to   = (cost.convertToCurrency || '').trim().toUpperCase();
  if (from && to && from !== to && !cost.rateIsManual) {
    autoFetchRateForCost(cost.id);
  }
});
