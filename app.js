const STORAGE_KEY = 'cashflow-web-app:state:v1';

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

function loadInitialState() {
  const defaults = createDefaultState();
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
  cashflowEstimate: document.querySelector('#cashflowEstimate'),
  userGuideBtn: document.querySelector('#userGuideBtn'),
  userGuideModal: document.querySelector('#userGuideModal'),
  closeGuideBtn: document.querySelector('#closeGuideBtn'),
};

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

  // ── Contract row (table-style, mirrors cost elements) ──────────────────
  const contractFxRate = state.project.contractFxRate ?? 1;
  const contractRateIsManual = state.project.contractRateIsManual;
  const sameCurrency = (state.project.contractCurrency || '').toUpperCase() === (state.project.convertToCurrency || '').toUpperCase();
  const marginValue = String(state.project.manualMargin ?? '0%');
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
      <label class="inline-note" style="display:block;margin-bottom:6px;">Opportunity Name</label>
      <input class="cell-input" type="text" data-kind="project" data-field="opportunityName" value="${state.project.opportunityName || ''}" placeholder="Opportunity Name">
    </div>
    <div class="timeline-field">
      <label class="inline-note" style="display:block;margin-bottom:6px;">Quoted lead time (months)</label>
      <input class="cell-input" type="text" data-kind="project" data-field="quotedLeadTimeMonths" value="${state.project.quotedLeadTimeMonths}">
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
            <button class="ghost-button" style="padding:4px 8px;font-size:0.78rem;" type="button" data-kind="reset-contract-fx-rate" title="Reset to 3-month average">↺</button>
          </td>
          <td class="currency-cell ${sameCurrency ? 'muted-cell' : 'converted-value'}">${formatCurrency(clampNumber(state.project.contractValue) * clampNumber(contractFxRate, 1), state.project.convertToCurrency)}</td>
          <td class="number-cell"><input class="cell-input" type="text" data-kind="project" data-field="manualMargin" value="${marginValue}" placeholder="e.g. 28%"></td>
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
        <button class="ghost-button" style="padding:4px 8px;font-size:0.78rem;" type="button" data-kind="reset-fx-rate" data-id="${cost.id}" title="Reset to 3-month average">↺</button>
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

function getSummaryItems(model) {
  const contractValueConverted = clampNumber(state.project.contractValue) * clampNumber(state.project.contractFxRate, 1);
  const marginText = String(state.project.manualMargin || '').trim() || '0%';
  const horizon = getHorizonMonths();

  return [
    { label: 'Contract Value', value: formatCurrency(contractValueConverted), note: '' },
    { label: 'Margin', value: marginText, note: '' },
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
    return `<circle cx="${x}" cy="${y}" r="4.5" fill="#0f7b78"></circle>`;
  }).join('');

  const xLabels = model.months.map((month, index) => {
    const x = padding.left + index * stepX + stepX / 2;
    return `<text x="${x}" y="${height - 40}" text-anchor="middle" font-size="12" fill="#314f63">${formatMonthLabel(month)}</text>`;
  }).join('');

  svg.innerHTML = `
    <rect x="0" y="0" width="${width}" height="${height}" fill="transparent"></rect>
    ${gridLines}
    <line x1="${padding.left}" y1="${zeroY}" x2="${width - padding.right}" y2="${zeroY}" stroke="rgba(23,33,43,0.24)" />
    <path d="${cumulativePath}" fill="none" stroke="#0f7b78" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"></path>
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

      context.fillStyle = '#fffaf3';
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

  // ── Sheet 1: Inputs ──────────────────────────────────────────────────────
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

  // ── Sheet 2: Cashflow Forecast ───────────────────────────────────────────
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

  const safeName = (state.project.opportunityName || 'cashflow-forecast')
    .replace(/[^a-z0-9\-_ ]/gi, '').trim() || 'cashflow-forecast';
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

      // ── Opportunity Information ──────────────────────────────────────────
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

      // ── Find section start rows ──────────────────────────────────────────
      let msHeaderRow = -1, costHeaderRow = -1, plannedHeaderRow = -1;
      rows.forEach((row, i) => {
        const cell = String(row[0] || '').trim();
        if (cell === 'MILESTONES') msHeaderRow = i;
        if (cell === 'COST ELEMENTS') costHeaderRow = i;
        if (cell === 'PLANNED EXPENDITURES') plannedHeaderRow = i;
      });

      // ── Milestones ───────────────────────────────────────────────────────
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

      // ── Cost Elements ────────────────────────────────────────────────────
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

      // ── Planned Expenditures (progress %) ────────────────────────────────
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

      // ── Apply to state ───────────────────────────────────────────────────
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
    slide.background = { color: 'F7F1E8' };

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

    slide.addText('HMH', {
      x: 7.95,
      y: 0.56,
      w: 0.9,
      h: 0.2,
      fontFace: 'Arial',
      fontSize: 13,
      bold: true,
      color: '1F8AB8',
    });

    slide.addText('Forecast Summary', {
      x: 8.85,
      y: 0.55,
      w: 3.2,
      h: 0.24,
      fontFace: 'Arial',
      fontSize: 20,
      bold: true,
      color: '118CBE',
    });

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
        color: isNegative ? 'D75A2D' : '0F7B78',
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

    await pptx.writeFile({ fileName: 'cashflow-forecast.pptx' });
  } catch (error) {
    window.alert(error instanceof Error ? error.message : 'PowerPoint export failed.');
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = 'Export PPT';
    }
  }
}

// ── FX auto-fetch ─────────────────────────────────────────────────────────────
// Primary: open.er-api.com (free, no key, live rates)
// Fallback: Frankfurter latest (no key, may have older data)
async function fetchAvgFxRate(from, to) {
  const fromCode = from.trim().toUpperCase();
  const toCode   = to.trim().toUpperCase();
  if (!fromCode || !toCode || fromCode === toCode) return 1;
  if (!/^[A-Z]{3}$/.test(fromCode) || !/^[A-Z]{3}$/.test(toCode)) return null;

  try {
    // Primary: open.er-api.com — live rates, free, no key required
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
// so rerender() is NOT called on every keystroke — prevents focus loss.
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
  if (!(target instanceof HTMLInputElement)) return;

  const kind = target.dataset.kind;
  const field = target.dataset.field;

  // Skip text fields here — handled by 'change' below
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
  if (!(target instanceof HTMLInputElement) && !(target instanceof HTMLSelectElement)) return;

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
