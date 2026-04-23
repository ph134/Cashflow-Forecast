const WORKFLOW_STORAGE_KEY = 'workflow-studio:state:v1';
const WORKFLOW_ROLES = ['Engineering', 'Sourcing', 'Project Manager'];
const DEFAULT_ASSIGNEES = ['Alex Carter', 'Jordan Lee', 'Taylor Morgan', 'Riley Patel'];

function createDefaultState() {
  const roleDetails = {};
  const submissions = {};

  WORKFLOW_ROLES.forEach((role) => {
    roleDetails[role] = { assigneeName: '', assigneeContact: '', slaHours: 48, dueAt: '' };
    submissions[role] = { notes: '', status: 'pending', submittedAt: '' };
  });

  return {
    opportunityContext: '',
    scope: '',
    returnTo: 'Me',
    mode: 'sequence',
    selectedInputs: [...WORKFLOW_ROLES],
    sequenceOrder: [...WORKFLOW_ROLES],
    roleDetails,
    submissions,
    status: 'draft',
    startedAt: '',
    returnedAt: '',
    approval: { status: 'not-required', decisionBy: '', decisionAt: '', notes: '' },
    notifications: { enabled: false, channel: 'email', target: '', sentForRoles: [] },
    assigneeOptions: [...DEFAULT_ASSIGNEES],
    history: [],
    opportunityDetails: {
      bidManager: '',
      opportunityNumber: '',
      opportunityName: '',
      account: '',
      seller: '',
      rigName: '',
      askTicketNumber: '',
      orderType: '',
      forecastCategory: '',
      bidDueDate: '',
      opportunityScope: '',
      bidManagerComment: '',
      bidManagerAttachment: '',
      placeholderLink: '',
    },
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(WORKFLOW_STORAGE_KEY);
    if (!raw) return createDefaultState();
    const parsed = JSON.parse(raw);
    return normalizeState(parsed);
  } catch {
    return createDefaultState();
  }
}

function saveState() {
  localStorage.setItem(WORKFLOW_STORAGE_KEY, JSON.stringify(state));
}

function normalizeState(raw) {
  const base = createDefaultState();
  const next = { ...base, ...(raw || {}) };

  next.assigneeOptions = Array.isArray(raw?.assigneeOptions)
    ? Array.from(new Set(raw.assigneeOptions.map((x) => String(x || '').trim()).filter(Boolean))).slice(0, 100)
    : base.assigneeOptions;

  next.selectedInputs = Array.isArray(raw?.selectedInputs)
    ? raw.selectedInputs.filter((role) => WORKFLOW_ROLES.includes(role))
    : [...base.selectedInputs];

  next.sequenceOrder = Array.isArray(raw?.sequenceOrder)
    ? raw.sequenceOrder.filter((role) => next.selectedInputs.includes(role))
    : [...next.selectedInputs];

  next.selectedInputs.forEach((role) => {
    if (!next.sequenceOrder.includes(role)) next.sequenceOrder.push(role);
  });

  WORKFLOW_ROLES.forEach((role) => {
    next.roleDetails[role] = { ...base.roleDetails[role], ...(raw?.roleDetails?.[role] || {}) };
    next.submissions[role] = { ...base.submissions[role], ...(raw?.submissions?.[role] || {}) };
  });

  next.history = Array.isArray(raw?.history)
    ? raw.history.filter((h) => h && typeof h === 'object').slice(0, 200)
    : [];

  next.opportunityDetails = { ...base.opportunityDetails, ...(raw?.opportunityDetails || {}) };

  updateStatus(next);
  return next;
}

function addHistory(event, role = '', details = '') {
  state.history.unshift({ id: crypto.randomUUID(), at: new Date().toISOString(), event, role, details });
  if (state.history.length > 200) state.history.length = 200;
}

function orderedRoles() {
  return state.sequenceOrder.filter((r) => state.selectedInputs.includes(r));
}

function nextRole() {
  return orderedRoles().find((r) => state.submissions[r].status !== 'submitted');
}

function activeRoles() {
  if (state.mode === 'parallel') return state.selectedInputs.filter((r) => state.submissions[r].status !== 'submitted');
  const next = nextRole();
  return next ? [next] : [];
}

function dueAt(role) {
  const detail = state.roleDetails[role];
  if (detail.dueAt) return detail.dueAt;
  if (!state.startedAt) return '';
  const start = new Date(state.startedAt).getTime();
  const ms = Math.max(1, Number(detail.slaHours || 48)) * 60 * 60 * 1000;
  return new Date(start + ms).toISOString().slice(0, 16);
}

function isOverdue(role) {
  if (state.submissions[role].status === 'submitted') return false;
  const due = dueAt(role);
  return !!due && new Date(due).getTime() < Date.now();
}

function updateStatus(targetState) {
  const allSubmitted = targetState.selectedInputs.length > 0
    && targetState.selectedInputs.every((r) => targetState.submissions[r].status === 'submitted');

  const hasActivity = targetState.selectedInputs.some((r) => {
    const s = targetState.submissions[r];
    return (s.notes || '').trim().length > 0 || s.status === 'submitted';
  });

  if (allSubmitted) {
    targetState.status = targetState.approval.status === 'approved' ? 'returned' : 'awaiting-approval';
    if (!targetState.startedAt) targetState.startedAt = new Date().toISOString();
    if (targetState.status === 'returned' && !targetState.returnedAt) {
      targetState.returnedAt = targetState.approval.decisionAt || new Date().toISOString();
    }
    if (targetState.status === 'awaiting-approval') {
      targetState.approval.status = 'pending';
      targetState.returnedAt = '';
    }
    return;
  }

  targetState.returnedAt = '';
  targetState.approval.status = 'not-required';
  targetState.approval.decisionAt = '';

  if (hasActivity) {
    targetState.status = 'in-progress';
    if (!targetState.startedAt) targetState.startedAt = new Date().toISOString();
    return;
  }

  targetState.status = 'draft';
  targetState.startedAt = '';
}

function notifyIfNeeded() {
  if (!state.notifications.enabled) return;
  const currentlyActive = activeRoles();
  const alreadySent = new Set(state.notifications.sentForRoles || []);
  const newRoles = currentlyActive.filter((r) => !alreadySent.has(r));
  if (!newRoles.length) return;

  const message = `Workflow input requested: ${newRoles.join(', ')}`;
  const fallbackTargets = newRoles.map((r) => state.roleDetails[r].assigneeContact).filter(Boolean);
  const target = state.notifications.target || fallbackTargets.join(';');

  if ('Notification' in window && Notification.permission === 'granted') {
    new Notification('Quotation Workflow', { body: message });
  } else if ('Notification' in window && Notification.permission === 'default') {
    Notification.requestPermission().catch(() => {});
  }

  if (state.notifications.channel === 'email' && target) {
    const email = encodeURIComponent(target);
    const subject = encodeURIComponent('Workflow input request');
    const body = encodeURIComponent(`${message}\nPlease submit your role input.`);
    window.open(`mailto:${email}?subject=${subject}&body=${body}`, '_blank');
  }

  if (state.notifications.channel === 'teams' && target) {
    const users = encodeURIComponent(target.replace(/;/g, ','));
    const chatMessage = encodeURIComponent(`${message}. Please submit your role input.`);
    window.open(`https://teams.microsoft.com/l/chat/0/0?users=${users}&message=${chatMessage}`, '_blank');
  }

  state.notifications.sentForRoles = Array.from(new Set([...(state.notifications.sentForRoles || []), ...newRoles]));
  addHistory('Notification sent', newRoles.join(', '), `Channel: ${state.notifications.channel}`);
}

const state = loadState();

const dom = {
  opportunityContext: document.querySelector('#opportunityContext'),
  scope: document.querySelector('#scope'),
  returnTo: document.querySelector('#returnTo'),
  roleSelection: document.querySelector('#roleSelection'),
  sequenceOrder: document.querySelector('#sequenceOrder'),
  roleCards: document.querySelector('#roleCards'),
  notifyEnabled: document.querySelector('#notifyEnabled'),
  notifyChannel: document.querySelector('#notifyChannel'),
  notifyTarget: document.querySelector('#notifyTarget'),
  decisionBy: document.querySelector('#decisionBy'),
  approvalNotes: document.querySelector('#approvalNotes'),
  approveBtn: document.querySelector('#approveBtn'),
  reworkBtn: document.querySelector('#reworkBtn'),
  resetBtn: document.querySelector('#resetBtn'),
  exportBtn: document.querySelector('#exportBtn'),
  statusLine: document.querySelector('#statusLine'),
  historyList: document.querySelector('#historyList'),
  newAssignee: document.querySelector('#newAssignee'),
  addAssigneeBtn: document.querySelector('#addAssigneeBtn'),
  assigneeTags: document.querySelector('#assigneeTags'),
  odBidManager: document.querySelector('#odBidManager'),
  odOpportunityNumber: document.querySelector('#odOpportunityNumber'),
  odOpportunityName: document.querySelector('#odOpportunityName'),
  odAccount: document.querySelector('#odAccount'),
  odSeller: document.querySelector('#odSeller'),
  odRigName: document.querySelector('#odRigName'),
  odAskTicketNumber: document.querySelector('#odAskTicketNumber'),
  odOrderType: document.querySelector('#odOrderType'),
  odForecastCategory: document.querySelector('#odForecastCategory'),
  odBidDueDate: document.querySelector('#odBidDueDate'),
  odOpportunityScope: document.querySelector('#odOpportunityScope'),
  odBidManagerComment: document.querySelector('#odBidManagerComment'),
  odBidManagerAttachment: document.querySelector('#odBidManagerAttachment'),
  odPlaceholderLink: document.querySelector('#odPlaceholderLink'),
};

function esc(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function render() {
  const od = state.opportunityDetails;
  dom.odBidManager.value = od.bidManager || '';
  dom.odOpportunityNumber.value = od.opportunityNumber || '';
  dom.odOpportunityName.value = od.opportunityName || '';
  dom.odAccount.value = od.account || '';
  dom.odSeller.value = od.seller || '';
  dom.odRigName.value = od.rigName || '';
  dom.odAskTicketNumber.value = od.askTicketNumber || '';
  dom.odOrderType.value = od.orderType || '';
  dom.odForecastCategory.value = od.forecastCategory || '';
  dom.odBidDueDate.value = od.bidDueDate || '';
  dom.odOpportunityScope.value = od.opportunityScope || '';
  dom.odBidManagerComment.value = od.bidManagerComment || '';
  dom.odBidManagerAttachment.value = od.bidManagerAttachment || '';
  dom.odPlaceholderLink.value = od.placeholderLink || '';

  dom.opportunityContext.value = state.opportunityContext;
  dom.scope.value = state.scope;
  dom.returnTo.value = state.returnTo;

  const statusText = state.status === 'returned'
    ? `Returned to ${state.returnTo || 'you'}`
    : state.status === 'awaiting-approval'
      ? 'Awaiting final approval'
      : state.status === 'in-progress'
        ? 'In progress'
        : 'Draft';
  dom.statusLine.textContent = `Status: ${statusText}`;

  dom.notifyEnabled.checked = !!state.notifications.enabled;
  dom.notifyChannel.value = state.notifications.channel || 'email';
  dom.notifyTarget.value = state.notifications.target || '';

  dom.decisionBy.value = state.approval.decisionBy || state.returnTo || 'Me';
  dom.approvalNotes.value = state.approval.notes || '';

  dom.roleSelection.innerHTML = WORKFLOW_ROLES.map((role) => {
    const checked = state.selectedInputs.includes(role) ? 'checked' : '';
    return `<label><input type="checkbox" data-kind="select-role" data-role="${esc(role)}" ${checked}> ${esc(role)}</label>`;
  }).join('');

  if (state.mode === 'sequence') {
    dom.sequenceOrder.innerHTML = orderedRoles().map((role, i, arr) => `
      <div class="sequence-item">
        <span>${i + 1}. ${esc(role)}</span>
        <div class="inline-row">
          <button type="button" data-kind="move-role" data-role="${esc(role)}" data-dir="up" ${i === 0 ? 'disabled' : ''}>Up</button>
          <button type="button" data-kind="move-role" data-role="${esc(role)}" data-dir="down" ${i === arr.length - 1 ? 'disabled' : ''}>Down</button>
        </div>
      </div>
    `).join('');
  } else {
    dom.sequenceOrder.innerHTML = '<span class="field-label">Parallel mode: all selected roles are active together.</span>';
  }

  dom.assigneeTags.innerHTML = (state.assigneeOptions || []).map((name) => `
    <span class="tag">${esc(name)} <button type="button" data-kind="remove-assignee" data-name="${esc(name)}">x</button></span>
  `).join('') || '<span class="field-label">No names configured.</span>';

  const next = nextRole();
  dom.roleCards.innerHTML = orderedRoles().map((role) => {
    const detail = state.roleDetails[role];
    const sub = state.submissions[role];
    const roleId = role.toLowerCase().replace(/[^a-z0-9]+/g, '-');
    const isSubmitted = sub.status === 'submitted';
    const canSubmit = !isSubmitted && (state.mode === 'parallel' || next === role);
    const badge = isSubmitted ? 'done' : isOverdue(role) ? 'overdue' : canSubmit ? 'ready' : 'hold';
    const badgeText = isSubmitted ? 'Submitted' : isOverdue(role) ? 'Overdue' : canSubmit ? 'Ready' : 'On Hold';

    const options = (state.assigneeOptions || []).map((n) => `<option value="${esc(n)}"></option>`).join('');

    return `
      <article class="role-card ${isSubmitted ? 'submitted' : ''}">
        <div class="role-head">
          <strong>${esc(role)}</strong>
          <span class="badge ${badge}">${badgeText}</span>
        </div>
        <datalist id="assignees-${roleId}">${options}</datalist>
        <label>Assignee
          <input type="text" list="assignees-${roleId}" data-kind="assignee-name" data-role="${esc(role)}" value="${esc(detail.assigneeName)}" ${isSubmitted ? 'disabled' : ''}>
        </label>
        <label>Contact
          <input type="text" data-kind="assignee-contact" data-role="${esc(role)}" value="${esc(detail.assigneeContact)}" placeholder="email or Teams user" ${isSubmitted ? 'disabled' : ''}>
        </label>
        <label>SLA Hours
          <input type="number" min="1" step="1" data-kind="sla-hours" data-role="${esc(role)}" value="${detail.slaHours}">
        </label>
        <label>Due
          <input type="datetime-local" data-kind="due-at" data-role="${esc(role)}" value="${esc(dueAt(role))}">
        </label>
        <label>Input Notes
          <textarea rows="2" data-kind="role-notes" data-role="${esc(role)}" ${isSubmitted ? 'disabled' : ''}>${esc(sub.notes)}</textarea>
        </label>
        <div class="inline-row">
          <button type="button" data-kind="submit-role" data-role="${esc(role)}" ${canSubmit ? '' : 'disabled'}>Submit Input</button>
          <button type="button" data-kind="reopen-role" data-role="${esc(role)}" ${isSubmitted ? '' : 'disabled'}>Reopen</button>
        </div>
      </article>
    `;
  }).join('') || '<p class="field-label">Select at least one required role.</p>';

  dom.historyList.innerHTML = (state.history || []).slice(0, 12).map((entry) => {
    const roleText = entry.role ? ` [${entry.role}]` : '';
    const detailText = entry.details ? ` - ${entry.details}` : '';
    return `<li><strong>${new Date(entry.at).toLocaleString('en-US')}:</strong> ${esc(entry.event)}${esc(roleText)}${esc(detailText)}</li>`;
  }).join('') || '<li>No activity yet.</li>';

  saveState();
  notifyIfNeeded();
}

function rerenderWithHistory(event, role = '', details = '') {
  if (event) addHistory(event, role, details);
  updateStatus(state);
  render();
}

document.addEventListener('input', (event) => {
  const t = event.target;
  if (!(t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement || t instanceof HTMLSelectElement)) return;

  const kind = t.dataset.kind;
  const role = t.dataset.role;

  if (t.id === 'opportunityContext') {
    state.opportunityContext = t.value;
    saveState();
    return;
  }

  if (t.id === 'scope') {
    state.scope = t.value;
    saveState();
    return;
  }

  if (t.id === 'returnTo') {
    state.returnTo = t.value;
    saveState();
    return;
  }

  if (kind === 'od' && t.dataset.field && t.dataset.field in state.opportunityDetails) {
    state.opportunityDetails[t.dataset.field] = t.value;
    saveState();
    return;
  }

  if (kind === 'role-notes' && role) {
    state.submissions[role].notes = t.value;
    updateStatus(state);
    saveState();
    return;
  }

  if (kind === 'assignee-name' && role) {
    state.roleDetails[role].assigneeName = t.value;
    saveState();
    return;
  }

  if (kind === 'assignee-contact' && role) {
    state.roleDetails[role].assigneeContact = t.value;
    saveState();
    return;
  }
});

document.addEventListener('change', (event) => {
  const t = event.target;
  if (!(t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement || t instanceof HTMLSelectElement)) return;

  const kind = t.dataset.kind;
  const role = t.dataset.role;

  if (kind === 'od' && t.dataset.field && t.dataset.field in state.opportunityDetails) {
    state.opportunityDetails[t.dataset.field] = t.value;
    saveState();
    return;
  }

  if (t.name === 'flowMode') {
    state.mode = t.value === 'parallel' ? 'parallel' : 'sequence';
    state.notifications.sentForRoles = [];
    rerenderWithHistory('Flow mode changed', '', `Mode: ${state.mode}`);
    return;
  }

  if (kind === 'select-role' && role && t instanceof HTMLInputElement) {
    const set = new Set(state.selectedInputs);
    if (t.checked) {
      set.add(role);
      if (!state.sequenceOrder.includes(role)) state.sequenceOrder.push(role);
      state.submissions[role].status = 'pending';
      state.submissions[role].submittedAt = '';
    } else {
      set.delete(role);
      state.sequenceOrder = state.sequenceOrder.filter((r) => r !== role);
      state.submissions[role].status = 'pending';
      state.submissions[role].submittedAt = '';
      state.roleDetails[role].dueAt = '';
    }
    state.selectedInputs = Array.from(set);
    state.notifications.sentForRoles = [];
    rerenderWithHistory('Role selection changed', role, t.checked ? 'Included' : 'Removed');
    return;
  }

  if (kind === 'assignee-name' && role) {
    state.roleDetails[role].assigneeName = t.value;
    rerenderWithHistory('Assignee updated', role, t.value || 'Unassigned');
    return;
  }

  if (kind === 'assignee-contact' && role) {
    state.roleDetails[role].assigneeContact = t.value;
    render();
    return;
  }

  if (kind === 'sla-hours' && role && t instanceof HTMLInputElement) {
    state.roleDetails[role].slaHours = Math.max(1, Math.round(Number(t.value) || 48));
    if (!state.roleDetails[role].dueAt) state.roleDetails[role].dueAt = dueAt(role);
    rerenderWithHistory('SLA updated', role, `${state.roleDetails[role].slaHours} hours`);
    return;
  }

  if (kind === 'due-at' && role && t instanceof HTMLInputElement) {
    state.roleDetails[role].dueAt = t.value || '';
    rerenderWithHistory('Due date updated', role, t.value || 'Cleared');
    return;
  }

  if (t.id === 'notifyEnabled' && t instanceof HTMLInputElement) {
    state.notifications.enabled = t.checked;
    if (!t.checked) state.notifications.sentForRoles = [];
    rerenderWithHistory('Notifications updated', '', t.checked ? 'Enabled' : 'Disabled');
    return;
  }

  if (t.id === 'notifyChannel' && t instanceof HTMLSelectElement) {
    state.notifications.channel = t.value === 'teams' ? 'teams' : 'email';
    state.notifications.sentForRoles = [];
    rerenderWithHistory('Notification channel updated', '', state.notifications.channel);
    return;
  }

  if (t.id === 'notifyTarget') {
    state.notifications.target = t.value;
    render();
    return;
  }

  if (t.id === 'decisionBy') {
    state.approval.decisionBy = t.value;
    render();
    return;
  }

  if (t.id === 'approvalNotes') {
    state.approval.notes = t.value;
    render();
    return;
  }
});

document.addEventListener('click', (event) => {
  const t = event.target;
  if (!(t instanceof HTMLElement)) return;

  const kind = t.dataset.kind;
  const role = t.dataset.role;

  if (t.id === 'addAssigneeBtn') {
    const name = (dom.newAssignee.value || '').trim();
    if (!name) return;
    if (!state.assigneeOptions.includes(name)) {
      state.assigneeOptions.push(name);
      state.assigneeOptions = state.assigneeOptions.slice(0, 100);
      rerenderWithHistory('Assignee option added', '', name);
    }
    dom.newAssignee.value = '';
    return;
  }

  if (kind === 'remove-assignee' && t.dataset.name) {
    state.assigneeOptions = state.assigneeOptions.filter((x) => x !== t.dataset.name);
    rerenderWithHistory('Assignee option removed', '', t.dataset.name);
    return;
  }

  if (kind === 'move-role' && role && t.dataset.dir) {
    const order = [...state.sequenceOrder];
    const i = order.indexOf(role);
    if (i >= 0) {
      const j = t.dataset.dir === 'up' ? i - 1 : i + 1;
      if (j >= 0 && j < order.length) {
        [order[i], order[j]] = [order[j], order[i]];
        state.sequenceOrder = order;
        rerenderWithHistory('Sequence updated', role, `Moved ${t.dataset.dir}`);
      }
    }
    return;
  }

  if (kind === 'submit-role' && role) {
    const next = nextRole();
    if (state.mode === 'sequence' && role !== next) return;

    state.submissions[role].status = 'submitted';
    state.submissions[role].submittedAt = new Date().toISOString();
    state.notifications.sentForRoles = [];
    rerenderWithHistory('Input submitted', role, `Submitted by ${state.roleDetails[role].assigneeName || 'Unknown assignee'}`);
    return;
  }

  if (kind === 'reopen-role' && role) {
    state.submissions[role].status = 'pending';
    state.submissions[role].submittedAt = '';
    state.approval.status = 'not-required';
    state.approval.decisionAt = '';
    state.returnedAt = '';
    state.notifications.sentForRoles = [];
    rerenderWithHistory('Input reopened', role, 'Set back to pending');
    return;
  }

  if (t.id === 'approveBtn') {
    if (state.status !== 'awaiting-approval') return;
    state.approval.status = 'approved';
    state.approval.decisionAt = new Date().toISOString();
    if (!state.approval.decisionBy) state.approval.decisionBy = state.returnTo || 'Me';
    rerenderWithHistory('Workflow approved', '', `Approved by ${state.approval.decisionBy}`);
    return;
  }

  if (t.id === 'reworkBtn') {
    if (state.status !== 'awaiting-approval') return;
    state.approval.status = 'rework';
    state.approval.decisionAt = new Date().toISOString();
    state.status = 'in-progress';
    state.returnedAt = '';
    state.selectedInputs.forEach((r) => {
      state.submissions[r].status = 'pending';
      state.submissions[r].submittedAt = '';
    });
    state.notifications.sentForRoles = [];
    rerenderWithHistory('Rework requested', '', `Requested by ${state.approval.decisionBy || state.returnTo || 'Me'}`);
    return;
  }

  if (t.id === 'resetBtn') {
    const base = createDefaultState();
    Object.assign(state, base);
    rerenderWithHistory('Reset cycle', '', 'Workflow reset to defaults');
    return;
  }

  if (t.id === 'exportBtn') {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'workflow-studio-export.json';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }
});

render();
