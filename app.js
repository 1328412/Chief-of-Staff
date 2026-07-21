import { state } from './state.js';
import { m365Service } from './m365.js';
import { triageManager } from './triage.js';
import { scheduleManager } from './schedule.js';
import { briefingManager } from './briefing.js';
import { orchestrator } from './orchestrator.js';

class App {
  constructor() {
    this.m365 = m365Service;
    this.triage = triageManager;
    this.schedule = scheduleManager;
    this.briefing = briefingManager;
    this.orchestrator = orchestrator;

    this.connectButton = document.getElementById('connect-button');
    this.statusText = document.getElementById('cos-status-text');
    this.directiveInput = document.getElementById('master-directive-input');
    this.directiveButton = document.getElementById('btn-send-directive');
    this.diagnosticsButton = document.getElementById('btn-open-m365-diagnostics');
    this.diagnosticsPanel = document.getElementById('m365-diagnostics-panel');

    this.diagState = {
      lastCheck: 'No checks run yet.',
      lastCheckClass: '',
      deepChecks: {
        profile: { status: 'Not run', className: '' },
        mail: { status: 'Not run', className: '' },
        files: { status: 'Not run', className: '' }
      },
      report: 'Report not generated yet.'
    };
  }

  async init() {
    this.setupEventListeners();
    this.setupTabs();
    this.orchestrator.init();

    try {
      const account = await this.m365.init();
      if (account) {
        this.switchToLiveMode();
      } else {
        this.renderDemoMode();
      }
    } catch (error) {
      console.error('Initialization failed:', error);
      this.renderDemoMode();
    }

    this.renderAll();
    this.updateMetrics();
    this.renderM365Diagnostics();
  }

  setupEventListeners() {
    if (this.connectButton) {
      this.connectButton.addEventListener('click', () => this.login());
    }

    if (this.directiveButton) {
      this.directiveButton.addEventListener('click', () => this.runMasterDirective());
    }

    if (this.directiveInput) {
      this.directiveInput.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
          this.runMasterDirective();
        }
      });
    }

    if (this.diagnosticsButton) {
      this.diagnosticsButton.addEventListener('click', () => this.toggleM365Diagnostics());
    }

    const refreshDiagBtn = document.getElementById('btn-refresh-m365-diagnostics');
    if (refreshDiagBtn) {
      refreshDiagBtn.addEventListener('click', () => this.renderM365Diagnostics());
    }

    const runDiagBtn = document.getElementById('btn-run-m365-diagnostics');
    if (runDiagBtn) {
      runDiagBtn.addEventListener('click', () => this.runM365Diagnostics());
    }

    const runDeepDiagBtn = document.getElementById('btn-run-m365-deep-diagnostics');
    if (runDeepDiagBtn) {
      runDeepDiagBtn.addEventListener('click', () => this.runM365DeepDiagnostics());
    }

    const copyDiagBtn = document.getElementById('btn-copy-m365-diagnostics');
    if (copyDiagBtn) {
      copyDiagBtn.addEventListener('click', () => this.copyM365DiagnosticsReport());
    }

    document.querySelectorAll('.preset-chip').forEach((chip) => {
      chip.addEventListener('click', () => {
        const prompt = chip.getAttribute('data-prompt');
        if (prompt && this.directiveInput) {
          this.directiveInput.value = prompt;
          this.runMasterDirective();
        }
      });
    });
  }

  setupTabs() {
    document.querySelectorAll('.nav-item[data-tab]').forEach((item) => {
      item.addEventListener('click', () => {
        const tabName = item.getAttribute('data-tab');
        if (tabName) {
          this.switchTab(tabName);
        }
      });
    });
  }

  async login() {
    try {
      const account = await this.m365.login();
      if (!account) return;

      this.switchToLiveMode();
      await this.tryHydrateLiveInbox();
      this.renderAll();
      this.updateMetrics();
      this.renderM365Diagnostics();
    } catch (error) {
      console.error('Login failed:', error);
      alert('Microsoft 365 sign-in failed. Please try again.');
      this.diagState.lastCheck = `Login failed: ${error.message || String(error)}`;
      this.diagState.lastCheckClass = 'diag-error';
      this.renderM365Diagnostics();
    }
  }

  async tryHydrateLiveInbox() {
    try {
      const liveMessages = await this.m365.getLiveInboxMessages();
      if (!Array.isArray(liveMessages) || liveMessages.length === 0) return;

      state.inbox = liveMessages.map((msg, index) => {
        const senderName = msg?.from?.emailAddress?.name || 'Unknown Sender';
        const senderEmail = msg?.from?.emailAddress?.address || 'unknown@contoso.com';
        const subject = msg.subject || '(No subject)';
        const snippet = msg.bodyPreview || '';
        const isVip = state.executive.vipList.some((vip) => senderName.includes(vip.split(' (')[0]));

        let priority = 'MEDIUM';
        if (msg.importance === 'high') priority = 'HIGH';
        if (isVip && msg.isRead === false) priority = 'CRITICAL';

        return {
          id: msg.id || `live-msg-${index}`,
          sender: senderName,
          email: senderEmail,
          subject,
          timestamp: msg.receivedDateTime ? new Date(msg.receivedDateTime).toLocaleString() : 'Recently',
          priority,
          isVip,
          unread: msg.isRead === false,
          snippet,
          aiAnalysis: {
            sentiment: priority === 'CRITICAL' ? 'Urgent / High Consequence' : 'Action Needed',
            recommendedAction: priority === 'CRITICAL'
              ? 'Review immediately and draft executive response.'
              : 'Review and triage in order of urgency.',
            urgencyScore: priority === 'CRITICAL' ? 95 : priority === 'HIGH' ? 80 : 55
          }
        };
      });
    } catch (error) {
      console.warn('Live inbox hydration failed, continuing with demo data:', error);
    }
  }

  switchToLiveMode() {
    if (this.statusText) {
      const label = this.m365.account?.username || 'Connected';
      this.statusText.textContent = `CoS Agent (Live M365: ${label})`;
    }

    if (this.connectButton) {
      this.connectButton.textContent = 'Connected to Microsoft 365';
      this.connectButton.disabled = true;
    }

    this.renderM365Diagnostics();
  }

  renderDemoMode() {
    if (this.statusText) {
      this.statusText.textContent = 'CoS Agent (Demo Mode)';
    }

    if (this.connectButton) {
      this.connectButton.textContent = 'Connect Microsoft 365';
      this.connectButton.disabled = false;
    }

    this.renderM365Diagnostics();
  }

  toggleM365Diagnostics() {
    if (!this.diagnosticsPanel) return;

    const isOpen = this.diagnosticsPanel.style.display !== 'none';
    this.diagnosticsPanel.style.display = isOpen ? 'none' : 'block';

    if (this.diagnosticsButton) {
      this.diagnosticsButton.innerHTML = isOpen
        ? '<i class="fa-solid fa-stethoscope"></i> M365 Diagnostics'
        : '<i class="fa-solid fa-xmark"></i> Hide Diagnostics';
    }

    if (!isOpen) {
      this.renderM365Diagnostics();
    }
  }

  setDiagnosticValue(elementId, text, className = '') {
    const el = document.getElementById(elementId);
    if (!el) return;

    el.textContent = text;
    el.className = `m365-diag-value${el.classList.contains('mono') ? ' mono' : ''}`;
    if (className) {
      el.classList.add(className);
    }
  }

  renderM365Diagnostics() {
    const hasMsalLibrary = !!window.msal;
    const hasClient = !!this.m365.msalInstance;
    const hasAccount = !!this.m365.account;
    const liveMode = !!this.m365.isLiveMode;

    this.setDiagnosticValue('diag-msal-library', hasMsalLibrary ? 'Loaded' : 'Not loaded', hasMsalLibrary ? 'diag-ok' : 'diag-error');
    this.setDiagnosticValue('diag-msal-client', hasClient ? 'Initialized' : 'Not initialized', hasClient ? 'diag-ok' : 'diag-warn');
    this.setDiagnosticValue('diag-account', hasAccount ? (this.m365.account.username || this.m365.account.name || 'Signed in') : 'No active account', hasAccount ? 'diag-ok' : 'diag-warn');
    this.setDiagnosticValue('diag-live-mode', liveMode ? 'Enabled' : 'Disabled', liveMode ? 'diag-ok' : 'diag-warn');
    this.setDiagnosticValue('diag-scopes', this.m365.scopes.join(', '), '');
    this.setDiagnosticValue('diag-last-check', this.diagState.lastCheck, this.diagState.lastCheckClass);
    this.setDiagnosticValue('diag-deep-profile', this.diagState.deepChecks.profile.status, this.diagState.deepChecks.profile.className);
    this.setDiagnosticValue('diag-deep-mail', this.diagState.deepChecks.mail.status, this.diagState.deepChecks.mail.className);
    this.setDiagnosticValue('diag-deep-files', this.diagState.deepChecks.files.status, this.diagState.deepChecks.files.className);
    this.setDiagnosticValue('diag-report-preview', this.diagState.report, '');
  }

  async runM365Diagnostics() {
    const runButton = document.getElementById('btn-run-m365-diagnostics');
    if (runButton) {
      runButton.disabled = true;
      runButton.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Running...';
    }

    try {
      if (!window.msal) {
        this.diagState.lastCheck = 'MSAL library missing. Ensure msal-browser script is loaded.';
        this.diagState.lastCheckClass = 'diag-error';
        this.renderM365Diagnostics();
        return;
      }

      if (!this.m365.account) {
        this.diagState.lastCheck = 'No signed-in account. Click Connect Microsoft 365 first.';
        this.diagState.lastCheckClass = 'diag-warn';
        this.renderM365Diagnostics();
        return;
      }

      await this.m365.getAccessToken();
      const me = await this.m365.fetchGraphApi('/me');

      this.diagState.lastCheck = `Graph test succeeded for ${me?.userPrincipalName || me?.displayName || 'current user'} at ${new Date().toLocaleTimeString()}.`;
      this.diagState.lastCheckClass = 'diag-ok';
      this.diagState.report = this.buildDiagnosticsReport();
    } catch (error) {
      this.diagState.lastCheck = `Graph test failed at ${new Date().toLocaleTimeString()}: ${error.message || String(error)}`;
      this.diagState.lastCheckClass = 'diag-error';
      this.diagState.report = this.buildDiagnosticsReport();
    } finally {
      this.renderM365Diagnostics();
      if (runButton) {
        runButton.disabled = false;
        runButton.innerHTML = '<i class="fa-solid fa-play"></i> Run Graph Test';
      }
    }
  }

  async runM365DeepDiagnostics() {
    const runButton = document.getElementById('btn-run-m365-deep-diagnostics');
    if (runButton) {
      runButton.disabled = true;
      runButton.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Verifying...';
    }

    const mark = (key, status, className) => {
      this.diagState.deepChecks[key] = { status, className };
      this.renderM365Diagnostics();
    };

    try {
      if (!this.m365.account) {
        this.diagState.lastCheck = 'Cannot verify personal data access without sign-in. Connect Microsoft 365 first.';
        this.diagState.lastCheckClass = 'diag-warn';
        mark('profile', 'Blocked (not signed in)', 'diag-warn');
        mark('mail', 'Blocked (not signed in)', 'diag-warn');
        mark('files', 'Blocked (not signed in)', 'diag-warn');
        this.diagState.report = this.buildDiagnosticsReport();
        return;
      }

      mark('profile', 'Running...', 'diag-warn');
      mark('mail', 'Running...', 'diag-warn');
      mark('files', 'Running...', 'diag-warn');

      const profile = await this.safeGraphCheck('/me', (data) => {
        const name = data?.displayName || data?.userPrincipalName || 'Unknown';
        return `PASS (${name})`;
      });
      mark('profile', profile.status, profile.className);

      const mail = await this.safeGraphCheck('/me/messages?$top=1&$select=id,subject,receivedDateTime', (data) => {
        const count = Array.isArray(data?.value) ? data.value.length : 0;
        return `PASS (${count} message sample)`;
      });
      mark('mail', mail.status, mail.className);

      const files = await this.safeGraphCheck('/me/drive/root/children?$top=1&$select=id,name', (data) => {
        const count = Array.isArray(data?.value) ? data.value.length : 0;
        return `PASS (${count} item sample)`;
      });
      mark('files', files.status, files.className);

      const failed = [profile, mail, files].filter((r) => r.className === 'diag-error').length;
      const passed = 3 - failed;
      this.diagState.lastCheck = `Deep verification finished at ${new Date().toLocaleTimeString()}: ${passed}/3 checks passed.`;
      this.diagState.lastCheckClass = failed === 0 ? 'diag-ok' : 'diag-warn';
      this.diagState.report = this.buildDiagnosticsReport();
    } catch (error) {
      this.diagState.lastCheck = `Deep verification failed at ${new Date().toLocaleTimeString()}: ${error.message || String(error)}`;
      this.diagState.lastCheckClass = 'diag-error';
      this.diagState.report = this.buildDiagnosticsReport();
    } finally {
      this.renderM365Diagnostics();
      if (runButton) {
        runButton.disabled = false;
        runButton.innerHTML = '<i class="fa-solid fa-vial"></i> Verify My Data Access';
      }
    }
  }

  async safeGraphCheck(endpoint, onSuccess) {
    try {
      const data = await this.m365.fetchGraphApi(endpoint);
      return {
        status: onSuccess(data),
        className: 'diag-ok'
      };
    } catch (error) {
      return {
        status: `FAIL (${error.message || String(error)})`,
        className: 'diag-error'
      };
    }
  }

  buildDiagnosticsReport() {
    const accountName = this.m365.account?.username || this.m365.account?.name || 'None';
    const lines = [
      'Executive CoS M365 Diagnostics Report',
      `Timestamp: ${new Date().toLocaleString()}`,
      `MSAL Loaded: ${window.msal ? 'Yes' : 'No'}`,
      `MSAL Client Initialized: ${this.m365.msalInstance ? 'Yes' : 'No'}`,
      `Signed-In Account: ${accountName}`,
      `Live Mode: ${this.m365.isLiveMode ? 'Enabled' : 'Disabled'}`,
      `Scopes: ${this.m365.scopes.join(', ')}`,
      `Last Check: ${this.diagState.lastCheck}`,
      `Profile Access: ${this.diagState.deepChecks.profile.status}`,
      `Mailbox Access: ${this.diagState.deepChecks.mail.status}`,
      `OneDrive Access: ${this.diagState.deepChecks.files.status}`
    ];

    return lines.join('\n');
  }

  copyM365DiagnosticsReport() {
    const report = this.buildDiagnosticsReport();
    this.diagState.report = report;
    this.renderM365Diagnostics();

    navigator.clipboard.writeText(report)
      .then(() => {
        this.diagState.lastCheck = `Diagnostics report copied at ${new Date().toLocaleTimeString()}.`;
        this.diagState.lastCheckClass = 'diag-ok';
        this.renderM365Diagnostics();
      })
      .catch(() => {
        this.diagState.lastCheck = 'Failed to copy diagnostics report to clipboard.';
        this.diagState.lastCheckClass = 'diag-error';
        this.renderM365Diagnostics();
      });
  }

  renderAll() {
    this.briefing.renderBriefingView();
    this.schedule.renderScheduleView();
    this.triage.renderTriageView();
    this.renderAdminTasks();
  }

  updateMetrics() {
    const criticalCount = state.inbox.filter((m) => m.priority === 'CRITICAL' && m.unread).length;

    const criticalEl = document.getElementById('metric-critical-count');
    const meetingsEl = document.getElementById('metric-meetings-count');
    const tasksEl = document.getElementById('metric-tasks-count');

    if (criticalEl) criticalEl.textContent = String(criticalCount);
    if (meetingsEl) meetingsEl.textContent = String(state.schedule.length);
    if (tasksEl) tasksEl.textContent = String(state.adminTasks.length);
  }

  renderAdminTasks() {
    const tbody = document.getElementById('admin-tasks-tbody');
    if (!tbody) return;

    tbody.innerHTML = state.adminTasks.map((task) => `
      <tr>
        <td>${task.title}</td>
        <td>${task.assignedTo}</td>
        <td>${task.dueDate}</td>
        <td>${task.priority}</td>
        <td>${task.status}</td>
        <td><button class="btn btn-secondary" style="font-size: 0.75rem; padding: 0.25rem 0.5rem;">View</button></td>
      </tr>
    `).join('');
  }

  switchTab(tabName) {
    document.querySelectorAll('.nav-item[data-tab]').forEach((item) => {
      item.classList.toggle('active', item.getAttribute('data-tab') === tabName);
    });

    document.querySelectorAll('.tab-pane').forEach((pane) => {
      pane.classList.remove('active');
    });

    const activePane = document.getElementById(`tab-${tabName}`);
    if (activePane) activePane.classList.add('active');
  }

  runMasterDirective() {
    const query = this.directiveInput?.value?.trim();
    if (!query) return;
    this.orchestrator.runDirective(query);
  }

  refreshBriefing() {
    this.briefing.renderBriefingView();
  }

  optimizeSchedule() {
    this.switchTab('schedule');
    this.schedule.optimizeSchedule();
    this.updateMetrics();
  }

  triageCriticalEmail() {
    this.switchTab('triage');
    this.triage.filterInbox('CRITICAL');
  }

  filterInbox(filterType) {
    this.triage.filterInbox(filterType);
  }

  openEmailModal(msgId) {
    const msg = state.inbox.find((m) => m.id === msgId);
    if (!msg) return;

    const modal = document.getElementById('email-modal');
    const body = document.getElementById('email-modal-body');
    if (!modal || !body) return;

    body.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:0.75rem;">
        <div><strong>From:</strong> ${msg.sender} &lt;${msg.email}&gt;</div>
        <div><strong>Subject:</strong> ${msg.subject}</div>
        <div><strong>Priority:</strong> ${msg.priority}</div>
        <div><strong>Received:</strong> ${msg.timestamp}</div>
        <div style="padding:0.75rem; border:1px solid var(--border-color); border-radius:var(--radius-sm); background:rgba(255,255,255,0.02);">${msg.snippet}</div>
      </div>
    `;

    modal.classList.add('active');
    msg.unread = false;
    this.triage.renderTriageView();
    this.updateMetrics();
  }

  closeModal() {
    const modal = document.getElementById('email-modal');
    if (modal) modal.classList.remove('active');
  }

  draftAiReply(msgId) {
    this.triage.draftAiReply(msgId);
  }

  delegateTask(msgId) {
    this.triage.delegateTask(msgId);
    this.renderAdminTasks();
    this.updateMetrics();
  }

  copyReportText() {
    const area = document.getElementById('subagent-results-area');
    const text = area ? area.innerText : '';
    if (!text) return;

    navigator.clipboard.writeText(text)
      .then(() => alert('Executive report copied to clipboard.'))
      .catch(() => alert('Unable to copy report text.'));
  }

  draftResponseToChair() {
    alert('Draft response prepared and queued for Board Chair review.');
  }
}

window.CoSApp = new App();
window.addEventListener('DOMContentLoaded', () => window.CoSApp.init());
