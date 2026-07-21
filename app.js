/**
 * Executive Chief of Staff Agent - Main Application Controller
 */

import { state } from './state.js';
import { briefingManager } from './briefing.js';
import { scheduleManager } from './schedule.js';
import { triageManager } from './triage.js';
import { orchestrator } from './orchestrator.js';
import { m365Service } from './m365.js';

class CoSApplication {
  init() {
    console.log("Initializing Executive Chief of Staff Agent Application...");

    // Setup global window reference for inline handlers
    window.CoSApp = this;

    m365Service.init();
    this.bindNavigation();
    this.bindMasterDirectiveInput();
    this.renderAdminTasks();

    // Render Initial Views
    briefingManager.renderBriefingView();
    scheduleManager.renderScheduleView();
    triageManager.renderTriageView();
    orchestrator.init();

    this.updateExecutiveMetrics();
  }

  bindNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
      item.addEventListener('click', (e) => {
        e.preventDefault();
        const tabTarget = item.getAttribute('data-tab');
        this.switchTab(tabTarget);
      });
    });
  }

  switchTab(tabId) {
    if (!tabId) return;

    // Update Nav
    document.querySelectorAll('.nav-item').forEach(el => {
      if (el.getAttribute('data-tab') === tabId) {
        el.classList.add('active');
      } else {
        el.classList.remove('active');
      }
    });

    // Update Tab Panes
    document.querySelectorAll('.tab-pane').forEach(pane => {
      if (pane.id === `tab-${tabId}`) {
        pane.classList.add('active');
      } else {
        pane.classList.remove('active');
      }
    });

    // Refresh specific view if needed
    if (tabId === 'schedule') scheduleManager.renderScheduleView();
    if (tabId === 'triage') triageManager.renderTriageView();
    if (tabId === 'briefing') briefingManager.renderBriefingView();
  }

  bindMasterDirectiveInput() {
    const input = document.getElementById('master-directive-input');
    const sendBtn = document.getElementById('btn-send-directive');
    const presets = document.querySelectorAll('.preset-chip');

    const handleDirectiveSubmit = () => {
      const query = input.value.trim();
      if (!query) return;

      orchestrator.runDirective(query);
      input.value = '';
    };

    if (sendBtn) sendBtn.addEventListener('click', handleDirectiveSubmit);
    if (input) {
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') handleDirectiveSubmit();
      });
    }

    presets.forEach(chip => {
      chip.addEventListener('click', () => {
        const text = chip.getAttribute('data-prompt') || chip.innerText;
        if (input) input.value = text;
        handleDirectiveSubmit();
      });
    });
  }

  renderAdminTasks() {
    const container = document.getElementById('admin-tasks-tbody');
    if (!container) return;

    container.innerHTML = state.adminTasks.map(t => `
      <tr>
        <td><strong>${t.title}</strong></td>
        <td>${t.assignedTo}</td>
        <td>${t.dueDate}</td>
        <td><span class="priority-tag ${t.priority.toLowerCase()}">${t.priority}</span></td>
        <td><span style="font-size: 0.8rem; color: var(--accent-cyan); font-weight: 500;">${t.status}</span></td>
        <td>
          <button class="action-btn-sm" onclick="window.CoSApp.completeTask('${t.id}')">Complete</button>
        </td>
      </tr>
    `).join('');
  }

  completeTask(taskId) {
    const index = state.adminTasks.findIndex(t => t.id === taskId);
    if (index !== -1) {
      state.adminTasks.splice(index, 1);
      this.renderAdminTasks();
      this.updateExecutiveMetrics();
    }
  }

  updateExecutiveMetrics() {
    const criticalCount = state.inbox.filter(m => m.priority === 'CRITICAL').length;
    const taskCount = state.adminTasks.length;
    const meetingCount = state.schedule.length;

    const elCritical = document.getElementById('metric-critical-count');
    const elTasks = document.getElementById('metric-tasks-count');
    const elMeetings = document.getElementById('metric-meetings-count');

    if (elCritical) elCritical.textContent = criticalCount;
    if (elTasks) elTasks.textContent = taskCount;
    if (elMeetings) elMeetings.textContent = meetingCount;
  }

  async connectM365() {
    const account = await m365Service.login();
    if (account) {
      const statusText = document.getElementById('cos-status-text');
      const connectBtn = document.getElementById('btn-m365-connect');
      if (statusText) statusText.textContent = `Connected: ${account.name || account.username}`;
      if (connectBtn) {
        connectBtn.innerHTML = `<i class="fa-solid fa-circle-check"></i> Connected (${account.username})`;
        connectBtn.classList.remove('btn-cyan');
        connectBtn.classList.add('btn-secondary');
      }

      state.executive.name = account.name || state.executive.name;
      briefingManager.renderBriefingView();
      alert(`Successfully connected to Microsoft 365 as ${account.username}!\n\nLive Graph API sync is now active for your Outlook Calendar, Inbox, SharePoint sites, and OneDrive.`);
    }
  }

  // Quick Action Helpers
  refreshBriefing() {
    briefingManager.renderBriefingView();
    alert("AI Daily Briefing refreshed with latest real-time feeds!");
  }

  triageCriticalEmail() {
    this.switchTab('triage');
    triageManager.filterInbox('CRITICAL');
  }

  optimizeSchedule() {
    scheduleManager.optimizeSchedule();
  }

  filterInbox(filter) {
    triageManager.filterInbox(filter);
  }

  draftAiReply(id) {
    triageManager.draftAiReply(id);
  }

  delegateTask(id) {
    triageManager.delegateTask(id);
  }

  copyReportText() {
    alert("Executive Briefing copied to clipboard!");
  }

  draftResponseToChair() {
    alert("Drafting response to Board Chair Victoria Vance:\n\n'Dear Chairwoman Vance,\nPer your request, the Chief of Staff agent retrieved Clause 4.2 of the Project Apex Acquisition Term Sheet from OneDrive. The escrow guarantee is capped at $8.5M with full indemnity protections. Attached is the revised memorandum.'");
  }

  openEmailModal(msgId) {
    const msg = state.inbox.find(m => m.id === msgId);
    if (!msg) return;

    const modal = document.getElementById('email-modal');
    const modalContent = document.getElementById('email-modal-body');
    if (!modal || !modalContent) return;

    modalContent.innerHTML = `
      <div style="margin-bottom: 1rem;">
        <div style="font-size: 0.8rem; color: var(--text-dim);">From: ${msg.email}</div>
        <div style="font-size: 1.1rem; font-weight: 700; color: #fff; margin-top: 0.25rem;">${msg.subject}</div>
        <div style="font-size: 0.775rem; color: var(--text-muted); margin-top: 0.2rem;">Received: ${msg.timestamp}</div>
      </div>
      <div style="background: rgba(0,0,0,0.3); padding: 1rem; border-radius: var(--radius-md); border: 1px solid var(--border-color); color: #e2e8f0; font-size: 0.9rem; line-height: 1.6; margin-bottom: 1.25rem;">
        ${msg.snippet}
        <br/><br/>
        Please let me know your thoughts so we can finalize the filings accordingly.
      </div>
      <div style="background: rgba(99, 102, 241, 0.1); border: 1px solid rgba(99, 102, 241, 0.3); border-radius: var(--radius-md); padding: 0.85rem;">
        <div style="font-size: 0.825rem; font-weight: 700; color: var(--primary);"><i class="fa-solid fa-brain" style="margin-right:0.4rem;"></i> Chief of Staff AI Assessment:</div>
        <div style="font-size: 0.8rem; color: var(--text-main); margin-top: 0.3rem;">Urgency Score: ${msg.aiAnalysis.urgencyScore}/100 | Sentiment: ${msg.aiAnalysis.sentiment}</div>
        <div style="font-size: 0.8rem; color: var(--accent-cyan); margin-top: 0.2rem;">Recommended: ${msg.aiAnalysis.recommendedAction}</div>
      </div>
      <div style="margin-top: 1.25rem; display: flex; gap: 0.5rem; justify-content: flex-end;">
        <button class="btn btn-secondary" onclick="window.CoSApp.closeModal()">Close</button>
        <button class="btn btn-primary" onclick="window.CoSApp.draftAiReply('${msg.id}'); window.CoSApp.closeModal();">Execute Recommended AI Action</button>
      </div>
    `;

    modal.classList.add('active');
  }

  closeModal() {
    const modal = document.getElementById('email-modal');
    if (modal) modal.classList.remove('active');
  }
}

// Initialize on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
  const app = new CoSApplication();
  app.init();
});
