/**
 * Executive Chief of Staff - Inbox Triage & Urgent Communications Manager
 */

import { state } from './state.js';
import { orchestrator } from './orchestrator.js';

export class TriageManager {
  constructor() {
    this.currentFilter = 'ALL';
  }

  renderTriageView() {
    const container = document.getElementById('inbox-list-area');
    if (!container) return;

    let items = state.inbox;
    if (this.currentFilter === 'CRITICAL') {
      items = state.inbox.filter(m => m.priority === 'CRITICAL');
    } else if (this.currentFilter === 'VIP') {
      items = state.inbox.filter(m => m.isVip);
    } else if (this.currentFilter === 'ACTION') {
      items = state.inbox.filter(m => m.priority === 'HIGH' || m.priority === 'CRITICAL');
    }

    container.innerHTML = `
      <div class="triage-filters">
        <button class="filter-btn ${this.currentFilter === 'ALL' ? 'active' : ''}" onclick="window.CoSApp.filterInbox('ALL')">All Messages (${state.inbox.length})</button>
        <button class="filter-btn ${this.currentFilter === 'CRITICAL' ? 'active' : ''}" onclick="window.CoSApp.filterInbox('CRITICAL')">🔥 Critical Priority (${state.inbox.filter(m => m.priority === 'CRITICAL').length})</button>
        <button class="filter-btn ${this.currentFilter === 'VIP' ? 'active' : ''}" onclick="window.CoSApp.filterInbox('VIP')">⭐ VIP Contacts (${state.inbox.filter(m => m.isVip).length})</button>
        <button class="filter-btn ${this.currentFilter === 'ACTION' ? 'active' : ''}" onclick="window.CoSApp.filterInbox('ACTION')">⚡ Action Needed</button>
      </div>

      <div class="email-list">
        ${items.map(msg => `
          <div class="email-card ${msg.unread ? 'unread' : ''} ${msg.priority.toLowerCase()}" onclick="window.CoSApp.openEmailModal('${msg.id}')">
            <div class="email-header-line">
              <div class="email-sender">
                ${msg.isVip ? '<i class="fa-solid fa-star" style="color: var(--accent-amber); font-size: 0.8rem;"></i>' : ''}
                <span>${msg.sender}</span>
                <span class="priority-tag ${msg.priority.toLowerCase()}">${msg.priority}</span>
              </div>
              <span class="email-time">${msg.timestamp}</span>
            </div>

            <div class="email-subject">${msg.subject}</div>
            <div class="email-snippet">${msg.snippet}</div>

            <div class="email-ai-triage">
              <div class="ai-recommendation">
                <i class="fa-solid fa-robot" style="color: var(--primary);"></i>
                <span><strong>AI Triage:</strong> ${msg.aiAnalysis.recommendedAction}</span>
              </div>
              <div class="email-actions" onclick="event.stopPropagation();">
                <button class="action-btn-sm" onclick="window.CoSApp.draftAiReply('${msg.id}')">
                  <i class="fa-solid fa-wand-magic-sparkles"></i> AI Reply
                </button>
                <button class="action-btn-sm" onclick="window.CoSApp.delegateTask('${msg.id}')">
                  <i class="fa-solid fa-share"></i> Delegate
                </button>
              </div>
            </div>
          </div>
        `).join('')}
      </div>
    `;
  }

  filterInbox(filterType) {
    this.currentFilter = filterType;
    this.renderTriageView();
  }

  draftAiReply(msgId) {
    const msg = state.inbox.find(m => m.id === msgId);
    if (!msg) return;

    if (msg.priority === 'CRITICAL') {
      // Trigger sub-agent document retrieval directive
      orchestrator.runDirective(`Find escrow guarantee Clause 4.2 in Project Apex acquisition term sheet from OneDrive and draft formal response to Chair ${msg.sender}`);
    } else {
      alert(`AI Chief of Staff is drafting executive response for ${msg.sender}:\n\n"Dear ${msg.sender.split(' ')[0]},\nThank you for the update on '${msg.subject}'. I have reviewed the materials and approved the recommended next steps.\n\nBest regards,\nAlex Mercer (via Chief of Staff AI)"`);
    }
  }

  delegateTask(msgId) {
    const msg = state.inbox.find(m => m.id === msgId);
    if (!msg) return;

    state.adminTasks.unshift({
      id: `task-${Date.now().toString().slice(-3)}`,
      title: `Follow up on: ${msg.subject}`,
      assignedTo: "Marcus Brody (CFO)",
      dueDate: "Tomorrow, 12:00 PM",
      priority: msg.priority === 'CRITICAL' ? 'High' : 'Medium',
      status: "Delegated to CFO",
      source: `Email from ${msg.sender}`
    });

    alert(`Task delegated to Marcus Brody (CFO) and added to Administrative Task Matrix!`);
  }
}

export const triageManager = new TriageManager();
