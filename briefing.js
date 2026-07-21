/**
 * Executive Chief of Staff - Daily Briefing Coordinator Module
 */

import { state } from './state.js';

export class BriefingManager {
  constructor() {
    this.speechSynth = window.speechSynthesis;
    this.isSpeaking = false;
  }

  renderBriefingView() {
    const container = document.getElementById('briefing-container');
    if (!container) return;

    const criticalEmail = state.inbox.find(m => m.priority === "CRITICAL");
    const conflictSchedule = state.schedule.find(s => s.conflict);
    const pendingTask = state.adminTasks[0];

    container.innerHTML = `
      <div class="glass-card briefing-hero glow-animated mb-4">
        <div class="card-header">
          <div class="card-title">
            <i class="fa-solid fa-crown" style="color: var(--accent-amber);"></i>
            <span>Executive Morning Briefing — ${new Date().toLocaleDateString('en-US', { weekday: 'long', month: 'short', day: 'numeric', year: 'numeric' })}</span>
          </div>
          <div style="display: flex; gap: 0.5rem;">
            <button class="btn btn-secondary" id="btn-audio-brief">
              <i class="fa-solid fa-volume-high"></i>
              <span id="audio-btn-label">Play Audio Briefing</span>
            </button>
            <button class="btn btn-primary" onclick="window.CoSApp.refreshBriefing()">
              <i class="fa-solid fa-arrows-rotate"></i> Refresh AI Synthesis
            </button>
          </div>
        </div>

        <p class="briefing-text">
          Good morning, <strong>${state.executive.name}</strong>. Your Executive Chief of Staff Agent has completed today's overnight triage across your schedule, inbox, and document repositories. You have <strong>${state.schedule.length} meetings</strong> scheduled today, <strong>1 calendar conflict</strong> requiring resolution, and <strong>1 CRITICAL message</strong> from Board Chair Victoria Vance.
        </p>

        <div class="key-takeaways-list">
          <div class="takeaway-item urgent">
            <div class="takeaway-content">
              <h5><i class="fa-solid fa-triangle-exclamation" style="color: var(--accent-rose); margin-right: 0.4rem;"></i> URGENT: Board Chair Clarification Needed (Before 11:00 AM)</h5>
              <p>${criticalEmail ? criticalEmail.snippet : 'Review Escrow guarantees for Project Apex acquisition term sheet in OneDrive.'}</p>
            </div>
            <button class="btn btn-secondary" style="font-size: 0.75rem; padding: 0.3rem 0.6rem; margin-left: auto;" onclick="window.CoSApp.triageCriticalEmail()">
              Triage Now
            </button>
          </div>

          <div class="takeaway-item info">
            <div class="takeaway-content">
              <h5><i class="fa-solid fa-calendar-check" style="color: var(--accent-cyan); margin-right: 0.4rem;"></i> Schedule Optimization Alert: 10:00 AM Conflict</h5>
              <p>${conflictSchedule ? conflictSchedule.conflictReason : 'Double-booked at 10:45 AM.'}</p>
            </div>
            <button class="btn btn-secondary" style="font-size: 0.75rem; padding: 0.3rem 0.6rem; margin-left: auto;" onclick="window.CoSApp.optimizeSchedule()">
              Auto-Resolve Conflict
            </button>
          </div>

          <div class="takeaway-item success">
            <div class="takeaway-content">
              <h5><i class="fa-solid fa-check-double" style="color: var(--accent-emerald); margin-right: 0.4rem;"></i> Sub-Agent Triage Status</h5>
              <p>SharePoint & OneDrive file sub-agents indexed 7 confidential executive documents and verified EU AI Act compliance sign-offs.</p>
            </div>
            <button class="btn btn-secondary" style="font-size: 0.75rem; padding: 0.3rem 0.6rem; margin-left: auto;" onclick="window.CoSApp.switchTab('subagents')">
              View Agent Logs
            </button>
          </div>
        </div>
      </div>
    `;

    // Attach audio button event
    const audioBtn = document.getElementById('btn-audio-brief');
    if (audioBtn) {
      audioBtn.addEventListener('click', () => this.toggleAudioBriefing());
    }
  }

  toggleAudioBriefing() {
    if (this.isSpeaking) {
      this.speechSynth.cancel();
      this.isSpeaking = false;
      this.updateAudioButton(false);
      return;
    }

    const textToSpeak = `Good morning, ${state.executive.name}. This is your Chief of Staff briefing for ${new Date().toLocaleDateString()}. You have 4 meetings today including the 10:00 AM Board Strategy session. There is one critical email from Board Chair Victoria Vance regarding the acquisition term sheet escrow guarantee. I recommend dispatching the OneDrive sub-agent to retrieve clause 4.2. Have a productive day.`;

    const utterance = new SpeechSynthesisUtterance(textToSpeak);
    utterance.rate = 1.0;
    utterance.pitch = 1.0;

    utterance.onend = () => {
      this.isSpeaking = false;
      this.updateAudioButton(false);
    };

    utterance.onerror = () => {
      this.isSpeaking = false;
      this.updateAudioButton(false);
    };

    this.speechSynth.speak(utterance);
    this.isSpeaking = true;
    this.updateAudioButton(true);
  }

  updateAudioButton(speaking) {
    const label = document.getElementById('audio-btn-label');
    const btn = document.getElementById('btn-audio-brief');
    if (!btn || !label) return;

    if (speaking) {
      label.textContent = "Stop Audio Briefing";
      btn.classList.add('btn-cyan');
      btn.classList.remove('btn-secondary');
    } else {
      label.textContent = "Play Audio Briefing";
      btn.classList.remove('btn-cyan');
      btn.classList.add('btn-secondary');
    }
  }
}

export const briefingManager = new BriefingManager();
