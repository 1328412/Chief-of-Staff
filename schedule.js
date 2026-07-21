/**
 * Executive Chief of Staff - Calendar & Schedule Management Module
 */

import { state } from './state.js';

export class ScheduleManager {
  renderScheduleView() {
    const container = document.getElementById('schedule-list-area');
    if (!container) return;

    container.innerHTML = `
      <div class="timeline-list">
        ${state.schedule.map(evt => `
          <div class="timeline-item">
            <div class="timeline-time">${evt.time}</div>
            <div class="timeline-dot ${evt.conflict ? 'conflict' : (evt.time.includes('08:30') ? 'current' : '')}"></div>
            <div class="timeline-card">
              <div class="timeline-header">
                <span class="timeline-title">${evt.title}</span>
                <span class="timeline-badge badge-${evt.badgeType}">${evt.badge}</span>
              </div>
              <div class="timeline-meta">
                <span><i class="fa-regular fa-clock" style="margin-right:0.3rem;"></i>${evt.category}</span>
                <span><i class="fa-solid fa-location-dot" style="margin-right:0.3rem;"></i>${evt.location}</span>
                <span><i class="fa-solid fa-users" style="margin-right:0.3rem;"></i>${evt.attendees.join(', ')}</span>
              </div>

              ${evt.conflict ? `
                <div style="background: rgba(244, 63, 94, 0.1); border: 1px solid rgba(244, 63, 94, 0.3); border-radius: var(--radius-sm); padding: 0.5rem 0.75rem; margin-top: 0.5rem; font-size: 0.8rem; color: var(--accent-rose); display: flex; align-items: center; justify-content: space-between;">
                  <span><i class="fa-solid fa-triangle-exclamation" style="margin-right:0.4rem;"></i><strong>CONFLICT:</strong> ${evt.conflictReason}</span>
                  <button class="btn btn-secondary" style="font-size: 0.725rem; padding: 0.2rem 0.5rem;" onclick="window.CoSApp.optimizeSchedule()">Resolve</button>
                </div>
              ` : ''}

              <div class="timeline-prep-docs">
                <i class="fa-solid fa-paperclip"></i>
                <strong>AI Prep Briefs Attached:</strong> ${evt.prepDocs.join(' | ')}
              </div>
            </div>
          </div>
        `).join('')}
      </div>
    `;
  }

  optimizeSchedule() {
    // Resolve the conflict
    const conflictItem = state.schedule.find(s => s.conflict);
    if (conflictItem) {
      conflictItem.conflict = false;
      conflictItem.time = "10:00 AM - 11:15 AM";
      conflictItem.conflictReason = null;
      conflictItem.badge = "Optimized by CoS";
      conflictItem.badgeType = "focus";
      
      // Add buffer block
      state.schedule.splice(2, 0, {
        id: "evt-buffer",
        time: "11:15 AM - 11:45 AM",
        title: "AI CoS Focus Block & Board Prep Digest",
        category: "Executive Focus",
        badge: "Buffer Time",
        badgeType: "focus",
        location: "Private Executive Suite",
        attendees: ["Alex Mercer"],
        conflict: false,
        prepDocs: ["SharePoint: Q3 Executive Brief"],
        status: "Auto-Scheduled"
      });
    }

    this.renderScheduleView();
    alert("AI Schedule Optimization Applied! Conflicting 10:45 AM Vendor sign-off deferred to VP Ops, and a 30-min Executive Buffer Block was inserted before lunch.");
  }
}

export const scheduleManager = new ScheduleManager();
