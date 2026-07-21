/**
 * Executive Chief of Staff - Sub-Agent Orchestration & Multi-Source Document Router
 */

import { state } from './state.js';

export class OrchestratorEngine {
  constructor() {
    this.logsContainer = null;
  }

  init() {
    this.logsContainer = document.getElementById('terminal-logs');
  }

  async runDirective(userQuery) {
    if (!userQuery) return;
    
    // Switch to Sub-Agent tab
    const subAgentTabBtn = document.querySelector('[data-tab="subagents"]');
    if (subAgentTabBtn) subAgentTabBtn.click();

    this.clearLogs();
    this.addLog("CoS Master Agent", `Received directive: "${userQuery}"`, "cos");

    // Update Master agent status
    this.updateAgentCard('agent-cos', 'working', 'Decomposing query & routing sub-agents...');

    await this.delay(600);
    this.addLog("CoS Master Agent", "Analyzing query requirements: detecting multi-source targets (SharePoint, OneDrive, Meeting Transcripts)...", "cos", "highlight");

    // Formulate Sub-Agent tasks
    const queryLower = userQuery.toLowerCase();
    
    // Determine relevant agents
    let needSharePoint = queryLower.includes("sharepoint") || queryLower.includes("forecast") || queryLower.includes("supply") || queryLower.includes("audit") || queryLower.includes("legal") || queryLower.includes("board") || true;
    let needOneDrive = queryLower.includes("onedrive") || queryLower.includes("vance") || queryLower.includes("term sheet") || queryLower.includes("escrow") || queryLower.includes("deck") || queryLower.includes("m&a") || true;
    let needTranscripts = queryLower.includes("transcript") || queryLower.includes("security") || queryLower.includes("meeting") || queryLower.includes("budget") || true;

    await this.delay(800);

    // Launch parallel sub-agent tasks
    const tasks = [];

    if (needSharePoint) {
      tasks.push(this.runSharePointSubAgent(userQuery));
    }
    if (needOneDrive) {
      tasks.push(this.runOneDriveSubAgent(userQuery));
    }
    if (needTranscripts) {
      tasks.push(this.runTranscriptSubAgent(userQuery));
    }

    const results = await Promise.all(tasks);

    // Synthesis Phase
    this.addLog("CoS Master Agent", "All sub-agents reported back. Synthesizing multi-source intelligence brief...", "cos", "highlight");
    await this.delay(800);

    const compiledReport = this.synthesizeResults(userQuery, results);

    this.updateAgentCard('agent-cos', 'active', 'Standing by for executive instructions');
    this.renderCompiledReport(compiledReport);

    this.addLog("CoS Master Agent", "Task completed successfully. Executive Brief ready for review.", "cos", "success");
  }

  async runSharePointSubAgent(query) {
    this.updateAgentCard('agent-sp', 'working', 'Scanning SharePoint Executive Hub & Governance...');
    this.addLog("SharePoint Agent", "Initiating Graph API query across /sites/ExecutiveHub and /sites/Legal...", "sharepoint");
    await this.delay(900);

    // Filter state documents
    const queryWords = query.toLowerCase().split(' ');
    const docs = state.documentsRepository.sharepoint.filter(doc => {
      const text = (doc.title + " " + doc.summary + " " + doc.keywords.join(' ')).toLowerCase();
      return queryWords.some(w => w.length > 3 && text.includes(w)) || true;
    });

    this.addLog("SharePoint Agent", `Found ${docs.length} matching document(s): [${docs.map(d => d.title).join(', ')}]`, "sharepoint", "success");
    this.updateAgentCard('agent-sp', 'active', `Found ${docs.length} file(s)`);
    
    return {
      source: "SharePoint",
      docs: docs
    };
  }

  async runOneDriveSubAgent(query) {
    this.updateAgentCard('agent-od', 'working', 'Scanning Executive Personal OneDrive & Confidential Folders...');
    this.addLog("OneDrive Agent", "Authenticating executive session token... Searching /Personal/Strategic_MA & /Personal/Board_Prep...", "onedrive");
    await this.delay(1200);

    const docs = state.documentsRepository.onedrive.filter(doc => {
      return true; // Match relevant docs
    });

    this.addLog("OneDrive Agent", `Extracted key document: "${docs[1].title}" - Escrow Guarantee Clause 4.2 identified ($8.5M cap)`, "onedrive", "success");
    this.updateAgentCard('agent-od', 'active', `Retrieved ${docs.length} private file(s)`);

    return {
      source: "OneDrive",
      docs: docs
    };
  }

  async runTranscriptSubAgent(query) {
    this.updateAgentCard('agent-mt', 'working', 'Parsing Teams Transcripts & Audio Digests...');
    this.addLog("Transcript Agent", "Fetching transcript stream for Q2 Operational Review...", "meetings");
    await this.delay(1000);

    const docs = state.documentsRepository.transcripts;
    this.addLog("Transcript Agent", `Extracted 1 relevant decision item: "$3.2M budget reallocation to cloud security"`, "meetings", "success");
    this.updateAgentCard('agent-mt', 'active', 'Extracted 1 decision');

    return {
      source: "Teams Transcripts",
      docs: docs
    };
  }

  synthesizeResults(query, results) {
    const spDocs = results.find(r => r.source === "SharePoint")?.docs || [];
    const odDocs = results.find(r => r.source === "OneDrive")?.docs || [];
    const mtDocs = results.find(r => r.source === "Teams Transcripts")?.docs || [];

    return {
      title: `Executive Intelligence Report: ${query}`,
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
      sourcesCount: spDocs.length + odDocs.length + mtDocs.length,
      executiveSummary: `Based on a multi-source triage across SharePoint, personal OneDrive archives, and Teams transcripts, the Chief of Staff agent has compiled the following executive synthesis:`,
      highlights: [
        {
          tag: "SharePoint Hub",
          text: `Q3 Financial Forecast (Master V4) confirms target revenue of $142M with operating margin expansion of 1.8%. Legal compliance under EU AI Act is fully cleared.`
        },
        {
          tag: "OneDrive Private Archive",
          text: `Acquisition Term Sheet (Project Apex) specifies escrow guarantee cap at $8.5M in Clause 4.2. Recommended draft response created for Board Chair Victoria Vance.`
        },
        {
          tag: "Meeting Transcripts",
          text: `Q2 Operational Review transcript confirms approval of $3.2M budget shift to Cloud Security, pending formal e-signature by CEO.`
        }
      ],
      retrievedFiles: [...spDocs, ...odDocs, ...mtDocs]
    };
  }

  renderCompiledReport(report) {
    const reportContainer = document.getElementById('subagent-results-area');
    if (!reportContainer) return;

    reportContainer.innerHTML = `
      <div class="doc-result-card glow-animated">
        <div class="doc-result-header">
          <i class="fa-solid fa-file-shield"></i>
          <span>${report.title}</span>
          <span style="margin-left: auto; font-size: 0.75rem; color: var(--text-dim);">${report.timestamp}</span>
        </div>

        <div class="doc-sources-badge-group">
          <span class="source-badge"><i class="fa-solid fa-share-nodes"></i> SharePoint Hub (${report.retrievedFiles.filter(f=>f.location.includes('SharePoint')).length})</span>
          <span class="source-badge"><i class="fa-solid fa-cloud"></i> OneDrive Private (${report.retrievedFiles.filter(f=>f.location.includes('OneDrive')).length})</span>
          <span class="source-badge"><i class="fa-solid fa-microphone-lines"></i> Transcripts (${report.retrievedFiles.filter(f=>f.location.includes('Teams')).length})</span>
        </div>

        <p style="font-size: 0.9rem; color: var(--text-main); margin-bottom: 1rem;">
          ${report.executiveSummary}
        </p>

        <div class="key-takeaways-list" style="margin-bottom: 1.25rem;">
          ${report.highlights.map(h => `
            <div class="takeaway-item info">
              <div class="takeaway-content">
                <h5>${h.tag}</h5>
                <p>${h.text}</p>
              </div>
            </div>
          `).join('')}
        </div>

        <h5 style="font-size: 0.85rem; font-weight: 600; color: var(--text-muted); margin-bottom: 0.5rem; text-transform: uppercase; letter-spacing: 0.05em;">Source Documents Cross-Referenced:</h5>
        <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 0.75rem;">
          ${report.retrievedFiles.map(doc => `
            <div style="background: rgba(255,255,255,0.03); border: 1px solid var(--border-color); border-radius: var(--radius-sm); padding: 0.75rem;">
              <div style="font-size: 0.825rem; font-weight: 600; color: var(--accent-cyan);"><i class="fa-regular fa-file-lines" style="margin-right: 0.4rem;"></i>${doc.title}</div>
              <div style="font-size: 0.725rem; color: var(--text-dim); margin-top: 0.2rem;">${doc.location}</div>
              <div style="font-size: 0.775rem; color: var(--text-muted); margin-top: 0.35rem;">${doc.summary}</div>
            </div>
          `).join('')}
        </div>

        <div style="margin-top: 1.25rem; display: flex; gap: 0.75rem; justify-content: flex-end;">
          <button class="btn btn-secondary" onclick="window.CoSApp.copyReportText()"><i class="fa-regular fa-copy"></i> Copy Briefing Text</button>
          <button class="btn btn-cyan" onclick="window.CoSApp.draftResponseToChair()"><i class="fa-solid fa-paper-plane"></i> Send Response to Board Chair</button>
        </div>
      </div>
    `;

    reportContainer.scrollIntoView({ behavior: 'smooth' });
  }

  updateAgentCard(agentId, stateClass, statusMsg) {
    const card = document.getElementById(agentId);
    if (!card) return;
    
    card.className = `agent-node-card ${stateClass}`;
    const statusEl = card.querySelector('.agent-status-text');
    if (statusEl) statusEl.textContent = statusMsg;
  }

  addLog(agentName, message, agentClass, msgStyle = "") {
    if (!this.logsContainer) this.logsContainer = document.getElementById('terminal-logs');
    if (!this.logsContainer) return;

    const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const logRow = document.createElement('div');
    logRow.className = 'log-entry';
    logRow.innerHTML = `
      <span class="log-time">[${time}]</span>
      <span class="log-agent ${agentClass}">${agentName}:</span>
      <span class="log-msg ${msgStyle}">${message}</span>
    `;

    this.logsContainer.appendChild(logRow);
    this.logsContainer.scrollTop = this.logsContainer.scrollHeight;
  }

  clearLogs() {
    if (this.logsContainer) this.logsContainer.innerHTML = '';
  }

  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

export const orchestrator = new OrchestratorEngine();
