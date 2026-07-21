/**
 * Executive Chief of Staff - State Management Engine & Data Store
 */

export const state = {
  executive: {
    name: "Alex Mercer",
    role: "Chief Executive Officer",
    title: "Executive Office",
    avatar: "AM",
    vipList: ["Victoria Vance (Board Chair)", "Marcus Brody (CFO)", "Elena Rostova (CTO)", "David Sterling (Legal Counsel)"],
    preferences: {
      briefingTime: "08:00 AM",
      autoTriageStrictness: "High",
      conciseMode: true
    }
  },

  schedule: [
    {
      id: "evt-101",
      time: "08:30 AM - 09:15 AM",
      title: "Daily Executive Standup & CoS Sync",
      category: "Internal Sync",
      badge: "Focus Block",
      badgeType: "focus",
      location: "Executive Boardroom / Teams",
      attendees: ["Alex Mercer", "Chief of Staff AI", "VP Ops"],
      conflict: false,
      prepDocs: ["SharePoint: Executive Ops Dashboard", "OneDrive: Daily Action List"],
      status: "Upcoming"
    },
    {
      id: "evt-102",
      time: "10:00 AM - 11:30 AM",
      title: "Q3 Board Strategy & Financial Review",
      category: "Board Meeting",
      badge: "VIP High Priority",
      badgeType: "vip",
      location: "Main Auditorium & Hybrid",
      attendees: ["Victoria Vance (Chair)", "Marcus Brody (CFO)", "Alex Mercer"],
      conflict: true,
      conflictReason: "Double-booked with Operations Vendor Sign-off at 10:45 AM",
      prepDocs: ["OneDrive: Board_Meeting_Deck_2026.pptx", "SharePoint: Q3_Financial_Forecast_V4.pdf"],
      status: "Conflict Flagged"
    },
    {
      id: "evt-103",
      time: "01:30 PM - 02:15 PM",
      title: "M&A Intelligence Sync - Enterprise Expansion",
      category: "Strategic M&A",
      badge: "Confidential",
      badgeType: "board",
      location: "Secure Room 4B",
      attendees: ["Elena Rostova (CTO)", "Corporate Dev Team"],
      conflict: false,
      prepDocs: ["OneDrive: Confidential_Acquisition_TermSheet.docx"],
      status: "Upcoming"
    },
    {
      id: "evt-104",
      time: "03:30 PM - 04:15 PM",
      title: "Global Supply Chain Audit & Risk Review",
      category: "Operations",
      badge: "Action Needed",
      badgeType: "vip",
      location: "Virtual Conference Room A",
      attendees: ["Head of Supply Chain", "Risk Officer"],
      conflict: false,
      prepDocs: ["SharePoint: Supply_Chain_Disruption_Audit_2026.xlsx"],
      status: "Upcoming"
    }
  ],

  inbox: [
    {
      id: "msg-201",
      sender: "Victoria Vance (Board Chair)",
      email: "v.vance@enterprise-board.org",
      subject: "URGENT: Board Committee Feedback on Q3 Acquisition Target",
      timestamp: "07:42 AM Today",
      priority: "CRITICAL",
      isVip: true,
      unread: true,
      snippet: "Alex, the Audit Committee reviewed the term sheet last night. We need clarification on the escrow guarantees before 11:00 AM...",
      aiAnalysis: {
        sentiment: "Urgent / High Consequence",
        recommendedAction: "Dispatch OneDrive Agent to fetch Term Sheet Clause 4.2 & draft formal response to Chair.",
        urgencyScore: 98
      }
    },
    {
      id: "msg-202",
      sender: "Marcus Brody (CFO)",
      email: "m.brody@company.com",
      subject: "Updated Q3 Financial Forecast & Tax Adjustment Model",
      timestamp: "08:15 AM Today",
      priority: "HIGH",
      isVip: true,
      unread: true,
      snippet: "Attached is the latest revision from our tax strategy advisors. Net operating margins increased by 1.8%...",
      aiAnalysis: {
        sentiment: "Action Needed",
        recommendedAction: "Cross-reference with SharePoint Q3 Master Budget and attach to 10:00 AM Board prep packet.",
        urgencyScore: 88
      }
    },
    {
      id: "msg-203",
      sender: "David Sterling (Legal Counsel)",
      email: "d.sterling@company.com",
      subject: "SharePoint Governance Compliance Clearance Sign-off",
      timestamp: "Yesterday, 05:40 PM",
      priority: "MEDIUM",
      isVip: true,
      unread: false,
      snippet: "The regulatory filing for EU AI Act compliance is ready for executive signature in the legal repository...",
      aiAnalysis: {
        sentiment: "Routine Approval",
        recommendedAction: "Route to Administrative Task Matrix for e-Signature approval.",
        urgencyScore: 65
      }
    },
    {
      id: "msg-204",
      sender: "Corporate IT Support",
      email: "it-notifications@company.com",
      subject: "Scheduled System Maintenance Window - Saturday 2 AM",
      timestamp: "Yesterday, 02:15 PM",
      priority: "LOW",
      isVip: false,
      unread: false,
      snippet: "Please be advised that SharePoint Online and OneDrive sync will undergo routine maintenance...",
      aiAnalysis: {
        sentiment: "Informational",
        recommendedAction: "Auto-archive or file under IT Updates.",
        urgencyScore: 20
      }
    }
  ],

  subAgents: [
    {
      id: "agent-sp",
      name: "SharePoint File Sub-Agent",
      type: "SharePoint Specialist",
      icon: "fa-folder-tree",
      color: "cyan",
      scope: "Corporate SharePoint Hubs, Governance Repositories, Department Portals",
      status: "Idle",
      activeTask: "Listening for document queries..."
    },
    {
      id: "agent-od",
      name: "OneDrive Personal File Sub-Agent",
      type: "OneDrive Specialist",
      icon: "fa-cloud",
      color: "emerald",
      scope: "Executive Personal OneDrive, Confidential Drafts, Slide Decks, Private Notes",
      status: "Idle",
      activeTask: "Listening for document queries..."
    },
    {
      id: "agent-mt",
      name: "Teams & Meeting Transcript Sub-Agent",
      type: "Transcript Specialist",
      icon: "fa-file-audio",
      color: "purple",
      scope: "Recorded Meeting Transcripts, Action Item Extracts, Live Stream Logs",
      status: "Idle",
      activeTask: "Listening for query dispatch..."
    }
  ],

  documentsRepository: {
    sharepoint: [
      {
        id: "sp-doc-1",
        title: "Q3_Financial_Forecast_Master_V4.pdf",
        location: "SharePoint > Executive Hub > Financials 2026",
        modified: "2026-07-20 18:30",
        author: "Finance Team",
        summary: "Comprehensive Q3 financial analysis showing revenue target of $142M (+14% YoY). Operating margin at 24.2%.",
        keywords: ["Q3", "Financial", "Forecast", "Revenue", "Budget"]
      },
      {
        id: "sp-doc-2",
        title: "Supply_Chain_Disruption_Audit_2026.xlsx",
        location: "SharePoint > Ops Hub > Global Logistics",
        modified: "2026-07-19 14:15",
        author: "Global Operations",
        summary: "Risk evaluation of Asia-Pacific freight routes. Highlights 3 potential bottleneck suppliers with risk mitigation plans.",
        keywords: ["Supply Chain", "Audit", "Logistics", "Vendor", "Risk"]
      },
      {
        id: "sp-doc-3",
        title: "EU_AI_Act_Compliance_Framework_2026.docx",
        location: "SharePoint > Legal Hub > Regulatory Compliance",
        modified: "2026-07-21 07:10",
        author: "David Sterling (Legal)",
        summary: "Legal compliance assessment for enterprise AI deployments. All internal agents certified under Tier 2 transparency guidelines.",
        keywords: ["Compliance", "Legal", "AI", "EU", "Governance"]
      }
    ],
    onedrive: [
      {
        id: "od-doc-1",
        title: "Board_Meeting_Deck_2026_FINAL.pptx",
        location: "OneDrive > Alex Mercer > Private > Board Prep",
        modified: "2026-07-21 06:45",
        author: "Alex Mercer",
        summary: "Key presentation deck for Board Chair Victoria Vance. Covers strategic acquisitions, Q3 targets, and executive hiring roadmap.",
        keywords: ["Board", "Deck", "Strategy", "Vance", "Acquisition"]
      },
      {
        id: "od-doc-2",
        title: "Confidential_Acquisition_TermSheet_ProjectApex.docx",
        location: "OneDrive > Alex Mercer > Strategic M&A",
        modified: "2026-07-21 07:15",
        author: "M&A Advisory Group",
        summary: "Binding term sheet for Project Apex acquisition. Escrow guarantee clause 4.2 specifies $8.5M indemnity cap.",
        keywords: ["Acquisition", "Term Sheet", "Project Apex", "Escrow", "M&A", "Vance"]
      },
      {
        id: "od-doc-3",
        title: "CEO_Executive_Notes_Q3_Priorities.txt",
        location: "OneDrive > Alex Mercer > Personal Notes",
        modified: "2026-07-20 22:10",
        author: "Alex Mercer",
        summary: "Personal memos regarding organizational restructuring, leadership transition in APAC, and board alignment strategies.",
        keywords: ["Notes", "Restructuring", "APAC", "Priorities"]
      }
    ],
    transcripts: [
      {
        id: "tr-doc-1",
        title: "Transcript_Q2_Operational_Review_Sync.txt",
        location: "Teams Transcripts > Boardroom Alpha",
        modified: "2026-07-18 16:00",
        author: "Teams AI Transcriber",
        summary: "Key decision: Approved $3.2M budget reallocation to cloud security infrastructure. Action item assigned to CTO Elena Rostova.",
        keywords: ["Transcript", "Operations", "Budget", "Security", "Elena"]
      }
    ]
  },

  adminTasks: [
    {
      id: "task-01",
      title: "Approve Q3 Cloud Security Infrastructure Budget ($3.2M)",
      assignedTo: "Elena Rostova (CTO)",
      dueDate: "Today, 5:00 PM",
      priority: "High",
      status: "Pending Signature",
      source: "Teams Transcript Extract"
    },
    {
      id: "task-02",
      title: "Send Escrow Clarification to Board Chair Victoria Vance",
      assignedTo: "Chief of Staff AI",
      dueDate: "Today, 11:00 AM",
      priority: "High",
      status: "In Progress (Sub-Agent Triage)",
      source: "Email Msg-201"
    },
    {
      id: "task-03",
      title: "Resolve 10:45 AM Schedule Conflict for Operations Sign-off",
      assignedTo: "Chief of Staff AI",
      dueDate: "Today, 09:30 AM",
      priority: "Medium",
      status: "Ready for Auto-Optimize",
      source: "Executive Calendar"
    }
  ]
};
