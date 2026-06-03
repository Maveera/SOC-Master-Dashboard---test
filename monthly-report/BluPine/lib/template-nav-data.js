/** Human-readable labels for slide navigation (matches PPT slide titles) */
const SLIDE_NAV_LABELS = {
  cover: "Cover Page",
  revision: "Document Revision History",
  engagement: "The Engagement",
  "alert-quad": "Alert Summary",
  "tot-pot": "Total Potential Alerts",
  trend: "Alert Tickets Trend",
  "rule-severity": "Rule Severity Categories",
  "pot-incidents": "Potential Incidents",
  eps: "EPS Trend Plot",
  "eps-events": "EPS Events",
  "top-eps": "Top 10 Devices (EPS)",
  risks: "Risks Mitigated",
  "sla-monthly": "Response Time SLA",
  "fortisiem-alerts": "FortiSIEM Alerts",
  "tfpos-alerts": "TFPOS Alerts",
  "support-tickets": "Support Tickets",
  inventory: "Device Inventory",
  "key-points": "Key Points Summary",
  contact: "Contact"
};

/** Preview section order per template (matches export / HTML slide order) */
const TEMPLATE_REPORT_SECTIONS = {
  apraava: [
    "cover", "revision", "engagement", "alert-quad", "trend", "pot-incidents",
    "eps", "top-eps", "risks", "sla-monthly", "inventory", "key-points", "contact"
  ],
  nocil: [
    "cover", "revision", "engagement", "tot-pot", "trend", "risks", "rule-severity",
    "sla-monthly", "inventory", "contact"
  ],
  blupine: [
    "cover", "revision", "engagement", "tot-pot", "trend", "risks", "rule-severity", "sla-monthly",
    "fortisiem-alerts", "tfpos-alerts", "eps", "eps-events", "support-tickets",
    "inventory", "key-points", "contact"
  ],
  rmz: [
    "cover", "revision", "engagement", "tot-pot", "trend", "risks", "rule-severity", "sla-monthly",
    "fortisiem-alerts", "tfpos-alerts", "eps", "eps-events",
    "inventory", "key-points", "contact"
  ],
  custom: null
};
