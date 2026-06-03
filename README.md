# SOC Command Center

A single-page **Security Operations Center (SOC) dashboard** that unifies daily SOC tools in one dark-themed command center. The layout uses a **30% / 70%** split: sidebar navigation on the left and a workspace panel on the right where tools load inline or open externally when embedding is not supported.

**Repository:** [github.com/Maveera/SOC-Master-Dashboard](https://github.com/Maveera/SOC-Master-Dashboard)  
**Developer:** [Maveera](https://maveera.tech)

---

## Features

- **30% / 70% split-screen layout** — Sidebar navigation + main workspace iframe panel
- **Dark cybersecurity theme** — Deep grays with red accents (`#0d1117`, `#11141a`, `#ff4a4a`)
- **Collapsible sidebar** — On desktop, the ☰ button hides the nav bar completely and expands the workspace to full width; click again to restore the 30% / 70% split
- **Resizable navigation** — Drag the handle between the SOC Command sidebar and workspace to adjust panel width (saved in browser storage)
- **Mobile overlay drawer** — On phones and tablets (≤900px), the sidebar slides in as a full-width overlay drawer (up to 320px) so all nav links and labels are visible
- **Mobile iframe fit** — SOC Daily Report, SOC Monthly Report, and SOC Master Consolidation scale to fit narrow screens without horizontal overflow
- **Self-hosted Monthly Report** — Apraava, Nocil, BluPine, and RMZ Corp templates with a **30% form / 70% preview** layout and a **draggable resize handle** between form and preview
- **Live search** — Filter sidebar tools as you type
- **Online / Offline status** — Green **Online** or red **Offline** labels based on configured URLs
- **Tool info popup** — About dialog on each tab switch with **Don't show again** checkbox (saved per tool in `localStorage`)
- **Fast tab switching** — Cached iframes; only the active panel is visible
- **External tool support** — Tools that block iframes show an **Open Tool** panel for new-tab launch
- **110% default zoom** — Slightly enlarged UI on desktop for better readability (reset to 100% on screens ≤900px)
- **Responsive design** — Optimized for laptop, tablet, and mobile

---

## Project Structure

```
SOC Main dashboard/
├── index.html              # Dashboard markup; sidebar + backdrop live outside .app
├── css/
│   └── styles.css          # Theme, layout, modal, mobile drawer, and zoom styles
├── js/
│   └── app.js              # Tool config, navigation, popups, panel logic
├── monthly-report/         # Self-hosted SOC Monthly Report (Apraava, Nocil, BluPine, RMZCorp)
│   ├── index.html        # Template picker landing page
│   ├── shared/           # Shared JS/CSS (charts, nav, theme)
│   ├── samples/          # Sample CSV uploads per customer (see below)
│   │   ├── UPLOAD-FIELD-MAP.txt
│   │   ├── apraava/      # 10 sample CSVs
│   │   ├── nocil/        # 4 sample CSVs
│   │   ├── blupine/      # 8 sample CSVs (RMZ-parity template)
│   │   └── rmzcorp/      # 8 sample CSVs
│   ├── Apraava/
│   ├── Nocil/
│   ├── BluPine/
│   └── RMZCorp/
├── .gitignore
└── README.md
```

---

## Getting Started

### Option 1: Open directly

Double-click `index.html` or open it in any modern browser (Chrome, Edge, Firefox).

### Option 2: Local server (recommended)

Some embedded tools behave better when served over HTTP:

```bash
cd "SOC Main dashboard"
python -m http.server 8080
```

Then open: [http://localhost:8080/index.html](http://localhost:8080/index.html)

---

## Sidebar Tools

### Operations & Metrics

| Tool | URL | Mode | Status |
|------|-----|------|--------|
| SOC Daily report | [soc-daily-report.vercel.app](https://soc-daily-report.vercel.app/) | iframe | Online |
| SOC Incident Analysis | [soc-incident-analysis.vercel.app](https://soc-incident-analysis.vercel.app/) | iframe | Online |
| SOC Monthly Report | `./monthly-report/` (self-hosted) | iframe | Online |
| SOC Master Consolidation | [soc-master-consolidation-report.vercel.app](https://soc-master-consolidation-report.vercel.app/) | iframe | Online |
| FortiSIEM Full NoteBook | [forti-siem-book.vercel.app](https://forti-siem-book.vercel.app/) | iframe | Online |
| SOC MTTR and MTTA Calculator | [itil-tracker.vercel.app](https://itil-tracker.vercel.app/) | iframe | Online |

### Intelligence & Dev

| Tool | URL | Mode | Status |
|------|-----|------|--------|
| SOC Parser learning | [fortisieimparser.free.nf](https://fortisieimparser.free.nf/) | external | Online |
| Viperintel Threat Intel | [viper-intel.streamlit.app](https://viper-intel.streamlit.app/#threat-map) | external | Online |

---

## Configuring Tools

Edit `js/app.js` and update the `TOOL_CONFIG` object:

```javascript
const TOOL_CONFIG = {
  "SOC Daily report": {
    url: "https://soc-daily-report.vercel.app/",
    mode: "iframe",
    mobileFit: true,
    about: "Short description shown in the tool info popup."
  },
  "SOC Monthly Report": {
    url: "./monthly-report/",
    mode: "iframe",
    mobileFit: true,
    about: "SOC monthly reporting console for Apraava, Nocil, RMZ, and Blue Pine."
  },
  "SOC Parser learning": {
    url: "https://fortisieimparser.free.nf/",
    mode: "external",
    message: "Message shown in the workspace when the tool opens externally.",
    about: "Description for the info popup."
  }
};
```

### Config fields

| Field | Purpose |
|-------|---------|
| `url` | Tool URL. Use `#` if not configured (shows as **Offline**) |
| `mode` | `iframe` (embedded) or `external` (new tab) |
| `mobileFit` | Optional. When `true`, scales the iframe on mobile to fit the screen width (for desktop-first embedded tools) |
| `about` | Text shown in the tool info popup |
| `message` | Optional workspace message for `external` tools |

### Modes

| Mode | Behavior |
|------|----------|
| `iframe` | Loads the tool inside the workspace panel |
| `external` | Shows an **Open Tool** button; opens URL in a new browser tab |

Use `external` for sites that block iframe embedding (e.g. Streamlit) or require third-party cookies.

### Add a new tool

1. Add a nav link in `index.html` with a unique `data-tool` name
2. Add a matching entry in `TOOL_CONFIG` inside `js/app.js`

---

## Mobile Usage

On screens **900px and below**:

- The sidebar is hidden by default and opens as a slide-in drawer via the ☰ button
- The drawer uses **full mobile width** (up to **320px**) — not the desktop 30% share — so tool names and status labels are not clipped
- The drawer fills the viewport height (`100dvh`) with safe-area padding for notched phones
- Sidebar open/close state is toggled on **`body`** (`sidebar-nav-open`) because the nav lives outside `.app`
- Tap outside the drawer, press **Escape**, or select a tool to close the menu
- The workspace toolbar stays compact; tool names truncate with ellipsis in the toolbar only
- Tools with `mobileFit: true` render at a 1024px reference width and scale down to fit the phone screen

Currently enabled for:

- **SOC Daily report**
- **SOC Monthly Report**
- **SOC Master Consolidation**

To enable mobile fit for another embedded tool, add `mobileFit: true` to its entry in `TOOL_CONFIG`.

Mobile layout uses `MOBILE_MQ = (max-width: 900px)` in `js/app.js`.

Mobile drawer CSS (in `css/styles.css`):

```css
@media (max-width: 900px) {
  .sidebar {
    position: fixed;
    width: min(92vw, 320px) !important;
    max-width: min(92vw, 320px) !important;
    height: 100dvh;
  }

  body.sidebar-nav-open .sidebar {
    transform: translate3d(0, 0, 0);
  }
}
```

The viewport meta tag in `index.html` includes `viewport-fit=cover` for safe-area support on iOS.

---

## Desktop Navigation

On screens **wider than 900px**:

- The sidebar is visible by default at **30%** width
- Click **☰** in the workspace toolbar to collapse the sidebar — the workspace expands to **100%**
- Click **☰** again to restore the sidebar
- **Drag the red edge** between the sidebar and workspace to resize the nav panel (preference saved locally)
- Collapse uses `sidebar-collapsed` on **`body`** and `.app`, with `flex: 0 0 0` so the nav bar does not keep reserving space
- `body` uses flex layout so the sidebar (sibling of `.app`) and workspace share the 30% / 70% split

---

## Monthly Report Layout

The SOC Monthly Report builder uses a **30% left form panel** and **70% right preview panel** on desktop. Layout rules live in each template's `styles.css` under `monthly-report/`:

- **Apraava, BluPine, RMZCorp** — CSS Grid: `grid-template-columns: var(--form-panel-width, 30%) 6px 1fr`
- **Nocil** — Flex sidebar at `30%` width with collapsible drawer

### Changing month columns

**Apraava** — **Report Month Name** for cover/engagement/charts; **Table Month Columns** for four-column EPS/SLA/alert tables:

| What | Where |
|------|--------|
| **Report month (cover, [MONTH])** | **Report Month Name** (`reportMonthLabel`) |
| **Table column headers (×4)** | **Table Month Columns** (`monthColumns`) — leave blank to auto-fill last 4 months from report end date |
| **JavaScript logic** | `getReportMonthColumnLabels()` in `monthly-report/Apraava/main.js` |
| **Default slot keys** | `MONTH_SLOT_KEYS` (field IDs like `epsJan`, `slaHighJanCount` stay as Jan–Apr internally) |

**Apraava** — **Report Month Name** (`reportMonthLabel`) for cover, engagement `[MONTH]`, and chart titles; **Table Month Columns** (`monthColumns`) for four-column EPS/SLA/alert tables.

**Nocil, BluPine, RMZCorp** — **Report Month Name** (`tableMonthLabel`) updates the cover (`Monthly SOC Report | June 2026`), engagement `[MONTH]`, chart titles, and monthly table column headers. Type `June 2026` or just `June` (year is taken from **Report Date Range** when omitted). Leave blank to use the date range only.

To adjust the split, edit `--sidebar-width` in `css/styles.css` (dashboard), drag the resize handle on desktop, or edit the grid/flex rules in `monthly-report/*/styles.css`.

When collapsed on desktop, the sidebar uses `flex: 0 0 0` and `max-width: 0` so the hamburger toggle fully hides the nav bar.

---

## Monthly Report — Sample CSV Uploads

Sample upload files live under `monthly-report/samples/`. Each customer has its **own folder** with CSV files named using the **exact PPT slide heading** (same text as shown on each slide). Use them to test CSV imports in the report builder.

| Customer | Folder | CSV count | Template path |
|----------|--------|-----------|---------------|
| Apraava | `samples/apraava/` | 10 | `monthly-report/Apraava/` |
| Nocil | `samples/nocil/` | 4 | `monthly-report/Nocil/` |
| BluPine | `samples/blupine/` | 8 | `monthly-report/BluPine/` |
| RMZ Corp | `samples/rmzcorp/` | 8 | `monthly-report/RMZCorp/` |

**Full field mapping:** see [`monthly-report/samples/UPLOAD-FIELD-MAP.txt`](monthly-report/samples/UPLOAD-FIELD-MAP.txt).

### How to use

1. Open a customer template from `monthly-report/index.html`.
2. Enter **Customer Name**, **Report Date Range**, and other metadata — forms start **blank** (no pre-filled engagement data).
3. Set **Report Date Range** to match your reporting period (e.g. `01-01-2026 to 30-04-2026` for four monthly columns).
4. Use each section’s **file upload** to import CSVs from `samples/<customer>/` (reference format only) or your own exports.
5. Fill remaining fields manually (SLA counts, logos, key points, narratives, etc.).

### Apraava sample files

| PPT slide | Sample file |
|-----------|-------------|
| Potential Incidents and Alert Summary | `Potential Incidents and Alert Summary - Overall Alerts.csv` |
| Potential Incidents and Alert Summary | `Potential Incidents and Alert Summary - True Positive Alerts.csv` |
| Potential Incidents and Alert Summary | `Potential Incidents and Alert Summary - False Positive Alerts.csv` |
| Potential Incidents and Alert Summary | `Potential Incidents and Alert Summary - Potential Incidents.csv` |
| Potential Incident Tickets Trend | `Potential Incident Tickets Trend.csv` |
| Potential Incidents | `Potential Incidents.csv` |
| EPS Trend Plot | `EPS Trend Plot.csv` |
| Top 10 Devices Contributing To Highest EPS | `Top 10 Devices Contributing To Highest EPS.csv` |
| Potential Incidents – Risks Mitigated | `Potential Incidents - Risks Mitigated.csv` |
| Integrated Device Inventory | `Integrated Device Inventory.csv` |

### Nocil sample files

| PPT slide | Sample file |
|-----------|-------------|
| Potential Alert Tickets Trend | `Potential Alert Tickets Trend.csv` |
| Potential Alerts - Risks Mitigated | `Potential Alerts - Risks Mitigated.csv` |
| Rule-Based Severity Categories For Potential Alerts | `Rule-Based Severity Categories For Potential Alerts.csv` |
| Integrated Device Inventory | `Integrated Device Inventory.csv` |

### BluPine & RMZ Corp sample files

BluPine is aligned with **RMZ Corp** slide layout, charts, tables, and PPT export behavior (BluPine branding and titles such as **Potential Alerts - Risks Mitigated** are unchanged).

| PPT slide | Sample file |
|-----------|-------------|
| Potential Incident Tickets Trend | `Potential Incident Tickets Trend.csv` |
| Potential Alerts - Risks Mitigated | `Potential Alerts - Risks Mitigated.csv` |
| Rule-Based Severity Categories For Potential Incidents | `Rule-Based Severity Categories For Potential Incidents.csv` |
| EPS Trend Plot (right chart — Top Hosts) | `Top Hosts EPS.csv` |
| Highest EPS Consuming Events For Mentioned Firewalls | `Highest EPS Consuming Events For Mentioned Firewalls.csv` |
| Overall Support Ticket Handled By SNS (Firewall Support) | `Overall Support Ticket Handled By SNS (Firewall Support) - Major.csv` |
| Overall Support Ticket Handled By SNS (Firewall Support) | `Overall Support Ticket Handled By SNS (Firewall Support) - Minor.csv` |
| Integrated Device Inventory | `Integrated Device Inventory.csv` |

Sample CSVs are **not loaded automatically** — use the file upload controls in each template. Each sample file has a header plus **two generic example rows** so you can verify CSV shape before pasting real data.

**Apraava optional hints:** On first open, `Apraava/sample-data.js` may pre-fill **only empty** fields with minimal examples (e.g. two-row alert quad CSV, two-row EPS table). All customer names, dates, metrics, and narratives must be entered manually or imported via CSV/PPTX.

**BluPine engagement puzzle:** default graphic is `monthly-report/BluPine/final_puzzle.png` (upload override via **Puzzle Graphic** in the form panel).

---

## Tool Info Popup

Each time you switch to a tool, an info popup appears with that tool's `about` text.

- **OK** — Closes the popup; it will appear again next visit
- **Don't show again** + **OK** — Saves preference per tool in browser `localStorage`

Reset all popup preferences in the browser console:

```javascript
localStorage.removeItem("soc-dashboard-tool-info-dismissed");
```

---

## Customization

### Layout split (30% / 70%)

In `css/styles.css`:

```css
:root {
  --sidebar-share: 30%;
  --workspace-share: 70%;
}

/* Collapsed desktop sidebar */
body.sidebar-collapsed .sidebar {
  width: 0;
  max-width: 0;
  flex: 0 0 0;
}
```

### Default zoom (110%)

In `css/styles.css`:

```css
html {
  zoom: 1.1;   /* 110% on desktop — mobile resets to 1 */
}
```

### Mobile iframe fit width

In `js/app.js`:

```javascript
const MOBILE_FIT_WIDTH = 1024;   /* reference width before scale-down */
```

### Theme colors

CSS variables at the top of `css/styles.css`:

```css
:root {
  --bg-deep: #0d1117;
  --bg-panel: #11141a;
  --accent: #ff4a4a;
}
```

### Fonts

Loaded from Google Fonts in `index.html`:

- **Orbitron** — Headings
- **Inter** — Body text
- **JetBrains Mono** — Footer and clock

---

## Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| Tool shows blank / "refused to connect" | Site blocks iframe embedding | Set `mode: "external"` in `TOOL_CONFIG` |
| "Cookies are not enabled" in iframe | Third-party cookie restrictions | Use `external` mode and open in a new tab |
| Previous tool still visible after switching | Stale iframe panel | Fixed via active-panel logic; hard-refresh with `Ctrl+F5` |
| Dashboard nested inside itself | Tool URL points to this dashboard | Never set tool URLs to `index.html` or `/` |
| Offline tool shows last visited page | Invalid URL (`#`) | Expected — shows offline placeholder instead |
| Embedded tool cut off on mobile | Tool is desktop-first | Add `mobileFit: true` in `TOOL_CONFIG` |
| Sidebar covers content on mobile | Expected drawer behavior | Tap ☰ to open/close, or select a tool to auto-close |
| ☰ does not hide sidebar on desktop | Flex basis still reserves sidebar width | Hard-refresh; collapsed state must set `flex: 0 0 0` (included in current `styles.css`) |
| Mobile nav drawer clipped or too narrow | Desktop 30% width or stale CSS cached | Hard-refresh or use incognito; drawer uses `body.sidebar-nav-open` and `width: min(92vw, 320px)` |
| Mobile nav shows truncated labels (SO…, Search t…) | Drawer open state CSS not matching sidebar DOM | Ensure `body.sidebar-nav-open .sidebar` is used — sidebar is outside `.app` |
| Monthly Report form panel too wide | Old 360px fixed layout | Use the self-hosted `./monthly-report/` copy in this repo |

---

## Browser Support

Works best in modern browsers with iframe and ES6 support:

- Google Chrome
- Microsoft Edge
- Mozilla Firefox

> **Note:** `zoom: 1.1` is supported in Chrome and Edge. Firefox may render at 100% unless an alternate scale method is added. Mobile iframe scaling uses CSS `transform: scale()` and works in all modern mobile browsers.

---

## License

Internal SOC tooling — developed for Security Operations Center workflows.

**Developed by [Maveera](https://maveera.tech). All rights reserved.**
