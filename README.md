# 📋 PCms Quality Tracker

A lightweight, browser-based web app that converts the horizontal PCms Excel QA table into a clear, filterable breach dashboard — no backend required.

## 🚀 Live App

Host on **GitHub Pages**: go to *Settings → Pages → Branch: main → / (root)* and save.

---

## ✨ Features

| Feature | Details |
|---|---|
| **Excel Upload** | Drag & drop or browse your `.xlsx` PCms file |
| **Auto-parser** | Converts horizontal table (11 category columns) → vertical breach records |
| **Overview dashboard** | Stats cards + 4 charts: by month, by agent, by category, weekly trend |
| **KO-only tab** | Filterable table showing only KO breaches |
| **Dedup logic** | Same INC + same agent + same category = counted as **1 breach** |
| **Multi-agent flag** | If a ticket is breached by multiple agents, each agent gets their own breach counted as 1; flagged as **Multi-agent** |
| **Export** | Download filtered KO data as **CSV** or **Excel** |

---

## 📂 File Structure

```
pcms-tracker/
├── index.html   # App shell + layout
├── style.css    # Dark theme styles
├── app.js       # All parsing, chart & export logic
└── README.md
```

---

## 📊 Expected Excel Format (Summary sheet)

The app looks for a sheet named **Summary** (falls back to first sheet). It expects:

| Week | Month | Ticket ID | 1. Valid Rejection | Agent | 2. Communication | Agent | … | 11. Other Errors | Agent |
|---|---|---|---|---|---|---|---|---|---|
| 1 | Jan 25 | INC000103507178 | 6. KO Missing SS translation | Verona | | | … | | |

- **Category columns** are detected by the `1.` … `11.` prefix.
- Each category column is followed by an **Agent** column (which may have a blank header).
- Multiple agents in one cell should be separated by `,` or `;`.

---

## 🔢 Counting Logic

```
Breach unit = unique (Ticket ID + Agent + Category)

If same ticket breached by different agents:
  → Each agent gets 1 breach (flagged as "Multi-agent")

If same ticket breached multiple times by same agent in same category:
  → Counted as 1 breach only (duplicate rows are hidden)
```

---

## 🛠 Running Locally

```bash
# No build step needed — just open in a browser
npx serve .
# or
python -m http.server 8080
```

Then open `http://localhost:8080`.
