# Excel Expert — Audit Assistant Add-in

A Big 4 / IFRS Excel Add-in for Singapore & Hong Kong auditors.
Runs entirely on your local machine. No data is ever sent to any server except Anthropic's API (and only the text of your question — never your raw spreadsheet data).

---

## First-Time Setup (5 minutes)

### Step 1 — Install Node.js
Download from https://nodejs.org (choose the LTS version) and install.

### Step 2 — Run SETUP.bat
Double-click `SETUP.bat`. This installs the required packages.

### Step 3 — Start the server
Double-click `START.bat`.
You should see: `✅ Excel Expert Add-in running at https://localhost:3000`
**Keep this window open while using Excel.**

### Step 4 — Sideload the Add-in in Excel
1. Open Excel
2. Go to **Insert → Get Add-ins → Upload My Add-in** (bottom of the dialog)
3. Browse to this folder and select `manifest.xml`
4. Click Upload

The Excel Expert pane will appear on the right side of Excel.

### Step 5 — Add your Anthropic API Key
In the add-in, go to the **⚙ Settings** tab and paste your API Key.
Get one at: https://console.anthropic.com

---

## Daily Use

1. Double-click `START.bat` (before opening Excel, or after)
2. Open Excel — the add-in loads automatically

To pin the add-in so it's always visible:
- Right-click the add-in in the task pane and choose **Pin**

---

## Features

| Tab | What it does |
|-----|-------------|
| **Ask** | Describe a problem → get formulas, steps, VBA. Includes ⬇ Read Selection (pastes your selected range into the question) and ⬆ Insert into cell (puts the formula directly into Excel). |
| **Test** | Upload ① Input Data CSV + ② Expected Result CSV → AI simulates your formula row-by-row and returns ✅ MATCH / ❌ MISMATCH / ⚠ PARTIAL with a final PASS/FAIL verdict. You can also pull data directly from your sheet using ⬇ From Sheet. |
| **KB** | Knowledge base — add company/client-specific rules (e.g. "Client ABC always has a TOTAL row at the bottom"). These are automatically injected into every question by keyword match. |
| **⚙** | API Key, history. |

---

## Troubleshooting

**Add-in shows "This site can't be reached"**
→ Make sure START.bat is running (check the terminal window).

**"Insert into selected cell" doesn't work**
→ Click a cell in Excel first, then press the button.

**SSL certificate warning in Excel**
→ Open https://localhost:3000 in your browser, accept the security warning once, then reload Excel.

**Knowledge base resets after restart**
→ This is normal if you only use localStorage. The server also saves to `knowledge-base.json` in this folder for persistence across sessions.

---

## Data Privacy

- Your spreadsheet data **never leaves your device**
- Only the text of your question (and KB entries you've added) are sent to Anthropic's API
- The API Key is stored in your browser's localStorage
- All files are local — nothing is uploaded anywhere
