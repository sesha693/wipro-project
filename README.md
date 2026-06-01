# Wipro Weekly Slide Generator

Automatically creates a polished PowerPoint presentation from your weekly BPM Excel file — no manual copy-paste needed.

---

## What It Does

Every week you upload the Excel file (`Week08 Netadd_BPM_ADH.xlsx`) and the tool generates:

- A **title slide** and a **summary slide**
- **Individual slides** for each account showing RU / RD / Netadd metrics with colour-coded gaps and bar charts
- **Group comparison slides** for accounts you tag together (e.g. "Tier 1", "Healthcare") — 3 slides per group, one for each metric

---

## One-Time Setup (Do This Once)

### Step 1 — Install Python

1. Go to **https://www.python.org/downloads/**
2. Download the latest **Python 3.13** installer for Windows
3. Run the installer — **tick the box that says "Add Python to PATH"** before clicking Install
4. Click Install Now and wait for it to finish

### Step 2 — Download the Project

If you received the project as a ZIP file:
1. Unzip it to a folder, e.g. `C:\Users\YourName\wipro-project`

Or if using Git:
```
git clone https://github.com/sesha693/wipro-project.git
```

### Step 3 — Install Dependencies

1. Open **Command Prompt** (search for `cmd` in the Start menu)
2. Navigate to the project folder:
   ```
   cd C:\Users\YourName\wipro-project
   ```
3. Run this command to install everything needed:
   ```
   pip install -r requirements.txt
   ```
4. Wait for it to finish (this only needs to be done once)

---

## Running the App (Every Week)

### Step 1 — Start the App

1. Open **Command Prompt**
2. Navigate to the project folder:
   ```
   cd C:\Users\YourName\wipro-project
   ```
3. Run:
   ```
   python -m streamlit run app.py
   ```
4. Your browser will automatically open to **http://localhost:8501**
   - If the browser doesn't open, manually go to that address

> **To stop the app:** go back to Command Prompt and press `Ctrl + C`

---

## Using the App

### Step 1 — Configure Settings (Left Sidebar)

On the left side of the screen you will see settings:

| Setting | What It Does |
|---|---|
| **Week Label** | Enter the current week, e.g. `WK09` |
| **Quarter Label** | Enter the quarter, e.g. `Q1'27` |
| **Metrics to include** | Tick/untick RU, RD, Netadd |
| **Slide Layout** | Choose how slides are organised (see below) |
| **Chart Type** | Bar + KPI Cards is recommended |
| **Output filename** | Name of the PowerPoint file that will be downloaded |

**Slide Layout options:**
- **Per Account (one slide each)** — each company gets its own slide per metric *(default)*
- **Per Metric (all accounts, ranked table)** — one slide per metric showing all companies
- **Per Account — Combined** — one slide per company with all 3 metrics

---

### Step 2 — Upload Your Excel File

1. Click the **Upload** button in the main area
2. Select your weekly file (e.g. `Week08 Netadd_BPM_ADH.xlsx`)
3. The file loads and you will see:
   - A green confirmation banner with the file name and size
   - A **data preview table** on the right showing all accounts and their metrics
   - Gaps are colour-coded: **red = behind plan**, **green = ahead of plan**

---

### Step 3 — Filter Accounts (Optional)

Under **🏢 Individual Slides** in the left column (appears after upload):

- **All accounts** — generates slides for every account in the file *(default)*
- **Select specific** — choose only the accounts you want

---

### Step 4 — Create Groups (Optional)

Groups let you put multiple companies together on **one comparison slide** with your custom label as the title. Each group generates **3 slides** (RU, RD, Netadd).

**Example:** You want a "Tier 1 Accounts" slide showing CISCO, GOOGLE, and AMAZON side by side.

1. Click **➕ Add Group**
2. Type a name in the **Group tag / name** box — e.g. `Tier 1 Accounts`
3. Click in the **Accounts in this group** box and select the companies
4. Repeat to add more groups (e.g. `Healthcare`, `APAC`, etc.)
5. To remove a group click the 🗑️ button on its panel

> **Tip:** You can have as many groups as you need. Each group name will appear as the title on the generated slides.

---

### Step 5 — Generate and Download

1. Scroll to the bottom of the page
2. Check the **Estimated slides** count shown next to the button
3. Click **▶ Generate Slides**
4. A progress bar shows the current slide being built
5. When it says ✅ Done, click the **⬇️ Download** button
6. The PowerPoint file saves to your Downloads folder

---

## Understanding the Slides

### Individual Account Slide

Each slide shows one company for one metric:

```
┌──────────────────────────────────────────────────────────────┐
│  RU  CISCO                                      WK08 | Q1'27 │
├──────────┬──────────┬──────────┬─────────────┬──────────────┤
│ Plan QTR │ WK Plan  │ WK Act   │ Gap         │ WoW          │
│    24    │    23    │    20    │  -4  (RED)  │  ▼ -2        │
├──────────┴──────────┴──────────┴─────────────┴──────────────┤
│  BPM Plan QTR  │  BPM WK Act  │  BPM Gap    │  BPM WoW     │
├────────────────────────────────────────────────────────────- ┤
│  Delta Reason:   [text from the Excel file]                  │
│  Recovery Plan:  [text from the Excel file]                  │
├──────────────────────────────┬───────────────────────────────┤
│  Bar Chart                   │  KPI Cards (Gap, WoW, etc.)  │
└──────────────────────────────┴───────────────────────────────┘
```

**Colour coding:**
- 🔴 Red = Gap is negative (behind plan)
- 🟢 Green = Gap is positive (ahead of plan)
- ▲ Arrow up = improvement week over week
- ▼ Arrow down = worsened week over week

### Group Comparison Slide

All tagged companies appear in one table side by side with the group name as the title and a gap comparison chart.

---

## Troubleshooting

**The browser doesn't open automatically**
→ Manually go to `http://localhost:8501` in your browser (Chrome or Edge recommended)

**"pip is not recognised" error**
→ Python was not added to PATH during installation. Reinstall Python and tick the "Add to PATH" checkbox.

**The app won't start — "No module named streamlit"**
→ Run `pip install -r requirements.txt` again from the project folder.

**Upload button does nothing**
→ Use Chrome or Edge. The app does not work well in Internet Explorer.

**Slides are generated but some accounts are missing**
→ Check that the account name in the Excel file matches exactly (spelling, spacing). The filter is case-insensitive.

**The download doesn't start**
→ Check that your browser is not blocking pop-ups or downloads from localhost.

---

## Weekly Workflow Summary

```
1. Open Command Prompt
2. cd wipro-project
3. python -m streamlit run app.py
4. Browser opens → upload this week's xlsx
5. Set week label (e.g. WK09) in the sidebar
6. (Optional) create groups, filter accounts
7. Click Generate Slides → Download
8. Close app with Ctrl+C when done
```

---

## File Structure

```
wipro-project/
├── app.py                  ← the web app (run this)
├── generate_slides.py      ← command-line version (advanced users)
├── config.yaml             ← default settings for command-line use
├── requirements.txt        ← list of dependencies
├── README.md               ← this file
└── src/
    ├── data_reader.py      ← reads the Excel file
    ├── chart_builder.py    ← generates the charts
    └── slide_builder.py    ← builds the PowerPoint slides
```
