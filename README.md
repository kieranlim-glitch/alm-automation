# ALM Daily Refresh — Automation

This repo automatically refreshes the ALM Excel file every morning when you upload the new raw CSV.

---

## 📁 Folder Structure

```
alm-automation/
├── input/
│   ├── raw_data.csv          ← YOU UPLOAD THIS every day
│   └── ALM_template.xlsx     ← your base ALM file (upload once, keep updated)
├── output/
│   └── ALM_DD_Mon_YYYY_8AM_SGT.xlsx   ← auto-generated output
├── scripts/
│   └── alm_refresh.py        ← the automation script
└── .github/workflows/
    └── alm_refresh.yml       ← GitHub Actions trigger
```

---

## 🚀 Daily Workflow (takes ~1 minute)

### Step 1 — Upload your new CSV

Go to your repo on GitHub → click `input/` folder → click `raw_data.csv` → click the **pencil/edit icon** → then drag & drop or paste the new file → **Commit changes**.

Or via Git on your computer:
```bash
cp ~/Downloads/Overview_Asset_and_AUM_Balances__10_.csv input/raw_data.csv
git add input/raw_data.csv
git commit -m "ALM data 14 Apr 2026"
git push
```

### Step 2 — Wait ~2 minutes

GitHub Actions runs automatically. You'll see a yellow dot → green tick on the repo homepage.

### Step 3 — Download your file

**Option A — Artifacts (always works):**
- Go to your repo → **Actions** tab → click the latest run → scroll down to **Artifacts** → download `ALM-refreshed-XX`

**Option B — Output folder (if you kept the commit-back step):**
- Go to `output/` folder in the repo → click the `.xlsx` file → **Download**

---

## ⚙️ One-Time Setup

### 1. Create the GitHub repo
- Go to [github.com/new](https://github.com/new)
- Name it `alm-automation` (private recommended)
- Click **Create repository**

### 2. Upload the template file
- Upload your current ALM Excel as `input/ALM_template.xlsx`
- This is the base file that gets all its RAW data replaced each day
- Update it whenever you make structural changes to other sheets

### 3. Enable GitHub Actions
- Go to **Settings** → **Actions** → **General** → set to "Allow all actions"
- Under **Workflow permissions** → select "Read and write permissions" (needed for the auto-commit step)

### 4. First push
Upload `input/raw_data.csv` and the workflow will run for the first time.

---

## 🔧 What the script does automatically

| Adjustment | Rule |
|---|---|
| **GALA → GALA (V1)** | Renames 4 On-Chain rows + inserts ALM-AUM summary row, all highlighted yellow |
| **stETH warm_wallet Customer** | If balance is negative/wrong → overrides to `2.06` with diff formulas, yellow |
| **TUSD warm_wallet Customer** | If balance is non-zero → sets to `0`, yellow |

All other sheets (Finance Pivot, Daily Changes, Segregation Ratio, etc.) are **preserved untouched**.

---

## 🛠️ Adding new adjustment rules

Edit `scripts/alm_refresh.py` and add a new block inside the `apply_adjustments()` function, following the same pattern as the TUSD fix. Then commit — the workflow will use the new logic next time it runs.
