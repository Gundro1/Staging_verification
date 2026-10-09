# BisViews Staging Verification & Verifier Allocation Tool (v5)

An automated Python utility for generating SOP-compliant QA verification workbooks from BisViews seed datasets (`.xlsx` or `.sql`).

This tool extracts unique franchise brands, calculates total branch counts, and distributes brands equally among QA verifiers using a **Round-Robin allocation system** (no brand shared or duplicated between verifiers). Verifiers use the generated Excel workbook to audit deployed data on `https://staging.bisviews.com`.

---

## 📋 Table of Contents
- [✨ Features](#-features)
- [🛠️ Installation](#%EF%B8%8F-installation)
- [🚀 How to Run (Step-by-Step)](#-how-to-run-step-by-step)
  - [Method 1: Interactive Mode](#method-1-interactive-mode-guided-prompts)
  - [Method 2: Command Line (CLI) Mode](#method-2-command-line-cli-mode)
- [📊 Understanding the Output Workbook](#-understanding-the-output-workbook)
- [🔍 How Verifiers Use the Output on Staging](#-how-verifiers-use-the-output-on-staging)

---

## ✨ Features
- **Dual Format Support**: Natively parses Excel seed datasets (`.xlsx`) and SQL database dump files (`.sql`).
- **Fair Round-Robin Allocation**: Automatically splits unique brands equally across named verifiers.
- **SOP-Compliant Excel Layout**:
  - **Master Overview**: Total dataset branch count, total unique brands, category & subcategory breakdown table.
  - **Verifier Tabs**: Individual verifier tabs containing test case IDs (`TC001`, `TC002`...), Expected Results (`Displayed`), and built-in interactive Excel dropdowns (`PASS`, `FAIL - Wrong Category`, `FAIL - Missing`, `FAIL - Count Mismatch`, etc.).

---

## 🛠️ Installation

### Step 1: Clone the Repository
```bash
git clone https://github.com/Gundro1/Staging_verification.git
cd Staging_verification
```

### Step 2: Install Python Dependencies
Ensure Python 3.9+ is installed, then run:
```bash
pip install -r requirements.txt
```

---

## 🚀 How to Run (Step-by-Step)

### Method 1: Interactive Mode (Guided Prompts)
If you prefer an interactive wizard that prompts you for the file path and verifier names:

```bash
python gundro_reference_generator.py
```
**Prompts:**
1. `Enter path to source file (.xlsx or .sql):` -> Paste your dataset path (e.g. `C:\path\to\dataset.xlsx`)
2. `How many verifiers?` -> Enter a number (e.g., `3`)
3. `Enter Verifier 1 name:` -> Enter first verifier (e.g., `Grace`)
4. `Enter Verifier 2 name:` -> Enter second verifier (e.g., `Abdul`)
5. `Enter Verifier 3 name:` -> Enter third verifier (e.g., `Siddik`)

---

### Method 2: Command Line (CLI) Mode

Run directly in a single command by passing your file path and verifier names:

#### **A. Using an Excel Seed File (`.xlsx`):**
```bash
python gundro_reference_generator.py --file "C:\path\to\BisViews_Algeria_Franchise_final.xlsx" --verifiers Grace Abdul Siddik
```

#### **B. Using a SQL Seed File (`.sql`):**
```bash
python gundro_reference_generator.py --file "C:\path\to\BisViews_Algeria_Franchise_final.sql" --verifiers Grace Abdul Siddik
```

---

## 📊 Understanding the Output Workbook

The generated Excel workbook will be saved in the `QA_Reference_Datasets/` folder with a timestamped filename (e.g., `QA_Verification_Algeria_DZ_20261009_1930.xlsx`).

### Structure:
1. **Master Overview Sheet**:
   - High-level KPIs: Total Branch Rows, Total Unique Brands, Number of Categories, Number of Subcategories.
   - Comprehensive Category & Subcategory Breakdown table.
2. **Individual Verifier Tabs** (e.g., `Grace`, `Abdul`, `Siddik`):
   - Assigned brands table sorted logically by Category and Subcategory.
   - Column `Actual Result`: Dropdown menu (`Displayed`, `Missing`, `Wrong Category`, `Wrong Subcategory`, `Duplicate`).
   - Column `Status`: Dropdown menu (`PASS`, `FAIL - Missing`, `FAIL - Wrong Category`, `FAIL - Wrong Subcategory`, `FAIL - Count Mismatch`).
   - Column `Notes / Defect Description`: Free-text field for logging discrepancies.

---

## 🔍 How Verifiers Use the Output on Staging

Once the workbook is generated, each verifier completes the verification on `https://staging.bisviews.com` using these steps:

1. **Open your assigned tab** (e.g. `Grace` tab).
2. **Go to Staging**: Open `https://staging.bisviews.com` in your browser.
3. **Search for your assigned brand** in the search bar.
4. **Compare details**:
   - Is the **Brand Name** displayed?
   - Does the **Category & Subcategory** match the breadcrumb?
   - Does the **Total Branch Count** match the number shown on staging?
5. **Log Status in Excel**:
   - If everything matches: Select **`PASS`** in the Status column.
   - If there is a mismatch: Select the appropriate **`FAIL - ...`** option and write details in the Notes column.

---

## 📜 License
MIT © Azeez Gundro & BisViews QA Team
