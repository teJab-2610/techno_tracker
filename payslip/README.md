# Payslip PDF Generator

Generate payslip PDFs from salary sheet Excel files used at **TECHNO SOLUTIONS**.

## Prerequisites

Python 3.8+ with the following packages:

```bash
pip install openpyxl reportlab qrcode pillow
```

## Quick Start

```bash
# Generate payslips for all employees
python3 generate_payslips.py "FEB_2026_ANAKAPALLI_TL4P SALARY SHEET.xlsx"

# Output: payslips_FEB_2026/<EmployeeName>_FEB_2026.pdf (one PDF per employee)
```

---

## Command Line Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `excel_file` | Yes | Path to the salary sheet `.xlsx` file |
| `--employees` | No | Comma-separated employee names to filter (case-insensitive, matches both Attendance name and Aadhar name) |
| `--designation` | No | Filter by designation (case-insensitive), e.g. `"PICKER / PACKER"` or `"SUPERVISOR"` |
| `--output-dir` | No | Custom output directory. Default: `payslips_<MONTH>_<YEAR>/` |
| `--bw` | No | Generate black and white payslips (no colored backgrounds, printer-friendly) |
| `--list` | No | List all employees and their designations from the sheet, then exit (no PDFs generated) |
| `--month` | No | Month filter label (currently auto-detected from the Attendance sheet cell D1) |

## Usage Examples

```bash
# Generate for all employees (color)
python3 generate_payslips.py "FEB_2026_ANAKAPALLI_TL4P SALARY SHEET.xlsx"

# Generate black and white versions
python3 generate_payslips.py "FEB_2026_ANAKAPALLI_TL4P SALARY SHEET.xlsx" --bw

# Filter by designation
python3 generate_payslips.py "FEB_2026_ANAKAPALLI_TL4P SALARY SHEET.xlsx" --designation "PICKER / PACKER"
python3 generate_payslips.py "FEB_2026_ANAKAPALLI_TL4P SALARY SHEET.xlsx" --designation "SUPERVISOR"

# Filter by specific employee names (case-insensitive, comma-separated)
python3 generate_payslips.py "FEB_2026_ANAKAPALLI_TL4P SALARY SHEET.xlsx" --employees "C SUVARNA LAXMI,G LAVANYA"

# Combine filters with custom output directory
python3 generate_payslips.py "FEB_2026_ANAKAPALLI_TL4P SALARY SHEET.xlsx" --designation "PICKER / PACKER" --bw --output-dir feb_bw_pickers

# List all employees to see available names and designations
python3 generate_payslips.py "FEB_2026_ANAKAPALLI_TL4P SALARY SHEET.xlsx" --list
```

### Employee Name Matching

- Matching is **case-insensitive**: `"j rahelu"`, `"J RAHELU"`, and `"J Rahelu"` all work.
- Matches against **both** name columns: "Name as per Attendance" (col D) and "Name as per AADHAR" (col E).
- The name must be an **exact full match** (after ignoring case). Partial names like `"RAHELU"` will not match `"J RAHELU"`.

---

## Required Excel Sheet Structure

The input Excel file **must** contain these two sheets with the exact names and column layouts described below. The script reads pre-computed values from the Excel formulas (opened with `data_only=True`).

### Sheet: `Attendence` (note the spelling)

The sheet name must be exactly `Attendence` (not "Attendance").

| Row | Description |
|-----|-------------|
| Row 1 | Cell `D1` contains the month as a date (e.g. `2026-02-01`). This is used to auto-detect the payslip month. |
| Row 2 | Column headers |
| Row 3 | Day-of-week labels (Sun, Mon, etc.) |
| Row 4+ | Employee attendance data (one row per employee) |

**Column layout (Row 4 onwards):**

| Column | Letter | Description |
|--------|--------|-------------|
| A | A | Serial number (must be a number — rows without a numeric S.No are skipped) |
| B | B | Employee Code (SCRUM ID, e.g. `PPRR01063947`) — used to link to Wage Sheet |
| C | C | Employee name (as per attendance) |
| D | D | Gender |
| E | E | Store Code |
| F | F | Date of Joining |
| G-AH | G-AH | Daily attendance marks: `P` (Present), `A` (Absent), `W/O` (Week Off), `PPH` (Present on Public Holiday) |
| AI | AI | Present Days (total) |
| AJ | AJ | Absent (total) |
| AK | AK | Week Offs (total) |
| AL | AL | Total Days |
| AM | AM | Public Holidays |
| AN | AN | TLLP (Total days including leave, public holidays) |

### Sheet: `Wage Sheet`

| Row | Description |
|-----|-------------|
| Row 1 | Company name |
| Row 2 | Period description |
| Row 3 | Location and other headers |
| Row 4 | Column headers |
| Row 5+ | Employee salary data (one row per employee) |

**Column layout (Row 5 onwards):**

| Column Index | Letter | Field | Used In Payslip As |
|-------------|--------|-------|-------------------|
| 0 | A | Serial number (must be numeric) | — |
| 2 | C | SCRUM ID / Employee Code | Links to Attendance sheet |
| 3 | D | Name as per Attendance | Employee name (fallback) |
| 4 | E | Name as per AADHAR | Employee name (primary) |
| 5 | F | Gender | — |
| 6 | G | UAN NUMBER | UAN NO |
| 7 | H | ESIC NUMBER | ESIC NO |
| 8 | I | Date of Joining | — |
| 9 | J | Location | Location |
| 12 | M | Designation | Designation |
| 14 | O | Month (date) | Month & Year |
| 15 | P | No. of days in month | No Of Days This Month |
| 16 | Q | No. of Working Days | Days Worked |
| 17 | R | OT Hours | — |
| 18 | S | Basic (monthly rate) | — |
| 29 | AD | Consolidated Basic/PF Wages (actual) | Basic - (A) |
| 30 | AE | HRA (actual) | HRA - (B) |
| 31 | AF | Other Allowance (actual) | Other Allowance (shown only if > 0) |
| 32 | AG | Conveyance (actual) | Conveyance - (C) |
| 33 | AH | Leave (actual) | PL / Leave - (E) |
| 34 | AI | OT Amount (actual) | OT Hours Amount (shown only if > 0) |
| 35 | AJ | Bonus (actual) | BONUS - (D) |
| 36 | AK | Gross Per Month (actual) | Total Earnings |
| 45 | AT | PF - 12% (employee share) | PF (On Basic-A) |
| 46 | AU | ESIC - 0.75% (employee share) | E.S.I.C (0.75%) |
| 47 | AV | Employee LWF | L W F |
| 48 | AW | Professional Tax | Professional Tax |
| 49 | AX | Total Deduction | Total Deduction |
| 50 | AY | Take Home | Net Salary |

**Important:** The script reads **pre-computed values** from the Excel file (using `data_only=True`). All salary calculations must already be done by the Excel formulas. The script does not recalculate anything — it only formats and presents the data.

---

## How Salary Calculations Work (in the Excel Sheet)

The script reads these values as-is. This section documents how the Excel sheet computes them, for reference.

### Monthly Rates (fixed per designation)

Each employee has fixed monthly rates based on their designation:

| Component | PICKER / PACKER (example) |
|-----------|--------------------------|
| Basic (col S) | 12,813 |
| DA (col T) | 0 |
| HRA (col V) | 0 |
| Leave Wages (col Y) | 1,602 |
| Bonus (col Z) | 1,067 |
| **Gross Per Month (col AA)** | **15,482** |

### Actual (Prorated) Amounts

Actual earnings are prorated based on days worked vs days in the month:

```
Actual Amount = (Monthly Rate / Days in Month) x Days Worked
```

Where "Days Worked" (col Q) = Present Days + Week Offs + Public Holidays (from Attendance sheet col AN / TLLP).

**Example — J RAHELU, Feb 2026 (26 payable days, 27 days worked including 1 public holiday):**

| Component | Monthly Rate | Actual | Calculation |
|-----------|-------------|--------|-------------|
| Basic (A) | 12,813 | 13,306 | 12,813 / 26 x 27 |
| Leave (E) | 1,602 | 1,664 | 1,602 / 26 x 27 |
| Bonus (D) | 1,067 | 1,108 | 1,067 / 26 x 27 |
| **Gross** | **15,482** | **16,078** | |

### Deductions

| Deduction | Rate / Rule | Example (J RAHELU) |
|-----------|-------------|---------------------|
| PF (Employee) | 12% of Basic actual (col AD) | 13,306 x 0.12 = 1,597 |
| ESIC (Employee) | 0.75% of Gross actual (col AK) | 16,078 x 0.0075 = 121 |
| LWF (Employee) | Fixed amount | 2.50 |
| Professional Tax | Slab-based (state rules) | 150 (for gross > 15,000) |
| **Total Deduction** | Sum of above | **1,870.50** |

### Net Salary

```
Net Salary = Gross Actual - Total Deduction
           = 16,078 - 1,870.50
           = 14,208
```

### Employer-Side Costs (not shown on payslip)

These are calculated in the Excel sheet but not displayed on the employee payslip:

| Component | Rate |
|-----------|------|
| PF (Employer, col AM) | 13% of Basic actual |
| ESIC (Employer, col AN) | 3.25% of Gross actual |
| LWF (Employer, col AO) | Fixed amount |

---

## QR Code Verification

Each payslip includes a **QR code** and a **verification code** (e.g. `TS-046F3EFC93B1`) at the bottom.

### What the QR code contains

Scanning the QR code reveals:

```
TECHNO SOLUTIONS - PAYSLIP
Employee: JAJULA SALOMI RAHELU
Code: PPRR01063947
UAN: 102241169775
Month: FEBRUARY 2026
Net Pay: 14208
Generated: 30-03-2026 17:25:53
Verify: TS-046F3EFC93B1
```

### How the verification code is generated

```
Input  = "{emp_code}|{emp_name}|{month}|{net_pay}|{timestamp}"
Hash   = SHA-256(Input)
Code   = "TS-" + first 12 hex characters (uppercased)
```

The hash includes the **generation timestamp** (date + time), so:
- Each PDF generation produces a unique code (even for the same employee)
- The code cannot be copied from one payslip to another
- Tampering with any field (name, amount, date) would produce a different hash

---

## Output

- PDFs are saved as `<EmployeeName>_<MONTH>_<YEAR>.pdf`
- Spaces in names are replaced with `_`, dots and `/` are removed
- Each employee gets a separate PDF file
- Default output directory: `payslips_<MONTH>_<YEAR>/`

### Output File Naming Examples

| Employee Name | Output File |
|---------------|-------------|
| C SUVARNA LAXMI | `C_SUVARNA_LAXMI_FEB_2026.pdf` |
| O.ANITHA DEVI | `OANITHA_DEVI_FEB_2026.pdf` |
| P.SAI SURENDRA BHANU | `PSAI_SURENDRA_BHANU_FEB_2026.pdf` |

---

## Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| `KeyError: 'Wage Sheet'` | Sheet name doesn't match exactly | Ensure the Excel file has a sheet named exactly `Wage Sheet` |
| `KeyError: 'Attendence'` | Sheet name doesn't match | Ensure the sheet is named `Attendence` (with this exact spelling) |
| `No employee data found` | No numeric serial numbers in col A starting from row 5 | Check that Wage Sheet has employee data starting at row 5 with numeric S.No in column A |
| All values show as 0 | Excel file was not opened/saved with calculated values | Open the file in Excel, let formulas calculate, save, then re-run the script |
| `No employees matched` | Name doesn't match exactly (case-insensitive) | Use `--list` to see the exact names available |
