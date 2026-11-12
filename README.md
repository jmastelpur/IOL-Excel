# IOL_Excel

Pulls the **IOL calculation** page out of Zeiss IOLMaster 700 PDF reports and puts it in an Excel workbook. Point it at a folder of PDFs and you get one workbook covering every patient in that folder.

For each PDF it extracts:

- the patient's name, ID and date of measurement
- the eye status for OD and OS (LS, VS, Ref, VA, LVC, target refraction, SIA)
- the biometric values (AL, ACD, LT, WTW, SE, K1, K2, ΔK, and TSE/TK1/TK2/ΔTK when they're present), including SDs and axes
- every formula box: the lens, the formula, the constants, all IOL/Ref lines, the selected (bold) line, and the emmetropia IOL

---

## Requirements

- Python 3.9 or newer
- These packages:

```bash
pip install pdfplumber pandas openpyxl
```

`tkinter` (used for the folder-picker window) comes with the standard Python installers for Windows and macOS. On Linux you may need `sudo apt install python3-tk`.

---

## How to run it

**Option 1: folder picker**

```bash
python IOL_Excel.py
```

A window opens. Choose the folder with your PDFs. The results are saved in that folder as `extracted_data.xlsx`.

**Option 2: give the folder on the command line**

```bash
python IOL_Excel.py "C:\path\to\pdfs"
```

**Option 3: choose the output file name or location**

```bash
python IOL_Excel.py "C:\path\to\pdfs" -o "C:\path\to\results.xlsx"
```

The console shows a line for each PDF as it goes:

```
OK      Example_6.pdf  (3 OD box(es), 3 OS box(es))
OK      BAKER_JORDAN_..._OU_....pdf  (4 OD box(es), 4 OS box(es); 8 pages, only page 1 read)
OK      ROUSSEAU_DENISE_..._OD_....pdf  (4 OD box(es), 0 OS box(es); OS has no IOL calculation or biometry (eye status only))

Done: 11/11 PDFs extracted -> C:\path\to\pdfs\extracted_data.xlsx
```

> **Note:** close `extracted_data.xlsx` in Excel before re-running. Otherwise Windows won't let the script overwrite it.

---

## What's in the workbook

### Summary

One row per eye, so each PDF gets an OD row and an OS row.

| Column group | Columns |
|---|---|
| Patient | File, Name, ID, Date of measurement, Eye |
| Eye status | LS, VS, Ref, VA, LVC, LVC mode, Target ref (D), SIA (D), SIA axis (°) |
| Biometry | AL (mm), ACD (mm), LT (mm), WTW (mm), SE (D), K1 (D) + axis, K2 (D) + axis, ΔK (D) + axis |
| Total keratometry | TSE (D), TK1, TK2, ΔTK (+ axes), only filled when the report has them |
| Standard deviations | AL SD (µm), ACD SD (µm), LT SD (µm), SE SD (D), TSE SD (D) |
| Flags | Values marked on the report: `(*)` = edited manually, `(!)` = borderline value |
| Box 1 to Box N | Lens, Formula, Constants, IOL (D), Ref (D), Emmetropia (D) |

For each box, **IOL (D)** and **Ref (D)** are the bold, highlighted line from the PDF. All numbers are stored as real numbers, so you can sort, filter and use them in formulas. Diopter values are shown with a sign, for example `+19.00` and `-0.05`.

The number of Box column sets matches the largest number of boxes found in any PDF. Eyes with fewer boxes leave the extra columns blank.

### IOL Tables

All five IOL/Ref lines from every box, **grouped by patient**:

- The sheet opens **collapsed**, with one bold row per PDF showing name, ID, date and how many boxes each eye has. Rows are sorted by name.
- Click **+** at the left of a patient's row to expand their lines, and **−** to collapse them again.
- The **1** / **2** buttons at the top-left collapse or expand everyone at once.
- Within a patient, OD boxes come first, then OS. Each box starts with a thin divider line. The selected (bold) line is shaded yellow and marked `Yes` in the *Selected* column. OS rows have a light grey background.

### Log

One line per PDF:

- **Status:** `OK` or `FAILED`
- **Notes:** the number of boxes found per eye, plus anything worth knowing, such as:
  - `8 pages, only page 1 read`
  - `OS is blank`
  - `OS has no IOL calculation or biometry (eye status only)`
  - `OD: highlighted row not found in some boxes; used Ref closest to target`
  - for failures, the error message (for example `No 'IOL calculation' section on page 1`)

One bad PDF never stops the run. It's listed as FAILED and the script moves on to the next file.

---

## Layouts it handles

- **4, 3, 2 or 1 formula box per eye.** Each box is found by its `IOL (D) / Ref (D)` header, not by its position on the page.
- **A blank eye, or an eye with status only.** OD-only and OS-only reports work.
- **Any lens maker or formula** (Barrett Universal II, Barrett TK, Barrett True-K, SRK/T, and so on).
- **Multi-page reports.** Only page 1 is read.

---

## How it works

The script reads each word's **position** on page 1 with `pdfplumber`, rather than the raw text stream:

1. **Header.** Patient name, ID and date are read from above the "IOL calculation" title. The clinic name and address block on the right side is ignored.
2. **OD / OS split.** The long vertical line in the middle of the table separates the left eye column (OD) from the right (OS).
3. **Eye status and biometry.** `Label: value` pairs (`AL:`, `K1:`, `Target ref.:`, …) are read from the top of each eye's column.
4. **Formula boxes.** Each `IOL (D)  Ref (D)` column header marks one box. The headers are arranged in a grid, and each box takes the words in its grid cell: the lens name and formula above the header, and the IOL/Ref lines down to the "Emmetropia" line.
5. **Selected line.** The IOLMaster draws a light-grey bar behind the bold line. The script uses whichever line sits on that bar. If no bar is found, it falls back to the line whose Ref is closest to the target refraction, and notes that in the Log.

---

## Troubleshooting

| Problem | What to check |
|---|---|
| `FAILED ... No 'IOL calculation' section on page 1` | Page 1 isn't an IOL calculation page (for example a biometry-only or keratometry printout). |
| A value is blank that shows on the PDF | The PDF may be a scanned image instead of a digital export. This script needs real text, so scanned PDFs need OCR first. |
| Name or ID looks wrong | Check that the report header uses the standard `Patient` / `Patient ID` labels. |
| `PermissionError` when saving | The output workbook is open in Excel. Close it and run again. |
| Same patient appears twice | Two PDFs of the same exam are both in the folder. Each file gets its own entry. |

If a report comes out wrong, keep the PDF. It's the quickest way to adjust the parsing for that layout.

---

## Files

| File | Purpose |
|---|---|
| `IOL_Excel.py` | The extractor script |
| `README.md` | This file |
| `extracted_data.xlsx` | Output, created in the PDF folder by default |