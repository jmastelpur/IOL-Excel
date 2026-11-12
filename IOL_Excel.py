"""
IOL_Excel.py
Extract the "IOL calculation" page (page 1) of Zeiss IOLMaster 700 PDFs into Excel.

Usage:
    python IOL_Excel.py                      -> pick a folder in a dialog
    python IOL_Excel.py "C:\\path\\to\\pdfs"   -> process that folder
    python IOL_Excel.py folder -o out.xlsx   -> choose the output file name

How it works:
    The page is read by POSITION, not by counting regex matches. The vertical
    divider in the middle of the table splits OD (left) from OS (right). Each
    formula box is found by its "IOL (D)  Ref (D)" column header, so pages with
    4, 3, 2, 1 or 0 boxes per eye, or with one eye left blank, all work.

Output workbook:
    Summary    - one row per eye: patient info, eye status, biometry, and for
                 each box the lens, formula, constants, the highlighted (bold)
                 IOL/Ref row and the emmetropia IOL.
    IOL Tables - every IOL/Ref line of every box, grouped per PDF: click + on a
                 patient row to expand it, - to collapse (1 / 2 at top-left
                 collapse / expand all). The bold row from the PDF is highlighted.
    Log        - one line per PDF: OK, warnings, or the error that stopped it.
"""

import argparse
import os
import re
import sys
from collections import OrderedDict

import pandas as pd
import pdfplumber

# --------------------------------------------------------------------------
# Labels that appear as "Label:" in the Eye status / Biometric values area.
# Longest first so "LVC mode" wins over "LVC".
# --------------------------------------------------------------------------
LABELS = ["Target ref.", "LVC mode", "LS", "VS", "Ref", "VA", "LVC", "SIA",
          "AL", "ACD", "LT", "WTW", "SE", "SD", "K1", "ΔK", "K2",
          "TSE", "TK1", "ΔTK", "TK2"]
LABEL_RE = re.compile(
    r"(?<!\S)(" + "|".join(re.escape(l) for l in sorted(LABELS, key=len, reverse=True)) + r"):"
)
NUM_RE = re.compile(r"[+-]?\d+(?:\.\d+)?")
NUM_ONLY_RE = re.compile(r"^[+-]?\d+(?:\.\d+)?$")
KERATO_MARKS = ("K", "TK", "SK", "TSK")

TEXT_FIELDS = ["LS", "VS", "Ref", "VA", "LVC", "LVC mode"]
# (label, column name, has axis?)
NUMERIC_FIELDS = [
    ("Target ref.", "Target ref (D)", False),
    ("SIA", "SIA (D)", True),
    ("AL", "AL (mm)", False),
    ("ACD", "ACD (mm)", False),
    ("LT", "LT (mm)", False),
    ("WTW", "WTW (mm)", False),
    ("SE", "SE (D)", False),
    ("K1", "K1 (D)", True),
    ("K2", "K2 (D)", True),
    ("ΔK", "ΔK (D)", True),
    ("TSE", "TSE (D)", False),
    ("TK1", "TK1 (D)", True),
    ("TK2", "TK2 (D)", True),
    ("ΔTK", "ΔTK (D)", True),
]
SD_FIELDS = [("AL", "AL SD (µm)"), ("ACD", "ACD SD (µm)"), ("LT", "LT SD (µm)"),
             ("SE", "SE SD (D)"), ("TSE", "TSE SD (D)")]


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def group_lines(words, tol=3.0):
    """Group words into text lines by vertical position. Returns list of lists."""
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(w["top"] - lines[-1][0]["top"]) <= tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [sorted(l, key=lambda w: w["x0"]) for l in lines]


def line_text(line):
    return " ".join(w["text"] for w in line)


def to_float(s):
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def parse_measure(raw):
    """'42.72D (!)' -> (42.72, None, ['!']);  '41.77D @ 71°' -> (41.77, 71, [])."""
    raw = (raw or "").strip()
    flags = re.findall(r"\((\*|!)\)", raw)
    raw = re.sub(r"\((\*|!)\)", "", raw).strip()
    axis = None
    m = re.search(r"@\s*(\d+)\s*°?", raw)
    if m:
        axis = int(m.group(1))
        raw = raw[:m.start()].strip()
    if raw in ("", "-", "---"):
        return None, axis, flags
    m = NUM_RE.search(raw)
    return (float(m.group(0)) if m else None), axis, flags


def find_phrase(lines, phrase):
    """Return the first line whose text contains phrase, or None."""
    for l in lines:
        if phrase in line_text(l):
            return l
    return None


# --------------------------------------------------------------------------
# Page geometry
# --------------------------------------------------------------------------
def find_divider_x(page):
    """x position of the long vertical line splitting OD from OS."""
    mid = page.width / 2
    candidates = []
    for obj in list(page.lines) + list(page.rects):
        if abs(obj["x1"] - obj["x0"]) < 3 and (obj["bottom"] - obj["top"]) > 150:
            x = (obj["x0"] + obj["x1"]) / 2
            if abs(x - mid) < page.width * 0.15:
                candidates.append(x)
    if candidates:
        return min(candidates, key=lambda x: abs(x - mid))
    return page.width * 0.515  # fallback that matches the IOLMaster 700 layout


# --------------------------------------------------------------------------
# Header (patient) section
# --------------------------------------------------------------------------
def parse_header(lines, page_width):
    info = {"Name": "", "ID": "", "Date of measurement": ""}
    for l in lines:
        # Ignore the clinic name/address block printed on the right side
        l = [w for w in l if w["x0"] < page_width * 0.58]
        if not l:
            continue
        t = line_text(l)
        if t.startswith("Patient ID"):
            rest = t[len("Patient ID"):].split()
            info["ID"] = rest[0] if rest else ""
        elif t.startswith("Patient") and not info["Name"]:
            info["Name"] = t[len("Patient"):].strip()
        elif t.startswith("Date of measurement:"):
            m = re.search(r"Date of measurement:\s*(\S+)", t)
            if m:
                info["Date of measurement"] = m.group(1)
    return info


# --------------------------------------------------------------------------
# Eye status + biometric values for one eye
# --------------------------------------------------------------------------
def parse_eye_fields(lines):
    raw = {}
    sd = {}
    for l in lines:
        t = line_text(l)
        parts = LABEL_RE.split(t)
        # parts = [before, label1, value1, label2, value2, ...]
        prev = None
        for i in range(1, len(parts) - 1, 2):
            label, value = parts[i], parts[i + 1].strip()
            if label == "SD":
                if prev:
                    sd[prev] = value
            else:
                raw.setdefault(label, value)
                prev = label

    out = OrderedDict()
    flags = []
    for lab in TEXT_FIELDS:
        v = raw.get(lab, "").strip()
        out[lab] = "" if v in ("---", "-") else v
    for lab, col, has_axis in NUMERIC_FIELDS:
        val, axis, fl = parse_measure(raw.get(lab))
        out[col] = val
        if has_axis:
            out[f"{col.split(' ')[0]} axis (°)"] = axis
        flags += [f"{lab}({f})" for f in fl]
    for lab, col in SD_FIELDS:
        out[col] = parse_measure(sd.get(lab))[0]
    out["Flags"] = ", ".join(flags)  # (*) = edited manually, (!) = borderline
    return out, bool(raw)


# --------------------------------------------------------------------------
# Formula boxes for one eye
# --------------------------------------------------------------------------
def find_box_headers(lines):
    """Lines containing 'IOL (D) ... Ref (D)'. A line may hold 1 or 2 headers."""
    headers = []
    for l in lines:
        for i, w in enumerate(l[:-1]):
            if w["text"] == "IOL" and l[i + 1]["text"] == "(D)":
                ref = next((r for r in l[i + 2:] if r["text"] == "Ref"), None)
                if ref is None:
                    continue
                ref_d = l[l.index(ref) + 1] if l.index(ref) + 1 < len(l) else ref
                headers.append({"top": w["top"], "x0": w["x0"], "x1": ref_d["x1"],
                                "split": (l[i + 1]["x1"] + ref["x0"]) / 2})
    return headers


def cluster(values, tol):
    """Cluster sorted numbers, returning the mean of each cluster."""
    groups = []
    for v in sorted(values):
        if groups and v - groups[-1][-1] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def parse_box(words, header, target_ref, highlights=()):
    lines = group_lines(words)
    above = [l for l in lines if l[0]["top"] < header["top"] - 2]
    below = [l for l in lines if l[0]["top"] > header["top"] + 2]

    # Lens name, formula, constants (lines above the column header)
    name_lines, formula, const_lines = [], "", []
    seen_formula = False
    for l in above:
        t = line_text(l)
        m = re.match(r"^-\s*(.+?)\s*-$", t)
        if m and not seen_formula:
            formula, seen_formula = m.group(1), True
        elif seen_formula:
            const_lines.append(t)
        else:
            # strip the little "K" / "TK" keratometry markers at the box corners
            # (a neighbouring box's marker can sit right at the cell edge)
            l = [w for k, w in enumerate(l)
                 if not (w["text"] in KERATO_MARKS and (k == 0 or k == len(l) - 1))]
            if l:
                name_lines.append(line_text(l))
    if not seen_formula and len(name_lines) > 1:
        const_lines, name_lines = name_lines[1:], name_lines[:1]

    # IOL / Ref rows and emmetropia line
    rows, emmetropia = [], None
    for l in below:
        t = line_text(l)
        nums = [w for w in l if NUM_ONLY_RE.match(w["text"])]
        if "Emmetropia" in t:
            if nums:
                emmetropia = float(nums[0]["text"])
            break
        iol = next((float(w["text"]) for w in nums
                    if (w["x0"] + w["x1"]) / 2 < header["split"]), None)
        ref = next((float(w["text"]) for w in nums
                    if (w["x0"] + w["x1"]) / 2 >= header["split"]), None)
        if iol is not None or ref is not None:
            rows.append({"IOL (D)": iol, "Ref (D)": ref,
                         "_y": (l[0]["top"] + l[0]["bottom"]) / 2})

    # Selected (bold) row: the one sitting on the grey highlight bar.
    # If no bar is found, use the row whose Ref is closest to target.
    sel, source = None, ""
    for i, r in enumerate(rows):
        if any(h["top"] - 1 <= r["_y"] <= h["bottom"] + 1
               and h["x0"] < header["x1"] and h["x1"] > header["x0"] for h in highlights):
            sel, source = i, "highlighted"
            break
    if sel is None:
        with_ref = [i for i, r in enumerate(rows) if r["Ref (D)"] is not None]
        if with_ref:
            t = target_ref if target_ref is not None else 0.0
            sel = min(with_ref, key=lambda i: abs(rows[i]["Ref (D)"] - t))
            source = "closest to target"
    for i, r in enumerate(rows):
        r["Selected"] = (i == sel)
        r.pop("_y", None)

    return {
        "Lens": " ".join(name_lines).strip(),
        "Formula": formula,
        "Constants": " ".join(const_lines).strip(),
        "Rows": rows,
        "Selected": rows[sel] if sel is not None else {},
        "Selected_source": source,
        "Emmetropia (D)": emmetropia,
    }


def find_highlights(page):
    """Grey bars the IOLMaster draws behind the selected (bold) IOL/Ref row."""
    out = []
    for r in page.rects:
        c = r.get("non_stroking_color")
        if not (6 < r["height"] < 16 and r["width"] > 40 and isinstance(c, (tuple, list)) and c):
            continue
        vals = list(c)
        if len(vals) in (1, 3) and all(0.85 <= v < 0.99 for v in vals):
            out.append(r)
    return out


def parse_boxes(eye_words, x_left, x_right, top_limit, bottom_limit, target_ref, highlights=()):
    lines = group_lines(eye_words)
    headers = find_box_headers(lines)
    if not headers:
        return []

    # Grid: rows by header top, columns by header centre.
    row_tops = cluster([h["top"] for h in headers], tol=5)
    col_centres = cluster([(h["x0"] + h["x1"]) / 2 for h in headers], tol=15)
    col_bounds = [x_left] + [(a + b) / 2 for a, b in zip(col_centres, col_centres[1:])] + [x_right]

    # Row boundaries: start just below the previous row's Emmetropia line
    # (or below the biometric block for the first row).
    row_bounds = [top_limit]
    for rt in row_tops[1:]:
        emm = [w["top"] for w in eye_words if w["text"] == "Emmetropia" and w["top"] < rt]
        row_bounds.append(max(emm) + 6 if emm else rt - 45)
    row_bounds.append(bottom_limit)

    boxes = []
    for r, rt in enumerate(row_tops):
        for c in range(len(col_centres)):
            h = next((h for h in headers
                      if abs(h["top"] - rt) <= 5
                      and col_bounds[c] <= (h["x0"] + h["x1"]) / 2 < col_bounds[c + 1]), None)
            if h is None:
                continue
            cell = [w for w in eye_words
                    if col_bounds[c] <= (w["x0"] + w["x1"]) / 2 < col_bounds[c + 1]
                    and row_bounds[r] <= w["top"] < row_bounds[r + 1]]
            boxes.append(parse_box(cell, h, target_ref, highlights))
    return boxes


# --------------------------------------------------------------------------
# Whole page
# --------------------------------------------------------------------------
def extract_page(page):
    words = [w for w in page.extract_words(x_tolerance=3, y_tolerance=3)
             if w.get("upright", True)]
    lines = group_lines(words)
    warnings = []

    title = find_phrase(lines, "IOL calculation")
    if title is None:
        raise ValueError("No 'IOL calculation' section on page 1")
    title_top = title[0]["top"]

    legend = find_phrase(lines, "Borderline value") or find_phrase(lines, "Comment")
    bottom = legend[0]["top"] - 1 if legend else page.height

    header = parse_header([l for l in lines if l[0]["top"] < title_top - 5], page.width)
    highlights = find_highlights(page)

    divider = find_divider_x(page)
    eyes = {}
    for eye, (xl, xr) in {"OD": (0, divider), "OS": (divider, page.width)}.items():
        ew = [w for w in words
              if title_top + 5 < w["top"] < bottom
              and xl <= (w["x0"] + w["x1"]) / 2 < xr
              and w["x1"] > 45]          # skip the rotated copyright margin text
        elines = group_lines(ew)
        hdrs = find_box_headers(elines)
        first_box_top = min(h["top"] for h in hdrs) if hdrs else bottom

        # Biometry/status = everything above the first box's name line.
        field_lines = [l for l in elines if l[0]["top"] < first_box_top - 50]
        if not hdrs:
            field_lines = elines
        fields, has_data = parse_eye_fields(field_lines)

        # Boxes start just below the last labelled line.
        label_tops = [l[0]["top"] for l in field_lines if LABEL_RE.search(line_text(l))]
        box_top = (max(label_tops) + 8) if label_tops else title_top + 5
        boxes = parse_boxes(ew, xl, xr, box_top, bottom, fields.get("Target ref (D)"), highlights)

        has_biometry = fields.get("AL (mm)") is not None or fields.get("K1 (D)") is not None
        if not has_data and not boxes:
            warnings.append(f"{eye} is blank")
        elif not boxes:
            warnings.append(f"{eye} has no IOL calculation"
                            + ("" if has_biometry else " or biometry (eye status only)"))
        if any(b.get("Selected_source") == "closest to target" for b in boxes):
            warnings.append(f"{eye}: highlighted row not found in some boxes; used Ref closest to target")
        eyes[eye] = {"fields": fields, "boxes": boxes}

    return header, eyes, warnings


def process_pdfs(folder):
    results, log = [], []
    files = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
    for fn in files:
        path = os.path.join(folder, fn)
        try:
            with pdfplumber.open(path) as pdf:
                header, eyes, warnings = extract_page(pdf.pages[0])
                n_pages = len(pdf.pages)
            results.append({"File": fn, "header": header, "eyes": eyes})
            note = f"{len(eyes['OD']['boxes'])} OD box(es), {len(eyes['OS']['boxes'])} OS box(es)"
            if n_pages > 1:
                note += f"; {n_pages} pages, only page 1 read"
            if warnings:
                note += "; " + "; ".join(warnings)
            log.append({"File": fn, "Status": "OK", "Notes": note})
            print(f"OK      {fn}  ({note})")
        except Exception as e:
            log.append({"File": fn, "Status": "FAILED", "Notes": str(e)})
            print(f"FAILED  {fn}: {e}")
    return results, log


# --------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------
def build_summary(results):
    max_boxes = max([len(e["boxes"]) for r in results for e in r["eyes"].values()] + [0])
    summary = []
    for r in results:
        h = r["header"]
        for eye in ("OD", "OS"):
            e = r["eyes"][eye]
            row = OrderedDict([("File", r["File"]), ("Name", h["Name"]), ("ID", h["ID"]),
                               ("Date of measurement", h["Date of measurement"]), ("Eye", eye)])
            row.update(e["fields"])
            for i in range(max_boxes):
                b = e["boxes"][i] if i < len(e["boxes"]) else {}
                p = f"Box {i + 1}"
                row[f"{p} Lens"] = b.get("Lens", "")
                row[f"{p} Formula"] = b.get("Formula", "")
                row[f"{p} Constants"] = b.get("Constants", "")
                row[f"{p} IOL (D)"] = b.get("Selected", {}).get("IOL (D)")
                row[f"{p} Ref (D)"] = b.get("Selected", {}).get("Ref (D)")
                row[f"{p} Emmetropia (D)"] = b.get("Emmetropia (D)")
            summary.append(row)
    return pd.DataFrame(summary)


DIOPTER_FMT = "+0.00;-0.00;0.00"


def write_iol_tables(wb, results):
    """
    One bold patient row per PDF with a +/- button; the IOL/Ref lines for
    that PDF sit underneath in a collapsible group (collapsed by default).
    The 1 / 2 buttons at the top-left collapse / expand everything at once.
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    ws = wb.create_sheet("IOL Tables")
    ws.sheet_properties.outlinePr.summaryBelow = False   # +/- on the patient row

    cols = ["Name", "ID", "Date", "Eye", "Box", "Lens", "Formula", "Constants",
            "IOL (D)", "Ref (D)", "Selected", "Emmetropia (D)", "File"]
    widths = [24, 13, 12, 6, 5, 28, 24, 22, 9, 9, 9, 14, 40]
    ws.append(cols)
    head_fill = PatternFill("solid", fgColor="DDE5F0")
    for c, w in zip(ws[1], widths):
        c.font = Font(bold=True)
        c.fill = head_fill
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"

    pt_fill = PatternFill("solid", fgColor="EEF2F7")
    sel_fill = PatternFill("solid", fgColor="FFF2CC")
    eye_fill = PatternFill("solid", fgColor="F7F7F7")
    thin = Side(style="thin", color="B7C3D3")
    box_top = Border(top=Side(style="thin", color="D9D9D9"))

    for r in sorted(results, key=lambda r: (r["header"]["Name"], r["header"]["Date of measurement"])):
        h = r["header"]
        nb = {e: len(r["eyes"][e]["boxes"]) for e in ("OD", "OS")}
        summary_txt = f"OD: {nb['OD']} box(es)   OS: {nb['OS']} box(es)"
        ws.append([h["Name"], h["ID"], h["Date of measurement"], "", "", summary_txt,
                   "", "", None, None, "", None, r["File"]])
        pr = ws.max_row
        for c in ws[pr]:
            c.font = Font(bold=True)
            c.fill = pt_fill
            c.border = Border(top=thin)
        ws.cell(pr, 6).font = Font(italic=True, color="555555")
        ws.cell(pr, 13).font = Font(color="555555")

        first = pr + 1
        for eye in ("OD", "OS"):
            boxes = r["eyes"][eye]["boxes"]
            if not boxes:
                ws.append(["", "", "", eye, "", "(no IOL calculation)"])
                ws.cell(ws.max_row, 6).font = Font(italic=True, color="888888")
                continue
            for bi, b in enumerate(boxes, start=1):
                for li, br in enumerate(b["Rows"]):
                    ws.append(["", "", "", eye, bi,
                               b["Lens"] if li == 0 else "",
                               b["Formula"] if li == 0 else "",
                               b["Constants"] if li == 0 else "",
                               br["IOL (D)"], br["Ref (D)"],
                               "Yes" if br["Selected"] else "",
                               b["Emmetropia (D)"] if li == 0 else None, ""])
                    rr = ws.max_row
                    for col in (9, 10, 12):
                        ws.cell(rr, col).number_format = DIOPTER_FMT
                    if li == 0:
                        for c in ws[rr]:
                            c.border = box_top
                    if br["Selected"]:
                        for col in range(4, 13):
                            ws.cell(rr, col).fill = sel_fill
                            ws.cell(rr, col).font = Font(bold=True)
                    elif eye == "OS":
                        for col in range(4, 13):
                            ws.cell(rr, col).fill = eye_fill
        last = ws.max_row
        if last >= first:
            ws.row_dimensions.group(first, last, outline_level=1, hidden=True)
            ws.row_dimensions[pr].collapsed = True

    for row in ws.iter_rows(min_row=1):
        for c in row:
            centred = c.column in (4, 5, 11) and c.row > 1
            c.alignment = Alignment(vertical="top", horizontal="center" if centred else None)


def export_to_excel(results, log, out_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils.dataframe import dataframe_to_rows

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="DDE5F0")

    def write_df(ws, df):
        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = head_fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            width = max([len(str(c.value)) for c in col[1:] if c.value is not None] + [0])
            width = max(width, min(len(str(col[0].value or "")), 18))
            ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 8), 40)

    ws = wb.active
    ws.title = "Summary"
    summary = build_summary(results)
    write_df(ws, summary)
    ws.freeze_panes = "F2"
    signed = [i + 1 for i, c in enumerate(summary.columns)
              if c.startswith("Target ref") or c.startswith("ΔK") or c.startswith("ΔTK")
              or c.endswith("IOL (D)") or c.endswith("Ref (D)") or c.endswith("Emmetropia (D)")]
    for row in ws.iter_rows(min_row=2):
        for ci in signed:
            row[ci - 1].number_format = DIOPTER_FMT

    write_iol_tables(wb, results)

    ws = wb.create_sheet("Log")
    write_df(ws, pd.DataFrame(log, columns=["File", "Status", "Notes"]))

    wb.save(out_path)

# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="IOLMaster 700 PDFs -> Excel")
    ap.add_argument("folder", nargs="?", help="folder containing the PDFs")
    ap.add_argument("-o", "--output", help="output .xlsx (default: <folder>/extracted_data.xlsx)")
    args = ap.parse_args()

    folder = args.folder
    if not folder:
        from tkinter import Tk, filedialog
        root = Tk()
        root.withdraw()
        folder = filedialog.askdirectory(title="Select the folder containing PDF files")
        root.destroy()
    if not folder:
        print("No folder selected.")
        return

    results, log = process_pdfs(folder)
    if not log:
        print("No PDF files found.")
        return
    out = args.output or os.path.join(folder, "extracted_data.xlsx")
    export_to_excel(results, log, out)
    ok = sum(1 for l in log if l["Status"] == "OK")
    print(f"\nDone: {ok}/{len(log)} PDFs extracted -> {out}")


if __name__ == "__main__":
    main()