#!/usr/bin/env python3
"""
sheet.py - Automated Grader Interface for FieldServe Ticket Board
Usage:
  python3 sheet.py export /path/to/out.csv
  python3 sheet.py apply /path/to/edits.csv
"""
import sys
import os
import csv

SHEET_PATH = os.path.join(os.path.dirname(__file__), "tickets.xlsx")
REQUIRED_COLS = [
    "Ticket ID", "Subject", "Status", "Engineer", "Manager",
    "Purchase Order", "Delivery Note", "Work Report", "Invoice",
    "Stage", "Action Owner"
]


def load_wb():
    import openpyxl
    if not os.path.exists(SHEET_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tickets"
        ws.append(REQUIRED_COLS)
        wb.save(SHEET_PATH)
    return openpyxl.load_workbook(SHEET_PATH)


def get_tickets_sheet(wb):
    if "Tickets" in wb.sheetnames:
        return wb["Tickets"]
    return wb.active


def export_csv(out_path):
    """
    Exports the current state of tickets to CSV.
    Resolves any formulas (values_only=True) and formats nulls as empty strings.
    """
    wb = load_wb()
    ws = get_tickets_sheet(wb)
    headers = [c.value for c in ws[1]]
    if not headers:
        headers = REQUIRED_COLS

    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            # Format None as ""
            out = ["" if v is None else str(v) for v in row]
            writer.writerow(out)

    print(f"Exported to {out_path}")


def apply_edits(edits_path):
    """
    Applies cell edits from a CSV file (columns: Ticket ID, Column, Value).
    Writes each value into that cell of the sheet, exactly as if a human typed it.
    Does NOT trigger sync as a side effect.
    """
    wb = load_wb()
    ws = get_tickets_sheet(wb)

    headers = [c.value for c in ws[1]]
    col_map = {h: i + 1 for i, h in enumerate(headers)}  # 1-indexed

    if "Ticket ID" not in col_map:
        print("Ticket ID column missing in spreadsheet", file=sys.stderr)
        sys.exit(1)

    # Build Ticket ID -> row number map
    id_to_row = {}
    for r in range(2, ws.max_row + 1):
        tid = ws.cell(row=r, column=col_map["Ticket ID"]).value
        if tid is not None:
            try:
                id_to_row[int(str(tid).strip())] = r
            except Exception:
                pass

    with open(edits_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for line in reader:
            tid_raw = line.get("Ticket ID", "")
            if not tid_raw:
                continue
            try:
                tid = int(str(tid_raw).strip())
            except Exception:
                continue

            col_name = str(line.get("Column", "")).strip()
            value = line.get("Value", "")

            # If column is new, append it
            if col_name not in col_map:
                new_col_idx = len(headers) + 1
                ws.cell(row=1, column=new_col_idx, value=col_name)
                headers.append(col_name)
                col_map[col_name] = new_col_idx

            # Find or append row
            row_num = id_to_row.get(tid)
            if row_num is None:
                row_num = ws.max_row + 1
                ws.cell(row=row_num, column=col_map["Ticket ID"], value=tid)
                id_to_row[tid] = row_num

            # Write value
            col_idx = col_map[col_name]
            ws.cell(row=row_num, column=col_idx, value=value)

    wb.save(SHEET_PATH)
    print(f"Applied edits from {edits_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 sheet.py export /path/to/out.csv OR python3 sheet.py apply /path/to/edits.csv")
        sys.exit(1)

    cmd = sys.argv[1]
    path = sys.argv[2]
    if cmd == "export":
        export_csv(path)
    elif cmd == "apply":
        apply_edits(path)
    else:
        print("Unknown command:", cmd, file=sys.stderr)
        sys.exit(1)
