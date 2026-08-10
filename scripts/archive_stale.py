# -*- coding: utf-8 -*-
"""Auto-archive stale tracker rows (jobs.xlsx -> "Archive" sheet).

Policy (age = today minus the date column, i.e. posting date when known):
  - "חדש" older than 21 days           -> archived (posting almost surely closed)
  - "נדחה"/"דילגתי" older than 30 days -> archived (declutter, kept for history)
  - "הוגש"/"ראיון" are NEVER auto-archived.
Archived rows keep every column + archive date + reason, stay in the dedupe set
(add_jobs.existing_keys reads Archive too) so scans never re-add them, and can
be restored by hand (ask Claude).

Usage: python archive_stale.py [--dry-run]
Also runs automatically at the start of every make_dashboard.py rebuild.
"""
import datetime
import sys
from pathlib import Path

from openpyxl import load_workbook

TRACKER = Path(__file__).resolve().parent.parent / "jobs.xlsx"
STALE_NEW_DAYS = 21     # status "חדש"
STALE_CLOSED_DAYS = 30  # status "נדחה" / "דילגתי"
ARCHIVE_SHEET = "Archive"
SUBMIT_DATE_COL = 13    # Jobs col "הוגש בתאריך" (cv_server writes it on ✅)

ARCHIVE_HEADERS = ["#", "תאריך", "חברה", "משרה", "מיקום", "ציון", "למה מתאים",
                   "קישור", "Job ID", "סטטוס", "קובץ CV", "הערות",
                   "הוגש בתאריך", "תאריך ארכוב", "סיבה"]


def parse_date(v):
    try:
        return datetime.date.fromisoformat(str(v)[:10])
    except (ValueError, TypeError):
        return None


def ensure_schema(wb):
    """Archive sheet + the submit-date column on Jobs. Returns (sheet, changed)."""
    changed = False
    ws = wb["Jobs"]
    if ws.cell(row=1, column=SUBMIT_DATE_COL).value is None:
        ws.cell(row=1, column=SUBMIT_DATE_COL, value="הוגש בתאריך")
        ws.column_dimensions["M"].width = 12
        changed = True
    if ARCHIVE_SHEET not in wb.sheetnames:
        ar = wb.create_sheet(ARCHIVE_SHEET)
        ar.sheet_view.rightToLeft = True
        for col, head in enumerate(ARCHIVE_HEADERS, start=1):
            ar.cell(row=1, column=col, value=head)
        changed = True
    return wb[ARCHIVE_SHEET], changed


def run(dry_run=False, today=None):
    """Move stale rows to Archive. Returns (count, ["#16 ISCAR (reason)", ...])."""
    today = today or datetime.date.today()
    wb = load_workbook(TRACKER)
    ws = wb["Jobs"]
    ar, schema_changed = ensure_schema(wb)
    stale = []  # (sheet_row_idx, values, reason)
    for row in ws.iter_rows(min_row=2):
        if row[0].value is None:
            continue
        if "שוחזר" in str(row[11].value or ""):
            continue  # user pulled it back from the archive on purpose - keep it
        status = row[9].value or "חדש"
        d = parse_date(row[1].value)
        if d is None:
            continue  # unparseable date: never auto-archive
        age = (today - d).days
        reason = None
        if status == "חדש" and age >= STALE_NEW_DAYS:
            reason = f"חדש {age} ימים בלי הגשה"
        elif status in ("נדחה", "דילגתי") and age >= STALE_CLOSED_DAYS:
            reason = f"{status}, בן {age} ימים"
        if reason:
            vals = [c.value for c in row[:13]]
            if row[7].hyperlink:  # legacy rows: real URL lives on the hyperlink
                vals[7] = row[7].hyperlink.target
            stale.append((row[0].row, vals, reason))

    moved = [f"#{v[0]} {v[2]} ({r})" for _, v, r in stale]
    if not dry_run:
        for _, vals, reason in stale:
            vals = list(vals) + [None] * (13 - len(vals))
            ar.append(vals + [today.isoformat(), reason])
        for idx, _, _ in sorted(stale, key=lambda s: s[0], reverse=True):
            ws.delete_rows(idx)
        if stale or schema_changed:
            wb.save(TRACKER)
    return len(stale), moved


def main():
    dry = "--dry-run" in sys.argv
    n, moved = run(dry_run=dry)
    tag = "would archive" if dry else "archived"
    print(f"{tag}: {n}")
    for m in moved:
        print(" ", m)


if __name__ == "__main__":
    main()
