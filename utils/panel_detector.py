"""
Reads row 1 of a worksheet from column G (PANEL_START_COL) onwards.
Returns an ordered list of (col_index, panel_name) tuples.
Stops at the first empty cell — the gap is intentional in this workbook:
it separates the panel columns from trailing summary columns (e.g.
'KEPL INTERNAL COST') which must not be treated as panels.
"""

from config import PANEL_ROW, PANEL_START_COL


def detect_panels(ws) -> list[tuple[int, str]]:
    panels: list[tuple[int, str]] = []
    col = PANEL_START_COL

    while True:
        value = ws.cell(row=PANEL_ROW, column=col).value
        if value is None or str(value).strip() == "":
            break
        panels.append((col, str(value).strip()))
        col += 1

    return panels