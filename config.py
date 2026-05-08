import os

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR  = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
LOGS_DIR   = os.path.join(BASE_DIR, "logs")

# ── File names ─────────────────────────────────────────────────────────────────
LOG_FILE        = "run_log.log"
OUTPUT_BOM      = "final_bom.xlsx"
EXCEPTIONS_FILE = "exceptions.xlsx"

# ── Source workbook layout (1-indexed, as Excel shows) ─────────────────────────
PANEL_ROW        = 1   # row 1 holds panel names
HEADER_ROW       = 4   # row 4 holds static column headers
PANEL_START_COL  = 7   # G = column 7; panel names begin here

# Column headers expected in row 4 of the source
STATIC_COLS = ["SNo", "DESCRIPTION", "SPEC", "MAKE", "UNIT", "CAT NO."]
PRIMARY_KEY = "CAT NO."

# ── Output layout (1-indexed) ──────────────────────────────────────────────────
OUT_ROW_PANELS  = 1   # panel names go in row 1
OUT_ROW_HEADER  = 4   # static headers go in row 4
OUT_DATA_START  = 5   # data begins on row 5

OUT_COL_SR    = 1   # A  → SR NO.
OUT_COL_DESC  = 2   # B  → DESCRIPTION
OUT_COL_SPEC  = 3   # C  → SPEC
OUT_COL_MAKE  = 4   # D  → MAKE
OUT_COL_CATNO = 5   # E  → CAT NO.
OUT_COL_UNIT  = 6   # F  → UNIT
OUT_COL_MODEL = 6   # F1 → "MODEL" label
OUT_PANEL_START_COL = 7  # G  → first panel column

# ── Styling hex codes (no leading #) ──────────────────────────────────────────
FILL_PINK  = "FFB6C1"   # row 4 header cells
FILL_GREEN = "92D050"   # row 1 panel name cells
FONT_BLACK = "000000"