"""
Entry point.

Usage:
    python main.py

Place the source workbook (.xlsx) inside the input/ directory.
Results are written to output/.  Logs go to logs/run_log.log.
"""

import logging
import os
import sys

from config import INPUT_DIR, LOG_FILE, LOGS_DIR, OUTPUT_DIR
from utils.consolidator import consolidate
from utils.excel_reader import read_workbook
from utils.exporter import export_bom, export_exceptions, safety_check


# ── Logger setup ──────────────────────────────────────────────────────────────

def _setup_logger() -> logging.Logger:
    os.makedirs(LOGS_DIR, exist_ok=True)
    log_path = os.path.join(LOGS_DIR, LOG_FILE)

    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("bom_consolidator")


# ── Main pipeline ─────────────────────────────────────────────────────────────

def main() -> None:
    logger = _setup_logger()
    logger.info("═══ BoM Consolidation run started ═══════════════════════════")

    for d in (INPUT_DIR, OUTPUT_DIR, LOGS_DIR):
        os.makedirs(d, exist_ok=True)

    # ── Find input workbook ───────────────────────────────────────────────
    candidates = [
        f for f in os.listdir(INPUT_DIR)
        if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")
    ]
    if not candidates:
        logger.error("No Excel workbook found in input/. Add your .xlsx file and re-run.")
        sys.exit(1)

    input_file = os.path.join(INPUT_DIR, candidates[0])
    logger.info(f"Input file  : {input_file}")

    # ── Step 1: Read ──────────────────────────────────────────────────────
    sheets_data, exceptions = read_workbook(input_file, logger)

    total_sheets = len(sheets_data)
    total_rows   = sum(len(s["data_rows"]) for s in sheets_data)
    logger.info(f"Sheets with valid BoM content : {total_sheets}")
    logger.info(f"Total valid rows extracted    : {total_rows}")

    if total_sheets == 0:
        logger.error("No usable sheets found. Check the input workbook structure.")
        # FIX: pass empty list for all_panels — consolidate() hasn't run yet
        export_exceptions(exceptions, [], logger)
        sys.exit(1)

    # ── Step 2: Consolidate ───────────────────────────────────────────────
    all_panels, sorted_rows = consolidate(sheets_data)
    logger.info(f"Unique panels collected       : {len(all_panels)}")
    logger.info(f"Rows after sort               : {len(sorted_rows)}")

    # ── Step 3: Export ────────────────────────────────────────────────────
    export_bom(all_panels, sorted_rows, logger)
    export_exceptions(exceptions, all_panels, logger)

    # ── Step 4: Safety check ──────────────────────────────────────────────
    ok = safety_check(all_panels, sorted_rows, sheets_data, logger)

    logger.info(f"Exceptions flagged            : {len(exceptions)}")
    logger.info("═══ Run complete ════════════════════════════════════════════")

    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()