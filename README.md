# Master Bill of Materials (BoM) Consolidation Engine

## Overview

The BoM Consolidation Engine is a Python-based automated data pipeline developed to ingest, normalize, and consolidate complex, multi-sheet Bill of Materials (BoM) workbooks. Designed specifically to handle intricate industrial rate masters (such as the Daikin Rate Master), the system seamlessly extracts component data while programmatically navigating unstructured visual gaps and embedded financial calculation matrices.

The engine translates disparate, loosely structured Excel sheets into a single, analytics-ready master ledger, optimizing workflow efficiency for KEPL operations.

## Core Architecture & Features

The system is built on a modular architecture, ensuring strict separation of concerns, high fault tolerance, and scalability.

* **Sentinel Termination Protocol:** Employs dynamic boundary detection. The parser identifies where structural BoM data ends and financial footers (e.g., "KEPL INTERNAL COST", "DAIKIN PROPOSED", "CONVERSION %") begin, automatically halting extraction to prevent data corruption and false-positive exceptions.
* **Composite Identity Hashing:** Components are merged across sheets using a deterministic identity hash derived from the `CAT NO.` and `DESCRIPTION` fields, ensuring accurate quantity aggregation without duplicating identical parts.
* **Dynamic Category Aggregation:** When a single component is classified under multiple domains across different sheets (e.g., "ELECTRICAL" vs. "WIRING"), the engine preserves the full context by concatenating the categories rather than truncating the data.
* **Void Bypass Logic:** The extraction loop contains built-in tolerance for human formatting inconsistencies, automatically bypassing blank rows and visual spacing without prematurely truncating the dataset.
* **Non-Fatal Exception Routing:** Invalid rows lacking critical identifiers do not interrupt the pipeline. They are isolated, logged, and exported to a dedicated `exceptions.xlsx` audit file for manual review.

## Directory Structure

```text
BoM Architecture/
├── config.py                 # Global constants, file paths, and layout definitions
├── main.py                   # System entry point and orchestrator
├── input/                    # Target directory for source .xlsx workbooks
├── output/
│   ├── final_bom.xlsx        # The consolidated master dataset
│   └── exceptions.xlsx       # Audit log for flagged data anomalies
├── logs/
│   └── run_log.log           # Real-time execution and error tracking logs
└── utils/
    ├── consolidator.py       # Identity hashing and grouping logic
    ├── excel_reader.py       # Data ingestion and Sentinel Termination logic
    ├── exporter.py           # Output formatting, styling, and safety validation
    ├── panel_detector.py     # Dynamic column mapping for panel headers
    └── row_cleaner.py        # Data sanitization and validation functions

```

## Prerequisites

* **Python 3.12+**
* **Dependencies:** Defined in `requirements.txt`. Install via:
```bash
pip install -r requirements.txt

```


*(Core dependency: `openpyxl`)*

## Execution Protocol

1. **Initialization:** Place the target Excel workbook into the `input/` directory. Ensure it is the only `.xlsx` or `.xls` file in the folder to prevent ambiguity.
2. **Execution:** Run the primary orchestrator from the terminal:
```bash
python main.py

```


3. **Monitoring:** Review the terminal output or `logs/run_log.log` for real-time processing metrics, including sheets parsed, panels detected, and sentinel terminations triggered.
4. **Retrieval:** Access the compiled master dataset in `output/final_bom.xlsx`. Review `output/exceptions.xlsx` to audit any skipped rows.

## Design Methodology

This architecture addresses the inherent risks of manual data consolidation—human error, fatigue, and systemic redundancy. By abstracting the consolidation process into an autonomous, identity-based engine, the system guarantees mathematical precision across large datasets, eliminating manual cross-referencing and streamlining the generation of internal cost proposals.
