# BoM Consolidator (bom-python)

An automated Python pipeline designed to extract, clean, and consolidate Bill of Materials (BoM) data spread across multiple Excel worksheets. 

This tool scans standard BoM templates, dynamically identifies target panels and item categories, merges duplicate components based on Catalog Numbers or Descriptions, aggregates their quantities, and exports a unified master BoM along with an exception log.

## 📂 Project Structure

```text
bom-python/
├── input/                  # Place your source .xlsx workbook here
├── output/                 # The script saves final_bom.xlsx and exceptions.xlsx here
├── logs/                   # Execution logs are saved here (run_log.log)
├── utils/                  # Core processing modules
├── config.py               # Settings for layout, headers, and output mapping
├── main.py                 # The main script to run the program
└── requirements.txt        # Required Python libraries

```

## 🚀 Getting Started (Beginner Friendly)

Follow these steps to set up and run the script on your machine.

### Step 1: Install Python

Ensure you have Python 3.12 or higher installed. You can download it from [python.org](https://www.python.org/).

### Step 2: Set Up the Environment

Open your terminal (Command Prompt/PowerShell on Windows, Terminal on Mac/Linux), navigate to the `bom-python` folder, and run these commands to install the necessary tools:

```bash
# 1. Create a virtual environment to keep things clean
python -m venv .venv

# 2. Activate the virtual environment
# On Windows:
.venv\Scripts\activate
# On Mac/Linux:
source .venv/bin/activate

# 3. Install the required libraries
pip install -r requirements.txt

```

### Step 3: Run the Script

1. Place **one** Excel file (`.xlsx`) that you want to process inside the `input/` folder.
2. Run the main script:
```bash
python main.py

```


3. Check the `output/` folder for your consolidated `final_bom.xlsx` and `exceptions.xlsx`.

---

## ⚙️ How to Customize `config.py`

If your input Excel files have a different layout, you don't need to rewrite the code. You can simply adjust the variables in `config.py` to match your spreadsheet's structure.

Open `config.py` in a text editor and modify the following sections based on your data:

### 1. Source Workbook Layout

These variables tell the script where to look for data in your input Excel sheets (using 1-based indexing, just like Excel rows/columns).

* `PANEL_ROW = 1`
* **Change this if:** The names of your panels (e.g., "Main Board", "Sub Panel A") are not in Row 1.


* `HEADER_ROW = 4`
* **Change this if:** The static column headers (like SNo, DESCRIPTION, MAKE) are located on a different row.


* `PANEL_START_COL = 7`
* **Change this if:** Your panel quantities don't start at Column G (which is the 7th letter of the alphabet). If they start at Column J, change this to `10`.


* `STATIC_COLS = ["SNo", "DESCRIPTION", "SPEC", "MAKE", "UNIT", "CAT NO."]`
* **Change this if:** Your input columns have slightly different names. **Note:** The script looks for exact matches (case-insensitive) of these names in the `HEADER_ROW`.



### 2. Output Layout

These variables control how the final `final_bom.xlsx` is formatted.

* `OUT_DATA_START = 5`
* **Change this if:** You want the actual item data to start lower down on the page (e.g., leaving more room for custom headers).


* `OUT_COL_DESC = 2`, `OUT_COL_SPEC = 3`, etc.
* **Change this if:** You want to rearrange the order of the columns in the final output file.



---

## 🛠️ Advanced Architectures & Code Quality Improvements

To push this tool from a functional script to a robust, scalable system, consider implementing the following architectural shifts:

### 1. Data Contract Enforcement via Pydantic

Currently, row data is passed through the pipeline as raw dictionaries (`list[dict]`). This lacks type safety and runtime validation. Implementing `pydantic.BaseModel` will codify the data contract, ensuring strict schema validation during the extraction phase.

```python
from pydantic import BaseModel, Field

class BoMRow(BaseModel):
    category: str = "Uncategorized"
    description: str
    spec: str = ""
    make: str = ""
    unit: str = ""
    cat_no: str
    panel_quantities: dict[str, float] = Field(default_factory=dict)

```

### 2. Trimming Dependency Bloat

Your `pyproject.toml` and `requirements.txt` mandate `pandas` and `numpy`. However, the core extraction relies entirely on `openpyxl` and standard library data structures (`collections.OrderedDict`). Unless vectorized operations are introduced, strip `pandas` and `numpy` to significantly reduce the environment footprint and initialization time.

### 3. Generator Patterns for State Management

In `utils/excel_reader.py`, `current_category` state is tracked using variables scoped outside a continuous `while True` loop, heightening cyclomatic complexity. Abstracting row extraction into a generator (`yield`) isolates the state machine logic, making the parsing pipeline strictly functional and much easier to unit test without mocking the entire `openpyxl.Worksheet` object.

```
