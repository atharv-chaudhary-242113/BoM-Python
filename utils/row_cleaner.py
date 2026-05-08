"""
Validates and cleans a single raw row dict extracted by excel_reader.
"""


def clean_row(
        raw: dict,
        sheet_name: str,
        source_row_num: int,
) -> tuple[bool, dict | None, dict | None]:
    raw_cat = raw.get("CAT NO.")
    cat_no = str(raw_cat).strip() if raw_cat is not None else ""

    raw_desc = raw.get("DESCRIPTION")
    desc = str(raw_desc).strip() if raw_desc is not None else ""

    if not cat_no and not desc:
        exc = {
            "sheet": sheet_name,
            "row": source_row_num,
            "issue": "No identity (CAT NO. and DESCRIPTION both empty)",
            "description": "Row skipped: neither CAT NO. nor DESCRIPTION has a value.",
            "raw_row": raw,
        }
        return False, None, exc

    cleaned = {
        "CATEGORY": str(raw.get("CATEGORY", "")).strip(),
        "DESCRIPTION": desc,
        "SPEC": str(raw.get("SPEC")).strip() if raw.get("SPEC") is not None else "",
        "MAKE": str(raw.get("MAKE")).strip() if raw.get("MAKE") is not None else "",
        "UNIT": str(raw.get("UNIT")).strip() if raw.get("UNIT") is not None else "",
        "CAT NO.": cat_no,
        "panel_quantities": raw.get("panel_quantities", {}),
        "_sheet": sheet_name,
        "_source_row": source_row_num,
    }

    return True, cleaned, None