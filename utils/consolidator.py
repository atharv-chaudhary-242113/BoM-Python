"""
Collects all cleaned rows and groups them by composite identity hashes.
Aggregates category classifications for items spanning multiple domains.
"""

from collections import OrderedDict


def _identity_key(row: dict) -> str:
    cat_no = str(row.get("CAT NO.") or "").strip()
    if cat_no:
        return f"cat::{cat_no}"
    return f"desc::{str(row.get('DESCRIPTION', '')).strip()}"


def consolidate(
    sheets_data: list[dict],
) -> tuple[list[str], list[dict]]:

    all_panels: list[str] = []
    seen_panels: set[str] = set()

    for sheet in sheets_data:
        for _col_idx, panel_name in sheet["panels"]:
            if panel_name not in seen_panels:
                all_panels.append(panel_name)
                seen_panels.add(panel_name)

    merged_store: OrderedDict[str, dict] = OrderedDict()

    for sheet in sheets_data:
        for row in sheet["data_rows"]:
            key = _identity_key(row)

            if key not in merged_store:
                merged_store[key] = {
                    "CATEGORY":         str(row.get("CATEGORY", "")).strip(),
                    "DESCRIPTION":      row["DESCRIPTION"],
                    "SPEC":             row["SPEC"],
                    "MAKE":             row["MAKE"],
                    "UNIT":             row["UNIT"],
                    "CAT NO.":          row["CAT NO."],
                    "panel_quantities": {p: None for p in all_panels},
                }
            else:
                # Append unique categories to prevent arbitrary truncation
                existing_cats = merged_store[key]["CATEGORY"].split(" | ")
                new_cat = str(row.get("CATEGORY", "")).strip()
                if new_cat and new_cat not in existing_cats and new_cat != "Uncategorized":
                    if merged_store[key]["CATEGORY"] == "Uncategorized":
                        merged_store[key]["CATEGORY"] = new_cat
                    else:
                        merged_store[key]["CATEGORY"] += f" | {new_cat}"

            merged = merged_store[key]

            for field in ("DESCRIPTION", "SPEC", "MAKE", "UNIT"):
                if not merged[field] and row.get(field):
                    merged[field] = row[field]

            for panel_name, qty in row.get("panel_quantities", {}).items():
                if qty is None or qty == "":
                    continue
                try:
                    numeric = float(qty)
                except (TypeError, ValueError):
                    continue

                current = merged["panel_quantities"].get(panel_name)
                if current is None:
                    merged["panel_quantities"][panel_name] = numeric
                else:
                    merged["panel_quantities"][panel_name] = current + numeric

    sorted_rows = sorted(
        merged_store.values(),
        key=lambda r: (
            r.get("CATEGORY", "").upper(),
            r["CAT NO."].upper() if r["CAT NO."] else r["DESCRIPTION"].upper(),
        ),
    )

    return all_panels, sorted_rows