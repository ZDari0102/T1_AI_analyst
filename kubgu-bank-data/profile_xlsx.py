"""Load a bank portfolio XLSX and print its validation results as JSON."""

import argparse
from collections import Counter
from datetime import date, datetime, time
from decimal import ROUND_HALF_UP, Decimal
import json
import math
from pathlib import Path

from openpyxl import load_workbook


SHEET = "positions"
SCHEMA = {
    "as_of_date": "date",
    "snapshot_kind": "text",
    "side": "text",
    "product": "text",
    "currency": "text",
    "maturity_bucket": "text",
    "balance_ccy": "number",
    "fx_rub_per_unit": "number",
    "balance_rub": "number",
    "rate_annual": "number",
}
KEY = ("as_of_date", "side", "product", "currency", "maturity_bucket")
DOMAINS = {
    "snapshot_kind": {"MONTH_END", "DAILY"},
    "side": {"ASSET", "LIABILITY"},
    "currency": {"RUB", "USD", "EUR"},
}


def cell_type(cell):
    if cell.data_type in {"e", "f"}:
        return {"e": "error", "f": "formula"}[cell.data_type]
    value = cell.value
    if value is None or isinstance(value, str) and not value.strip():
        return "blank"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, date) and (
        not isinstance(value, datetime) or value.time() == time.min
    ):
        return "date"
    if isinstance(value, (int, float)) and math.isfinite(value):
        return "number"
    if isinstance(value, str):
        return "text"
    return "unsupported"


def load_profile(source):
    report = {
        "source": str(source.resolve()),
        "sheet": SHEET,
        "records": 0,
        "structure_errors": [],
        "issue_counts": {},
        "examples": [],
    }
    workbook = load_workbook(source, read_only=True, data_only=False)
    try:
        if SHEET not in workbook.sheetnames:
            report["structure_errors"].append(f"Missing sheet: {SHEET}")
            return report
        sheet_rows = iter(workbook[SHEET].iter_rows())
        header = [cell.value for cell in next(sheet_rows, ())]
        report["columns"] = len(header)
        report["headers"] = header
        if header != list(SCHEMA):
            report["structure_errors"].append("Headers differ from the ordered schema")

        observed = {field: Counter() for field in SCHEMA}
        issues = Counter(dict.fromkeys(
            ["missing", "type_error", "duplicate_row", "duplicate_key", "domain_error"], 0
        ))
        report["types_by_column"] = observed
        report["issue_counts"] = issues
        seen_rows, seen_keys, rates_by_date_currency = set(), set(), {}

        def flag(kind, row_number, field):
            issues[kind] += 1
            if len(report["examples"]) < 20:
                report["examples"].append({"kind": kind, "row": row_number, "field": field})

        for row_number, cells in enumerate(sheet_rows, start=2):
            values = tuple(cell.value for cell in cells)
            if all(value is None for value in values):
                continue
            report["records"] += 1
            if report["structure_errors"]:
                continue
            if values in seen_rows:
                flag("duplicate_row", row_number, None)
            seen_rows.add(values)

            valid = {}
            for (field, expected), cell in zip(SCHEMA.items(), cells):
                actual = cell_type(cell)
                observed[field][actual] += 1
                if actual == "blank":
                    flag("missing", row_number, field)
                elif actual != expected:
                    flag("type_error", row_number, field)
                else:
                    valid[field] = cell.value
                    if field in DOMAINS and cell.value not in DOMAINS[field]:
                        flag("domain_error", row_number, field)

            if all(field in valid for field in KEY):
                key = tuple(valid[field] for field in KEY)
                if key in seen_keys:
                    flag("duplicate_key", row_number, None)
                seen_keys.add(key)

            if "rate_annual" in valid and not 0 <= valid["rate_annual"] <= 1:
                flag("domain_error", row_number, "rate_annual")
            if "balance_ccy" in valid and valid["balance_ccy"] < 0:
                flag("domain_error", row_number, "balance_ccy")
            if "fx_rub_per_unit" in valid:
                rate = valid["fx_rub_per_unit"]
                if rate <= 0 or valid.get("currency") == "RUB" and rate != 1:
                    flag("domain_error", row_number, "fx_rub_per_unit")
                if "as_of_date" in valid and "currency" in valid:
                    rate_key = (valid["as_of_date"], valid["currency"])
                    if rates_by_date_currency.setdefault(rate_key, rate) != rate:
                        flag("domain_error", row_number, "fx_rub_per_unit")
            if {"balance_ccy", "fx_rub_per_unit", "balance_rub"} <= valid.keys():
                amount, rate, rubles = (
                    Decimal(str(valid[field]))
                    for field in ("balance_ccy", "fx_rub_per_unit", "balance_rub")
                )
                expected = (amount * rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                if expected != rubles:
                    flag("domain_error", row_number, "balance_rub")
        if report["records"] == 0:
            report["structure_errors"].append("No data rows")
        return report
    finally:
        workbook.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("xlsx", type=Path)
    args = parser.parse_args()
    try:
        report = load_profile(args.xlsx)
    except Exception as error:
        report = {"source": str(args.xlsx.resolve()), "error": str(error), "valid": False}
    else:
        report["valid"] = not (report["structure_errors"] or any(report["issue_counts"].values()))
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    return 0 if report["valid"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
