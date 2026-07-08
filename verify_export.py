#!/usr/bin/env python3
"""
=============================================================================
Solace Cloud / SAP AEM — Export Verification Helper
=============================================================================
Description : Sanity-checks the output of a solace_cloud_export_script.py
              run. Auto-detects the most recent output/<yyyymmddhhMMss>/
              subdirectory (or a specific one via --dir) and verifies:
              - all three output files exist (csv, xlsx, json) for the
                formats present
              - row/record counts agree across CSV, JSON, and the Excel
                "All Users" sheet
              - the Excel "Admins" sheet contains exactly the rows where
                is_admin is true, and nothing else
              - the Excel "Role Summary" sheet is non-empty, has the
                expected columns, and every user_count is between 1 and
                the total user count

Usage:
    python3 verify_export.py                  # checks the latest run
    python3 verify_export.py --dir output/20260708142530
    python3 verify_export.py --output-dir ./output/sap-aem

Exit status: 0 if all checks pass, 1 otherwise (so this can be chained in
a shell script, e.g. `python3 solace_cloud_export_script.py && python3
verify_export.py`).
=============================================================================
"""

import argparse
import csv
import json
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("❌ Missing: python3 -m pip install pandas openpyxl")


CSV_NAME   = "solace_users_roles.csv"
XLSX_NAME  = "solace_users_roles.xlsx"
JSON_NAME  = "solace_users_roles.json"


def parse_args():
    parser = argparse.ArgumentParser(
        description="Verify a solace_cloud_export_script.py output set"
    )
    parser.add_argument("--output-dir", default="./output",
        help="Parent output directory to search for run subdirectories (default: ./output)")
    parser.add_argument("--dir", default=None,
        help="Verify a specific run directory instead of auto-detecting the latest")
    return parser.parse_args()


def find_latest_run_dir(output_dir: str) -> Path:
    """
    Run subdirectories are named with a 14-digit yyyymmddhhMMss timestamp,
    which sorts lexicographically in chronological order — no need to
    inspect mtimes.
    """
    root = Path(output_dir)
    if not root.exists():
        sys.exit(f"❌ Output directory not found: {root}")

    candidates = sorted(
        (p for p in root.iterdir() if p.is_dir() and p.name.isdigit() and len(p.name) == 14),
        key=lambda p: p.name,
    )
    if not candidates:
        sys.exit(f"❌ No timestamped run subdirectories (yyyymmddhhMMss) found under {root}")
    return candidates[-1]


def check(label: str, condition: bool, detail: str = "") -> bool:
    icon = "✅" if condition else "❌"
    print(f"  {icon} {label}" + (f" — {detail}" if detail and not condition else ""))
    return condition


def main():
    args = parse_args()
    run_dir = Path(args.dir) if args.dir else find_latest_run_dir(args.output_dir)

    print("=" * 60)
    print("  Solace Cloud / SAP AEM — Export Verification")
    print("=" * 60)
    print(f"  Verifying: {run_dir.resolve()}")
    print("=" * 60)

    ok = True

    csv_path  = run_dir / CSV_NAME
    xlsx_path = run_dir / XLSX_NAME
    json_path = run_dir / JSON_NAME

    present = {p.name: p.exists() for p in (csv_path, xlsx_path, json_path)}
    found_any = any(present.values())
    if not found_any:
        sys.exit(f"❌ None of {CSV_NAME}, {XLSX_NAME}, {JSON_NAME} found in {run_dir} "
                  f"— was a different --format used, or is this the wrong directory?")

    print("\n1. File presence (only formats actually generated are required)")
    for name, exists in present.items():
        print(f"  {'✅' if exists else '⏭ '} {name}" + ("" if exists else "  (not generated this run — skipping its checks)"))

    # ---- CSV ----
    csv_count = None
    if present[CSV_NAME]:
        with open(csv_path, encoding="utf-8-sig") as f:
            csv_count = sum(1 for _ in csv.DictReader(f))
        ok &= check(f"CSV has data rows ({csv_count})", csv_count > 0)

    # ---- JSON ----
    json_count = None
    if present[JSON_NAME]:
        data = json.loads(json_path.read_text(encoding="utf-8"))
        ok &= check("JSON has 'exported_at', 'total_users', 'users' keys",
                    {"exported_at", "total_users", "users"} <= data.keys())
        json_count = len(data.get("users", []))
        ok &= check(f"JSON total_users matches len(users) ({data.get('total_users')} == {json_count})",
                    data.get("total_users") == json_count)

    # ---- Excel ----
    xlsx_count = None
    if present[XLSX_NAME]:
        xl = pd.ExcelFile(xlsx_path)
        expected_sheets = {"All Users", "Admins", "Role Summary"}
        ok &= check(f"Excel has sheets {sorted(expected_sheets)}",
                    expected_sheets <= set(xl.sheet_names),
                    detail=f"found {xl.sheet_names}")

        all_users_df = xl.parse("All Users")
        xlsx_count = len(all_users_df)
        ok &= check(f"Excel 'All Users' has data rows ({xlsx_count})", xlsx_count > 0)

        if "is_admin" in all_users_df.columns:
            admins_df = xl.parse("Admins")
            expected_admin_count = int((all_users_df["is_admin"] == True).sum())  # noqa: E712
            ok &= check(
                f"Excel 'Admins' row count matches is_admin==True in 'All Users' "
                f"({len(admins_df)} == {expected_admin_count})",
                len(admins_df) == expected_admin_count,
            )
            if len(admins_df):
                ok &= check("Every row in 'Admins' actually has is_admin == True",
                            bool((admins_df["is_admin"] == True).all()))  # noqa: E712

        role_df = xl.parse("Role Summary")
        ok &= check("Excel 'Role Summary' is non-empty", len(role_df) > 0)
        ok &= check("Excel 'Role Summary' has columns ['role', 'user_count']",
                     list(role_df.columns) == ["role", "user_count"],
                     detail=f"found {list(role_df.columns)}")
        if len(role_df) and xlsx_count:
            in_range = role_df["user_count"].between(1, xlsx_count).all()
            ok &= check(f"Every Role Summary user_count is between 1 and total users ({xlsx_count})",
                        bool(in_range))

    # ---- Cross-format consistency ----
    counts = {k: v for k, v in {"CSV": csv_count, "JSON": json_count, "Excel": xlsx_count}.items() if v is not None}
    if len(counts) > 1:
        distinct = set(counts.values())
        ok &= check(f"Row/record counts agree across generated formats {counts}", len(distinct) == 1)

    print("\n" + "=" * 60)
    if ok:
        print("✅ All checks passed.")
    else:
        print("❌ One or more checks failed — see ❌ lines above.")
    print("=" * 60)

    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
