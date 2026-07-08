# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A single-purpose Python tool that exports all users and their assigned roles from a Solace Cloud (or SAP AEM) organisation via the Solace Cloud REST API, writing the result as CSV, Excel, and/or JSON. There is exactly one script: `solace_cloud_export_script.py`. There is no package structure, no test suite, and no build step — it's a standalone CLI script plus config plumbing.

`README.md` is the canonical, detailed reference (API discovery notes, live results, SAP AEM cross-platform validation, troubleshooting). Read it before making non-trivial changes — this file only covers what's needed to be immediately productive.

## Commands

```bash
# Setup
python3 -m venv .venv && source .venv/bin/activate
python3 -m pip install -r requirements.txt
cp config.example.yaml config.yaml   # optional — edit with a real token

# Run
export SOLACE_API_TOKEN="..."                      # simplest form
python3 solace_cloud_export_script.py
python3 solace_cloud_export_script.py --profile sap_aem   # uses config.yaml profiles.sap_aem
python3 solace_cloud_export_script.py --help       # full flag list

# Sanity-check changes without hitting the live API
python3 -m py_compile solace_cloud_export_script.py
```

There are no automated tests, linter, or CI configured. When changing `load_config()`/`resolve_settings()`, verify the precedence logic manually (see "Configuration precedence" below) rather than assuming — it's easy to get the merge order backwards.

## Architecture

Everything lives in `solace_cloud_export_script.py`, structured as:

- `load_config(path)` — loads `config.yaml` if present, returns `{}` if not. The config file is always optional; the script must keep working with CLI flags / env var alone.
- `resolve_settings(args, config)` — merges CLI args, an optional named `--profile` section from `config.yaml`, the `SOLACE_API_TOKEN` env var (token only), and built-in defaults, in that precedence order (CLI > config/profile > env var > default). This is the piece most likely to need care when adding a new setting — every new setting needs a slot in `config.example.yaml`, in `resolve_settings()`, and in `parse_args()` with a `None` default (so "not passed" is distinguishable from "passed the same as the default").
- `fetch_all_users(base_url, token, role_sep)` — pagination loop against `GET /api/v2/platform/users`, stopping when `meta.pagination.nextPage` is `null`. Roles come embedded in each user object; there is no per-user role lookup.
- `write_csv()` / `write_excel()` / `write_json()` — the three output writers. Excel writes three sheets (All Users, Admins, Role Summary) via `_format_sheet()` for consistent styling.
- `main()` — wires the above together: resolve settings, fail fast if no token, fetch, write.

### Configuration model (`config.yaml` / `config.example.yaml`)

`config.yaml` is gitignored and holds real tokens; `config.example.yaml` is the committed template with the same shape. Top-level keys (`token`, `base_url`, `output_dir`, `format`, `role_separator`) are the defaults; `profiles.<name>` sections override any subset of those keys for a specific org (e.g. `profiles.sap_aem` for a SAP AEM tenant using a different token). Selecting `--profile sap_aem` merges that profile over the top-level defaults, not over CLI args — CLI args always win over both.

### Solace Cloud / SAP AEM API notes that affect the code

- `/api/v2/platform/users` is the correct endpoint; `/api/v2/iam/users` (seen in some generic API reference docs) 404s on this API and should not be used.
- SAP AEM uses the *identical* API base URL and endpoints as Solace Cloud — only the bearer token differs. Confirmed live; see README section 14.
- SAP AEM's top-level admin role is named `sap-organization-administrator`, not `administrator`. Code that filters/labels admins (`is_admin` in `fetch_all_users`) currently only checks `administrator` — this is a known gap if you're working across both platforms (see README "Findings, Gaps & Improvements").
- Regional base URLs beyond US are documented inconsistently across Solace's own sources (see README section 3) — treat any non-US `--base-url` as unverified until checked live.

## Output and data handling

`output/` is gitignored in full — it contains real exported user PII (names, emails, roles) from live runs, not fixtures. Don't assume its contents are safe sample data, and don't add code that would cause it to be committed.

## File-sharing convention

The README documents a manual convention of zipping `README.md` + `solace_cloud_export_script.py` together (`API test README and script.zip`) for sharing outside of git. That zip is not kept in the repo (gitignored via `**.zip`) — regenerate it on demand per README section 17 rather than resurrecting a stale copy.
