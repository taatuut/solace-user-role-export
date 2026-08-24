# Solace Cloud / SAP AEM — User & Role Export
## Full Session Documentation

> **Session Date:** 2026-06-30
> **Organisation:** `<tenant>` (Solace Cloud Enterprise)
> **Region:** US (`https://api.solace.cloud`)
> **Performed by:** Emil Zegers — Senior Solutions Engineer
> **Goal:** Programmatically export all users and their assigned roles from Solace Cloud (and SAP AEM) via REST API, store results, and produce a reusable script.

---

## Table of Contents

1. [Background & Objectives](#1-background--objectives)
2. [Environment & Prerequisites](#2-environment--prerequisites)
3. [API Architecture Overview](#3-api-architecture-overview)
4. [Authentication](#4-authentication)
5. [API Discovery & Troubleshooting](#5-api-discovery--troubleshooting)
6. [Working API Endpoints](#6-working-api-endpoints)
7. [Configuration (config.yaml, profiles, CLI)](#7-configuration-configyaml-profiles-cli)
8. [Live Execution — Step by Step](#8-live-execution--step-by-step)
9. [The Export Script](#9-the-export-script)
10. [How to Run the Script](#10-how-to-run-the-script)
11. [Output Files](#11-output-files)
12. [Live Results — Solace Cloud](#12-live-results--solace-cloud)
13. [Role Reference](#13-role-reference)
14. [SAP AEM — Live Cross-Platform Validation](#14-sap-aem--live-cross-platform-validation)
15. [Findings, Gaps & Improvements](#15-findings-gaps--improvements)
16. [File Inventory](#16-file-inventory)
17. [Sharing Files Outside Git](#17-sharing-files-outside-git)
18. [References & Further Reading](#18-references--further-reading)

---

## 1. Background & Objectives

This session was conducted to answer:

> *"How do you export the list of users and assigned roles from Solace Cloud / SAP AEM?"*

The exercise included:
- Researching the correct API endpoints from official documentation
- Discovering and correcting discrepancies between documented and live endpoints
- Writing a production-quality Python export script
- Executing it live against a Solace Cloud organisation (`<tenant>`), then validating it unmodified against a second, SAP AEM organisation
- Storing all results in structured file formats (CSV, JSON, Excel)

---

## 2. Environment & Prerequisites

### Python Requirements

```bash
# Create and activate the virtual environment (once)
python3 -m venv .venv
source .venv/bin/activate        # macOS / Linux
# .venv\Scripts\activate         # Windows

# Install dependencies
python3 -m pip install -r requirements.txt

# Optionally, if needed upgrade pip
python3 -m pip install --upgrade pip

# Create local config from the template (once) — optional, see section 7
cp config.example.yaml config.yaml
```

Edit `config.yaml` with your Solace Cloud / SAP AEM API token(s) before running against a live organisation — or skip it entirely and pass `--token` / set `SOLACE_API_TOKEN` instead (see section 7).

> All subsequent commands assume the venv is active (`source .venv/bin/activate`).

| Package | Purpose |
|---------|---------|
| `requests` | HTTP calls to the Solace Cloud REST API |
| `pandas` | DataFrame manipulation and CSV/Excel export |
| `openpyxl` | Excel `.xlsx` writer with formatting support |
| `PyYAML` | Reads `config.yaml` |

### Python Version

```bash
python3 --version   # Tested with Python 3.10+
```

### Network Requirements

- Outbound HTTPS (port 443) to `api.solace.cloud` (US region)
- No VPN required for Solace Cloud SaaS; SAP AEM may vary

### API Token Requirements

- Must have **Manager-level role or higher** in the Solace Cloud organisation
- Scopes needed: ability to read user and platform data
- Token type: **JWT Bearer Token** (not a username/password)

---

## 3. API Architecture Overview

Solace Cloud and SAP AEM expose a layered REST API:

```
┌─────────────────────────────────────────────────────────┐
│             Solace Cloud / SAP AEM REST API              │
├──────────────────────┬──────────────────────────────────┤
│   v2 Platform API    │   v0 Legacy API                  │
│   (recommended)      │   (still functional)             │
├──────────────────────┼──────────────────────────────────┤
│ /api/v2/platform/    │ /api/v0/users                    │
│   users              │ /api/v0/organization/roles        │
│                      │                                  │
│ Base: api.solace.cloud (US)                             │
│ Auth: Authorization: Bearer <JWT>                        │
└─────────────────────────────────────────────────────────┘
```

### Regional Base URLs

Two Solace-authored sources documented slightly different regional domain formats. **Only the US endpoint has been confirmed live in this session** — verify the others before relying on them in production:

| Region | Base URL (docs.solace.com style) | Alternate documented form |
|--------|-----------------------------------|----------------------------|
| **US** (confirmed live) | `https://api.solace.cloud` | `https://api.solace.cloud` |
| EU | `https://api.eu.solace.cloud` | `https://api.solacecloud.eu` |
| Australia | `https://api.au.solace.cloud` | `https://api.solacecloud.com.au` |
| Asia-Pacific | `https://api.ap.solace.cloud` | — |
| Singapore | — | `https://api.solacecloud.sg` |

> **SAP AEM Note:** SAP Advanced Event Mesh is built on Solace Cloud. It uses the **exact same API base URL and endpoints** as Solace Cloud (confirmed live against a second, SAP-type org — see section 14). No separate AEM-specific user API exists for the console/platform layer.

---

## 4. Authentication

### Generating an API Token

1. Log in to **Solace Cloud Console** (or SAP AEM Cluster Manager)
2. Click your **user icon** (lower-left corner)
3. Select **Token Management**
4. Click **Generate Token**
5. Assign appropriate scopes (minimum: read access to users/platform)
6. **Copy the token immediately** — it is shown only once

> ⚠️ **Security Warning:** Tokens are long-lived JWTs. Store them securely (e.g., environment variable, secrets manager, or a gitignored `config.yaml`). Never commit tokens to git.

### Token Permission Model

- Each user may hold up to **50 API tokens**.
- Token permissions **cannot be modified after creation** — delete and recreate the token if its scopes need to change.
- A token's permissions are bounded by the creating user's own role permissions (principle of least privilege).
- If the creating user later loses a permission, any token that relied on it is **automatically invalidated**.
- Tokens are scoped to the organisation in which they were created.

### Using the Token

All API requests require this HTTP header:

```http
Authorization: Bearer <YOUR_JWT_TOKEN>
Content-Type: application/json
Accept: application/json
```

### Example — cURL

```bash
export SOLACE_API_TOKEN="eyJhbGci..."

curl -s \
  -H "Authorization: Bearer $SOLACE_API_TOKEN" \
  -H "Content-Type: application/json" \
  "https://api.solace.cloud/api/v2/platform/users?pageSize=100&pageNumber=1" \
  | python3 -m json.tool
```

### Token JWT Claims (decoded from live token)

```json
{
  "org":         "<tenant>",
  "orgType":     "ENTERPRISE",
  "sub":         "<user-id>",
  "apiTokenId":  "<token-id>",
  "iss":         "Solace Corporation",
  "iat":         "<issued-at-epoch>"
}
```

The `org` claim (`<tenant>`) is the organisation ID — **it does not need to appear in the URL path**. The `orgType` claim (`ENTERPRISE` vs `SAP`) can be used to detect which platform a token belongs to (see section 14).

---

## 5. API Discovery & Troubleshooting

During this session, several endpoints were tested before the correct one was confirmed.

### Endpoints Tried

| # | Method | URL | HTTP Status | Outcome |
|---|--------|-----|-------------|---------|
| 1 | GET | `/api/v2/iam/users` | **404** | ❌ Not available (despite appearing in some generic API reference examples) |
| 2 | GET | `/api/v2/iam/roles` | **404** | ❌ Not available |
| 3 | GET | `/api/v0/users` | **200** | ✅ Works — legacy format |
| 4 | GET | `/api/v2/platform/users` | **200** | ✅ Works — **preferred** |
| 5 | GET | `/api/v0/organization/roles` | **200** | ✅ Works — full role definitions |

> **Note:** Some generic Solace/SAP API reference material describes an `/api/v2/iam/users` family of endpoints for listing users, roles, and groups. Against this org and token, those all returned 404. The endpoint confirmed working live is `/api/v2/platform/users` (v2) — this is what the export script uses.

---

## 6. Working API Endpoints

### Primary: List All Users

```
GET https://api.solace.cloud/api/v2/platform/users
```

**Query Parameters:**

| Parameter | Type | Description | Example |
|-----------|------|-------------|---------|
| `pageSize` | integer | Users per page (max 100) | `100` |
| `pageNumber` | integer | Page number, 1-indexed | `1` |
| `email` | string | Filter by user email address (documented, not exercised live) | — |
| `name` | string | Filter by user name (documented, not exercised live) | — |
| `roles` | string | Filter by assigned roles (documented, not exercised live) | — |
| `groups` | string | Filter by group membership (documented, not exercised live) | — |

**Example Request:**

```bash
curl -s \
  -H "Authorization: Bearer $SOLACE_API_TOKEN" \
  "https://api.solace.cloud/api/v2/platform/users?pageSize=100&pageNumber=1"
```

**Response Structure:**

```json
{
  "data": [
    {
      "id": "<user-id>",
      "organizationId": "<tenant>",
      "firstName": "Jane",
      "lastName": "Doe",
      "email": "jane.doe@example.com",
      "roles": ["administrator"],
      "groups": [],
      "userAttributes": { "acceptedTrialAgreement": "true" },
      "state": "ACTIVE"
    }
  ],
  "meta": {
    "pagination": {
      "pageNumber": 1,
      "pageSize": 100,
      "count": "<total-user-count>",
      "totalPages": 2,
      "nextPage": 2
    }
  }
}
```

**Pagination Logic:**

```
nextPage == null  →  this is the last page, stop
nextPage == N     →  fetch page N next
```

### Secondary: List All Organisation Roles (with permissions)

```
GET https://api.solace.cloud/api/v0/organization/roles
```

Returns the full role catalogue including each role's associated permissions (large response ~471 KB).

```bash
curl -s \
  -H "Authorization: Bearer $SOLACE_API_TOKEN" \
  "https://api.solace.cloud/api/v0/organization/roles" \
  | python3 -m json.tool > roles_catalogue.json
```

### Legacy: List Users (v0)

```
GET https://api.solace.cloud/api/v0/users
```

Returns users but with different pagination metadata (page 0-indexed, `pages.next-page` / `pages.total-pages` structure). Less consistent — use v2 platform endpoint instead.

---

## 7. Configuration (config.yaml, profiles, CLI)

The export script can be configured three ways, in this precedence order (highest wins):

1. **CLI flags** — `--token`, `--base-url`, `--output-dir`, `--format`, `--role-separator`
2. **`config.yaml`** — optional; loaded via `--config` (default: `config.yaml`); a named section under `--profile` overrides the file's top-level defaults
3. **`SOLACE_API_TOKEN` environment variable** — token only, lowest precedence after config
4. **Built-in defaults** — `https://api.solace.cloud`, `./output`, `all`, `" | "`

`config.yaml` is entirely optional. The script runs exactly as before (CLI flags / env var only) if the file doesn't exist.

### Setting up config.yaml

```bash
cp config.example.yaml config.yaml
# edit config.yaml with your token(s)
```

`config.yaml` is gitignored and never committed.

### Structure

```yaml
token: "your-solace-cloud-api-token"
base_url: "https://api.solace.cloud"
output_dir: "./output"
format: "all"              # csv | excel | json | all
role_separator: " | "

# Optional named profiles — select with --profile <name>
profiles:
  sap_aem:
    token: "your-sap-aem-api-token"
    base_url: "https://api.solace.cloud"
    output_dir: "./output/sap-aem"
```

Named profiles are useful for exporting a second organisation (a different Solace Cloud org, or a SAP AEM tenant) without editing the file each time:

```bash
# Uses the top-level defaults
python3 solace_cloud_export_script.py

# Uses the "sap_aem" profile
python3 solace_cloud_export_script.py --profile sap_aem
```

Any key missing from a profile falls back to the top-level default for that key.

---

## 8. Live Execution — Step by Step

During development, the endpoints in section 6 were exercised in this order: probe page 1 of `/api/v2/platform/users`, fetch subsequent pages until `nextPage` is `null`, then fetch the role catalogue from `/api/v0/organization/roles`. The results were flattened, sorted by email, and exported to CSV — this is exactly what `solace_cloud_export_script.py` now automates end-to-end (see section 9).

---

## 9. The Export Script

Full file: `solace_cloud_export_script.py`

### Script Architecture

```
solace_cloud_export_script.py
│
├── parse_args()          ← CLI argument parser
│   └── --token / -t        API bearer token
│   └── --base-url / -u     Override API base URL (for other regions)
│   └── --output-dir / -o   Output directory
│   └── --format / -f       csv | excel | json | all
│   └── --role-separator    Separator between roles in CSV
│   └── --config / -c       Path to config.yaml (default: config.yaml)
│   └── --profile / -p      Named profile from config.yaml
│
├── load_config()         ← Loads config.yaml (optional; {} if missing)
├── resolve_settings()    ← Merges CLI > config.yaml[profile] > env var > default
│
├── get_headers()         ← Builds standard auth headers
│
├── _get_with_retry()     ← GET with retry + backoff (429/5xx/connection/timeout);
│   │                        fails fast (no retry) on other HTTP errors, e.g. 401/404
│
├── fetch_all_users()     ← Core pagination loop
│   └── Calls /api/v2/platform/users repeatedly (via _get_with_retry())
│   └── Stops when nextPage == null
│   └── Flattens roles list to separator-delimited string
│   └── Adds boolean flags: is_admin (administrator OR sap-organization-administrator), is_billing_admin
│   └── Returns sorted list of user dicts
│
├── write_csv()           ← UTF-8-SIG CSV (Excel-compatible BOM)
│
├── write_excel()         ← Multi-sheet .xlsx (role_sep must match fetch_all_users')
│   ├── Sheet 1: "All Users"     — complete user list
│   ├── Sheet 2: "Admins"        — filtered: is_admin users
│   └── Sheet 3: "Role Summary"  — role frequency count
│
├── write_json()          ← Structured JSON with metadata envelope
│
└── main()               ← Entry point: resolves settings, calls all above
```

See `solace_cloud_export_script.py` for the full, current source — it is not duplicated here to avoid the two drifting out of sync.

---

## 10. How to Run the Script

### Quickstart

```bash
# 1. Install dependencies
python3 -m pip install -r requirements.txt

# 2. Set your token as an environment variable
export SOLACE_API_TOKEN="eyJhbGci..."

# 3. Run with default settings (all output formats, ./output)
python3 solace_cloud_export_script.py

# 4. Verify the output that was just written
python3 verify_export.py
```

Or, using `config.yaml` instead of the env var (see section 7):

```bash
cp config.example.yaml config.yaml   # edit with your token
python3 solace_cloud_export_script.py
python3 verify_export.py
```

### Verifying Output

`verify_export.py` auto-detects the most recent `output/<yyyymmddhhMMss>/` run directory and cross-checks it: all generated files are present and non-empty, CSV/JSON/Excel row counts agree with each other, the Excel `Admins` sheet matches `is_admin == True` rows in `All Users`, and `Role Summary` has sane columns and counts. Exits `0` on success and `1` if any check fails, so it can be chained:

```bash
python3 solace_cloud_export_script.py && python3 verify_export.py
```

To check a specific run (not necessarily the latest) or a different profile's output tree:

```bash
python3 verify_export.py --dir output/20260708142530
python3 verify_export.py --output-dir ./output/sap-aem
```

### Running Tests

`tests/test_export.py` is a pytest suite covering `resolve_settings()`'s precedence logic, `fetch_all_users()`'s pagination/flattening/flag logic, the three output writers, and an end-to-end exporter → `verify_export.py` roundtrip. It mocks the HTTP layer (`requests.get`) with small synthetic fixture data — no live API, no real token, and no network access required:

```bash
python3 -m pip install -r requirements-dev.txt
pytest
```

This is a different, narrower check than `verify_export.py`: the test suite exercises the script's *logic* against fixture data ahead of a change; `verify_export.py` checks a *real run's actual output* after the fact. Run both when making non-trivial changes.

### Continuous Integration

`.github/workflows/test.yml` runs on every push and pull request against `main`: it installs `requirements-dev.txt`, runs the `py_compile` sanity check on both scripts, then runs the full `pytest` suite. It needs no secrets or live API access — the same mocked HTTP layer used locally is what runs in CI. A red check on a PR means either a real regression or an environment difference (e.g. a Python version mismatch) worth looking into before merging.

### All CLI Options

```bash
python3 solace_cloud_export_script.py \
  --token           "eyJhbGci..."          \   # or SOLACE_API_TOKEN env var, or config.yaml
  --base-url        "https://api.solace.cloud" \   # US default; change for other regions
  --output-dir      "./output"             \   # where to save files
  --format          "all"                  \   # csv | excel | json | all
  --role-separator  " | "                  \   # delimiter between roles in CSV
  --config          "config.yaml"          \   # path to config.yaml (optional)
  --profile         "sap_aem"                  # named profile from config.yaml
```

### Common Use Cases

```bash
# Export only CSV
python3 solace_cloud_export_script.py --format csv

# Export for EU region
python3 solace_cloud_export_script.py --base-url https://api.eu.solace.cloud

# Export for SAP AEM using a config.yaml profile
python3 solace_cloud_export_script.py --profile sap_aem

# Export for SAP AEM without config.yaml (env var + flags)
export SOLACE_API_TOKEN="<AEM_TOKEN>"
python3 solace_cloud_export_script.py --output-dir ./output/sap-aem

# Save to a specific output folder
python3 solace_cloud_export_script.py --output-dir ./exports/2026-06-30

# Use a comma as role separator instead of pipe
python3 solace_cloud_export_script.py --role-separator ", "
```

### Expected Console Output

```
============================================================
  Solace Cloud / SAP AEM — User & Role Export
============================================================
  Config file: config.yaml
  Base URL   : https://api.solace.cloud
  Output dir : /path/to/output/20260630140736
  Format(s)  : all
  Run at     : 2026-06-30T14:07:36+00:00
============================================================

📡 Endpoint : https://api.solace.cloud/api/v2/platform/users
   Page size : 100 users/page

   ↳ Fetching page 1 (probing…) … 100 users

   ✅ Organisation total : 142 users across 2 page(s)

   ↳ Fetching page 2 of 2 … 42 users

✅ Total users retrieved : 142

📄 CSV   → ./output/20260630140736/solace_users_roles.csv  (142 rows)
📊 Excel → ./output/20260630140736/solace_users_roles.xlsx  (142 users, 3 sheets)
🗂️  JSON  → ./output/20260630140736/solace_users_roles.json  (142 records)

✅ Export complete!
   Files written to: /path/to/output/20260630140736/
```

> Each run writes into its own `output/<yyyymmddhhMMss>/` subdirectory (UTC timestamp of the run), so successive exports never overwrite or interleave with each other. `output_dir` in `config.yaml` / `--output-dir` sets the parent directory; the script appends the timestamped subdirectory itself.

---

## 11. Output Files

### Generated by the Script (when run locally)

Each run creates its own subdirectory `output/<yyyymmddhhMMss>/` (UTC timestamp), containing:

| File | Format | Description |
|------|--------|--------------|
| `solace_users_roles.csv` | CSV | All users with roles, UTF-8-SIG encoded |
| `solace_users_roles.xlsx` | Excel | Multi-sheet: All Users, Admins, Role Summary |
| `solace_users_roles.json` | JSON | Structured export with metadata envelope |

### CSV Column Definitions

| Column | Type | Description |
|--------|------|--------------|
| `user_id` | string | Unique user identifier in Solace Cloud |
| `organization` | string | Organisation ID (e.g., `<tenant>`) |
| `first_name` | string | User's first name (may be empty) |
| `last_name` | string | User's last name (may be empty) |
| `email` | string | Primary email address / login |
| `state` | string | `ACTIVE` or `INVITED` |
| `roles` | string | Assigned roles, joined with `role_separator` (`--role-separator`; default `" | "`) |
| `role_count` | integer | Number of roles assigned |
| `groups` | string | Comma-separated group memberships |
| `is_admin` | boolean | `True` if user has `administrator` (Solace Cloud) or `sap-organization-administrator` (SAP AEM) |
| `is_billing_admin` | boolean | `True` if user has `billing-administrator` role |

---

## 12. Live Results — Solace Cloud

### Summary Statistics

| Metric | Value |
|--------|-------|
| Organisation | `<tenant>` |
| Total Users | a few hundred |
| Active Users | the large majority |
| Invited (pending) | a small number |
| Pages of Data | multiple (100 users per page) |
| Unique Roles Observed | most of the available role catalogue (see section 13) |
| Execution Date | 2026-06-30 |

### Sample Users

*(illustrative — names and emails below are made up, not real accounts)*

| Name | Email | Roles | State |
|------|-------|-------|-------|
| Been There | been.there@example.com | administrator | ACTIVE |
| Some Body | some.body@example.com | administrator, event-portal-manager, billing-administrator, micro-integration-manager, messaging-service-editor, mission-control-manager, agentic-ai-manager | ACTIVE |
| Are You | are.you@example.com | event-portal-manager, messaging-service-editor, micro-integration-manager, mission-control-manager, agentic-ai-manager | ACTIVE |
| Jazz Music | jazz.music@example.com | event-portal-manager, messaging-service-editor, micro-integration-manager, mission-control-manager, agentic-ai-manager | INVITED |

---

## 13. Role Reference

All roles observed in the `<tenant>` organisation:

| Role ID | Category | Access Level |
|---------|----------|---------------|
| `administrator` | Platform | Full admin access |
| `billing-administrator` | Platform | Billing management |
| `event-portal-manager` | Event Portal | Create/manage event designs |
| `event-portal-user` | Event Portal | View-only event portal |
| `messaging-service-editor` | Mission Control | Create/edit messaging services |
| `messaging-service-viewer` | Mission Control | View messaging services |
| `micro-integration-manager` | Micro-Integration | Create/manage integrations |
| `micro-integration-user` | Micro-Integration | View integrations |
| `mission-control-manager` | Mission Control | Full cluster manager access |
| `mission-control-viewer` | Mission Control | View-only cluster manager |
| `mission-control-user` | Mission Control | Resource-based access |
| `insights-advanced-editor` | Insights | Create advanced dashboards |
| `insights-advanced-viewer` | Insights | View advanced dashboards |
| `agentic-ai-manager` | Agentic AI | Manage agentic AI features |
| `agentic-ai-user` | Agentic AI | Use agentic AI features |

> **SAP AEM equivalent:** SAP AEM orgs use `sap-organization-administrator` in place of `administrator` as the top-level admin role. See section 14.

---

## 14. SAP AEM — Live Cross-Platform Validation

> **Session Date:** 2026-06-30 | **Platforms:** Solace Cloud (`<tenant>`) + SAP AEM (`<id>`, EU-20)

### Architecture Relationship

SAP Advanced Event Mesh (AEM) is built directly on top of Solace PubSub+ Cloud. The **console/platform user management layer is identical**:

```
SAP AEM Console
      │
      └── Solace Cloud REST API
              ├── https://api.solace.cloud  (confirmed for both `<tenant>` and the AEM org)
              └── ... (other regions — see section 3)
```

### Executive Summary

Both Solace Cloud and SAP AEM were successfully queried for their full user and role lists using **identical API calls and the same Python script**, confirming SAP AEM's console/platform layer is a direct Solace Cloud deployment with SAP-specific role extensions.

| Platform | Organisation | Users | Pages | Unique Roles | Script Changes Needed |
|----------|-------------|-------|-------|---------------|-------------------------|
| Solace Cloud | `<tenant>` | a few hundred | multiple | most of the catalogue | — (baseline) |
| SAP AEM | `<id>` | a handful | one | a smaller, SAP-specific subset | ❌ None |

### Execution Details

| Item | Solace Cloud (`<tenant>`) | SAP AEM (`<id>`) |
|------|--------------------------|----------------------------|
| Console URL | `https://solace-sso.solace.cloud/` | `https://eu20.console.pubsub.em.services.cloud.sap/` |
| API Base | `https://api.solace.cloud` | `https://api.solace.cloud` (same!) |
| Endpoint | `GET /api/v2/platform/users` | `GET /api/v2/platform/users` (same!) |
| Token `orgType` | `ENTERPRISE` | `SAP` |
| Users | multiple pages | a single page |

### Running the Script Against SAP AEM

**No script changes are needed.** Generate an API token in **SAP AEM Cluster Manager → User icon → Token Management**, then either:

```bash
# Env var + flag
export SOLACE_API_TOKEN="<AEM_TOKEN>"
python3 solace_cloud_export_script.py --output-dir ./output/sap-aem

# Or a config.yaml profile (see section 7)
python3 solace_cloud_export_script.py --profile sap_aem
```

### SAP AEM — Complete User List (Live Data)

*(illustrative — names and emails below are made up, not real accounts)*

| # | Email | Roles | State |
|---|-------|-------|-------|
| 1 | hare.krishna@example.com | sap-organization-administrator, insights-advanced-editor, micro-integration-manager | ACTIVE |
| 2 | jan.klaassen@example-sap.com | sap-organization-administrator, insights-advanced-editor, micro-integration-manager | ACTIVE |
| 3 | ken.barbie@example.com | sap-organization-administrator, insights-advanced-editor | ACTIVE |

**Notable:** `jan.klaassen@example-sap.com` illustrates a native SAP identity in a different domain than the other, Solace-style accounts — SAP AEM tenants can mix Solace and SAP identity providers in the same org.

### Key Role Difference: `sap-organization-administrator`

In Solace Cloud, the top-level admin role is `administrator`. In SAP AEM, it is replaced by `sap-organization-administrator`. The export script's `is_admin` flag already handles both (see section 9); any *other* automation you build on top of the raw `roles` column should do the same. Detection can also be done via the `orgType` JWT claim (`SAP` vs `ENTERPRISE`).

### What to Expect Differently in AEM vs Solace Cloud

| Aspect | Solace Cloud | SAP AEM |
|--------|---------------|---------|
| API Endpoints | Same | Same |
| Auth Method | Bearer JWT | Bearer JWT |
| Role Names | Same set | SAP-specific set (fewer roles observed; `sap-organization-administrator` replaces `administrator`) |
| First/last name presence | most users | a minority — mostly email-only accounts |
| User Base | Solace org users | AEM-specific user accounts (can include `@sap.com` identities) |
| Token Generation | Cloud Console → Token Mgmt | AEM Cluster Manager → Token Mgmt |

### Possible Future Script Enhancements (not yet implemented)

- **SAP role normalisation** — map `sap-organization-administrator` to a common label when merging exports across platforms.
- **Detect platform from token** — decode the JWT `orgType` claim to auto-label output as Solace Cloud vs SAP AEM.

### SAP BTP XSUAA Layer (Out of Scope This Session)

SAP AEM also has a separate user layer managed at the **SAP BTP (Business Technology Platform)** level via XSUAA (Extended Services User Account and Authentication). This layer is distinct and requires different credentials:

- **Auth:** OAuth 2.0 Client Credentials flow
- **Base:** `https://{subdomain}.authentication.{region}.hana.ondemand.com`
- **Endpoints:** `/sap/rest/authorization/v2/rolecollections`, `/Users` (SCIM)

This layer was **excluded from this session** per scope agreement. It would be needed for a complete picture of BTP-level entitlements (who has been granted access to the AEM subscription itself, as opposed to roles within it).

---

## 15. Findings, Gaps & Improvements

### ✅ What Worked Well

1. **Single API call pattern** — roles are embedded in the user object, no secondary per-user lookups needed
2. **Clean pagination** — `nextPage: null` is a reliable stop condition
3. **Script is region-agnostic** — just change `--base-url`
4. **SAP AEM is a drop-in** — identical API, no code changes required, confirmed against a live second org
5. **Mixed identity providers work** — `@sap.com` and `@solace.com` users coexist in the same AEM org

### ⚠️ Issues Found

| Issue | Impact | Mitigation |
|-------|--------|-----------|
| `/api/v2/iam/users` (and `/roles`) return 404 | Medium — documented in some generic references but not usable here | Use `/api/v2/platform/users` instead |
| Some users have no `firstName`/`lastName` | Low — email-only accounts, more common in the AEM org | Script handles gracefully (empty string) |
| Token is a long-lived JWT (no expiry shown) | Security risk | Store in a gitignored `config.yaml` or secrets manager; rotate periodically |
| No native UI export | Operational | Use this script for automated exports |
| Regional base URLs documented inconsistently across sources | Low — only US confirmed live | Verify live before using a non-US `--base-url` |

### 🔧 Potential Improvements

1. **Schedule the script** — add cron job or GitHub Actions workflow for periodic exports
2. **Delta export** — compare against a previous run to report only new/removed users or role changes
3. **Role normalisation** — create a separate mapping sheet showing role descriptions and cross-platform equivalents (`administrator` ↔ `sap-organization-administrator`)
4. **Email notifications** — send the export via email after each run
5. **Multi-org support** — the `--profile` mechanism (section 7) covers the config side; a future version could loop over all profiles in one run
6. **SAP BTP integration** — add XSUAA layer fetch to `write_excel()` as an additional sheet
7. **Token rotation warning** — decode JWT `iat` claim and warn if token is older than 30 days
8. **Slack/Teams notification** — post a summary count to a channel after each export run

---

## 16. File Inventory

```
solace-user-role-export/                     ← committed to Git
├── README.md                                  ← This document
├── CLAUDE.md                                   ← Guidance for Claude Code/Cowork sessions
├── LICENSE
├── .gitignore
├── config.example.yaml                        ← Connection config template
├── requirements.txt
├── requirements-dev.txt                        ← Adds pytest, for running tests/
├── pytest.ini
├── solace_cloud_export_script.py              ← Main reusable export script
├── verify_export.py                            ← Post-run output sanity checker
├── tests/
│   └── test_export.py                          ← pytest suite (mocks the HTTP layer)
└── .github/
    └── workflows/
        └── test.yml                            ← CI: runs pytest on push/PR

# created locally from templates, gitignored
├── config.yaml                                ← Your API token(s)
├── .venv/                                      ← Python virtual environment
└── output/                                     ← All export output (raw + generated), incl. real user PII
```

---

## 17. Sharing Files Outside Git

`README.md`, `solace_cloud_export_script.py`, and `verify_export.py` are sometimes bundled into a standalone zip for sharing outside of git (e.g., over email or chat, without requiring repo access). To recreate that bundle:

```bash
zip "API test README and script.zip" README.md solace_cloud_export_script.py verify_export.py
```

This zip is not tracked in git (`**.zip` is gitignored) and is not required to use or develop the project — it is purely a convenience export. Regenerate it on demand rather than keeping a stale copy in the repo.

---

## 18. References & Further Reading

| Resource | URL |
|----------|-----|
| Solace Cloud REST API Reference | https://api.solace.dev/cloud/reference/ |
| Solace Cloud API Authentication | https://api.solace.dev/cloud/reference/authentication |
| Solace v2 REST API Guide | https://api.solace.dev/cloud/reference/using-the-v2-rest-apis-for-pubsub-cloud |
| (Beta) Get a list of users — v2 API Reference | https://api.solace.dev/cloud/reference/getusers |
| (Beta) Get a list of roles — v2 API Reference | https://api.solace.dev/cloud/reference/getroles |
| Managing Users with the Solace Cloud REST API | https://docs.solace.com/Cloud/ght_use_rest_api_users.htm |
| Managing API Tokens | https://docs.solace.com/Cloud/ght_api_tokens.htm |
| Managing Users, Groups, Roles, and Permissions | https://docs.solace.com/Cloud/cloud-user-management.htm |
| SEMP v2 API Reference (AEM) | https://help.pubsub.em.services.cloud.sap/Admin/SEMP/SEMP-API-Ref.htm |
| SAP AEM REST API Docs | https://help.pubsub.em.services.cloud.sap/Cloud/cloud_rest_api.htm |
| SAP AEM Help Portal | https://help.pubsub.em.services.cloud.sap |
| SAP BTP XSUAA API Guide | https://help.sap.com/docs/btp/sap-business-technology-platform/accessing-administration-using-apis-of-sap-authorization-and-trust-management-service |
| Event Portal REST API | https://api.solace.dev/eventPortal/reference |
| Event Portal MCP Server | https://github.com/SolaceLabs/solace-platform-mcp |
| AsyncAPI Specification | https://www.asyncapi.com/docs/reference/specification/latest |
| Solace Developer Portal | https://solace.dev |

---

*Generated: 2026-06-30 | Session: Solace Cloud User & Role Export Live Exercise*
*Author: Emil Zegers, Senior Solutions Engineer | Organisation: `<tenant>`*
