# Payment Schedule Statement Generator

Generates one PDF payment-schedule statement per unit, from data synced into a shared Supabase database, and publishes them for the separate [Notification Service](../notification-service/README.md) to pick up.

## What it does

1. **Sync** — pulls the latest `payment_schedule` rows from the main dashboard's live database (read-only) and refreshes the Supabase mirror. This is the only step that ever touches the main dashboard DB.
2. **Fetch** — reads from the Supabase mirror (with optional project/customer filters).
3. **Generate** — builds one PDF per `Unit REF ID`, with the customer's installment rows for that unit laid out in a table, plus branch logos/watermark if present.
4. **Save to database** — writes each generated PDF (as bytes) into `generated_documents` in Supabase, tagged with a month folder and document type, so the Notification Service can find it later without any shared folder or filename matching.

Drive upload and SMS sending do **not** happen here — that's the Notification Service's job, triggered at the point a notification is actually sent.

## Project structure

```
app.py             Streamlit UI — Sync panel, Fetch, Generate, Save to database
core.py            PDF generation, DB read/write, filename logic
db.py              Supabase connection (get_engine) + main dashboard connection (get_master_engine)
sync_service.py    Pulls from the main dashboard DB, refreshes the Supabase mirror
launcher.py        Desktop launcher (used when packaging as a .exe with PyInstaller)
assets/            Optional branding: GHR.png, CED.png, CPlus.png (logos/watermark)
```

## Setup

### 1. Install dependencies

```bash
pip install -r requirements-generator.txt
```

### 2. Configure `.env`

Create a `.env` file next to `app.py`:

```env
# Supabase — where this app reads/writes (payment_schedule mirror, generated_documents)
DB_HOST=your-project.supabase.co
DB_PORT=5432
DB_NAME=postgres
DB_USER=postgres
DB_PASSWORD=your_supabase_db_password
DB_SSLMODE=require

# Main dashboard's LIVE database — READ-ONLY, used only by the Sync button.
# Use a database user with only SELECT granted on payment_schedule if you can.
MASTER_DB_HOST=your-main-dashboard-db-host
MASTER_DB_PORT=5432
MASTER_DB_NAME=your_main_dashboard_db
MASTER_DB_USER=readonly_user
MASTER_DB_PASSWORD=your_readonly_password
MASTER_DB_SSLMODE=require
```

`DB_HOST` and `MASTER_DB_HOST` must be the bare hostname only — no `user:password@` prefix. `DB_SSLMODE`/`MASTER_DB_SSLMODE` can be left blank for a database that doesn't use SSL (e.g. local/on-prem Postgres).

### 3. Set up the shared database

Run `schema.sql` and (optionally) `insert_test_data.sql` against Supabase — see the [top-level README](../README.md#shared-database-setup-do-this-once).

### 4. Point `sync_service.py` at your real schema

Open `sync_service.py` and edit the column names in `fetch_master_payment_schedule()` to match your main dashboard's actual `payment_schedule` table — the placeholders there (`branch`, `project`, `customer`, ...) were written without visibility into that schema, so they almost certainly need adjusting. The `AS "BRANCH"` etc. aliases must stay exactly as-is; only the left-hand column/table names should change.

### 5. Run it

```bash
streamlit run app.py
```

Or, packaged as a desktop app:

```bash
python launcher.py
```

## Using the app

1. **Sync payment schedule from main dashboard** (expander at the top) — click "Sync now" to refresh Supabase from the live dashboard DB. Shows the last sync time and row count.
2. **Fetch data from database** — optionally filter by project/customer, then fetch. A preview table shows row/customer/unit counts.
3. **Generate PDFs** — builds one PDF per unit in memory.
4. Once generated, you can:
   - **Download all as ZIP**, or download PDFs individually
   - **Save to a folder** on this computer
   - **Save to database** — the important one for the Notification Service. Pick the month folder (e.g. `2026-08`) and document type, then save. The "Skip units already saved for this month/type" checkbox (on by default) prevents accidentally re-saving/duplicating documents you already published for that period.

## Key design notes

- **One PDF per unit, not per customer.** A customer with multiple units gets multiple PDFs, one per `Unit REF ID`, each named after that row's `FILE NAME` value exactly (sanitized for the filesystem, `.pdf` appended) — no random suffix. If `FILE NAME` is missing, it falls back to the unit ref, then the customer name.
- **`document_type`** on `generated_documents` lets the same table hold different kinds of documents (`payment_schedule`, `welcome_letter`, `offer_letter`, ...) without needing a separate table per type. This app only produces `payment_schedule`-style statements from installment data; other letter types would need their own generation step or manual upload later.
- **Full-refresh sync.** `sync_payment_schedule()` truncates and reloads the Supabase `payment_schedule` mirror on every run rather than doing an incremental upsert — simpler, and safe here since nothing has a foreign key into that table by numeric ID (everything downstream keys off `customer`/`unit_ref_id` text values).
- **Required columns** (`core.REQUIRED_COLS`): `BRANCH`, `PROJECT`, `CUSTOMER`, `PHONE NO`, `Unit REF ID`, `S NO`, `INSTALLMENT NO`, `INSTALLMENT AMT`, `DUE DATE`, `PAID AMT`, `OUTSTANDING`. `PHONE NO` is required because it flows straight through to the Notification Service.

## Troubleshooting

- **`could not translate host name "...@db...."`** — a value in `.env` got mixed up (e.g. password embedded in `DB_HOST`). Check that `DB_HOST`/`MASTER_DB_HOST` are bare hostnames and passwords are in their own `_PASSWORD` variables only. `db.py` uses `sqlalchemy.engine.URL.create()`, which safely percent-encodes special characters (`@`, `$`, `:`, `/`, `#`) in usernames/passwords, so unusual passwords are fine as long as they're in the right field.
- **Sync fails with a column-does-not-exist error** — the placeholder query in `sync_service.py` hasn't been updated to your main dashboard's real column names yet (see setup step 4).
- **Missing branding assets warning** — harmless; PDFs still generate without the branch logos/watermark. Drop `GHR.png`, `CED.png`, `CPlus.png` into `assets/` to enable them.
