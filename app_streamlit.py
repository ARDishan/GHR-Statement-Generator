"""
Payment Schedule Statement Generator — Streamlit front end.
-------------------------------------------------------------
Run:
    pip install -r requirements.txt
    streamlit run app.py

This still runs entirely on your own machine — Streamlit just uses the
browser as the UI. Data comes from the shared Supabase database (kept in
sync with the main dashboard via the Sync panel below); one PDF is
generated per Unit REF ID.
"""

import io
import os
import zipfile
from datetime import datetime

import streamlit as st

import core
from db import get_engine
import sync_service

st.set_page_config(page_title="Statement Generator", page_icon="📄", layout="centered")

st.title("📄 Payment Schedule Statement Generator")
st.caption("Generate one PDF per unit, then optionally publish them for the Notification Service.")

missing_assets = [
    n for n in ("GHR.png", "CED.png", "CPlus.png")
    if not os.path.exists(os.path.join(core.ASSETS_DIR, n))
]
if missing_assets:
    st.info(
        "Optional branding assets not found in the `assets/` folder: "
        + ", ".join(missing_assets)
        + ". PDFs will still generate, just without those logos/watermark."
    )

if "gen_results" not in st.session_state:
    st.session_state["gen_results"] = None

with st.expander("🔄 Sync payment schedule from main dashboard", expanded=False):
    st.caption(
        "Pulls the latest payment_schedule rows from the main dashboard's live "
        "database (read-only) and refreshes the Supabase copy this app actually "
        "works from. The live dashboard database is never written to."
    )

    try:
        last_sync = sync_service.get_last_sync_info()
    except Exception as e:
        last_sync = None
        st.warning(f"Could not read last sync info: {e}")

    if last_sync:
        st.write(
            f"Last synced: **{last_sync['last_synced_at']}** "
            f"({last_sync['row_count']} row(s))"
        )
    else:
        st.write("Never synced yet.")

    if st.button("Sync now from main dashboard"):
        try:
            with st.spinner("Syncing from main dashboard..."):
                count = sync_service.sync_payment_schedule()
            st.success(f"Synced {count} row(s) into Supabase.")
        except Exception as e:
            st.error(f"Sync failed: {e}")

st.caption("Reads from the Supabase mirror using DB_HOST / DB_NAME / etc. from your .env file.")

col_a, col_b = st.columns(2)
with col_a:
    filter_project = st.text_input("Filter by project (optional)", value="")
with col_b:
    filter_customer = st.text_input("Filter by customer (optional)", value="")

df = None

if st.button("Fetch data from database"):
    try:
        engine = get_engine()

        query = 'SELECT branch AS "BRANCH", project AS "PROJECT", customer AS "CUSTOMER", ' \
                'phone_no AS "PHONE NO", unit_ref_id AS "Unit REF ID", s_no AS "S NO", ' \
                'installment_no AS "INSTALLMENT NO", installment_amt AS "INSTALLMENT AMT", ' \
                'due_date AS "DUE DATE", paid_amt AS "PAID AMT", outstanding AS "OUTSTANDING", ' \
                'file_name AS "FILE NAME" FROM payment_schedule WHERE 1=1'
        params = {}

        if filter_project.strip():
            query += " AND project ILIKE :project"
            params["project"] = f"%{filter_project.strip()}%"

        if filter_customer.strip():
            query += " AND customer ILIKE :customer"
            params["customer"] = f"%{filter_customer.strip()}%"

        query += " ORDER BY customer, unit_ref_id, s_no"

        fetched_df = core.load_from_db(engine, query=query, params=params)
        st.session_state["db_df"] = fetched_df
        st.success(f"Fetched {len(fetched_df)} row(s).")
    except Exception as e:
        st.error(f"Could not fetch from database: {e}")

df = st.session_state.get("db_df")

# ---------------------------------------------------------------------------
# Preview + Generate (shared for both sources)
# ---------------------------------------------------------------------------
if df is not None:

    missing_cols = [c for c in core.REQUIRED_COLS if c not in df.columns]

    st.subheader("Preview")
    st.dataframe(df.head(10), use_container_width=True)

    n_customers = df["CUSTOMER"].nunique() if "CUSTOMER" in df.columns else "?"
    n_units = df["Unit REF ID"].nunique() if "Unit REF ID" in df.columns else "?"
    st.caption(f"{len(df)} row(s) · {n_customers} customer(s) · {n_units} unit(s)")

    if missing_cols:
        st.error(f"Missing required column(s): {missing_cols}")
    else:
        if st.button("Generate PDFs", type="primary"):
            progress_bar = st.progress(0.0)
            status = st.empty()

            def on_progress(idx, total, name):
                progress_bar.progress(idx / total)
                status.write(f"[{idx}/{total}] Generated: {name}")

            try:
                results = core.generate_pdfs_in_memory_from_df(df, progress_fn=on_progress)
            except Exception as e:
                st.error(f"Generation failed: {e}")
                results = []

            st.session_state["gen_results"] = results

# ---------------------------------------------------------------------------
# Results: download + publish to DB
# ---------------------------------------------------------------------------
results = st.session_state.get("gen_results")

if results:

    st.success(f"Done — {len(results)} PDF(s) generated (one per unit).")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            zf.writestr(r["filename"], r["pdf_bytes"])
    zip_buf.seek(0)

    st.download_button(
        "⬇️ Download all as ZIP",
        data=zip_buf,
        file_name="statements.zip",
        mime="application/zip",
        type="primary",
    )

    with st.expander(f"Download individually ({len(results)} file(s))"):
        for r in results:
            st.download_button(
                f"⬇️ {r['filename']}",
                data=r["pdf_bytes"],
                file_name=r["filename"],
                mime="application/pdf",
                key=f"dl_{r['filename']}",
            )

    st.divider()
    st.subheader("Save straight to a folder on this computer")
    save_dir = st.text_input("Folder path", placeholder=r"C:\Statements\2026-08 or /Users/me/Statements")
    if st.button("Save to folder"):
        if not save_dir:
            st.warning("Enter a folder path first.")
        else:
            try:
                os.makedirs(save_dir, exist_ok=True)
                for r in results:
                    with open(os.path.join(save_dir, r["filename"]), "wb") as f:
                        f.write(r["pdf_bytes"])
                st.success(f"Saved {len(results)} PDF(s) to {save_dir}")
            except Exception as e:
                st.error(f"Could not save to that folder: {e}")

    st.divider()
    st.subheader("📤 Publish for the Notification Service")
    st.caption(
        "Saves each PDF into the shared database so the separate Notification "
        "Service app can upload it to Drive and send SMS — no manual folder needed."
    )

    col_m, col_t = st.columns(2)
    with col_m:
        month_folder = st.text_input(
            "Month folder (Drive subfolder name)",
            value=datetime.now().strftime("%Y-%m"),
        )
    with col_t:
        document_type = st.selectbox(
            "Document type",
            ["payment_schedule", "due_reminder", "welcome_letter", "offer_letter"],
        )

    skip_existing = st.checkbox(
        "Skip units already saved for this month/type (avoid duplicates)",
        value=True,
        help="If unchecked, existing rows for the same customer/unit/month/type "
             "are overwritten instead of skipped — useful if you corrected data "
             "and want to regenerate.",
    )

    if st.button("Save to database", type="primary"):
        if not month_folder.strip():
            st.warning("Enter a month folder value first.")
        else:
            month_folder_clean = month_folder.strip()

            try:
                engine = get_engine()
                to_save = results
                skipped = []

                if skip_existing:
                    existing_keys = core.find_existing_document_keys(
                        month_folder_clean, document_type, engine=engine
                    )
                    to_save = [
                        r for r in results
                        if (r["customer"], r["unit_ref_id"]) not in existing_keys
                    ]
                    skipped = [
                        r for r in results
                        if (r["customer"], r["unit_ref_id"]) in existing_keys
                    ]

                if to_save:
                    core.save_documents_to_db(
                        to_save, month_folder_clean, engine=engine, document_type=document_type
                    )

                msg = f"Saved {len(to_save)} document(s) under month '{month_folder_clean}' as '{document_type}'."
                if skipped:
                    msg += f" Skipped {len(skipped)} unit(s) already saved for this month/type."
                msg += " They're now ready for the Notification Service, which will upload each PDF to Drive when notifications are sent."
                st.success(msg)
            except Exception as e:
                st.error(f"Could not save to database: {e}")