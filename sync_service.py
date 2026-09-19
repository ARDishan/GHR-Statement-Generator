"""
sync_service.py
------------------------------------------------------------
Pulls the current payment_schedule rows from the MAIN DASHBOARD's
live database (read-only) and refreshes the mirrored copy in
Supabase.

This is the ONLY code in either app that connects to the live
dashboard DB. Everything else — PDF generation, generated_documents,
notification_log, status tracking — reads and writes Supabase only,
via db.get_engine(). The main dashboard DB is never written to here,
only SELECTed from.

Strategy: full refresh (TRUNCATE + re-insert) on every sync, not an
incremental upsert. This is deliberately simple, and safe here because
generated_documents / notification_log reference customer/unit_ref_id
(text), not payment_schedule's row id — so replacing the mirror table
doesn't break any foreign keys or history.
"""

import pandas as pd
from sqlalchemy import text

from db import get_engine, get_master_engine


def fetch_master_payment_schedule(master_engine=None):
    """
    Read from the MAIN DASHBOARD's live database.
    """

    if master_engine is None:
        master_engine = get_master_engine()

    query = """
        SELECT
            branch          AS "BRANCH",
            project         AS "PROJECT",
            customer        AS "CUSTOMER",
            phone_no        AS "PHONE NO",
            unit_ref_id     AS "Unit REF ID",
            s_no            AS "S NO",
            installment_no  AS "INSTALLMENT NO",
            installment_amt AS "INSTALLMENT AMT",
            due_date        AS "DUE DATE",
            paid_amt        AS "PAID AMT",
            outstanding     AS "OUTSTANDING",
            file_name       AS "FILE NAME"
        FROM payment_schedule
    """

    return pd.read_sql(query, master_engine)


def sync_payment_schedule(master_engine=None, supabase_engine=None):
    """
    Pull from the main dashboard DB and replace Supabase's mirrored
    payment_schedule table in one transaction. Returns the row count.
    """

    df = fetch_master_payment_schedule(master_engine)

    if supabase_engine is None:
        supabase_engine = get_engine()

    with supabase_engine.begin() as conn:

        conn.execute(text("TRUNCATE TABLE payment_schedule"))

        if not df.empty:

            df_db = df.rename(columns={
                "BRANCH": "branch",
                "PROJECT": "project",
                "CUSTOMER": "customer",
                "PHONE NO": "phone_no",
                "Unit REF ID": "unit_ref_id",
                "S NO": "s_no",
                "INSTALLMENT NO": "installment_no",
                "INSTALLMENT AMT": "installment_amt",
                "DUE DATE": "due_date",
                "PAID AMT": "paid_amt",
                "OUTSTANDING": "outstanding",
                "FILE NAME": "file_name",
            })

            df_db.to_sql(
                "payment_schedule",
                conn,
                if_exists="append",
                index=False,
            )

        conn.execute(
            text(
                """
                INSERT INTO sync_meta (key, last_synced_at, row_count)
                VALUES ('payment_schedule', now(), :count)
                ON CONFLICT (key) DO UPDATE SET
                    last_synced_at = EXCLUDED.last_synced_at,
                    row_count = EXCLUDED.row_count
                """
            ),
            {"count": len(df)},
        )

    return len(df)


def get_last_sync_info(engine=None):
    """
    Returns {"last_synced_at": ..., "row_count": ...} or None if
    a sync has never run yet.
    """

    if engine is None:
        engine = get_engine()

    with engine.connect() as conn:
        row = conn.execute(
            text(
                "SELECT last_synced_at, row_count FROM sync_meta "
                "WHERE key = 'payment_schedule'"
            )
        ).fetchone()

    if row is None:
        return None

    return {"last_synced_at": row[0], "row_count": row[1]}
