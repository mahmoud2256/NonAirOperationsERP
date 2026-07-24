"""
Import page for Bright Star Non-Air Operations ERP.

Lets the user upload an .xlsx or .csv file in the external invoice
export format (Client Name, Traveller Name, Invoice date, Invoice no.,
... TOTAL, etc.) and inserts the rows into the `invoices` table,
mapped via column_mapping.COLUMN_MAPPING.

HOW TO WIRE THIS INTO YOUR APP:
1. Run migration.sql once against your Supabase database.
2. Copy column_mapping.py and this file into the same folder as your
   main app.py.
3. In app.py, add:  from import_invoices import show_import
4. Add "Import" to the `pages` list in the sidebar section.
5. Add:  elif st.session_state.page == "Import": show_import()
"""

import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st

from column_mapping import COLUMN_MAPPING, NUMERIC_COLUMNS, DATE_COLUMNS


def _clean_numeric(value):
    """Turn '7,017.54' / '' / NaN into a float (0.0 if empty/unparseable)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "")
    if s == "" or s.lower() == "nan":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _clean_date(value):
    """Convert 'dd-mm-yyyy' (or similar) to 'yyyy-mm-dd' text. Leaves
    unparseable values as-is so nothing is silently lost."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return ""
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s  # leave as-is, flagged in preview for manual check


def _map_row(row: dict) -> dict:
    """Map one source row (dict keyed by original headers) to db columns."""
    mapped = {}
    for source_col, value in row.items():
        source_col_clean = str(source_col).strip()
        db_col = COLUMN_MAPPING.get(source_col_clean)
        if not db_col:
            continue  # unmapped column — ignored (add it to COLUMN_MAPPING to include)
        if db_col in NUMERIC_COLUMNS:
            mapped[db_col] = _clean_numeric(value)
        elif db_col in DATE_COLUMNS:
            mapped[db_col] = _clean_date(value)
        else:
            mapped[db_col] = "" if (isinstance(value, float) and pd.isna(value)) else str(value).strip()
    mapped.setdefault("payment_status", "Pending")
    mapped["source"] = "import"
    return mapped


def show_import(get_cursor):
    """Render the Import tab. Pass your existing get_cursor function in."""

    st.markdown("""
    <div style="background:#1A3A5C; padding:16px 24px; border-radius:6px; margin-bottom:24px;">
        <span style="color:white; font-size:20px; font-weight:700;">📥 Bulk Import</span>
        <span style="color:#93C5FD; font-size:13px; margin-left:16px;">Import invoices from Excel/CSV export</span>
    </div>
    """, unsafe_allow_html=True)

    st.caption(
        "Upload a file with the same headers as your external export "
        "(Client Name, Traveller Name, Invoice date, Invoice no., ... TOTAL, etc.)"
    )

    uploaded = st.file_uploader("Choose file", type=["xlsx", "csv"])
    if not uploaded:
        return

    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded, dtype=str)
    else:
        df = pd.read_excel(uploaded, dtype=str)

    st.write(f"**{len(df)}** rows found in file.")

    # Show which columns matched vs were ignored
    file_columns = [str(c).strip() for c in df.columns]
    matched = [c for c in file_columns if c in COLUMN_MAPPING]
    unmatched = [c for c in file_columns if c not in COLUMN_MAPPING]

    col1, col2 = st.columns(2)
    with col1:
        st.success(f"✅ {len(matched)} columns matched")
    with col2:
        if unmatched:
            st.warning(f"⚠️ {len(unmatched)} columns not mapped (will be skipped)")
            with st.expander("Show unmapped columns"):
                st.write(unmatched)

    mapped_rows = [_map_row(row) for row in df.to_dict(orient="records")]
    preview_df = pd.DataFrame(mapped_rows)
    st.markdown("#### Preview (first 10 rows, mapped)")
    st.dataframe(preview_df.head(10), use_container_width=True, height=300)

    # Check for invoice_no duplicates already in DB
    incoming_invoice_nos = [
        r.get("invoice_no") for r in mapped_rows if r.get("invoice_no")
    ]
    duplicates = []
    if incoming_invoice_nos:
        cur = get_cursor()
        cur.execute(
            "SELECT invoice_no FROM invoices WHERE invoice_no = ANY(%s)",
            (incoming_invoice_nos,),
        )
        duplicates = [row[0] for row in cur.fetchall()]

    if duplicates:
        st.error(
            f"⚠️ {len(duplicates)} invoice number(s) already exist in the "
            f"database and will be SKIPPED: {', '.join(duplicates[:10])}"
            + (" ..." if len(duplicates) > 10 else "")
        )

    to_insert = [r for r in mapped_rows if r.get("invoice_no") not in duplicates]
    st.info(f"**{len(to_insert)}** new rows will be inserted.")

    if st.button("✅ Confirm Import", type="primary", use_container_width=True):
        if not to_insert:
            st.warning("Nothing to insert.")
            return

        cur = get_cursor()
        inserted = 0
        errors = []
        for row in to_insert:
            columns = list(row.keys())
            values = [row[c] for c in columns]
            placeholders = ", ".join(["%s"] * len(columns))
            col_names = ", ".join(columns)
            query = f"INSERT INTO invoices ({col_names}) VALUES ({placeholders})"
            try:
                cur.execute(query, values)
                inserted += 1
            except Exception as exc:
                errors.append((row.get("invoice_no", "?"), str(exc)))

        st.success(f"✅ Imported {inserted} invoice(s) successfully!")
        if errors:
            st.error(f"❌ {len(errors)} row(s) failed:")
            st.dataframe(pd.DataFrame(errors, columns=["Invoice No", "Error"]))
        st.rerun()
