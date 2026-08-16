"""
Payment Schedule Statement Generator — Streamlit front end.
-------------------------------------------------------------
Run:
    pip install -r requirements.txt
    streamlit run app_streamlit.py

This opens a browser tab. It still runs entirely on your own machine
(nothing is uploaded to the internet) — Streamlit just uses the browser
as the UI instead of a native window.
"""

import io
import os
import zipfile

import pandas as pd
import streamlit as st

import core

st.set_page_config(page_title="Statement Generator", page_icon="📄", layout="centered")

st.title("📄 Payment Schedule Statement Generator")
st.caption("Upload the Excel file, generate one PDF per customer, download them.")

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

uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx", "xls"])

if uploaded is not None:
    try:
        preview_df = pd.read_excel(uploaded)
        preview_df.columns = [c.strip() for c in preview_df.columns]
        missing_cols = [c for c in core.REQUIRED_COLS if c not in preview_df.columns]

        st.subheader("Preview")
        st.dataframe(preview_df.head(10), use_container_width=True)
        st.caption(f"{len(preview_df)} row(s) · {preview_df['CUSTOMER'].nunique() if 'CUSTOMER' in preview_df.columns else '?'} customer(s)")

        if missing_cols:
            st.error(f"Missing required column(s): {missing_cols}")
        else:
            if st.button("Generate PDFs", type="primary"):
                uploaded.seek(0)
                progress_bar = st.progress(0.0)
                status = st.empty()

                def on_progress(idx, total, name):
                    progress_bar.progress(idx / total)
                    status.write(f"[{idx}/{total}] Generated: {name}")

                try:
                    results = core.generate_pdfs_in_memory(uploaded, progress_fn=on_progress)
                except Exception as e:
                    st.error(f"Generation failed: {e}")
                    results = []

                if results:
                    st.success(f"Done — {len(results)} PDF(s) generated.")

                    # Build a single ZIP for one-click download of everything
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for name, pdf_bytes in results:
                            zf.writestr(name, pdf_bytes)
                    zip_buf.seek(0)

                    st.download_button(
                        "⬇️ Download all as ZIP",
                        data=zip_buf,
                        file_name="statements.zip",
                        mime="application/zip",
                        type="primary",
                    )

                    with st.expander(f"Download individually ({len(results)} file(s))"):
                        for name, pdf_bytes in results:
                            st.download_button(
                                f"⬇️ {name}",
                                data=pdf_bytes,
                                file_name=name,
                                mime="application/pdf",
                                key=f"dl_{name}",
                            )

                    st.divider()
                    st.subheader("Or save straight to a folder on this computer")
                    st.caption(
                        "Since this app runs locally, you can also save the PDFs directly "
                        "to a folder path instead of downloading the ZIP."
                    )
                    save_dir = st.text_input("Folder path", placeholder=r"C:\Statements\2026-08 or /Users/me/Statements")
                    if st.button("Save to folder"):
                        if not save_dir:
                            st.warning("Enter a folder path first.")
                        else:
                            try:
                                os.makedirs(save_dir, exist_ok=True)
                                for name, pdf_bytes in results:
                                    with open(os.path.join(save_dir, name), "wb") as f:
                                        f.write(pdf_bytes)
                                st.success(f"Saved {len(results)} PDF(s) to {save_dir}")
                            except Exception as e:
                                st.error(f"Could not save to that folder: {e}")

    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")
else:
    st.info("Upload an Excel file to get started.")
    with st.expander("Required columns"):
        st.write(", ".join(core.REQUIRED_COLS))
