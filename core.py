"""
Core PDF-generation logic for the Payment Schedule Statement Generator.
Shared by both front ends: app_tkinter.py (desktop) and app_streamlit.py (browser).
Contains no UI code.
"""

import os
import re
import sys

import pandas as pd

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, HRFlowable
)
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------------------------------------------------------------------------
# Paths (works both as a plain script and as a PyInstaller --onefile exe)
# ---------------------------------------------------------------------------
def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
ASSETS_DIR = os.path.join(APP_DIR, "assets")

NAVY = colors.HexColor("#2B2A72")
LIGHT_GREY = colors.HexColor("#F2F2F5")
DARK_TEXT = colors.HexColor("#222222")

REQUIRED_COLS = [
    "BRANCH", "PROJECT", "CUSTOMER", "Unit REF ID", "S NO", "INSTALLMENT NO",
    "INSTALLMENT AMT", "DUE DATE", "PAID AMT", "OUTSTANDING", "FILE NAME"
]


def register_fonts():

    font_files = {
        "Calibri": "calibri.ttf",
        "Calibri-Bold": "calibri-bold.ttf",
        "Calibri-Italic": "calibri-italic.ttf",
        "Calibri-BoldItalic": "calibri-bold-italic.ttf",
    }

    missing_fonts = []

    for font_name, filename in font_files.items():
        font_path = os.path.join(ASSETS_DIR, filename)

        if not os.path.exists(font_path):
            missing_fonts.append(font_path)
            continue

        pdfmetrics.registerFont(
            TTFont(font_name, font_path)
        )

    if missing_fonts:
        raise FileNotFoundError(
            "The following font file(s) are missing:\n\n"
            + "\n".join(missing_fonts)
            + "\n\nPlease make sure these files exist inside the assets folder."
        )

# Register fonts when this module loads
register_fonts()

styles = getSampleStyleSheet()
style_company = ParagraphStyle(
    "Company", parent=styles["Normal"], fontName="Calibri-Bold",
    fontSize=16, textColor=NAVY, leading=19,
)
style_doc_title = ParagraphStyle(
    "DocTitle", parent=styles["Normal"], fontName="Calibri-Bold",
    fontSize=13, textColor=DARK_TEXT, alignment=TA_RIGHT, leading=16,
)
style_label = ParagraphStyle(
    "Label", parent=styles["Normal"], fontName="Calibri-Bold",
    fontSize=9, textColor=colors.grey,
)
style_value = ParagraphStyle(
    "Value", parent=styles["Normal"], fontName="Calibri-Bold",
    fontSize=11, textColor=DARK_TEXT, leading=14,
)
style_footer = ParagraphStyle(
    "Footer", parent=styles["Normal"], fontName="Calibri-Italic",
    fontSize=8, textColor=colors.grey, alignment=TA_CENTER,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def currency(val):
    try:
        return f"{float(val):,.2f}"
    except (TypeError, ValueError):
        return str(val)


def safe_filename(name: str) -> str:
    name = str(name).strip()
    # Replace slash with hyphen
    name = name.replace("/", "-")
    name = name.replace("\\", "-")
    # Replace invalid characters
    name = re.sub(r"[^A-Za-z0-9._ -]+","_", name)
    # Replace multiple spaces
    name = re.sub(r"\s+","_",name)
    return name or "customer"


def get_logo_path(branch, assets_dir=None):
    assets_dir = assets_dir or ASSETS_DIR
    branch = str(branch).strip().upper()
    if branch == "CORALS EDGE (PVT) LTD":
        fname = "CED.png"
    elif branch == "GLOBAL HOUSING & REAL ESTATE LTD":
        fname = "GHR.png"
    else:
        fname = "GHR.png"
    path = os.path.join(assets_dir, fname)
    return path if os.path.exists(path) else None


def get_watermark_path(assets_dir=None):
    assets_dir = assets_dir or ASSETS_DIR
    path = os.path.join(assets_dir, "CPlus.png")
    return path if os.path.exists(path) else None


def _make_watermark_fn(assets_dir):
    def draw_watermark(canvas, doc):
        wm_path = get_watermark_path(assets_dir)
        if not wm_path:
            return
        canvas.saveState()
        try:
            canvas.setFillAlpha(0.08)
        except Exception:
            pass
        canvas.drawImage(
            ImageReader(wm_path),
            x=40 * mm, y=50 * mm, width=130 * mm, height=130 * mm,
            preserveAspectRatio=True, mask="auto",
        )
        canvas.restoreState()
    return draw_watermark


def build_header_table(branch, assets_dir=None):
    logo_path = get_logo_path(branch, assets_dir)
    logo_cell = Image(logo_path, width=18 * mm, height=18 * mm) if logo_path else Spacer(18 * mm, 18 * mm)

    top_row = Table(
        [[logo_cell, Paragraph(str(branch), style_company),
          Paragraph("PAYMENT SCHEDULE<br/>STATEMENT", style_doc_title)]],
        colWidths=[20 * mm, 100 * mm, 60 * mm],
    )
    top_row.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 0), (2, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (0, 0), 0),
        ("LEFTPADDING", (1, 0), (1, 0), 8),
    ]))
    return top_row


def build_info_block(customer, unit_ref, project):
    data = [
        [Paragraph("CUSTOMER", style_label), Paragraph(str(customer), style_value)],
        [Paragraph("UNIT REF ID", style_label), Paragraph(str(unit_ref), style_value)],
        [Paragraph("PROJECT", style_label), Paragraph(str(project), style_value)],
    ]
    t = Table(data, colWidths=[40 * mm, 140 * mm])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (0, -1), "Calibri-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
        ("BACKGROUND", (0, 0), (0, -1), LIGHT_GREY),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t


def build_installment_table(group_df):
    header = ["S NO", "INSTALLMENT NO", "INSTALLMENT AMT", "DUE DATE", "PAID AMT", "OUTSTANDING"]
    rows = [header]

    total_installment = 0.0
    total_paid = 0.0
    total_outstanding = 0.0

    for _, r in group_df.iterrows():
        due_date = r["DUE DATE"].strftime("%Y-%m-%d") if pd.notna(r["DUE DATE"]) else ""
        rows.append([
            str(r["S NO"]),
            str(r["INSTALLMENT NO"]),
            currency(r["INSTALLMENT AMT"]),
            due_date,
            currency(r["PAID AMT"]),
            currency(r["OUTSTANDING"]),
        ])
        total_installment += float(r["INSTALLMENT AMT"] or 0)
        total_paid += float(r["PAID AMT"] or 0)
        total_outstanding += float(r["OUTSTANDING"] or 0)

    rows.append(["", "TOTAL", currency(total_installment), "", currency(total_paid), currency(total_outstanding)])

    col_widths = [12 * mm, 38 * mm, 38 * mm, 25 * mm, 36 * mm, 36 * mm]
    t = Table(rows, colWidths=col_widths, repeatRows=1)

    n_data_rows = len(rows) - 2
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Calibri-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),

        ("FONTNAME", (0, 1), (-1, -2), "Calibri"),
        ("FONTSIZE", (0, 1), (-1, -2), 9),
        ("ALIGN", (0, 1), (0, -2), "CENTER"),
        ("ALIGN", (1, 1), (1, -2), "CENTER"),
        ("ALIGN", (2, 1), (2, -2), "RIGHT"),
        ("ALIGN", (3, 1), (3, -2), "CENTER"),
        ("ALIGN", (4, 1), (5, -2), "RIGHT"),

        ("FONTNAME", (0, -1), (-1, -1), "Calibri-Bold"),
        ("FONTSIZE", (0, -1), (-1, -1), 9),
        ("BACKGROUND", (0, -1), (-1, -1), LIGHT_GREY),
        ("ALIGN", (0, -1), (1, -1), "CENTER"),
        ("ALIGN", (2, -1), (2, -1), "RIGHT"),
        ("ALIGN", (3, -1), (3, -1), "CENTER"),
        ("ALIGN", (4, -1), (5, -1), "RIGHT"),

        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    for i in range(1, n_data_rows + 1):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_GREY))

    t.setStyle(TableStyle(style_cmds))
    return t


def load_and_validate(input_excel_path_or_buffer):
    """Reads the Excel file/buffer, validates columns, normalizes types."""
    df = pd.read_excel(input_excel_path_or_buffer)
    df.columns = [c.strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing expected column(s) in Excel file: {missing}")

    df["DUE DATE"] = pd.to_datetime(df["DUE DATE"], errors="coerce")
    return df


def build_pdf_bytes(customer, cust_df, assets_dir=None):
    """Builds one customer's PDF and returns it as raw bytes (no disk write)."""
    import io
    buf = io.BytesIO()
    branch = cust_df["BRANCH"].iloc[0]

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        topMargin=15 * mm, bottomMargin=15 * mm,
        leftMargin=15 * mm, rightMargin=15 * mm,
        title=f"Payment Schedule - {customer}",
    )

    story = [
        build_header_table(branch, assets_dir),
        Spacer(1, 6),
        HRFlowable(width="100%", thickness=1.2, color=NAVY),
        Spacer(1, 10),
    ]

    for unit_ref, unit_df in cust_df.groupby("Unit REF ID", sort=False):
        project = unit_df["PROJECT"].iloc[0]

        story.append(Paragraph(
            "<b>Dear Sir/Madam,</b><br/><br/>"
            "This is a friendly reminder that you have an outstanding balance due this month. "
            "We kindly request you to settle the payment on or before the due date to avoid any inconvenience.<br/><br/>"
            "If you have already made the payment, please disregard this reminder. "
            "Thank you for your prompt attention to this matter.",
            styles["BodyText"],
        ))
        story.append(Spacer(1, 12))
        story.append(build_info_block(customer, unit_ref, project))
        story.append(Spacer(1, 12))
        story.append(build_installment_table(unit_df))
        story.append(Spacer(1, 14))

    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        "This is a system-generated payment schedule statement. "
        "For any discrepancies, please contact the branch office."
        "© 2026 CEYPLUS ERP By eBizClouds",
        style_footer,
    ))

    watermark_fn = _make_watermark_fn(assets_dir)
    doc.build(story, onFirstPage=watermark_fn, onLaterPages=watermark_fn)
    return buf.getvalue()


def generate_pdfs(input_excel, output_dir, log_fn=None, stop_flag=None, assets_dir=None):
    """Disk-writing entry point used by the Tkinter app: writes one PDF per
    customer directly into output_dir and returns the list of paths written."""
    log_fn = log_fn or (lambda msg: None)
    os.makedirs(output_dir, exist_ok=True)

    log_fn(f"Reading: {input_excel}")
    df = load_and_validate(input_excel)

    generated_files = []
    customers = list(df.groupby("CUSTOMER", sort=False))
    total = len(customers)

    for idx, (customer, cust_df) in enumerate(customers, start=1):
        if stop_flag is not None and stop_flag.is_set():
            log_fn("Cancelled by user.")
            break

        file_name = cust_df["FILE NAME"].iloc[0]
        out_name = f"{safe_filename(file_name)}.pdf"
        out_path = os.path.join(output_dir, out_name)

        pdf_bytes = build_pdf_bytes(customer, cust_df, assets_dir)
        with open(out_path, "wb") as f:
            f.write(pdf_bytes)

        generated_files.append(out_path)
        log_fn(f"[{idx}/{total}] Generated: {out_name}")

    return generated_files


def generate_pdfs_in_memory(input_excel, assets_dir=None, progress_fn=None):
    """Used by the Streamlit app: returns a list of (filename, pdf_bytes) tuples
    without touching disk, so it works both locally and when deployed/hosted."""
    df = load_and_validate(input_excel)

    results = []
    customers = list(df.groupby("CUSTOMER", sort=False))
    total = len(customers)

    for idx, (customer, cust_df) in enumerate(customers, start=1):
        file_name = cust_df["FILE NAME"].iloc[0]
        out_name = f"{safe_filename(file_name)}.pdf"
        pdf_bytes = build_pdf_bytes(customer, cust_df, assets_dir)
        results.append((out_name, pdf_bytes))
        if progress_fn:
            progress_fn(idx, total, out_name)

    return results
