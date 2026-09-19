"""
core.py
------------------------------------------------------------
Core PDF generation logic for Payment Schedule Statement Generator.

Supports:
    - Excel input
    - One PDF per customer
    - Multiple units per customer
    - Branch-specific logos
    - External assets folder
    - Calibri fonts
    - EBIZ watermark
    - Excel "File Name" column
    - UNIT REF ID + random hash filename fallback
    - Customer / Unit Ref ID / Project row-wise layout
    - Installment table alignment
    - Streamlit in-memory generation
    - Desktop/local folder generation
    - Cancellation support
    - Progress callbacks
"""

import os
import sys
import re
import hashlib
import secrets
from io import BytesIO

import pandas as pd
from sqlalchemy import text

from db import get_engine

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle,
)
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image,
    HRFlowable,
    PageBreak,
)


# ============================================================
# APPLICATION / ASSETS PATH
# ============================================================

def get_application_dir():
    """
    Return the directory where the application is running.

    Development:
        directory containing core.py

    PyInstaller:
        directory containing StatementGenerator.exe

    This is important because the assets folder is kept
    OUTSIDE the executable.
    """

    if getattr(sys, "frozen", False):
        return os.path.dirname(
            os.path.abspath(sys.executable)
        )

    return os.path.dirname(
        os.path.abspath(__file__)
    )


APP_DIR = get_application_dir()

ASSETS_DIR = os.path.join(
    APP_DIR,
    "assets"
)


# ============================================================
# ASSET PATHS
# ============================================================

GHR_LOGO = os.path.join(
    ASSETS_DIR,
    "GHR.png"
)

CED_LOGO = os.path.join(
    ASSETS_DIR,
    "CED.png"
)

WATERMARK_IMAGE = os.path.join(
    ASSETS_DIR,
    "CPlus.png"
)

CALIBRI = os.path.join(
    ASSETS_DIR,
    "calibri.ttf"
)

CALIBRI_BOLD = os.path.join(
    ASSETS_DIR,
    "calibri-bold.ttf"
)

CALIBRI_ITALIC = os.path.join(
    ASSETS_DIR,
    "calibri-italic.ttf"
)

CALIBRI_BOLD_ITALIC = os.path.join(
    ASSETS_DIR,
    "calibri-bold-italic.ttf"
)


# ============================================================
# REQUIRED EXCEL COLUMNS
# ============================================================

REQUIRED_COLS = [
    "BRANCH",
    "PROJECT",
    "CUSTOMER",
    "PHONE NO",
    "Unit REF ID",
    "S NO",
    "INSTALLMENT NO",
    "INSTALLMENT AMT",
    "DUE DATE",
    "PAID AMT",
    "OUTSTANDING",
]


# ============================================================
# COLORS
# ============================================================

NAVY = colors.HexColor("#2B2A72")

LIGHT_GREY = colors.HexColor(
    "#F2F2F5"
)

DARK_TEXT = colors.HexColor(
    "#222222"
)

BORDER_GREY = colors.HexColor(
    "#CCCCCC"
)

LABEL_GREY = colors.HexColor(
    "#666666"
)


# ============================================================
# FONT REGISTRATION
# ============================================================

def register_fonts():
    """
    Register Calibri fonts if they exist.

    Falls back to Helvetica if the font files are missing.
    """

    registered = []

    font_files = [
        ("Calibri", CALIBRI),
        ("Calibri-Bold", CALIBRI_BOLD),
        ("Calibri-Italic", CALIBRI_ITALIC),
        ("Calibri-BoldItalic", CALIBRI_BOLD_ITALIC),
    ]

    for font_name, font_path in font_files:

        if os.path.exists(font_path):

            try:
                pdfmetrics.registerFont(
                    TTFont(
                        font_name,
                        font_path
                    )
                )

                registered.append(font_name)

            except Exception as exc:

                print(
                    f"Warning: could not register "
                    f"{font_name}: {exc}"
                )

    # Return fonts actually available
    if "Calibri" in registered:
        normal = "Calibri"
    else:
        normal = "Helvetica"

    if "Calibri-Bold" in registered:
        bold = "Calibri-Bold"
    else:
        bold = "Helvetica-Bold"

    if "Calibri-Italic" in registered:
        italic = "Calibri-Italic"
    else:
        italic = "Helvetica-Oblique"

    if "Calibri-BoldItalic" in registered:
        bold_italic = "Calibri-BoldItalic"
    else:
        bold_italic = "Helvetica-BoldOblique"

    return {
        "normal": normal,
        "bold": bold,
        "italic": italic,
        "bold_italic": bold_italic,
    }


FONTS = register_fonts()


FONT_NORMAL = FONTS["normal"]
FONT_BOLD = FONTS["bold"]
FONT_ITALIC = FONTS["italic"]
FONT_BOLD_ITALIC = FONTS["bold_italic"]


# ============================================================
# STYLES
# ============================================================

styles = getSampleStyleSheet()


style_company = ParagraphStyle(
    "Company",
    parent=styles["Normal"],
    fontName=FONT_BOLD,
    fontSize=15,
    textColor=NAVY,
    leading=18,
)


style_doc_title = ParagraphStyle(
    "DocTitle",
    parent=styles["Normal"],
    fontName=FONT_BOLD,
    fontSize=12,
    textColor=DARK_TEXT,
    alignment=TA_RIGHT,
    leading=15,
)


style_label = ParagraphStyle(
    "Label",
    parent=styles["Normal"],
    fontName=FONT_BOLD,
    fontSize=8.5,
    textColor=LABEL_GREY,
    leading=11,
)


style_value = ParagraphStyle(
    "Value",
    parent=styles["Normal"],
    fontName=FONT_BOLD,
    fontSize=10.5,
    textColor=DARK_TEXT,
    leading=13,
)


style_message = ParagraphStyle(
    "Message",
    parent=styles["Normal"],
    fontName=FONT_NORMAL,
    fontSize=9.5,
    textColor=DARK_TEXT,
    leading=14,
    alignment=TA_LEFT,
)


style_footer = ParagraphStyle(
    "Footer",
    parent=styles["Normal"],
    fontName=FONT_ITALIC,
    fontSize=7.5,
    textColor=LABEL_GREY,
    alignment=TA_CENTER,
    leading=10,
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean_value(value):
    """
    Convert Excel values safely to strings.
    """

    if pd.isna(value):
        return ""

    return str(value).strip()


def currency(value):
    """
    Format numeric values with commas and 2 decimals.
    """

    if value is None:
        return "0.00"

    try:

        if pd.isna(value):
            return "0.00"

        return f"{float(value):,.2f}"

    except (
        TypeError,
        ValueError,
    ):

        return str(value)


def numeric_value(value):
    """
    Safely convert an Excel value to float.
    """

    try:

        if value is None or pd.isna(value):
            return 0.0

        return float(value)

    except (
        TypeError,
        ValueError,
    ):

        return 0.0


def safe_filename(name):
    """
    Make a filename Windows-safe.
    """

    name = clean_value(name)

    # Windows does not allow these characters
    name = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        name
    )

    # Remove control characters
    name = re.sub(
        r"[\x00-\x1f]",
        "",
        name
    )

    # Collapse spaces
    name = re.sub(
        r"\s+",
        "_",
        name
    )

    name = name.strip(
        " ._"
    )

    return name or "statement"


def generate_random_hash(value=""):
    """
    Generate a 10-character deterministic-looking random hash.

    Uses a random token combined with the supplied value.
    """

    random_part = secrets.token_hex(8)

    raw = (
        str(value)
        + "_"
        + random_part
    )

    return hashlib.sha256(
        raw.encode("utf-8")
    ).hexdigest()[:10]


def get_pdf_filename(row):
    """
    Determine PDF filename — this should equal the payment_schedule
    table's FILE NAME column exactly (just sanitized for the filesystem
    and given a .pdf extension), so a customer/unit's PDF is always
    named the same thing no matter how many times it's regenerated.

    Priority:

    1. "FILE NAME" column (from payment_schedule / Excel)
    2. Unit REF ID (fallback if FILE NAME is missing)
    3. CUSTOMER (last resort)

    Note:
    Windows does not allow '/' in filenames, so '/' is
    converted to '_'.
    """

    if "FILE NAME" in row.index:

        file_name = clean_value(
            row["FILE NAME"]
        )

        if file_name:

            file_name = safe_filename(
                file_name
            )

            if not file_name.lower().endswith(
                ".pdf"
            ):
                file_name += ".pdf"

            return file_name

    unit_ref = clean_value(
        row.get(
            "Unit REF ID",
            ""
        )
    )

    if unit_ref:

        return (
            f"{safe_filename(unit_ref)}.pdf"
        )

    customer = clean_value(
        row.get(
            "CUSTOMER",
            "statement"
        )
    )

    return (
        f"{safe_filename(customer)}.pdf"
    )


# ============================================================
# LOGO SELECTION
# ============================================================

def get_logo(branch):
    """
    Select logo based on BRANCH.

    CORALS EDGE (PVT) LTD
        -> CED.png

    GLOBAL HOUSING & REAL ESTATE LTD
        -> GHR.png
    """

    branch_normalized = (
        clean_value(branch)
        .upper()
    )

    if (
        branch_normalized
        == "CORALS EDGE (PVT) LTD"
    ):

        if os.path.exists(CED_LOGO):
            return CED_LOGO

    elif (
        branch_normalized
        == "GLOBAL HOUSING & REAL ESTATE LTD"
    ):

        if os.path.exists(GHR_LOGO):
            return GHR_LOGO

    # Default
    if os.path.exists(GHR_LOGO):
        return GHR_LOGO

    if os.path.exists(CED_LOGO):
        return CED_LOGO

    return None


# ============================================================
# WATERMARK
# ============================================================

def draw_watermark(canvas, doc):
    """
    Draw EBIZ.png as a light watermark behind the PDF content.

    The image is deliberately kept subtle.
    """

    if not os.path.exists(
        WATERMARK_IMAGE
    ):
        return

    try:

        canvas.saveState()

        page_width, page_height = A4

        watermark_width = 120 * mm
        watermark_height = 120 * mm

        x = (
            page_width
            - watermark_width
        ) / 2

        y = (
            page_height
            - watermark_height
        ) / 2

        # Transparency
        try:
            canvas.setFillAlpha(0.08)
        except Exception:
            pass

        canvas.drawImage(
            WATERMARK_IMAGE,
            x,
            y,
            width=watermark_width,
            height=watermark_height,
            preserveAspectRatio=True,
            mask="auto",
        )

        canvas.restoreState()

    except Exception as exc:

        print(
            f"Warning: watermark could not be drawn: {exc}"
        )


# ============================================================
# HEADER
# ============================================================

def build_header_table(branch):
    """
    Build PDF header.

    Logo + company name + statement title.
    """

    logo_file = get_logo(
        branch
    )

    if logo_file:

        try:

            logo_img = Image(
                logo_file,
                width=18 * mm,
                height=18 * mm,
            )

        except Exception:

            logo_img = Spacer(
                18 * mm,
                18 * mm
            )

    else:

        logo_img = Spacer(
            18 * mm,
            18 * mm
        )

    top_row = Table(
        [
            [
                logo_img,

                Paragraph(
                    clean_value(branch),
                    style_company
                ),

                Paragraph(
                    "PAYMENT SCHEDULE<br/>"
                    "STATEMENT",
                    style_doc_title
                ),
            ]
        ],
        colWidths=[
            20 * mm,
            100 * mm,
            60 * mm,
        ],
    )

    top_row.setStyle(
        TableStyle(
            [
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),

                (
                    "ALIGN",
                    (2, 0),
                    (2, 0),
                    "RIGHT",
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (0, 0),
                    0,
                ),

                (
                    "LEFTPADDING",
                    (1, 0),
                    (1, 0),
                    8,
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),
            ]
        )
    )

    return top_row


# ============================================================
# CUSTOMER / UNIT / PROJECT INFORMATION
# ============================================================

def build_info_block(
    customer,
    unit_ref,
    project
):
    """
    Display:

        CUSTOMER
        UNIT REF ID
        PROJECT

    Row-wise rather than column-wise.
    """

    data = [
        [
            Paragraph(
                "CUSTOMER",
                style_label
            ),

            Paragraph(
                clean_value(customer),
                style_value
            ),
        ],

        [
            Paragraph(
                "UNIT REF ID",
                style_label
            ),

            Paragraph(
                clean_value(unit_ref),
                style_value
            ),
        ],

        [
            Paragraph(
                "PROJECT",
                style_label
            ),

            Paragraph(
                clean_value(project),
                style_value
            ),
        ],
    ]

    table = Table(
        data,
        colWidths=[
            40 * mm,
            140 * mm,
        ],
    )

    table.setStyle(
        TableStyle(
            [
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),

                (
                    "BACKGROUND",
                    (0, 0),
                    (0, -1),
                    LIGHT_GREY,
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    8,
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    8,
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    5,
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    5,
                ),

                (
                    "LINEBELOW",
                    (0, 0),
                    (-1, -1),
                    0.25,
                    colors.HexColor(
                        "#DDDDDD"
                    ),
                ),
            ]
        )
    )

    return table


# ============================================================
# MESSAGE
# ============================================================

def build_message():
    """
    Friendly reminder message.
    """

    text = (
        "Dear Sir/Madam,<br/><br/>"

        "This is a friendly reminder regarding "
        "your outstanding balance due this month. "
        "If you have already made the payment, "
        "please disregard this message."
    )

    return Paragraph(
        text,
        style_message
    )


# ============================================================
# INSTALLMENT TABLE
# ============================================================

def build_installment_table(
    group_df
):
    """
    Build installment grid.

    Caption alignment:
        S NO              CENTER
        INSTALLMENT NO    CENTER
        INSTALLMENT AMT   CENTER
        DUE DATE          CENTER
        PAID AMT          CENTER
        OUTSTANDING       CENTER

    Data row alignment:
        S NO              CENTER
        INSTALLMENT NO    CENTER
        INSTALLMENT AMT   RIGHT
        DUE DATE          CENTER
        PAID AMT          RIGHT
        OUTSTANDING       RIGHT
    """

    header = [
        "S NO",
        "INSTALLMENT NO",
        "INSTALLMENT AMT",
        "DUE DATE",
        "PAID AMT",
        "OUTSTANDING",
    ]

    rows = [
        header
    ]

    total_installment = 0.0
    total_paid = 0.0
    total_outstanding = 0.0

    for _, row in group_df.iterrows():

        due_date = ""

        if pd.notna(
            row["DUE DATE"]
        ):

            try:

                due_date = pd.to_datetime(
                    row["DUE DATE"]
                ).strftime(
                    "%Y-%m-%d"
                )

            except Exception:

                due_date = clean_value(
                    row["DUE DATE"]
                )

        installment_amt = numeric_value(
            row["INSTALLMENT AMT"]
        )

        paid_amt = numeric_value(
            row["PAID AMT"]
        )

        outstanding = numeric_value(
            row["OUTSTANDING"]
        )

        rows.append(
            [
                clean_value(
                    row["S NO"]
                ),

                clean_value(
                    row["INSTALLMENT NO"]
                ),

                currency(
                    installment_amt
                ),

                due_date,

                currency(
                    paid_amt
                ),

                currency(
                    outstanding
                ),
            ]
        )

        total_installment += (
            installment_amt
        )

        total_paid += paid_amt

        total_outstanding += (
            outstanding
        )

    # Total row
    rows.append(
        [
            "TOTAL",
            "",
            currency(
                total_installment
            ),
            "",
            currency(
                total_paid
            ),
            currency(
                total_outstanding
            ),
        ]
    )

    # --------------------------------------------------------
    # Column widths
    # --------------------------------------------------------

    # S NO deliberately trimmed
    col_widths = [
        15 * mm,   # S NO
        28 * mm,   # INSTALLMENT NO
        38 * mm,   # INSTALLMENT AMT
        28 * mm,   # DUE DATE
        35 * mm,   # PAID AMT
        36 * mm,   # OUTSTANDING
    ]

    table = Table(
        rows,
        colWidths=col_widths,
        repeatRows=1,
    )

    # Number of actual data rows
    data_row_count = len(rows) - 2

    style_commands = [
        # ----------------------------------------------------
        # Header
        # ----------------------------------------------------

        (
            "BACKGROUND",
            (0, 0),
            (-1, 0),
            NAVY,
        ),

        (
            "TEXTCOLOR",
            (0, 0),
            (-1, 0),
            colors.white,
        ),

        (
            "FONTNAME",
            (0, 0),
            (-1, 0),
            FONT_BOLD,
        ),

        (
            "FONTSIZE",
            (0, 0),
            (-1, 0),
            8,
        ),

        # All captions centered
        (
            "ALIGN",
            (0, 0),
            (-1, 0),
            "CENTER",
        ),

        # ----------------------------------------------------
        # Data rows
        # ----------------------------------------------------

        (
            "FONTNAME",
            (0, 1),
            (-1, -2),
            FONT_NORMAL,
        ),

        (
            "FONTSIZE",
            (0, 1),
            (-1, -2),
            8.5,
        ),

        # S NO center
        (
            "ALIGN",
            (0, 1),
            (0, -2),
            "CENTER",
        ),

        # INSTALLMENT NO center
        (
            "ALIGN",
            (1, 1),
            (1, -2),
            "CENTER",
        ),

        # INSTALLMENT AMT right
        (
            "ALIGN",
            (2, 1),
            (2, -2),
            "RIGHT",
        ),

        # DUE DATE center
        (
            "ALIGN",
            (3, 1),
            (3, -2),
            "CENTER",
        ),

        # PAID AMT right
        (
            "ALIGN",
            (4, 1),
            (4, -2),
            "RIGHT",
        ),

        # OUTSTANDING right
        (
            "ALIGN",
            (5, 1),
            (5, -2),
            "RIGHT",
        ),

        # ----------------------------------------------------
        # Total row
        # ----------------------------------------------------

        (
            "FONTNAME",
            (0, -1),
            (-1, -1),
            FONT_BOLD,
        ),

        (
            "FONTSIZE",
            (0, -1),
            (-1, -1),
            8.5,
        ),

        (
            "BACKGROUND",
            (0, -1),
            (-1, -1),
            LIGHT_GREY,
        ),

        (
            "ALIGN",
            (0, -1),
            (1, -1),
            "CENTER",
        ),

        (
            "ALIGN",
            (2, -1),
            (2, -1),
            "RIGHT",
        ),

        (
            "ALIGN",
            (3, -1),
            (3, -1),
            "CENTER",
        ),

        (
            "ALIGN",
            (4, -1),
            (5, -1),
            "RIGHT",
        ),

        # ----------------------------------------------------
        # Grid
        # ----------------------------------------------------

        (
            "GRID",
            (0, 0),
            (-1, -1),
            0.5,
            BORDER_GREY,
        ),

        (
            "VALIGN",
            (0, 0),
            (-1, -1),
            "MIDDLE",
        ),

        (
            "TOPPADDING",
            (0, 0),
            (-1, -1),
            5,
        ),

        (
            "BOTTOMPADDING",
            (0, 0),
            (-1, -1),
            5,
        ),

        (
            "LEFTPADDING",
            (0, 0),
            (-1, -1),
            4,
        ),

        (
            "RIGHTPADDING",
            (0, 0),
            (-1, -1),
            4,
        ),
    ]

    # Zebra striping
    for index in range(
        1,
        data_row_count + 1
    ):

        if index % 2 == 0:

            style_commands.append(
                (
                    "BACKGROUND",
                    (0, index),
                    (-1, index),
                    LIGHT_GREY,
                )
            )

    table.setStyle(
        TableStyle(
            style_commands
        )
    )

    return table


# ============================================================
# EXCEL LOADING
# ============================================================

def load_excel(excel_source):
    """
    Read Excel from:

        - file path
        - Streamlit UploadedFile
        - BytesIO
    """

    if isinstance(
        excel_source,
        (str, os.PathLike)
    ):

        df = pd.read_excel(
            excel_source
        )

    else:

        # Reset stream position if possible
        try:
            excel_source.seek(0)
        except Exception:
            pass

        df = pd.read_excel(
            excel_source
        )

    # Strip whitespace from column names
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    missing = [
        column
        for column in REQUIRED_COLS
        if column not in df.columns
    ]

    if missing:

        raise ValueError(
            "Missing required column(s): "
            + ", ".join(missing)
        )

    # Date conversion
    df["DUE DATE"] = pd.to_datetime(
        df["DUE DATE"],
        errors="coerce"
    )

    return df


# ============================================================
# BUILD PDF STORY
# ============================================================

def build_story(
    customer,
    unit_ref,
    unit_df
):
    """
    Build ReportLab story for ONE unit belonging to one customer.

    unit_df contains only the installment rows for this single
    Unit REF ID (one PDF is generated per unit, not per customer).
    """

    branch = clean_value(
        unit_df["BRANCH"].iloc[0]
    )

    project = clean_value(
        unit_df["PROJECT"].iloc[0]
    )

    story = []

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

    story.append(
        build_header_table(
            branch
        )
    )

    story.append(
        Spacer(
            1,
            6
        )
    )

    story.append(
        HRFlowable(
            width="100%",
            thickness=1.2,
            color=NAVY
        )
    )

    story.append(
        Spacer(
            1,
            8
        )
    )

    # --------------------------------------------------------
    # Reminder message
    # --------------------------------------------------------

    story.append(
        build_message()
    )

    story.append(
        Spacer(
            1,
            10
        )
    )

    # --------------------------------------------------------
    # Customer / Unit / Project block
    # --------------------------------------------------------

    story.append(
        build_info_block(
            customer,
            unit_ref,
            project
        )
    )

    # Space between PROJECT and grid
    story.append(
        Spacer(
            1,
            14
        )
    )

    # Installment grid
    story.append(
        build_installment_table(
            unit_df
        )
    )

    story.append(
        Spacer(
            1,
            18
        )
    )

    # --------------------------------------------------------
    # Footer
    # --------------------------------------------------------

    story.append(
        Spacer(
            1,
            8
        )
    )

    story.append(
        HRFlowable(
            width="100%",
            thickness=0.5,
            color=colors.grey
        )
    )

    story.append(
        Spacer(
            1,
            4
        )
    )

    story.append(
        Paragraph(
            "This is a system-generated payment "
            "schedule statement. For any discrepancies, "
            "please contact the branch office.",
            style_footer
        )
    )

    return story


# ============================================================
# GENERATE SINGLE PDF TO BYTES
# ============================================================

def generate_unit_pdf(
    customer,
    unit_ref,
    unit_df
):
    """
    Generate one unit's PDF (one customer may have several of these,
    one per Unit REF ID) and return bytes.
    """

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,

        topMargin=15 * mm,
        bottomMargin=15 * mm,
        leftMargin=15 * mm,
        rightMargin=15 * mm,

        title=(
            f"Payment Schedule - "
            f"{customer} - {unit_ref}"
        ),
    )

    story = build_story(
        customer,
        unit_ref,
        unit_df
    )

    doc.build(
        story,
        onFirstPage=draw_watermark,
        onLaterPages=draw_watermark,
    )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# GENERATE PDF FILES
# ============================================================

def generate_pdfs(
    excel_path,
    output_dir,
    log_fn=None,
    stop_flag=None
):
    """
    Generate PDFs directly to a folder.

    Used by desktop applications.

    Returns:
        list of generated file paths
    """

    os.makedirs(
        output_dir,
        exist_ok=True
    )

    df = load_excel(
        excel_path
    )

    generated_files = []

    unit_groups = list(
        df.groupby(
            ["CUSTOMER", "Unit REF ID"],
            sort=False
        )
    )

    total = len(
        unit_groups
    )

    for index, (
        (customer, unit_ref),
        unit_df
    ) in enumerate(
        unit_groups,
        start=1
    ):

        # Cancellation
        if (
            stop_flag is not None
            and stop_flag.is_set()
        ):
            break

        customer = clean_value(
            customer
        )

        # Generate bytes
        pdf_bytes = generate_unit_pdf(
            customer,
            unit_ref,
            unit_df
        )

        # Determine filename from first row
        first_row = unit_df.iloc[0]

        filename = get_pdf_filename(
            first_row
        )

        output_path = os.path.join(
            output_dir,
            filename
        )

        # Avoid accidental overwrite
        output_path = make_unique_path(
            output_path
        )

        with open(
            output_path,
            "wb"
        ) as file:

            file.write(
                pdf_bytes
            )

        generated_files.append(
            output_path
        )

        if log_fn:

            log_fn(
                f"[{index}/{total}] "
                f"Generated: {filename}"
            )

    return generated_files


# ============================================================
# GENERATE PDFs IN MEMORY
# ============================================================

def generate_pdfs_in_memory_from_df(
    df,
    progress_fn=None
):
    """
    Generate PDFs in memory — one PDF per unit (Unit REF ID).

    Works with a DataFrame from either load_excel() or load_from_db(),
    since both return the same column shape.

    Returns a list of dicts, one per generated PDF:

        [
            {
                "filename": "...",
                "pdf_bytes": b"...",
                "customer": "...",
                "unit_ref_id": "...",
                "project": "...",
                "branch": "...",
                "phone_no": "...",
            },
            ...
        ]
    """

    results = []

    unit_groups = list(
        df.groupby(
            ["CUSTOMER", "Unit REF ID"],
            sort=False
        )
    )

    total = len(
        unit_groups
    )

    for index, (
        (customer, unit_ref),
        unit_df
    ) in enumerate(
        unit_groups,
        start=1
    ):

        customer = clean_value(
            customer
        )

        unit_ref = clean_value(
            unit_ref
        )

        pdf_bytes = generate_unit_pdf(
            customer,
            unit_ref,
            unit_df
        )

        first_row = unit_df.iloc[0]

        filename = get_pdf_filename(
            first_row
        )

        # Ensure names inside a single ZIP/batch are unique
        existing_names = {
            r["filename"]
            for r in results
        }

        filename = make_unique_filename(
            filename,
            existing_names
        )

        results.append(
            {
                "filename": filename,
                "pdf_bytes": pdf_bytes,
                "customer": customer,
                "unit_ref_id": unit_ref,
                "project": clean_value(first_row.get("PROJECT", "")),
                "branch": clean_value(first_row.get("BRANCH", "")),
                "phone_no": clean_value(first_row.get("PHONE NO", "")),
            }
        )

        if progress_fn:

            progress_fn(
                index,
                total,
                filename
            )

    return results


def generate_pdfs_in_memory(
    excel_source,
    progress_fn=None
):
    """
    Backward-compatible wrapper: load an Excel source, then generate.

    Used by Streamlit's "Excel upload" flow.

    Returns the same list-of-dicts shape as generate_pdfs_in_memory_from_df.
    """

    df = load_excel(
        excel_source
    )

    return generate_pdfs_in_memory_from_df(
        df,
        progress_fn=progress_fn
    )


# ============================================================
# DATABASE LOADING (replaces Excel as the data source)
# ============================================================

def load_from_db(engine=None, query=None, params=None):
    """
    Pull payment schedule rows from Postgres into the same
    shape load_excel() produces.

    engine: SQLAlchemy engine. If None, db.get_engine() is used.
    query:  optional custom SQL (e.g. filtered by month/branch).
    params: optional dict of query parameters.
    """

    if engine is None:
        engine = get_engine()

    if query is None:
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
            ORDER BY customer, unit_ref_id, s_no
        """

    df = pd.read_sql(
        query,
        engine,
        params=params
    )

    missing = [
        c
        for c in REQUIRED_COLS
        if c not in df.columns
    ]

    if missing:
        raise ValueError(
            "Missing required column(s) from DB query: "
            + ", ".join(missing)
        )

    df["DUE DATE"] = pd.to_datetime(
        df["DUE DATE"],
        errors="coerce"
    )

    return df


# ============================================================
# SAVE GENERATED PDFs TO DB (for the notification service to pick up)
# ============================================================

def save_documents_to_db(results, month_folder, engine=None, document_type="payment_schedule"):
    """
    Insert/update generated_documents rows so the separate
    Notification Service app can find these PDFs without any
    shared folder or filename matching.

    results: the list of dicts returned by generate_pdfs_in_memory_from_df.
    month_folder: e.g. "2026-08" — also used as the Google Drive subfolder.
    document_type: lets the same table hold other document kinds later
        (e.g. "welcome_letter", "offer_letter") without a new table —
        this app currently only produces "payment_schedule" documents.

    Re-running for the same (customer, unit_ref_id, month_folder, document_type)
    updates the existing row and clears any previous drive_file_id/uploaded_at,
    so a re-generated PDF gets re-uploaded on the next notification run.
    """

    if engine is None:
        engine = get_engine()

    with engine.begin() as conn:

        for r in results:

            conn.execute(
                text(
                    """
                    INSERT INTO generated_documents
                        (branch, project, customer, unit_ref_id, phone_no,
                         document_type, pdf_filename, pdf_data, month_folder)
                    VALUES
                        (:branch, :project, :customer, :unit_ref_id, :phone_no,
                         :document_type, :pdf_filename, :pdf_data, :month_folder)
                    ON CONFLICT (customer, unit_ref_id, month_folder, document_type)
                    DO UPDATE SET
                        branch = EXCLUDED.branch,
                        project = EXCLUDED.project,
                        phone_no = EXCLUDED.phone_no,
                        pdf_filename = EXCLUDED.pdf_filename,
                        pdf_data = EXCLUDED.pdf_data,
                        generated_at = now(),
                        drive_file_id = NULL,
                        drive_web_url = NULL,
                        uploaded_at = NULL
                    """
                ),
                {
                    "branch": r["branch"],
                    "project": r["project"],
                    "customer": r["customer"],
                    "unit_ref_id": r["unit_ref_id"],
                    "phone_no": r["phone_no"],
                    "document_type": document_type,
                    "pdf_filename": r["filename"],
                    "pdf_data": r["pdf_bytes"],
                    "month_folder": month_folder,
                },
            )


def find_existing_document_keys(month_folder, document_type, engine=None):
    """
    Return a set of (customer, unit_ref_id) tuples that already have a
    generated_documents row for this month_folder + document_type.

    Used to avoid re-saving (and therefore re-uploading/re-notifying)
    units that were already published for this exact month and type —
    the main lever for controlling duplicate PDF generation.
    """

    if engine is None:
        engine = get_engine()

    query = text(
        """
        SELECT customer, unit_ref_id
        FROM generated_documents
        WHERE month_folder = :month_folder
          AND document_type = :document_type
        """
    )

    with engine.connect() as conn:
        rows = conn.execute(
            query,
            {"month_folder": month_folder, "document_type": document_type},
        ).fetchall()

    return {(r[0], r[1]) for r in rows}


def mark_documents_uploaded(uploads, engine=None):
    """
    Record Google Drive upload results on generated_documents rows,
    right after generation — so the Notification Service later finds
    them already uploaded and just sends the SMS.

    uploads: list of dicts, one per uploaded PDF:
        {
            "customer": ..., "unit_ref_id": ..., "month_folder": ...,
            "document_type": ..., "drive_file_id": ..., "drive_web_url": ...,
        }
    """

    if engine is None:
        engine = get_engine()

    with engine.begin() as conn:

        for u in uploads:

            conn.execute(
                text(
                    """
                    UPDATE generated_documents
                    SET drive_file_id = :drive_file_id,
                        drive_web_url = :drive_web_url,
                        uploaded_at = now()
                    WHERE customer = :customer
                      AND unit_ref_id = :unit_ref_id
                      AND month_folder = :month_folder
                      AND document_type = :document_type
                    """
                ),
                {
                    "drive_file_id": u["drive_file_id"],
                    "drive_web_url": u["drive_web_url"],
                    "customer": u["customer"],
                    "unit_ref_id": u["unit_ref_id"],
                    "month_folder": u["month_folder"],
                    "document_type": u["document_type"],
                },
            )


# ============================================================
# UNIQUE FILE NAME
# ============================================================

def make_unique_filename(
    filename,
    existing_names
):
    """
    Prevent duplicate filenames in Streamlit ZIP.
    """

    if filename not in existing_names:
        return filename

    base, ext = os.path.splitext(
        filename
    )

    counter = 2

    while True:

        candidate = (
            f"{base}_{counter}{ext}"
        )

        if candidate not in existing_names:
            return candidate

        counter += 1


def make_unique_path(
    path
):
    """
    Prevent overwriting an existing PDF.
    """

    if not os.path.exists(path):
        return path

    directory = os.path.dirname(
        path
    )

    filename = os.path.basename(
        path
    )

    base, ext = os.path.splitext(
        filename
    )

    counter = 2

    while True:

        candidate = os.path.join(
            directory,
            f"{base}_{counter}{ext}"
        )

        if not os.path.exists(
            candidate
        ):
            return candidate

        counter += 1