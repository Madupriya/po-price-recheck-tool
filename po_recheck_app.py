import io
import re
import streamlit as st
import openpyxl
from pypdf import PdfReader
from PIL import Image, ImageDraw, ImageFont

# -------------------------------------------------------------------------
# 1. PAGE CONFIG
# -------------------------------------------------------------------------
st.set_page_config(
    page_title="VPG PO Price Recheck Studio",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -------------------------------------------------------------------------
# 2. CLEAN WHITE-OUTPUT & MODERN UI STYLING
# -------------------------------------------------------------------------
st.markdown(
    """
    <style>
        .stApp {
            background-color: #f8fafc;
            color: #0f172a;
        }
        [data-testid="stSidebar"] {
            background-color: #ffffff;
            border-right: 1px solid #e2e8f0;
        }
        .upload-card {
            background-color: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 12px;
            padding: 16px 18px;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
            margin-bottom: 12px;
            color: #0f172a;
        }
        .result-paper {
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #000000;
            border-radius: 10px;
            padding: 18px 20px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.06);
        }
        .result-paper h3, .result-paper p, .result-paper span, .result-paper div {
            color: #000000 !important;
        }
        div[data-testid="stDataFrame"] > div {
            background-color: #ffffff !important;
            color: #000000 !important;
            border: 1px solid #000000 !important;
            border-radius: 8px;
        }
        div[data-testid="stDataFrame"] * {
            color: #000000 !important;
            background-color: #ffffff !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------------------------------
# 3. AI-STYLE HEADER BANNER IMAGE GENERATOR (NO EXTERNAL ASSET NEEDED)
# -------------------------------------------------------------------------
@st.cache_data
def generate_header_banner() -> Image.Image:
    width, height = 1400, 210
    img = Image.new("RGB", (width, height), color=(15, 23, 42))
    draw = ImageDraw.Draw(img)

    # Subtle gradient-like horizontal bands & geometric industrial/AI accents
    for x in range(0, width, 4):
        ratio = x / width
        r = int(15 + ratio * 24)
        g = int(23 + ratio * 56)
        b = int(42 + ratio * 92)
        draw.line([(x, 0), (x, height)], fill=(r, g, b), width=4)

    # Decorative circuit / harness nodes on the right
    nodes = [
        (1080, 55), (1180, 55), (1280, 95),
        (1130, 145), (1240, 155), (1330, 55),
    ]
    edges = [
        (nodes[0], nodes[1]), (nodes[1], nodes[2]),
        (nodes[0], nodes[3]), (nodes[3], nodes[4]),
        (nodes[2], nodes[4]), (nodes[1], nodes[5]),
    ]
    for p1, p2 in edges:
        draw.line([p1, p2], fill=(56, 189, 248), width=3)
    for cx, cy in nodes:
        draw.ellipse((cx - 8, cy - 8, cx + 8, cy + 8), fill=(125, 211, 252), outline=(255, 255, 255), width=2)

    # Accent vertical bar
    draw.rectangle((42, 42, 54, height - 42), fill=(56, 189, 248))

    font_title = ImageFont.load_default()
    draw.text((78, 56), "VPG PURCHASE ORDER — PRICE RECONCILIATION STUDIO", fill=(255, 255, 255), font=font_title)
    draw.text((78, 96), "Cable1 & Harness Price Matrix Validation  |  Multi-PO Batch Engine", fill=(186, 230, 253), font=font_title)
    draw.text((78, 134), "White-Sheet Black-Text Output  •  Instant Discrepancy & Excel Export", fill=(226, 232, 240), font=font_title)

    return img


# -------------------------------------------------------------------------
# 4. EXCEL MATRIX LOADER & VPG PDF PARSER (PLUG YOUR EXACT RULES HERE)
# -------------------------------------------------------------------------
def load_price_matrix(excel_file, matrix_profile: str) -> dict:
    """
    Returns dict keyed by normalized item_number:
    {
        "PART-001": {"matrix_price": 12.50, "description": "Harness Assy ...", "sheet": "Harness"}
    }
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_file.getvalue()), data_only=True)
    matrix_lookup = {}

    target_sheets = wb.sheetnames
    if matrix_profile in ("Cable1", "Harness"):
        matched = [s for s in wb.sheetnames if matrix_profile.lower() in s.lower()]
        if matched:
            target_sheets = matched

    for sheet_name in target_sheets:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            item_no = str(row[0]).strip()
            desc = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
            price_val = None
            for cell in row[2:]:
                if isinstance(cell, (int, float)):
                    price_val = float(cell)
                    break
                if isinstance(cell, str):
                    cleaned = re.sub(r"[^\d.]", "", cell)
                    if cleaned:
                        try:
                            price_val = float(cleaned)
                            break
                        except ValueError:
                            pass

            if item_no and price_val is not None:
                matrix_lookup[item_no.upper()] = {
                    "matrix_price": price_val,
                    "description": desc,
                    "sheet": sheet_name,
                }

    return matrix_lookup


def extract_vpg_po_items(pdf_file) -> list[dict]:
    """
    Extracts PO line items from a single VPG PDF file.
    Replace/extend pattern matching below with your exact VPG PDF layout rules.
    """
    reader = PdfReader(io.BytesIO(pdf_file.getvalue()))
    full_text = "\n".join(page.extract_text() or "" for page in reader.pages)

    items = []
    # Pattern example: Item Number, Description, Unit Price
    line_pattern = re.compile(
        r"([A-Z0-9][A-Z0-9\-_/]{2,24})\s+(.+?)\s+\$?\s*(\d+(?:\.\d{1,4})?)\s*$",
        re.MULTILINE,
    )
    for match in line_pattern.finditer(full_text):
        item_no = match.group(1).strip().upper()
        desc = " ".join(match.group(2).split()).strip()
        po_price = float(match.group(3))
        if item_no.lower() in {"page", "total", "subtotal", "date", "vpg"}:
            continue
        items.append(
            {
                "PO File": pdf_file.name,
                "Item Number": item_no,
                "Item Description": desc,
                "PO Price": po_price,
            }
        )

    return items


def reconcile_items(po_items: list[dict], matrix_lookup: dict) -> list[dict]:
    rows = []
    for item in po_items:
        item_no = item["Item Number"].upper()
        po_price = item["PO Price"]
        pdf_desc = item.get("Item Description", "")

        m_entry = matrix_lookup.get(item_no)
        if m_entry is not None:
            matrix_price = m_entry["matrix_price"]
            final_desc = m_entry["description"] or pdf_desc
            diff = round(po_price - matrix_price, 4)
            status = "MATCHED" if abs(diff) < 1e-6 else "MISMATCH"
        else:
            matrix_price = None
            final_desc = pdf_desc
            diff = None
            status = "NOT IN MATRIX"

        rows.append(
            {
                "PO File": item["PO File"],
                "Item Number": item_no,
                "Item Description": final_desc,
                "Matrix Price": matrix_price,
                "PO Price": po_price,
                "Price Difference": diff,
                "Status": status,
            }
        )
    return rows


# -------------------------------------------------------------------------
# 5. HEADER BANNER & SIDE
