import io
import streamlit as st
import openpyxl
from pypdf import PdfReader

# 1. Page Configuration
st.set_page_config(
    page_title="PO Price Recheck Tool",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. Clean & Subtle Custom CSS (Works smoothly in Light & Dark modes)
st.markdown(
    """
    <style>
        .app-header {
            padding: 1.1rem 1.4rem;
            border-radius: 12px;
            border: 1px solid rgba(128, 128, 128, 0.22);
            background: rgba(128, 128, 128, 0.05);
            margin-bottom: 1.2rem;
        }
        .app-title {
            font-size: 1.65rem;
            font-weight: 700;
            margin: 0;
            line-height: 1.25;
        }
        .app-subtitle {
            font-size: 0.95rem;
            opacity: 0.8;
            margin-top: 0.35rem;
            margin-bottom: 0;
        }
        .section-card {
            padding: 1rem 1.15rem;
            border-radius: 10px;
            border: 1px solid rgba(128, 128, 128, 0.2);
            margin-bottom: 0.9rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------------------------------
# EXISTING BACKEND LOGIC (Replace / connect your actual VPG & Matrix functions)
# -------------------------------------------------------------------------
def extract_po_items_from_pdf(pdf_file) -> list[dict]:
    """Plug in your existing VPG PDF extraction logic using PdfReader."""
    reader = PdfReader(io.BytesIO(pdf_file.getvalue()))
    _pages_text = [page.extract_text() or "" for page in reader.pages]

    # TODO: Replace sample rows below with your real VPG parser output
    # Each row example: {"part_no": "...", "category": "Cable1", "po_price": 12.50, "qty": 100}
    return []


def load_price_matrix_from_excel(excel_file) -> dict:
    """Plug in your existing Cable1 & Harness Excel matrix loading logic using openpyxl."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_file.getvalue()), data_only=True)
    _sheet_names = wb.sheetnames

    # TODO: Replace with your real Cable1 / Harness lookup structure
    return {}


def run_price_recheck(po_items: list[dict], matrix_data: dict, category_filter: str) -> list[dict]:
    """Plug in your existing price comparison logic."""
    results = []
    for item in po_items:
        part = item.get("part_no", "")
        cat = item.get("category", "Unknown")
        po_price = item.get("po_price")

        if category_filter != "All" and cat != category_filter:
            continue

        # TODO: Replace with your real matrix lookup & tolerance check
        expected_price = matrix_data.get(part)

        if expected_price is None:
            status = "Unmatched"
            diff = None
        elif po_price == expected_price:
            status = "Matched"
            diff = 0.0
        else:
            status = "Mismatch"
            diff = round(float(po_price) - float(expected_price), 4)

        results.append(
            {
                "Part No": part,
                "Category": cat,
                "PO Price": po_price,
                "Matrix Price": expected_price,
                "Difference": diff,
                "Status": status,
            }
        )
    return results


# -------------------------------------------------------------------------
# HEADER BANNER
# -------------------------------------------------------------------------
st.markdown(
    """
    <div class="app-header">
        <p class="app-title">🧾 PO Price Reconciliation Tool</p>
        <p class="app-subtitle">
            Automated price validation for VPG PDF Purchase Orders against Cable1 &amp; Harness Price Matrices
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------------------------------
# SIDEBAR CONTROLS & FILE UPLOADS
# -------------------------------------------------------------------------
with st.sidebar:
    st.header("📂 Input Files")
    matrix_file = st.file_uploader(
        "1. Upload Price Matrix (Excel)",
        type=["xlsx", "xlsm"],
        help="Workbook containing Cable1 and Harness pricing sheets",
    )
    po_file = st.file_uploader(
        "2. Upload VPG Purchase Order (PDF)",
        type=["pdf"],
        help="VPG-style PDF PO document",
    )

    st.divider()
    st.header("⚙️ Recheck Options")
    matrix_type = st.selectbox(
        "Matrix Profile",
        options=["Auto / All", "Cable1", "Harness"],
        index=0,
    )
    view_filter = st.radio(
        "Show Results",
        options=["All Items", "Mismatches Only", "Unmatched Only", "Matched Only"],
        index=0,
    )

# -------------------------------------------------------------------------
# MAIN WORKSPACE
# -------------------------------------------------------------------------
col_left, col_right = st.columns([1, 1], gap="medium")

with col_left:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("📊 Price Matrix Status")
    if matrix_file:
        st.success(f"Loaded: **{matrix_file.name}**")
    else:
        st.info("Waiting for Price Matrix Excel file...")
    st.markdown("</div>", unsafe_allow_html=True)

with col_right:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("📄 Purchase Order Status")
    if po_file:
        st.success(f"Loaded: **{po_file.name}**")
    else:
        st.info("Waiting for VPG PO PDF file...")
    st.markdown("</div>", unsafe_allow_html=True)

run_btn = st.button(
    "🚀 Run Price Recheck",
    type="primary",
    disabled=not (matrix_file and po_file),
    use_container_width=False,
)

if run_btn and matrix_file and po_file:
    with st.spinner("Extracting PO items and validating against Price Matrix..."):
        matrix_data = load_price_matrix_from_excel(matrix_file)
        po_items = extract_po_items_from_pdf(po_file)
        all_results = run_price_recheck(po_items, matrix_data, matrix_type)

    total_count = len(all_results)
    matched_count = sum(1 for r in all_results if r["Status"] == "Matched")
    mismatch_count = sum(1 for r in all_results if r["Status"] == "Mismatch")
    unmatched_count = sum(1 for r in all_results if r["Status"] == "Unmatched")

    st.divider()
    st.subheader("📌 Summary")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total PO Items", total_count)
    m2.metric("Matched", matched_count)
    m3.metric("Price Mismatches", mismatch_count)
    m4.metric("Unmatched Parts", unmatched_count)

    # Filter rows based on user selection
    if view_filter == "Mismatches Only":
        filtered_results = [r for r in all_results if r["Status"] == "Mismatch"]
    elif view_filter == "Unmatched Only":
        filtered_results = [r for r in all_results if r["Status"] == "Unmatched"]
    elif view_filter == "Matched Only":
        filtered_results = [r for r in all_results if r["Status"] == "Matched"]
    else:
        filtered_results = all_results

    st.subheader("📋 Reconciliation Details")
    if filtered_results:
        st.dataframe(filtered_results, use_container_width=True, hide_index=True)

        # Export Excel report using openpyxl
        out_wb = openpyxl.Workbook()
        out_ws = out_wb.active
        out_ws.title = "Price Recheck Report"
        headers = list(filtered_results[0].keys())
        out_ws.append(headers)
        for row in filtered_results:
            out_ws.append([row.get(h) for h in headers])

        excel_buffer = io.BytesIO()
        out_wb.save(excel_buffer)
        excel_buffer.seek(0)

        st.download_button(
            label="⬇️ Download Report (.xlsx)",
            data=excel_buffer,
            file_name="po_price_recheck_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("No items match the selected filter.")
