import io
import re
import streamlit as st
import openpyxl
from pypdf import PdfReader

st.set_page_config(
    page_title="VPG PO Price Recheck Studio",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Light, clean styling without breaking Streamlit's internal DOM
st.markdown(
    """
    <style>
        .main-banner {
            background: linear-gradient(90deg, #0f172a 0%, #1e3a8a 100%);
            color: #ffffff;
            padding: 20px 24px;
            border-radius: 12px;
            margin-bottom: 18px;
            border-left: 6px solid #38bdf8;
        }
        .main-banner h1 {
            color: #ffffff;
            font-size: 1.6rem;
            margin: 0 0 6px 0;
        }
        .main-banner p {
            color: #e2e8f0;
            margin: 0;
            font-size: 0.95rem;
        }
        .info-box {
            background-color: #ffffff;
            color: #111827;
            border: 1px solid #d1d5db;
            border-radius: 10px;
            padding: 14px 16px;
            margin-bottom: 12px;
        }
        .white-output-header {
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #000000;
            border-bottom: none;
            border-radius: 10px 10px 0 0;
            padding: 12px 16px;
            font-weight: 700;
            font-size: 1.05rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def load_price_matrix(excel_file, matrix_profile: str) -> dict:
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
    reader = PdfReader(io.BytesIO(pdf_file.getvalue()))
    full_text = "\n".join(page.extract_text() or "" for page in reader.pages)

    items = []
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


# Banner
st.markdown(
    """
    <div class="main-banner">
        <h1>⚡ VPG Purchase Order — Price Reconciliation Studio</h1>
        <p>Multi-PO Batch Validation against Cable1 &amp; Harness Price Matrices • Multi-Sheet Excel Export</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Sidebar
with st.sidebar:
    st.header("📁 Upload Inputs")
    matrix_file = st.file_uploader(
        "1. Price Matrix Workbook (.xlsx)",
        type=["xlsx", "xlsm"],
    )
    po_files = st.file_uploader(
        "2. Upload VPG Purchase Orders",
        type=["pdf"],
        accept_multiple_files=True,
    )

    st.divider()
    st.header("🎛️ Filter & View")
    matrix_profile = st.selectbox("Price Matrix Sheet", ["Auto (All Sheets)", "Cable1", "Harness"])
    status_filter = st.radio(
        "Display Rows",
        ["All Rows", "Mismatches Only", "Matched Only", "Not in Matrix"],
    )

# Status cards
col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="info-box">', unsafe_allow_html=True)
    st.markdown("**📊 Price Matrix Excel**")
    st.write(matrix_file.name if matrix_file else "Not uploaded yet.")
    st.markdown("</div>", unsafe_allow_html=True)

with col2:
    st.markdown('<div class="info-box">', unsafe_allow_html=True)
    st.markdown("**📄 VPG Purchase Orders (PDF)**")
    if po_files:
        st.write(f"{len(po_files)} file(s): " + ", ".join(f.name for f in po_files))
    else:
        st.write("Not uploaded yet.")
    st.markdown("</div>", unsafe_allow_html=True)

run_clicked = st.button(
    "▶ Run Multi-PO Price Recheck",
    type="primary",
    disabled=not (matrix_file and po_files),
)

if run_clicked and matrix_file and po_files:
    with st.spinner("Validating POs against Price Matrix..."):
        matrix_data = load_price_matrix(matrix_file, matrix_profile)
        combined_po_items = []
        for pdf in po_files:
            combined_po_items.extend(extract_vpg_po_items(pdf))
        reconciled = reconcile_items(combined_po_items, matrix_data)

    if status_filter == "Mismatches Only":
        display_rows = [r for r in reconciled if r["Status"] == "MISMATCH"]
    elif status_filter == "Matched Only":
        display_rows = [r for r in reconciled if r["Status"] == "MATCHED"]
    elif status_filter == "Not in Matrix":
        display_rows = [r for r in reconciled if r["Status"] == "NOT IN MATRIX"]
    else:
        display_rows = reconciled

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Lines", len(reconciled))
    m2.metric("Matched", sum(1 for r in reconciled if r["Status"] == "MATCHED"))
    m3.metric("Price Mismatches", sum(1 for r in reconciled if r["Status"] == "MISMATCH"))
    m4.metric("Not in Matrix", sum(1 for r in reconciled if r["Status"] == "NOT IN MATRIX"))

    output_table = [
        {
            "PO File": r["PO File"],
            "Item Number": r["Item Number"],
            "Matrix Price": r["Matrix Price"],
            "PO Price": r["PO Price"],
            "Price Difference": r["Price Difference"],
            "Item Description": r["Item Description"],
            "Status": r["Status"],
        }
        for r in display_rows
    ]

    st.markdown('<div class="white-output-header">📋 Reconciliation Output</div>', unsafe_allow_html=True)

    if output_table:
        # Styled HTML table guarantees crisp white background + pure black text
        headers = list(output_table[0].keys())
        th_html = "".join(
            f"<th style='border:1px solid #000000; padding:8px 10px; background:#f1f5f9; color:#000000; text-align:left;'>{h}</th>"
            for h in headers
        )
        body_rows = []
        for row in output_table:
            tds = "".join(
                f"<td style='border:1px solid #000000; padding:7px 10px; background:#ffffff; color:#000000;'>{'' if row[h] is None else row[h]}</td>"
                for h in headers
            )
            body_rows.append(f"<tr>{tds}</tr>")

        table_html = (
            "<div style='overflow-x:auto; background:#ffffff; border:1px solid #000000; border-radius:0 0 10px 10px; padding:8px;'>"
            "<table style='width:100%; border-collapse:collapse; background:#ffffff; color:#000000; font-size:0.9rem;'>"
            f"<thead><tr>{th_html}</tr></thead>"
            f"<tbody>{''.join(body_rows)}</tbody>"
            "</table></div>"
        )
        st.markdown(table_html, unsafe_allow_html=True)

        out_wb = openpyxl.Workbook()
        combined_ws = out_wb.active
        combined_ws.title = "All_POs_Recheck"
        combined_ws.append(headers)
        for r in output_table:
            combined_ws.append([r[h] for h in headers])

        unique_pos = []
        for r in output_table:
            if r["PO File"] not in unique_pos:
                unique_pos.append(r["PO File"])

        for po_name in unique_pos:
            safe_title = re.sub(r"[\\/*?:\[\]]", "_", po_name)[:28] or "PO_Sheet"
            if safe_title in out_wb.sheetnames:
                safe_title = f"{safe_title[:24]}_{len(out_wb.sheetnames)}"
            ws_po = out_wb.create_sheet(title=safe_title)
            ws_po.append(headers)
            for r in output_table:
                if r["PO File"] == po_name:
                    ws_po.append([r[h] for h in headers])

        buf = io.BytesIO()
        out_wb.save(buf)
        buf.seek(0)

        st.download_button(
            label="⬇ Download Multi-Sheet Excel Report (.xlsx)",
            data=buf,
            file_name="vpg_po_price_recheck_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.markdown(
            "<div style='background:#ffffff; color:#000000; border:1px solid #000000; border-radius:0 0 10px 10px; padding:14px;'>"
            "No matching items found for the selected filter."
            "</div>",
            unsafe_allow_html=True,
        )
