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


def _find_header_and_cols(ws) -> tuple[int, dict]:
    for r in range(1, min(10, ws.max_row + 1)):
        row_map = {}
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val is not None:
                row_map[str(val).strip().lower()] = c
        if "customer p/n" in row_map:
            return r, row_map
    return 2, {}


def _pick_price_column(row_map: dict, price_mode: str) -> int | None:
    preferred_lists = {
        "Current Selling Price / Agreed Price": [
            "current selling price",
            "agreed price 2024",
            "current sales price usd/m",
            "new price july 2026 $/m",
            "suggested sales price validity sept/oct 2026",
            "suggested sales price for 2026/27",
        ],
        "Suggested Sales Price (Sept/Oct 2026)": [
            "suggested sales price validity sept/oct 2026",
            "new price july 2026 $/m",
            "suggested sales price for 2026/27",
            "current selling price",
            "agreed price 2024",
        ],
        "Suggested Sales Price (2026/27)": [
            "suggested sales price for 2026/27",
            "new price july 2026 $/m",
            "suggested sales price validity sept/oct 2026",
            "current selling price",
            "agreed price 2024",
        ],
    }
    candidates = preferred_lists.get(price_mode, preferred_lists["Current Selling Price / Agreed Price"])
    for key in candidates:
        if key in row_map:
            return row_map[key]
    return None


def load_price_matrix(excel_file, matrix_profile: str, price_mode: str) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(excel_file.getvalue()), data_only=True)
    matrix_lookup = {}

    target_sheets = [s for s in wb.sheetnames if s.lower() != "main sheet"]
    if matrix_profile in ("Cable1", "Harness"):
        matched = [s for s in target_sheets if matrix_profile.lower() in s.lower()]
        if matched:
            target_sheets = matched

    for sheet_name in target_sheets:
        ws = wb[sheet_name]
        header_row, row_map = _find_header_and_cols(ws)
        pn_col = row_map.get("customer p/n")
        desc_col = row_map.get("item description")
        price_col = _pick_price_column(row_map, price_mode)

        if not pn_col or not price_col:
            continue

        for r in range(header_row + 1, ws.max_row + 1):
            pn_val = ws.cell(r, pn_col).value
            price_val = ws.cell(r, price_col).value
            if pn_val is None or price_val is None:
                continue

            item_no = str(pn_val).strip().upper()
            if not re.match(r"^[A-Z0-9]+-[A-Z0-9\-]+$", item_no):
                continue

            try:
                m_price = float(str(price_val).replace(",", "").strip())
            except ValueError:
                continue

            desc = ""
            if desc_col and ws.cell(r, desc_col).value is not None:
                desc = " ".join(str(ws.cell(r, desc_col).value).split()).strip()

            if item_no not in matrix_lookup or (
                matrix_profile != "Auto (All Sheets)" and matrix_profile.lower() in sheet_name.lower()
            ):
                matrix_lookup[item_no] = {
                    "matrix_price": m_price,
                    "description": desc,
                    "sheet": sheet_name,
                }

    return matrix_lookup


def extract_vpg_po_items(pdf_file) -> list[dict]:
    reader = PdfReader(io.BytesIO(pdf_file.getvalue()))
    extracted = []

    item_start_re = re.compile(r"^\s*(\d{1,3})\s+([A-Z0-9][A-Z0-9\-_./]{2,25})\s+(.+)$")
    num_token_re = re.compile(r"(?<!\d)(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?)")
    date_re = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")

    stop_tokens = (
        "inco terms",
        "total amount",
        "bill to",
        "ship to",
        "page ",
        "vendor",
        "forwarder",
        "gstin",
        "phone",
        "discount",
        "vat",
        "shipping instructions",
        "terms and conditions",
    )

    for page in reader.pages:
        txt = page.extract_text() or ""
        lines = txt.splitlines()
        i = 0
        while i < len(lines):
            clean_line = " ".join(lines[i].replace("|", " ").split()).strip()
            low = clean_line.lower()

            if not clean_line or any(tok in low for tok in stop_tokens):
                i += 1
                continue

            m = item_start_re.match(clean_line)
            if m:
                item_no = m.group(2).strip().upper()
                rest = m.group(3).strip()

                if date_re.fullmatch(item_no) or item_no.replace(".", "").isdigit():
                    i += 1
                    continue

                rest_no_dates = date_re.sub(" ", rest)
                nums = [
                    float(tok.replace(",", ""))
                    for tok in num_token_re.findall(rest_no_dates)
                ]

                if len(nums) >= 2:
                    po_price = nums[-2] if len(nums) >= 3 else nums[-1]

                    desc_part1 = re.split(
                        r"\b(?:EA|M|KG|PC|PCS|MTR|MTRS|SET|ROLL|NOS?)\b",
                        rest,
                        maxsplit=1,
                    )[0]
                    desc_part1 = date_re.sub("", desc_part1)
                    desc_part1 = re.sub(
                        r"\s+\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:\s+\d{1,3}(?:,\d{3})*(?:\.\d+)?)*\s*$",
                        "",
                        desc_part1,
                    )
                    desc_part1 = re.sub(r"\s+[A-Z0-9]{1,2}\s*$", "", desc_part1).strip()

                    desc_parts = [desc_part1] if desc_part1 else []
                    if i + 1 < len(lines):
                        nxt = " ".join(lines[i + 1].replace("|", " ").split()).strip()
                        nxt_low = nxt.lower()
                        if (
                            nxt
                            and not item_start_re.match(nxt)
                            and not any(tok in nxt_low for tok in stop_tokens)
                            and not date_re.search(nxt)
                        ):
                            desc_parts.append(nxt)
                            i += 1

                    extracted.append(
                        {
                            "PO File": pdf_file.name,
                            "Item Number": item_no,
                            "Item Description": " ".join(desc_parts).strip(),
                            "PO Price": po_price,
                        }
                    )
            i += 1

    return extracted


def reconcile_items(po_items: list[dict], matrix_lookup: dict) -> list[dict]:
    rows = []
    for item in po_items:
        item_no = item["Item Number"].upper()
        po_price = item["PO Price"]
        pdf_desc = item.get("Item Description", "")

        m_entry = matrix_lookup.get(item_no)
        if m_entry is not None:
            matrix_price = m_entry["matrix_price"]
            final_desc = pdf_desc or m_entry["description"]
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
                "Matrix Price": matrix_price,
                "PO Price": po_price,
                "Price Difference": diff,
                "Item Description": final_desc,
                "Status": status,
            }
        )
    return rows


def render_white_table(rows_list: list[dict], headers: list[str]) -> str:
    if not rows_list:
        return (
            "<div style='background:#ffffff; color:#000000; padding:12px; border:1px solid #000000;'>"
            "No line items found for this PO / filter."
            "</div>"
        )
    th_html = "".join(
        f"<th style='border:1px solid #000000; padding:8px 10px; background:#f1f5f9; color:#000000; text-align:left;'>{h}</th>"
        for h in headers
    )
    body_rows = []
    for row in rows_list:
        tds = "".join(
            f"<td style='border:1px solid #000000; padding:7px 10px; background:#ffffff; color:#000000;'>{'' if row[h] is None else row[h]}</td>"
            for h in headers
        )
        body_rows.append(f"<tr>{tds}</tr>")
    return (
        "<div style='overflow-x:auto; background:#ffffff; padding:8px; border:1px solid #000000;'>"
        "<table style='width:100%; border-collapse:collapse; background:#ffffff; color:#000000; font-size:0.9rem;'>"
        f"<thead><tr>{th_html}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )


st.markdown(
    """
    <div class="main-banner">
        <h1>⚡ VPG Purchase Order — Price Reconciliation Studio</h1>
        <p>Multi-PO Batch Validation against Cable1 &amp; Harness Price Matrices • Multi-Sheet Excel Export</p>
    </div>
    """,
    unsafe_allow_html=True,
)

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
    price_mode = st.selectbox(
        "Matrix Price Column",
        [
            "Current Selling Price / Agreed Price",
            "Suggested Sales Price (Sept/Oct 2026)",
            "Suggested Sales Price (2026/27)",
        ],
    )
    status_filter = st.radio(
        "Display Rows",
        ["All Rows", "Mismatches Only", "Matched Only", "Not in Matrix"],
    )

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
        matrix_data = load_price_matrix(matrix_file, matrix_profile, price_mode)
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

    for pdf in po_files:
        found_count = sum(1 for r in reconciled if r["PO File"] == pdf.name)
        if found_count == 0:
            with st.expander(f"⚠️ 0 items extracted from {pdf.name} — View Raw PDF Text"):
                dbg_reader = PdfReader(io.BytesIO(pdf.getvalue()))
                st.text("\n".join((p.extract_text() or "") for p in dbg_reader.pages[:2]))

    st.markdown('<div class="white-output-header">📋 Reconciliation Output (All POs &amp; Per-PO Sheets)</div>', unsafe_allow_html=True)

    headers = [
        "PO File",
        "Item Number",
        "Matrix Price",
        "PO Price",
        "Price Difference",
        "Item Description",
        "Status",
    ]

    uploaded_po_names = [f.name for f in po_files]
    tabs = st.tabs(["All POs"] + uploaded_po_names)

    with tabs[0]:
        st.markdown(render_white_table(display_rows, headers), unsafe_allow_html=True)

    for idx, po_name in enumerate(uploaded_po_names, start=1):
        po_specific_rows = [r for r in display_rows if r["PO File"] == po_name]
        with tabs[idx]:
            st.markdown(render_white_table(po_specific_rows, headers), unsafe_allow_html=True)

    out_wb = openpyxl.Workbook()
    ws_all = out_wb.active
    ws_all.title = "All_POs"
    ws_all.append(headers)
    for r in display_rows:
        ws_all.append([r[h] for h in headers])

    used_names = {"All_POs"}
    for po_name in uploaded_po_names:
        base_name = re.sub(r"[\\/*?:\[\]]", "_", po_name.rsplit(".", 1)[0])[:28] or "PO_Sheet"
        sheet_title = base_name
        counter = 1
        while sheet_title in used_names:
            suffix = f"_{counter}"
            sheet_title = f"{base_name[:28 - len(suffix)]}{suffix}"
            counter += 1
        used_names.add(sheet_title)

        ws_po = out_wb.create_sheet(title=sheet_title)
        ws_po.append(headers)
        for r in [x for x in reconciled if x["PO File"] == po_name]:
            ws_po.append([r[h] for h in headers])

    buf = io.BytesIO()
    out_wb.save(buf)
