
import streamlit as st
import openpyxl
import io
import re

try:
    from pypdf import PdfReader
    PDF_AVAILABLE = True
except Exception:
    PDF_AVAILABLE = False

st.set_page_config(page_title="PO Price Reconciliation Tool", layout="wide")

st.title("PO vs Price Matrix Reconciliation Tool")
st.write("Supports Cable1 (Col B = Customer P/N, Col D = Description) and Harness (Col C = Customer P/N, Col E = Description) with 5th August 2026 pricing.")

def parse_price(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    cleaned = re.sub(r"[^0-9.\-]", "", str(val).strip())
    if cleaned in ("", "-", ".", "-."):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None

def format_price(num_val):
    if num_val is None:
        return "-"
    s = f"{num_val:.4f}".rstrip("0").rstrip(".")
    if "." not in s:
        return f"{s}.00"
    dec_part = s.split(".")[1]
    if len(dec_part) < 2:
        return f"{s}{'0' * (2 - len(dec_part))}"
    return s

def find_aug5_price_col(ws_val, ws_raw):
    # Search for "5th August 2026" or "August 2026" or default to Col AF (32)
    for r in range(1, min(6, ws_val.max_row + 1)):
        for c in range(1, ws_val.max_column + 1):
            v = ws_val.cell(row=r, column=c).value
            if v is None:
                v = ws_raw.cell(row=r, column=c).value
            if v is not None and ("5TH AUGUST 2026" in str(v).upper() or "AUG 2026" in str(v).upper()):
                return c
    return 32 if ws_val.max_column >= 32 else ws_val.max_column

def extract_po_items_from_pdf(pdf_bytes, file_name, known_pns_upper):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    extracted = []

    block_pattern = re.compile(
        r"\b([A-Z]{1,4}\d{3,6}-\d{2,6}[A-Z0-9\-]*)\b.*?"
        r"([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2})\s*(?:EA|PCS|PC|M|MTR|MTRS|KG|KGS|SET|SETS)\s+"
        r"(\d{2}/\d{2}/\d{4})\s+"
        r"([0-9]+(?:\.[0-9]{1,4})?)\s+"
        r"([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2})",
        re.IGNORECASE | re.DOTALL
    )

    fallback_pattern = re.compile(
        r"\b([A-Z]{1,4}\d{3,6}-\d{2,6}[A-Z0-9\-]*)\b.*?"
        r"([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2})\s*(?:EA|PCS|PC|M|MTR|MTRS|KG|KGS|SET|SETS)\s+"
        r"([0-9]+(?:\.[0-9]{1,4})?)\s+"
        r"([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2})",
        re.IGNORECASE | re.DOTALL
    )

    for page_idx, page in enumerate(reader.pages, start=1):
        txt = page.extract_text() or ""
        normalized = re.sub(r"\s+", " ", txt)
        seen_in_page = set()

        for m in block_pattern.finditer(normalized):
            pn_found = m.group(1).strip().upper()
            qty_val = parse_price(m.group(2))
            unit_price_val = parse_price(m.group(4))
            amount_val = parse_price(m.group(5))

            if qty_val and unit_price_val and amount_val:
                expected_amt = qty_val * unit_price_val
                if abs(expected_amt - amount_val) > max(1.0, amount_val * 0.05):
                    continue

            key_uniq = (pn_found, unit_price_val, amount_val)
            if key_uniq not in seen_in_page:
                seen_in_page.add(key_uniq)
                extracted.append({
                    "file": file_name,
                    "page": page_idx,
                    "pn": pn_found,
                    "po_price": unit_price_val
                })

        if not seen_in_page:
            for m in fallback_pattern.finditer(normalized):
                pn_found = m.group(1).strip().upper()
                qty_val = parse_price(m.group(2))
                unit_price_val = parse_price(m.group(3))
                amount_val = parse_price(m.group(4))

                key_uniq = (pn_found, unit_price_val, amount_val)
                if key_uniq not in seen_in_page:
                    seen_in_page.add(key_uniq)
                    extracted.append({
                        "file": file_name,
                        "page": page_idx,
                        "pn": pn_found,
                        "po_price": unit_price_val
                    })

        for k_pn in known_pns_upper:
            if re.search(rf"(?<![A-Z0-9\-]){re.escape(k_pn)}(?![A-Z0-9\-])", normalized.upper()):
                already = any(x["pn"] == k_pn and x["page"] == page_idx for x in extracted)
                if not already:
                    extracted.append({
                        "file": file_name,
                        "page": page_idx,
                        "pn": k_pn,
                        "po_price": None
                    })

    return extracted

col_up1, col_up2 = st.columns(2)
with col_up1:
    price_file = st.file_uploader("1. Upload Price Matrix Excel (.xlsx)", type=["xlsx"], key="price_matrix")
with col_up2:
    po_files = st.file_uploader(
        "2. Upload Customer PO Files (.pdf or .xlsx)",
        type=["pdf", "xlsx"],
        accept_multiple_files=True,
        key="po_files"
    )

if price_file is not None:
    try:
        p_bytes = price_file.getvalue()
        wb_val = openpyxl.load_workbook(io.BytesIO(p_bytes), data_only=True)
        wb_raw = openpyxl.load_workbook(io.BytesIO(p_bytes), data_only=False)

        available_sheets = wb_val.sheetnames
        default_sheets = [s for s in available_sheets if s.lower() in ("cable1", "harness")]
        if not default_sheets:
            default_sheets = available_sheets

        selected_sheets = st.multiselect(
            "Price Matrix Sheets:",
            options=available_sheets,
            default=default_sheets
        )

        price_map = {}
        sheet_counts = {}

        for s_name in selected_sheets:
            ws_v = wb_val[s_name]
            ws_r = wb_raw[s_name]

            # Cable1 has Customer P/N in Col B (2) and Description in Col D (4)
            # Harness has Customer P/N in Col C (3) and Description in Col E (5)
            if s_name.strip().lower() == "cable1":
                pn_col_idx = 2
                desc_col_idx = 4
            else:
                pn_col_idx = 3
                desc_col_idx = 5

            price_col_idx = find_aug5_price_col(ws_v, ws_r)
            loaded_here = 0

            for r in range(3, ws_v.max_row + 1):
                pn_v = ws_v.cell(row=r, column=pn_col_idx).value
                pn_r = ws_r.cell(row=r, column=pn_col_idx).value
                chosen_pn = pn_v if pn_v is not None else pn_r

                if chosen_pn is not None and str(chosen_pn).strip() != "":
                    pn_str = str(chosen_pn).strip()
                    if pn_str.upper() in ("CUSTOMER P/N", "P/N", "PART NO", "ITEM NO", "SL.NO"):
                        continue

                    d_v = ws_v.cell(row=r, column=desc_col_idx).value
                    d_r = ws_r.cell(row=r, column=desc_col_idx).value
                    chosen_desc = d_v if d_v is not None else d_r
                    desc_str = "-" if chosen_desc is None or str(chosen_desc).strip() == "" else str(chosen_desc).strip()

                    pr_v = ws_v.cell(row=r, column=price_col_idx).value
                    pr_r = ws_r.cell(row=r, column=price_col_idx).value
                    chosen_price = pr_v if pr_v is not None else pr_r
                    num_price = parse_price(chosen_price)

                    key = pn_str.upper()
                    if key not in price_map:
                        price_map[key] = {
                            "pn": pn_str,
                            "desc": desc_str,
                            "price_num": num_price,
                            "price_str": format_price(num_price) if num_price is not None else (str(chosen_price).strip() if chosen_price is not None else "-"),
                            "sheet": s_name
                        }
                        loaded_here += 1

            sheet_counts[s_name] = loaded_here

        st.success(f"Loaded {len(price_map)} unique P/Ns -> {sheet_counts}")
        st.divider()

        if po_files:
            results = []
            match_count = 0
            mismatch_count = 0
            not_found_count = 0

            for f in po_files:
                fname = f.name
                fbytes = f.getvalue()

                if fname.lower().endswith(".pdf") and PDF_AVAILABLE:
                    extracted_items = extract_po_items_from_pdf(fbytes, fname, list(price_map.keys()))

                    for item in extracted_items:
                        po_pn_str = item["pn"]
                        po_price_num = item["po_price"]
                        k = po_pn_str.upper()

                        if k in price_map:
                            ref = price_map[k]
                            ref_price_num = ref["price_num"]

                            if po_price_num is not None and ref_price_num is not None:
                                diff = round(po_price_num - ref_price_num, 4)
                                if abs(diff) < 0.0005:
                                    status = "MATCH"
                                    match_count += 1
                                else:
                                    status = "PRICE MISMATCH"
                                    mismatch_count += 1
                                diff_str = f"{diff:+.4f}".rstrip("0").rstrip(".")
                                if diff_str in ("", "+"):
                                    diff_str = "0.00"
                            else:
                                status = "CHECK MANUALLY"
                                diff_str = "-"

                            results.append({
                                "PO File": fname,
                                "Page": f"Page {item['page']}",
                                "PO Item Number": po_pn_str,
                                "Source Sheet": ref["sheet"],
                                "Item Description": ref["desc"],
                                "PO Price": format_price(po_price_num),
                                "5th Aug 2026 Price": ref["price_str"],
                                "Difference (PO - Matrix)": diff_str,
                                "Status": status
                            })
                        else:
                            not_found_count += 1
                            results.append({
                                "PO File": fname,
                                "Page": f"Page {item['page']}",
                                "PO Item Number": po_pn_str,
                                "Source Sheet": "-",
                                "Item Description": "Item not found",
                                "PO Price": format_price(po_price_num),
                                "5th Aug 2026 Price": "-",
                                "Difference (PO - Matrix)": "-",
                                "Status": "NOT FOUND"
                            })

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total PO Items Found", len(results))
            m2.metric("Exact Price Match", match_count)
            m3.metric("Price Mismatch", mismatch_count)
            m4.metric("Not Found", not_found_count)

            filter_opt = st.radio(
                "Filter Results:",
                ["All Items", "Only Mismatches & Not Found", "Only Matches"],
                horizontal=True
            )

            if filter_opt == "Only Mismatches & Not Found":
                filtered_results = [x for x in results if x["Status"] in ("PRICE MISMATCH", "NOT FOUND", "CHECK MANUALLY")]
            elif filter_opt == "Only Matches":
                filtered_results = [x for x in results if x["Status"] == "MATCH"]
            else:
                filtered_results = results

            st.dataframe(filtered_results, width="stretch")

            if results:
                out_wb = openpyxl.Workbook()
                out_ws = out_wb.active
                out_ws.title = "PO_Recheck_Report"
                headers_out = list(results[0].keys())
                out_ws.append(headers_out)
                for row_dict in results:
                    out_ws.append([row_dict[h] for h in headers_out])

                out_buf = io.BytesIO()
                out_wb.save(out_buf)
                out_buf.seek(0)

                st.download_button(
                    label="Download Full Comparison Report (.xlsx)",
                    data=out_buf.getvalue(),
                    file_name="PO_Price_Comparison_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error processing files: {e}")