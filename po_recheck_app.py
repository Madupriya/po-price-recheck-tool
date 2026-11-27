def extract_vpg_po_items(pdf_file) -> list[dict]:
    reader = PdfReader(io.BytesIO(pdf_file.getvalue()))
    extracted = []

    # Flexible line pattern: <LineNo> <PartCode> <MiddleTextWithQtyPriceAmount>
    # Examples:
    # 1 S2700-0029 CBL 0.210 (50128&55065) R 3,000.00 EA 25/09/2026 3.017 9,051.00
    # 1 | S2700-0084 | CBL0.210 ... | 100.00 EA | 25/09/2026 | 4.748 | 474.80
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
            # Normalize extra spaces and pipe separators from pypdf
            clean_line = " ".join(lines[i].replace("|", " ").split()).strip()
            low = clean_line.lower()

            if not clean_line or any(tok in low for tok in stop_tokens):
                i += 1
                continue

            m = item_start_re.match(clean_line)
            if m:
                item_no = m.group(2).strip().upper()
                rest = m.group(3).strip()

                # Ignore obvious non-part tokens like dates or pure numbers
                if date_re.fullmatch(item_no) or item_no.replace(".", "").isdigit():
                    i += 1
                    continue

                # Extract numeric tokens (excluding dates) to get Unit Price and Amount
                rest_no_dates = date_re.sub(" ", rest)
                nums = [
                    float(tok.replace(",", ""))
                    for tok in num_token_re.findall(rest_no_dates)
                ]

                # A PO item line typically has Qty, Price, Amount (at least 2-3 numbers)
                if len(nums) >= 2:
                    # Usually last two numbers are Price and Amount (or if only 2, Price then Amount/Qty)
                    if len(nums) >= 3:
                        po_price = nums[-2]
                    else:
                        po_price = nums[-1]

                    # Clean description by removing trailing numbers/dates/units
                    desc_part1 = re.split(r"\b(?:EA|M|KG|PC|PCS|MTR|MTRS|SET|ROLL|NOS?)\b", rest, maxsplit=1)[0]
                    desc_part1 = date_re.sub("", desc_part1)
                    desc_part1 = re.sub(r"\s+\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:\s+\d{1,3}(?:,\d{3})*(?:\.\d+)?)*\s*$", "", desc_part1)
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
