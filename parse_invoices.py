#!/usr/bin/env python3
"""
parse_invoices.py
-----------------
CLI tool to parse distributor invoices (PDF) into a normalized CSV/XLSX and map each line to a GL code.

USAGE
-----
python parse_invoices.py \
  --invoices "/path/to/folder/with/pdfs" \
  --headers "/path/to/Invoice-Headers.xlsx" \
  --gl "/path/to/ChartofAccounts.xlsx" \
  --out_csv "/path/to/out.csv" \
  --out_xlsx "/path/to/out.xlsx" \
  [--vendor "Big Geyser Inc."] \
  [--terms "CHAIN 30"] \
  [--recursive]

Notes
-----
- Designed/tested on "Big Geyser" style weekly PDFs. Other layouts will produce fewer/missing fields unless extended.
- Missing/unknown fields are left blank (NULL).
- Basic fuzzy matching for GL mapping is enabled via difflib.get_close_matches.
"""

import argparse
import re
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

import pandas as pd
from difflib import get_close_matches

# --- PDF text extraction ---
def extract_text_from_pdf(path: Path) -> str:
    try:
        import PyPDF2
    except Exception as e:
        print("ERROR: PyPDF2 is required. Install with: pip install PyPDF2", file=sys.stderr)
        raise
    try:
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            texts = []
            for p in reader.pages:
                try:
                    texts.append(p.extract_text() or "")
                except Exception:
                    texts.append("")
            return "\\n".join(texts)
    except Exception as e:
        # Fallback: empty text
        return ""

# --- Parse Big Geyser-style blocks ---
def parse_big_geyser(text: str) -> List[Dict[str, Any]]:
    # Loosen the pattern slightly around whitespace
    pattern = re.compile(
        r'(?:[A-Z][a-z]{2}\\s+[A-Z][a-z]{2}\\s+\\d{1,2},\\s+\\d{4}.*?)?'
        r'Account:\\s*(\\d+)\\s*Invoice#:\\s*([0-9A-Z]+)(.*?)(?=Account:\\s*\\d+\\s*Invoice#:|\\Z)',
        re.S
    )
    blocks = list(pattern.finditer(text))
    results = []
    for m in blocks:
        acct = m.group(1)
        inv_no = m.group(2)
        block = m.group(3)
        m_date = re.search(r'([A-Z][a-z]{2}\\s+[A-Z][a-z]{2}\\s+\\d{1,2},\\s+\\d{4})', block)
        inv_date = None
        if m_date:
            try:
                inv_date = pd.to_datetime(m_date.group(1))
            except Exception:
                inv_date = None
        items = []
        m_items = re.search(r'ITEM#\\s*DESCRIPTION\\s*QTY\\s*-+\\s*(.*?)\\s*(?:Cases:|FREE GOODS|\\Z)', block, re.S|re.I)
        if m_items:
            for raw in m_items.group(1).splitlines():
                raw = raw.strip()
                if not raw:
                    continue
                mline = re.match(r'(?P<sku>\\d{3,})\\s+(?P<desc>.*?)(?P<qty>\\d+)$', raw)
                if mline:
                    sku = mline.group("sku")
                    desc = mline.group("desc").strip()
                    qty = int(mline.group("qty"))
                else:
                    qty = int(raw[-1]) if raw and raw[-1].isdigit() else 1
                    rest = raw[:-1].strip()
                    m2 = re.match(r'(?P<sku>\\d{3,})\\s+(?P<desc>.+)', rest)
                    if m2:
                        sku = m2.group("sku")
                        desc = m2.group("desc").strip()
                    else:
                        sku=None; desc=rest
                items.append({"sku": sku, "description": desc, "quantity": qty})
        if not items and re.search(r'FREE GOODS', block, re.I):
            items = [{"sku": None, "description": "FREE GOODS - NO CHARGE TO CUSTOMER", "quantity": 1}]
        results.append({
            "account": acct,
            "invoice_number": inv_no,
            "invoice_date": inv_date,
            "items": items
        })
    return results

# --- GL mapping ---
def build_gl_index(gl_df: pd.DataFrame):
    gl_rows = gl_df[["Number","Account (invoices)"]].dropna().copy()
    gl_rows["Account (invoices)"] = gl_rows["Account (invoices)"].astype(str)
    gl_rows["Number"] = gl_rows["Number"].astype(str)
    acct_list = gl_rows["Account (invoices)"].tolist()
    acct_by_name = {a: n for a, n in zip(gl_rows["Account (invoices)"], gl_rows["Number"])}
    # keyword map
    kw_map = {}
    for _, r in gl_rows.iterrows():
        acct = r["Account (invoices)"].lower()
        num  = r["Number"].strip()
        if "sample" in acct: kw_map.setdefault("sample", num)
        if "free" in acct: kw_map.setdefault("free goods", num)
        if "advertis" in acct or "pos" in acct or "marketing" in acct: kw_map.setdefault("advertising", num)
        if "rebate" in acct: kw_map.setdefault("rebate", num)
        if "invasion fee" in acct or "slotting" in acct: kw_map.setdefault("invasion", num)
        if "allowance" in acct or "discount" in acct or "off invoice" in acct: kw_map.setdefault("allowance", num)
        if "incentive" in acct: kw_map.setdefault("incentive", num)
    return acct_list, acct_by_name, kw_map

def gl_map_for_description(desc: str, acct_list, acct_by_name, kw_map) -> str:
    if not desc:
        return "Unmapped"
    d = desc.lower()
    if any(w in d for w in ["free goods","no charge","sample","samples","donation"]):
        return kw_map.get("sample", "Unmapped")
    if "advertis" in d or "promo" in d or "pos" in d or "display" in d:
        return kw_map.get("advertising", "Unmapped")
    if "rebate" in d:
        return kw_map.get("rebate", "Unmapped")
    if "slotting" in d or "invasion" in d:
        return kw_map.get("invasion", "Unmapped")
    if "allowance" in d or "discount" in d or "off invoice" in d:
        return kw_map.get("allowance", "Unmapped")
    if "incentive" in d:
        return kw_map.get("incentive", "Unmapped")
    match = get_close_matches(desc, acct_list, n=1, cutoff=0.86)
    if match:
        return acct_by_name.get(match[0], "Unmapped")
    return "Unmapped"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--invoices", required=True, help="Folder containing PDF invoices")
    ap.add_argument("--headers", required=True, help="Invoice-Headers.xlsx path")
    ap.add_argument("--gl", required=True, help="Chart of Accounts (xlsx) path")
    ap.add_argument("--out_csv", required=True, help="Output CSV path")
    ap.add_argument("--out_xlsx", required=False, help="Output XLSX path")
    ap.add_argument("--vendor", default="Big Geyser Inc.", help="Vendor name to populate")
    ap.add_argument("--terms", default="CHAIN 30", help="Terms to populate")
    ap.add_argument("--recursive", action="store_true", help="Recurse into subfolders")
    args = ap.parse_args()

    inv_dir = Path(args.invoices)
    headers_xlsx = Path(args.headers)
    gl_xlsx = Path(args.gl)

    # Load required headers list
    hdr_df = pd.read_excel(headers_xlsx, header=None)
    required_headers = hdr_df[0].dropna().tolist()
    if required_headers and required_headers[0] == "Field Name":
        required_headers = required_headers[1:]

    # Load GL chart
    xls = pd.ExcelFile(gl_xlsx)
    sheet = "ChartofAccounts" if "ChartofAccounts" in xls.sheet_names else xls.sheet_names[0]
    gl_df = pd.read_excel(gl_xlsx, sheet_name=sheet)
    acct_list, acct_by_name, kw_map = build_gl_index(gl_df)

    rows = []
    pdf_iter = inv_dir.rglob("*.pdf") if args.recursive else inv_dir.glob("*.pdf")
    total_invoices = 0
    for pdf_path in pdf_iter:
        text = extract_text_from_pdf(pdf_path)
        if ("Account:" in text) and ("Invoice#:" in text):
            invs = parse_big_geyser(text)
        else:
            invs = []
        print(f"[{pdf_path.name}] parsed_invoices={len(invs)}")
        total_invoices += len(invs)
        for inv in invs:
            tranId = inv["invoice_number"]
            inv_date = inv["invoice_date"]
            items = inv["items"] or [{"sku": None, "description": None, "quantity": 1}]
            for it in items:
                desc = it.get("description")
                qty = it.get("quantity") or 1
                rate = 0.0
                amt = rate * qty
                gl_code = gl_map_for_description(desc or "", acct_list, acct_by_name, kw_map)
                row = {
                    "tranId": tranId,
                    "postingPeriodRef": inv_date.strftime("%b %Y") if isinstance(inv_date, pd.Timestamp) and not pd.isna(inv_date) else None,
                    "vendorRef": args.vendor,
                    "tranDate": inv_date.strftime("%m/%d/%Y") if isinstance(inv_date, pd.Timestamp) and not pd.isna(inv_date) else None,
                    "payableAccountRef": None,
                    "termsRef": args.terms,
                    "memo": None,
                    "purchaseItemline_itemRef": it.get("sku"),
                    "purchaseItemline_quantity": qty,
                    "purchaseItemline_serialNumbers": None,
                    "purchaseitemline_unitsRef": "CASE" if it.get("sku") else None,
                    "purchaseItemLine_rate": rate,
                    "purchaseItemLine_amount": amt,
                    "purchaseItemLine_memo": desc,
                    "purchaseItemLine_departmentRef": None,
                    "purchaseItemLine_classRef": gl_code,
                    "purchaseItemLine_locationRef": None,
                    "purchaseItemLine_customerRef": None,
                    "purchaseItemLine_isBillable": False,
                    "purchaseItemLine_taxCodeRef": None,
                    "purchaseItemLine_taxCodeAmount": 0.0
                }
                for h in required_headers:
                    row.setdefault(h, None)
                rows.append(row)

    out_df = pd.DataFrame(rows, columns=required_headers)
    out_csv = Path(args.out_csv)
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_csv(out_csv, index=False)
    if args.out_xlsx:
        out_xlsx = Path(args.out_xlsx)
        out_xlsx.parent.mkdir(parents=True, exist_ok=True)
        out_df.to_excel(out_xlsx, index=False)

    print(f"Found {total_invoices} invoice blocks across PDFs.")
    print(f"Wrote {len(out_df)} rows to {out_csv}")
    if args.out_xlsx:
        print(f"Wrote XLSX to {out_xlsx}")

if __name__ == "__main__":
    main()
