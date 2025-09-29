#!/usr/bin/env python3
"""
parse_invoices_v3.py
- Verbose logging
- Fallback extractor: pdfplumber (layout-aware)
USAGE:
python parse_invoices_v3.py --invoices ./invoices --headers ./Invoice-Headers.xlsx --gl ./ChartofAccounts.xlsx --out_csv ./out.csv --out_xlsx ./out.xlsx --recursive --verbose --use-plumber
"""
import argparse, re, sys
from pathlib import Path
from typing import List, Dict, Any
import pandas as pd
from difflib import get_close_matches

def log(msg, verbose):
    if verbose:
        print(msg, flush=True)

def extract_text_pypdf2(path: Path, verbose=False) -> str:
    try:
        import PyPDF2
        txts = []
        with open(path, "rb") as f:
            r = PyPDF2.PdfReader(f)
            for i, p in enumerate(r.pages):
                try:
                    t = p.extract_text() or ""
                except Exception:
                    t = ""
                txts.append(t)
        text = "\n".join(txts)
        log(f"[PyPDF2] {path.name}: {len(text)} chars", verbose)
        return text
    except Exception as e:
        log(f"[PyPDF2] {path.name}: ERROR {e}", verbose)
        return ""

def extract_text_plumber(path: Path, verbose=False) -> str:
    try:
        import pdfplumber
    except Exception as e:
        log("[pdfplumber] not installed. pip install pdfplumber", verbose)
        return ""
    out = []
    try:
        with pdfplumber.open(str(path)) as pdf:
            for i, page in enumerate(pdf.pages):
                try:
                    t = page.extract_text() or ""
                except Exception:
                    t = ""
                out.append(t)
        text = "\n".join(out)
        log(f"[pdfplumber] {path.name}: {len(text)} chars", verbose)
        return text
    except Exception as e:
        log(f"[pdfplumber] {path.name}: ERROR {e}", verbose)
        return ""

def parse_big_geyser(text: str, verbose=False):
    # Search both raw and flattened text
    flat = re.sub(r'\s+', ' ', text)
    patt = re.compile(r'Account:\s*(\d+)\s*Invoice#:\s*([0-9A-Z]+)(.*?)(?=Account:\s*\d+\s*Invoice#:|\Z)', re.S)
    blocks = list(patt.finditer(flat))
    results = []
    for m in blocks:
        acct = m.group(1)
        inv_no = m.group(2)
        block = m.group(3)
        # date (best-effort)
        m_date = re.search(r'([A-Z][a-z]{2}\s+[A-Z][a-z]{2}\s+\d{1,2},\s+\d{4})', block)
        inv_date = None
        if m_date:
            try:
                inv_date = pd.to_datetime(m_date.group(1))
            except Exception:
                inv_date = None
        items = []
        # Try to find "ITEM# ... DESCRIPTION ... QTY" section
        m_items = re.search(r'ITEM#\s*DESCRIPTION\s*QTY\s*-+\s*(.*?)\s*(?:Cases:|FREE GOODS|\Z)', block, re.S|re.I)
        if m_items:
            for raw in re.split(r'\s*\n', m_items.group(1)):
                raw = raw.strip()
                if not raw: 
                    continue
                mline = re.match(r'(?P<sku>\d{3,})\s+(?P<desc>.*?)(?P<qty>\d+)$', raw)
                if mline:
                    sku = mline.group("sku"); desc = mline.group("desc").strip(); qty = int(mline.group("qty"))
                else:
                    qty = int(raw[-1]) if raw and raw[-1].isdigit() else 1
                    rest = raw[:-1].strip()
                    m2 = re.match(r'(?P<sku>\d{3,})\s+(?P<desc>.+)', rest)
                    if m2:
                        sku = m2.group("sku"); desc = m2.group("desc").strip()
                    else:
                        sku=None; desc=rest
                items.append({"sku": sku, "description": desc, "quantity": qty})
        if not items and re.search(r'FREE GOODS', block, re.I):
            items = [{"sku": None, "description": "FREE GOODS - NO CHARGE TO CUSTOMER", "quantity": 1}]
        results.append({"account": acct, "invoice_number": inv_no, "invoice_date": inv_date, "items": items})
    log(f"[parser] found {len(results)} invoice blocks", verbose)
    return results

def build_gl_index(gl_df: pd.DataFrame):
    gl_rows = gl_df[["Number","Account (invoices)"]].dropna().copy()
    gl_rows["Account (invoices)"] = gl_rows["Account (invoices)"].astype(str)
    gl_rows["Number"] = gl_rows["Number"].astype(str)
    acct_list = gl_rows["Account (invoices)"].tolist()
    acct_by_name = {a: n for a, n in zip(gl_rows["Account (invoices)"], gl_rows["Number"])}
    kw_map = {}
    for _, r in gl_rows.iterrows():
        acct = r["Account (invoices)"].lower(); num = r["Number"].strip()
        if "sample" in acct: kw_map.setdefault("sample", num)
        if "free" in acct: kw_map.setdefault("free goods", num)
        if "advertis" in acct or "pos" in acct or "marketing" in acct: kw_map.setdefault("advertising", num)
        if "rebate" in acct: kw_map.setdefault("rebate", num)
        if "invasion fee" in acct or "slotting" in acct: kw_map.setdefault("invasion", num)
        if "allowance" in acct or "discount" in acct or "off invoice" in acct: kw_map.setdefault("allowance", num)
        if "incentive" in acct: kw_map.setdefault("incentive", num)
    return acct_list, acct_by_name, kw_map

def gl_map_for_description(desc: str, acct_list, acct_by_name, kw_map) -> str:
    if not desc: return "Unmapped"
    d = desc.lower()
    if any(w in d for w in ["free goods","no charge","sample","samples","donation"]): return kw_map.get("sample","Unmapped")
    if "advertis" in d or "promo" in d or "pos" in d or "display" in d: return kw_map.get("advertising","Unmapped")
    if "rebate" in d: return kw_map.get("rebate","Unmapped")
    if "slotting" in d or "invasion" in d: return kw_map.get("invasion","Unmapped")
    if "allowance" in d or "discount" in d or "off invoice" in d: return kw_map.get("allowance","Unmapped")
    if "incentive" in d: return kw_map.get("incentive","Unmapped")
    match = get_close_matches(desc, acct_list, n=1, cutoff=0.86)
    if match: return acct_by_name.get(match[0], "Unmapped")
    return "Unmapped"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--invoices", required=True)
    ap.add_argument("--headers", required=True)
    ap.add_argument("--gl", required=True)
    ap.add_argument("--out_csv", required=True)
    ap.add_argument("--out_xlsx", required=False)
    ap.add_argument("--vendor", default="Distributor")
    ap.add_argument("--terms", default="NET 30")
    ap.add_argument("--recursive", action="store_true")
    ap.add_argument("--verbose", action="store_true")
    ap.add_argument("--use-plumber", action="store_true", help="Try pdfplumber if PyPDF2 yields no text")
    args = ap.parse_args()

    inv_dir = Path(args.invoices)
    hdr_df = pd.read_excel(args.headers, header=None)
    required_headers = hdr_df[0].dropna().tolist()
    if required_headers and required_headers[0] == "Field Name":
        required_headers = required_headers[1:]

    gl_xls = pd.ExcelFile(args.gl)
    sheet = "ChartofAccounts" if "ChartofAccounts" in gl_xls.sheet_names else gl_xls.sheet_names[0]
    gl_df = pd.read_excel(args.gl, sheet_name=sheet)
    acct_list, acct_by_name, kw_map = build_gl_index(gl_df)

    pdf_iter = inv_dir.rglob("*.pdf") if args.recursive else inv_dir.glob("*.pdf")
    rows = []
    count_files = 0
    for pdf in pdf_iter:
        count_files += 1
        print(f"[file] {pdf.name}", flush=True) if args.verbose else None
        text = extract_text_pypdf2(pdf, args.verbose)
        if not text and args.use_plumber:
            text = extract_text_plumber(pdf, args.verbose)
        if not text:
            log(f"[warn] No text extracted from {pdf.name}", args.verbose)
            continue
        invs = parse_big_geyser(text, args.verbose)
        for inv in invs:
            inv_date = inv["invoice_date"]
            for it in (inv["items"] or [{"sku": None, "description": None, "quantity": 1}]):
                desc = it.get("description"); qty = it.get("quantity") or 1
                rate = 0.0; amt = rate * qty
                gl_code = gl_map_for_description(desc or "", acct_list, acct_by_name, kw_map)
                row = {
                    "tranId": inv["invoice_number"],
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
    out_df.to_csv(args.out_csv, index=False)
    if args.out_xlsx:
        out_df.to_excel(args.out_xlsx, index=False)
    log(f"[done] scanned_files={count_files}, rows={len(out_df)}", args.verbose)
    if len(out_df)==0:
        print("NOTE: No rows parsed. Try --use-plumber or ensure invoices folder contains text-based PDFs.", file=sys.stderr)

if __name__ == "__main__":
    main()
