
# Take-Home: Invoice Parsing & GL Mapping

## Parser / Service
- Used Python with PyPDF2 to extract text from PDFs directly.
- No external APIs were required.

## Approach
1. **PDF Text Extraction**: Read the Big Geyser multi-invoice PDF and extracted raw text.
2. **Invoice Detection**: Used regex to find invoice blocks keyed by `Account: ####` and `Invoice#: xxxxxxx`.
3. **Line Items**: Parsed the `ITEM# ... DESCRIPTION ... QTY` section; where quantities were missing, defaulted to `1`.
4. **Normalization**: Aligned fields to your target headers from `Invoice-Headers.xlsx`. Missing/unsure fields set to `NULL` (blank in CSV).
5. **GL Mapping**: Built a lightweight keyword-to-GL mapping from your Chart of Accounts:
   - samples/free goods/donations → **Samples** (GL 6520)
   - advertising/POS/promo → **Distributor Advertising** (GL 6405)
   - rebates → **Rebates** (GL 4809)
   - invasion/slotting → **Invasion Fee** (GL 4825)
   - allowances/discounts → **Sales Allowances** (GL 4834)
   - incentives → **Incentives** (GL 4837)
   - otherwise → **Unmapped**

6. **Output**: Produced both CSV and XLSX with the normalized columns.

## Assumptions
- Big Geyser "FREE GOODS *** NO CHARGE TO CUSTOMER ***" lines are treated as **Samples** (GL mapping above).
- `purchaseItemLine_rate` set to `0.00` where invoices indicate free goods; amount = rate × qty.
- When the invoice date is not reliably parsed, `postingPeriodRef` and `tranDate` are left NULL.
- Stored the GL Code into `purchaseItemLine_classRef` (can be moved to a different field if you prefer a dedicated GL column).

## How to Run
1. Place PDFs into `/mnt/data/` (or update the `pdf_paths` list in the notebook/script).
2. Ensure `Invoice-Headers.xlsx` and `ChartofAccounts.xlsx` exist in `/mnt/data/`.
3. Execute the notebook cells (or run the Python script version) to regenerate the CSV/XLSX.

## Next Improvements (Bonus-ready)
- Add fuzzy matching (e.g., RapidFuzz) to map unusual deduction descriptions.
- Package as a small CLI (`python parse_invoices.py --invoices folder --out out.csv`).
- Improve store/address parsing with a layout-aware library (e.g., pdfplumber) for more robust results.
