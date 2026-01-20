# ECHS Claim Automation — Windows COM (Word) Pipeline

This approach uses **Microsoft Word COM automation** on your Windows PC so Word itself performs the placeholder replacement and PDF export — preserving underline/bold/line-flow exactly as in the template.

## Why this works
- We do **not** edit DOCX structure via python-docx.
- We let **Word** do Find/Replace and layout.
- This avoids run-splitting, underline corruption, line-wrap regressions, and (a)/(b) renumbering issues.

---

## 1) Prerequisites (Windows)
1. Install **Microsoft Word** (Office/365).
2. Install Python 3.10+.
3. Install pywin32:
   ```bat
   py -m pip install --upgrade pip
   py -m pip install pywin32
   ```

---

## 2) Files
Place these files in a working folder, e.g. `C:\ECHS_claims\`

- `ECHS_Claim_template.docx`  (your authoritative template)
- `claim_values.json`         (values extracted/mapped)
- `run_claim_word_com.py`     (script below)

---

## 3) claim_values.json format

Example for the **A041997 / 14-01-2026** case:

```json
{
  "PATIENT_NAME": "ANURADHA TYAGI",
  "DATE_EXPENDITURE": "14-01-2026",
  "ROW1_AMOUNT": "210.00",
  "TOTAL_AMOUNT": "210",
  "ECHS_CARD_NO": "DL2000008879968",
  "INVOICE_NO.": "A041997",
  "DATE": "14-01-2026",
  "DIAGNOSIS": "F/U OF RA, HYPOTHYROIDISM",
  "MED_1": "HYDROXYCHLOROQUINE 200 MG",
  "FORM_MED_1": "TAB",
  "QTY_MED_1": "30",
  "MED_2": "PREDNISOLONE 2.5 MG",
  "FORM_MED_2": "TAB",
  "QTY_MED_2": "30",
  "MED_3": "",
  "FORM_MED_3": "",
  "QTY_MED_3": "",
  "MED_4": "",
  "FORM_MED_4": "",
  "QTY_MED_4": "",
  "MED_5": "",
  "FORM_MED_5": "",
  "QTY_MED_5": "",
  "AMT_1": "203.80",
  "AMT_2": "33.00",
  "AMT_3": "",
  "AMT_4": "",
  "AMT_5": "",
  "TOTAL_WO_DISCOUNT": "236.80",
  "AMOUNT_WORDS": "Two Hundred and Ten only",
  "CURRENT_MONTH_YEAR": "Jan 2026"
}
```

Notes:
- The JSON keys match placeholder names **without** braces.
  - `PATIENT_NAME` fills `{{PATIENT_NAME}}`
  - `INVOICE_NO.` fills `{{INVOICE_NO.}}`
- Keep values as strings.

---

## 4) Run script

```bat
cd /d C:\ECHS_claims
py run_claim_word_com.py ^
  --template ECHS_Claim_template.docx ^
  --values claim_values.json ^
  --outdocx out\ECHS_Claim_filled.docx ^
  --outpdf  out\ECHS_Claim_filled.pdf
```

---

## 5) What the script does (safely)
1. Opens the template in **Word**.
2. Replaces placeholders using Word’s `Find.Execute` (preserves formatting).
3. Removes empty medicine lines **only when the MED_n is empty** (deletes that paragraph; no “, , Qty –” residues).
4. Exports PDF using Word’s `ExportAsFixedFormat` (print-ready).
5. Saves the filled DOCX.

---

---

## Appendix — `claim_template_payload.json` vs `claim_values.json`

### What’s the difference?

- **`claim_template_payload.json`**
  - Produced by: `vision_claim_extractor.py`
  - Purpose: a **structured extraction payload** (rich, diagnostic, may contain extra fields)
  - Repo policy: keep one **reference sample** committed for reproducibility.

- **`claim_values.json`**
  - Consumed by: `run_claim_word_com.py`
  - Purpose: **exact placeholder value map** (keys match placeholders without braces)
  - Repo policy: treat as **per-claim input**, usually not committed (store samples in `examples/`).

### How to create / refresh `claim_template_payload.json`
1) Run the extractor on a known input set.
2) Save the JSON output as `claim_template_payload.json`.
3) Commit only when intentionally updating the reference example.

### How to create `claim_values.json`
Create a key/value JSON where each key matches a placeholder name **without** braces.

Example:
```json
{
  "PATIENT_NAME": "ANURADHA TYAGI",
  "ECHS_CARD_NO": "DL2000008879968",
  "INVOICE_NO.": "A041997",
  "...": "..."
}
```


## 6) Troubleshooting
- If `pywin32` is installed but COM fails, run:
  ```bat
  py -m pywin32_postinstall -install
  ```
- If output PDF is blank, ensure Word is installed and launches normally.