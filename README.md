# ECHS Claim Automation

👉 **Start here:** [ECHS_COM_AUTOMATION_GUIDE.md](ECHS_COM_AUTOMATION_GUIDE.md)

This repository automates generation of **ECHS Contingent Bill (Appx ‘A’, IAFA-155)** documents with a strict, non-negotiable constraint:

> **Formatting fidelity is more important than automation cleverness.**

To guarantee exact underline, bold, spacing, and line-flow fidelity, the final
document is rendered by **Microsoft Word itself**, via **Windows COM automation**.

---

## What this repository does (v1.0 scope)

- Takes **human-validated claim data** in `claim_values.json`
- Injects values into an **authoritative Word template**
- Produces **submission-ready DOCX and PDF**
- Ensures the output is **visually indistinguishable** from approved ECHS samples

There is **no OCR or vision extraction** in the current, supported workflow.

---

## Core files

- **`ECHS_Claim_template.docx`**  
  Authoritative ECHS claim template.  
  Formatting is locked; placeholders are embedded.

- **`place_holders_list.docx`**  
  Final, approved list of placeholders.  
  No new placeholders may be invented.

- **`run_claim_word_com.py`**  
  Canonical Word COM automation script.  
  This is the **only execution entry point**.

- **`examples/claim_values.sample.json`**  
  Sample input showing the expected structure of `claim_values.json`.

- **`out/`** *(gitignored)*  
  Generated DOCX/PDF outputs.

---

## `claim_values.json` (single source of truth)

`claim_values.json` is the **only input file** consumed by the automation.

- Keys match placeholder names **without braces**
- Values are **human-verified**
- One file per claim
- Typically **not committed** (samples live in `examples/`)

The provenance, rationale, and creation process for `claim_values.json`
are documented **in detail** in:

👉 [ECHS_COM_AUTOMATION_GUIDE.md](ECHS_COM_AUTOMATION_GUIDE.md)

---

## End-to-end workflow (canonical)

1. **Create `claim_values.json`**
   - By manual, human-validated reading of bill + prescription
   - Using `examples/claim_values.sample.json` as a starting point

2. **Render**
   ```bat
   py run_claim_word_com.py --template ECHS_Claim_template.docx --values claim_values.json --outdocx out\ECHS_Claim_filled.docx --outpdf out\ECHS_Claim_filled.pdf
