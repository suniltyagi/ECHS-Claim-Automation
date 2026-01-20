# ECHS Claim Automation

👉 **Start here:** [ECHS_COM_AUTOMATION_GUIDE.md](ECHS_COM_AUTOMATION_GUIDE.md)

This repository automates generation of **ECHS Contingent Bill (Appx ‘A’, IAFA‑155)** documents with a hard constraint:

> **Formatting fidelity is more important than automation cleverness.**

To preserve underline/bold/line‑flow exactly, the final document is rendered by **Microsoft Word via COM automation**.

---

## Repository components

### A) Extraction (Vision → JSON)
- **Script:** `vision_claim_extractor.py`
- **Inputs:** bill image(s) + prescription image(s)
- **Output:** a structured JSON payload you can inspect and validate.

#### `claim_template_payload.json` (what it is)
`claim_template_payload.json` is a **sample payload** produced by the extraction step and committed as a reference.

It exists for two reasons:
1) **Debug / audit:** you can see exactly what the extractor produced for a known input.
2) **Contract:** it documents the expected shape of extracted data before mapping into template placeholders.

#### How `claim_template_payload.json` was created
It was created by running the extractor against a known bill + prescription set and saving the extractor’s JSON output to this filename.

Typical workflow:
1. Put your inputs (bill + prescription images) in the repo (or a local working folder).
2. Run:
   ```bat
   py vision_claim_extractor.py
   ```
3. Save/rename the produced JSON output as:
   - `claim_template_payload.json` (reference sample, committed), and/or
   - `claim_values.json` (current run inputs for COM rendering; usually not committed).

> Note: `claim_template_payload.json` is intended as a **reference sample**; do not overwrite it unless you are intentionally updating the reference example in the repo.

---

### B) Rendering (JSON → DOCX/PDF via Word COM)
- **Script:** `run_claim_word_com.py`
- **Input:** `claim_values.json` (values mapped to template placeholders)
- **Outputs:** `out/ECHS_Claim_filled.docx` and `out/ECHS_Claim_filled.pdf`

This step is the one that produces the **printable/submittable** artefacts.

---

## Key files
- `ECHS_Claim_template.docx` — authoritative template (placeholders embedded; formatting locked)
- `place_holders_list.docx` — authoritative placeholder list (do not invent new ones)
- `run_claim_word_com.py` — Word COM renderer (format‑safe)
- `vision_claim_extractor.py` — extraction script (bill/prescription → JSON)
- `claim_template_payload.json` — **reference** extraction payload sample
- `examples/claim_values.sample.json` — sample `claim_values.json` for the COM step (safe to commit)
- `out/` — outputs (gitignored)

---

## Recommended workflow (end-to-end)
1) **Extract** with `vision_claim_extractor.py` → inspect JSON
2) **Map** extracted JSON into `claim_values.json` (placeholder values)
3) **Render** with `run_claim_word_com.py` → DOCX + PDF in `out/`

For full Windows setup and commands, see the guide:
- [ECHS_COM_AUTOMATION_GUIDE.md](ECHS_COM_AUTOMATION_GUIDE.md)
