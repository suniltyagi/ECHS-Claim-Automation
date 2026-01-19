\# ECHS Claim Automation (Vision-based)



This repository contains a \*\*single-script, vision-first pipeline\*\* to automate

ECHS medical claim preparation using \*\*GPT-4o Vision\*\*.



\## What it does

\- Takes a \*\*medicine bill image\*\* and \*\*ECHS prescription image\*\*

\- Uses GPT-4o Vision to extract structured data (JSON)

\- Fills the official \*\*ECHS claim DOCX template\*\*



No OCR engines, no regex parsing, no Google APIs.



\## Files

\- `vision\_claim\_extractor.py` – main script (one call, one output)

\- `bill.jpeg` – sample medicine bill image

\- `prescription.jpeg` – sample prescription image

\- `ECHS\_Claim\_template.docx` – claim template with placeholders

\- `ECHS\_Claim\_filled.docx` – generated claim

\- `claim\_template\_payload.json` – extracted structured data

\- `place\_holders\_list.docx` – template ↔ JSON contract



\## Requirements

\- Python 3.9+

\- `openai` Python SDK

\- Environment variable `OPENAI\_API\_KEY` set



\## Usage

```bash

python vision\_claim\_extractor.py



