import sys
import base64
import json
import os
from pathlib import Path
from docx import Document
from openai import OpenAI


# ---------- CONFIG ----------
BASE = Path(r"C:\Users\Admin\ECHS_claims")

if len(sys.argv) != 3:
    print("Usage: python vision_claim_extractor.py <bill_image> <prescription_image>")
    sys.exit(1)

BILL_IMG = Path(sys.argv[1])
PRES_IMG = Path(sys.argv[2])


# BILL_IMG = BASE / "bill.jpeg"
# PRES_IMG = BASE / "prescription.jpeg"
TEMPLATE_DOCX = BASE / "ECHS_Claim_template.docx"
OUT_JSON = BASE / "claim_template_payload.json"
OUT_DOCX = BASE / "ECHS_Claim_filled.docx"

MODEL = "gpt-4o"
# ----------------------------


def b64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def main():
    client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

    bill_b64 = b64(BILL_IMG)
    pres_b64 = b64(PRES_IMG)

    prompt = """
You are given TWO images:
1) A pharmacy medicine bill
2) An ECHS polyclinic prescription

TASK:
Extract data and return ONE JSON object ONLY.

RULES (MANDATORY):
- No markdown
- No explanations
- No missing keys
- Use empty string "" if a field is not present

DATE FORMAT:
- DD-MM-YYYY

TOTALS:
- TOTAL_WO_DISCOUNT = SUB TOTAL (before discount)
- TOTAL_AMOUNT = GRAND TOTAL / PAYABLE
- Both with 2 decimals

MEDICINES:
- Only medicines actually PURCHASED in the BILL
- Max 5 medicines
- FORM_MED_i must be TAB or CAP or ""
- QTY_MED_i must be numeric string (e.g. "30")
- AMT_i must be 2-decimal string (e.g. "192.00")

OUTPUT JSON KEYS (MUST MATCH EXACTLY):

{
  "PATIENT_NAME": "",
  "ECHS_CARD_NO": "",
  "DIAGNOSIS": "",
  "INVOICE_NO.": "",
  "DATE": "",
  "DATE_EXPENDITURE": "",
  "CURRENT_MONTH_YEAR": "",
  "TOTAL_WO_DISCOUNT": "",
  "TOTAL_AMOUNT": "",
  "AMOUNT_WORDS": "",
  "MED_1": "", "FORM_MED_1": "", "QTY_MED_1": "", "AMT_1": "",
  "MED_2": "", "FORM_MED_2": "", "QTY_MED_2": "", "AMT_2": "",
  "MED_3": "", "FORM_MED_3": "", "QTY_MED_3": "", "AMT_3": "",
  "MED_4": "", "FORM_MED_4": "", "QTY_MED_4": "", "AMT_4": "",
  "MED_5": "", "FORM_MED_5": "", "QTY_MED_5": "", "AMT_5": ""
}
""".strip()

    response = client.responses.create(
        model=MODEL,
        temperature=0.0,
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                    {"type": "input_image", "image_url": f"data:image/jpeg;base64,{bill_b64}"},
                    {"type": "input_image", "image_url": f"data:image/jpeg;base64,{pres_b64}"}
                ],
            }
        ],
    )

    text = response.output[0].content[0].text.strip()
    text = text[text.find("{"): text.rfind("}") + 1]
    data = json.loads(text)

    OUT_JSON.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"✅ JSON written: {OUT_JSON}")

    fill_docx(data)


def fill_docx(data: dict):
    doc = Document(TEMPLATE_DOCX)

    def replace(p):
        for k, v in data.items():
            p.text = p.text.replace(f"{{{{{k}}}}}", v)

    for p in doc.paragraphs:
        replace(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace(p)

    doc.save(OUT_DOCX)
    print(f"✅ DOCX created: {OUT_DOCX}")


if __name__ == "__main__":
    main()
