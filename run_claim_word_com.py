import argparse
import json
import re
from pathlib import Path

import win32com.client as win32
import pywintypes

# Word constants (avoid depending on win32com.constants)
WD_FIND_STOP = 0
WD_REPLACE_ALL = 2
WD_EXPORT_FORMAT_PDF = 17
WD_MAIN_TEXT_STORY = 1  # the editable document body


def word_find_replace(doc, find_text: str, replace_text: str) -> None:
    """Use Word's native Find/Replace to preserve formatting and layout."""
    find = doc.Content.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Execute(
        FindText=find_text,
        MatchCase=False,
        MatchWholeWord=False,
        MatchWildcards=False,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Forward=True,
        Wrap=WD_FIND_STOP,
        Format=False,
        ReplaceWith=replace_text,
        Replace=WD_REPLACE_ALL,
    )


def iter_main_story_paragraphs(doc):
    """
    Yield paragraphs only from the main editable story.
    Avoids 'Cannot edit Range.' from headers/footers/shapes/textboxes.
    """
    rng = doc.StoryRanges(WD_MAIN_TEXT_STORY)
    while rng is not None:
        for p in rng.Paragraphs:
            yield p
        rng = rng.NextStoryRange


def safe_delete_paragraph(p):
    """
    Delete a paragraph safely. If Word blocks deletion (protected/non-editable range),
    skip it.
    """
    try:
        p.Range.Delete()
        return True
    except pywintypes.com_error:
        return False


def delete_empty_medicine_paragraphs(doc, values: dict) -> None:
    """
    Delete the *entire paragraph* for MED_3..MED_5 (and AMT_3..AMT_5) when empty.

    IMPORTANT:
      - Must run BEFORE placeholder replacement (while {{MED_n}} tokens still exist).
      - Runs ONLY on the main story to avoid 'Cannot edit Range.' exceptions.
    """
    empty_med_ns = [n for n in (3, 4, 5) if not (values.get(f"MED_{n}") or "").strip()]

    empty_tokens = []
    for n in empty_med_ns:
        empty_tokens += [
            f"{{{{MED_{n}}}}}",
            f"{{{{FORM_MED_{n}}}}}",
            f"{{{{QTY_MED_{n}}}}}",
            f"{{{{AMT_{n}}}}}",
        ]

    for n in (3, 4, 5):
        if not (values.get(f"AMT_{n}") or "").strip():
            empty_tokens += [f"{{{{AMT_{n}}}}}"]

    if not empty_tokens:
        return

    # Iterate only main-story paragraphs
    for p in list(iter_main_story_paragraphs(doc)):
        txt = p.Range.Text
        if any(tok in txt for tok in empty_tokens):
            safe_delete_paragraph(p)


def build_placeholder_map(values: dict) -> dict:
    """
    Build placeholder map with Qty label injected in code,
    INCLUDING the leading comma.

      {{QTY_MED_1}} -> ', Qty – 30'
      {{QTY_MED_3}} -> ''  (empty; paragraph already deleted)
    """
    placeholder_map = {}

    for k, v in values.items():
        if v is None:
            v = ""
        v = str(v)

        if re.fullmatch(r"QTY_MED_\d+", k):
            if v.strip():
                v = f", Qty – {v.strip()}"
            else:
                v = ""

        placeholder_map[f"{{{{{k}}}}}"] = v

    return placeholder_map



def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--values", required=True)
    ap.add_argument("--outdocx", required=True)
    ap.add_argument("--outpdf", required=True)
    args = ap.parse_args()

    template_path = Path(args.template).resolve()
    values_path = Path(args.values).resolve()
    outdocx = Path(args.outdocx).resolve()
    outpdf = Path(args.outpdf).resolve()
    outdocx.parent.mkdir(parents=True, exist_ok=True)
    outpdf.parent.mkdir(parents=True, exist_ok=True)

    values = json.loads(values_path.read_text(encoding="utf-8"))

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = None
    try:
        doc = word.Documents.Open(str(template_path))

        # 1) Delete empty MED_3..5 paragraphs BEFORE replacement
        delete_empty_medicine_paragraphs(doc, values)

        # 2) Build placeholder map (adds 'Qty –' in code for QTY_MED_n)
        placeholder_map = build_placeholder_map(values)

        # 3) Replace placeholders
        for ph in sorted(placeholder_map.keys(), key=len, reverse=True):
            word_find_replace(doc, ph, placeholder_map[ph])

        doc.SaveAs(str(outdocx))
        doc.ExportAsFixedFormat(str(outpdf), WD_EXPORT_FORMAT_PDF)

    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        word.Quit()


if __name__ == "__main__":
    main()
