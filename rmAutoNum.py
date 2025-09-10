import sys
import os
import time
import win32com.client as win32
from win32com.client import constants

def is_question_paragraph(p):
    """
    Heuristic to identify question lines:
    - It is part of a numbered list (auto-numbered)
    - It is at top level (LevelNumber == 1)
    - The visible text does not start with 'Answer' or 'Explanation'
    Adjust as needed for the question set format.
    """
    text = p.Range.Text.strip()
    if text.lower().startswith("answer") or text.lower().startswith("explanation"):
        return False
    lf = p.Range.ListFormat
    try:
        # ListType will be non-zero for numbered/bulleted; LevelNumber is 1-based
        is_list = lf.ListType != constants.wdListNoNumbering
        is_numbered = lf.ListType in (
            constants.wdListOutlineNumbering,
            constants.wdListSimpleNumbering
        ) or lf.ListType == constants.wdListListNumOnly
        at_top_level = (lf.ListLevelNumber == 1)
        return is_list and is_numbered and at_top_level
    except Exception:
        return False

def convert_questions_to_text(doc):
    """
    Convert only top-level auto-numbered question paragraphs to literal text numbers
    and remove indentation for those paragraphs.
    """
    paragraphs = doc.Paragraphs
    # Collect candidates first to avoid collection mutation issues
    targets = []
    for i in range(1, paragraphs.Count + 1):
        p = paragraphs.Item(i)
        if is_question_paragraph(p):
            targets.append(p)

    for p in targets:
        # Convert only this paragraph's numbering to text
        try:
            p.Range.ListFormat.ConvertNumbersToText()
        except Exception:
            # If already plain text or not applicable, ignore
            pass

        # Remove left indent and first-line indent
        try:
            p.LeftIndent = 0
            p.FirstLineIndent = 0
        except Exception:
            pass

        # Optional: ensure a single space after the number if Word produced "21." without space
        # This looks for a leading pattern like "21." or "21)"; add a space if next char is a letter/digit
        rng = p.Range
        txt = rng.Text
        # Work only on the paragraph text (ignore end-of-paragraph mark)
        core = txt[:-1] if txt.endswith('\r') else txt
        import re
        m = re.match(r'^(\d+[\.\)]\s?)(.+)$', core)
        if m:
            lead, rest = m.groups()
            if not lead.endswith(' '):
                core = lead + ' ' + rest
            # Write back preserving paragraph end
            rng.Text = core + ('\r' if txt.endswith('\r') else '')

def main(input_path, output_path=None, visible=False):
    if not os.path.exists(input_path):
        print(f"File not found: {input_path}")
        sys.exit(1)

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = visible

    try:
        doc = word.Documents.Open(os.path.abspath(input_path))
        # Ensure list fields are up-to-date
        try:
            doc.Fields.Update()
        except Exception:
            pass

        convert_questions_to_text(doc)

        # Save output
        if not output_path:
            root, ext = os.path.splitext(input_path)
            output_path = root + "_numbered_plain.docx"
        doc.SaveAs(os.path.abspath(output_path), FileFormat=constants.wdFormatXMLDocument)
        print(f"Saved: {output_path}")

    finally:
        try:
            doc.Close(SaveChanges=False)
        except Exception:
            pass
        # Give Word a moment before quitting, avoids COM teardown race
        time.sleep(0.2)
        word.Quit()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python rmAutoNum.py <input.docx> [output.docx]")
        # Interactive fallback when no args are provided
        try:
            inp = input("Enter path to input .docx (or press Enter to exit): ").strip().strip('"')
        except EOFError:
            inp = ""
        if not inp:
            sys.exit(1)
        try:
            outp = input("Optional output path (press Enter to auto-name): ").strip().strip('"')
        except EOFError:
            outp = ""
        main(inp, outp if outp else None)
    else:
        input_path = sys.argv[1]
        output_path = sys.argv[2] if len(sys.argv) > 2 else None
        main(input_path, output_path)
