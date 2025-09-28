import os
import time
import re
import pythoncom
import win32com.client as win32
import win32com as win32com_root
import shutil
from win32com.client import dynamic
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


def _ensure_word_dispatch():
    """
    Create a Word COM automation object, handling a corrupted pywin32 gen_py cache.
    Tries EnsureDispatch first; if it fails with AttributeError caused by a bad
    generated cache (missing CLSIDToClassMap), it clears the cache and retries.
    Falls back to dynamic Dispatch as a last resort.
    """
    try:
        return win32.gencache.EnsureDispatch("Word.Application")
    except AttributeError:
        # First try to rebuild the cache programmatically
        try:
            win32.gencache.Rebuild()
            return win32.gencache.EnsureDispatch("Word.Application")
        except Exception:
            pass
        # Attempt to clear pywin32 generated cache directories and retry
        try:
            gen_paths = []
            try:
                gen_paths.append(getattr(win32com_root, "__gen_path__", None))
            except Exception:
                pass
            local = os.environ.get("LOCALAPPDATA")
            if local:
                gen_paths.append(os.path.join(local, "Temp", "gen_py"))
            for p in gen_paths:
                if p and os.path.isdir(p):
                    try:
                        shutil.rmtree(p, ignore_errors=True)
                    except Exception:
                        pass
            # Retry EnsureDispatch with a fresh cache
            try:
                return win32.gencache.EnsureDispatch("Word.Application")
            except Exception:
                pass
        except Exception:
            pass
        # Final fallback: dynamic dispatch avoids makepy/gen_py entirely
        return dynamic.Dispatch("Word.Application")


def convert_questions_to_text(doc, progress=None):
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

    total = len(targets)
    processed = 0
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

        # Ensure a single space after the number if Word produced "21." or "21)" without space
        rng = p.Range
        txt = rng.Text
        # Work only on the paragraph text (ignore end-of-paragraph mark)
        core = txt[:-1] if txt.endswith('\r') else txt
        m = re.match(r'^(\d+[\.\)]\s?)(.+)$', core)
        if m:
            lead, rest = m.groups()
            if not lead.endswith(' '):
                core = lead + ' ' + rest
            # Write back preserving paragraph end
            rng.Text = core + ('\r' if txt.endswith('\r') else '')

        processed += 1
        if progress and total > 0:
            try:
                pct = int((processed / total) * 100)
                progress(pct)
            except Exception:
                pass


def process_doc(input_path: str, output_path: str | None = None, visible: bool = False, progress=None) -> str:
    """
    Open the Word document at input_path, convert question numbering to plain text, remove indentation,
    and save the result to output_path (or auto-name if None). Returns the output path.
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File not found: {input_path}")

    # Initialize COM for this thread (Flask may serve requests in worker threads)
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = _ensure_word_dispatch()
        word.Visible = visible

        doc = word.Documents.Open(os.path.abspath(input_path))
        try:
            doc.Fields.Update()
        except Exception:
            pass

        convert_questions_to_text(doc, progress=progress)

        # Save output
        if not output_path:
            root, ext = os.path.splitext(input_path)
            output_path = root + "_numbered_plain.docx"
        doc.SaveAs(os.path.abspath(output_path), FileFormat=constants.wdFormatXMLDocument)
        return os.path.abspath(output_path)
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        # Ensure Word is quit if it was created
        try:
            if word is not None:
                time.sleep(0.2)
                word.Quit()
        except Exception:
            pass
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()
