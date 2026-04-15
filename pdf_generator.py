"""
Generates a receipt PDF from format.docx template.

Replaces custom fields (donor name, amount, date) and highlights them in yellow,
then converts the filled DOCX to PDF.
"""

import os
import copy
import subprocess
import platform
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn

# Placeholder values in format.docx that will be replaced
PLACEHOLDERS = {
    "Jacqueline Genevieve Hughes": "donor_name",
    "$1,000.00": "amount",
    "March 27, 2026": "date",
}

HIGHLIGHT_COLOR = "yellow"  # Word highlight color name


def set_highlight(run, color="yellow"):
    """Apply a highlight color to a run element."""
    # Map color names to Word highlight color IDs
    color_map = {
        "yellow": "7",
        "green": "4",
        "cyan": "5",
        "magenta": "6",
        "blue": "9",
        "red": "10",
        "darkBlue": "8",
        "darkCyan": "3",
        "darkGreen": "11",
        "darkMagenta": "12",
        "darkRed": "13",
        "darkYellow": "14",
        "lightGray": "15",
        "black": "16",
    }
    rPr = run._element.get_or_add_rPr()
    highlight = rPr.find(qn("w:highlight"))
    if highlight is None:
        highlight = copy.deepcopy(run._element.makeelement(qn("w:highlight"), {}))
        rPr.append(highlight)
    highlight.set(qn("w:val"), color)


def replace_and_highlight(doc, placeholder, replacement):
    """
    Find placeholder text in paragraphs and replace it with the
    replacement value, applying yellow highlight to the replaced text.
    """
    for paragraph in doc.paragraphs:
        full_text = paragraph.text
        if placeholder not in full_text:
            continue

        # Rebuild the paragraph runs with the replacement
        _replace_in_paragraph(paragraph, placeholder, replacement)

    # Also check inside table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if placeholder in paragraph.text:
                        _replace_in_paragraph(paragraph, placeholder, replacement)


def _replace_in_paragraph(paragraph, placeholder, replacement):
    """Replace placeholder text across runs in a paragraph, preserving formatting."""
    # Collect all run texts and their positions
    runs = paragraph.runs
    if not runs:
        return

    # Build a map of character positions to runs
    full_text = ""
    run_map = []  # (start_idx, end_idx, run_index)
    for i, run in enumerate(runs):
        start = len(full_text)
        full_text += run.text
        end = len(full_text)
        run_map.append((start, end, i))

    idx = full_text.find(placeholder)
    if idx == -1:
        return

    # Find which runs are affected
    ph_start = idx
    ph_end = idx + len(placeholder)

    # Clear the placeholder text from all affected runs
    for start, end, run_idx in run_map:
        run = runs[run_idx]
        r_start = max(ph_start, start) - start
        r_end = min(ph_end, end) - start

        if r_start < r_end:
            original = run.text
            run.text = original[:r_start] + original[r_end:]

    # Insert replacement text in the first affected run with yellow highlight
    for start, end, run_idx in run_map:
        if start <= ph_start < end:
            run = runs[run_idx]
            insert_pos = ph_start - start
            run.text = run.text[:insert_pos] + replacement + run.text[insert_pos:]
            set_highlight(run, HIGHLIGHT_COLOR)
            break


def convert_docx_to_pdf(docx_path, output_dir):
    """Convert a DOCX file to PDF using LibreOffice."""
    system = platform.system()

    if system == "Windows":
        lo_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        soffice = None
        for p in lo_paths:
            if os.path.exists(p):
                soffice = p
                break
        if not soffice:
            soffice = "soffice"
    else:
        soffice = "libreoffice"

    # Create an isolated user profile so LibreOffice doesn't conflict
    user_profile = os.path.join(output_dir, "libreoffice_profile")
    os.makedirs(user_profile, exist_ok=True)

    env = os.environ.copy()
    env["HOME"] = "/tmp"

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nofirststartwizard",
        f"-env:UserInstallation=file://{user_profile}",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path,
    ]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120, env=env)

    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

    if not os.path.exists(pdf_path):
        raise RuntimeError(
            f"LibreOffice conversion failed (exit code {result.returncode}):\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}\n"
            f"cmd: {' '.join(cmd)}"
        )

    return pdf_path


def generate_receipt_pdf(template_path, output_dir, donor_name, amount, date):
    """
    Generate a personalized receipt PDF:
    1. Open format.docx template
    2. Replace placeholders with actual values and highlight in yellow
    3. Convert to PDF
    """
    doc = Document(template_path)

    replacements = {
        "Jacqueline Genevieve Hughes": donor_name,
        "$1,000.00": amount,
        "March 27, 2026": date,
    }

    for placeholder, value in replacements.items():
        replace_and_highlight(doc, placeholder, value)

    filled_docx = os.path.join(output_dir, "receipt.docx")
    doc.save(filled_docx)

    pdf_path = convert_docx_to_pdf(filled_docx, output_dir)
    return pdf_path
