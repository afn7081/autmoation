"""
Generates personalized certificate PNGs from a PPTX template.

Flow:
  1. Open the PPTX template
  2. Replace "Name" placeholder with the donor's actual name
  3. Upload to ConvertAPI to convert PPTX → PNG
  4. Download and return the PNG bytes

Required environment variables:
  - CONVERTAPI_SECRET
"""

import os
import logging
import tempfile
import requests

from pptx import Presentation

TEMPLATE_DIR = os.path.dirname(os.path.abspath(__file__))
CERTIFICATE_TEMPLATE = os.path.join(
    TEMPLATE_DIR, "cert_template.pptx"
)


def _replace_text_preserving_format(shape, old_text, new_text):
    """Replace text in a shape while preserving font formatting."""
    for para in shape.text_frame.paragraphs:
        if old_text not in para.text:
            continue
        for run in para.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                return True
    return False


def generate_certificate_png(donor_name, date=""):
    """
    Generate a personalized certificate PNG using ConvertAPI.

    Args:
        donor_name: The donor's full name to place on the certificate.
        date: The donation date string (e.g. "April 16, 2026").

    Returns:
        bytes: The PNG image data of the personalized certificate.
    """
    api_secret = os.getenv("CONVERTAPI_SECRET")
    if not api_secret:
        raise ValueError("CONVERTAPI_SECRET environment variable is not set")

    # Format date to MM-DD-YY to match certificate style
    if not date:
        from datetime import datetime, timezone
        formatted_date = datetime.now(timezone.utc).strftime("%m-%d-%y")
    else:
        try:
            from datetime import datetime
            parsed = datetime.strptime(date, "%B %d, %Y")
            formatted_date = parsed.strftime("%m-%d-%y")
        except ValueError:
            formatted_date = date

    # Step 1: Verify template exists and is readable
    if not os.path.exists(CERTIFICATE_TEMPLATE):
        raise FileNotFoundError(f"Certificate template not found: {CERTIFICATE_TEMPLATE}")

    file_size = os.path.getsize(CERTIFICATE_TEMPLATE)
    logging.info(f"Template file: {CERTIFICATE_TEMPLATE} ({file_size} bytes)")

    # Step 2: Get a valid PPTX template (convert from legacy PPT if needed)
    import zipfile

    tpl_path = os.path.join(tempfile.gettempdir(), "cert_template_converted.pptx")

    # Use cached conversion if available
    if os.path.exists(tpl_path) and zipfile.is_zipfile(tpl_path):
        logging.info("Using cached converted PPTX template.")
    elif zipfile.is_zipfile(CERTIFICATE_TEMPLATE):
        # Already a valid PPTX — just copy
        with open(CERTIFICATE_TEMPLATE, "rb") as src, open(tpl_path, "wb") as dst:
            dst.write(src.read())
        logging.info("Template is valid PPTX, copied directly.")
    else:
        # Legacy OLE2 format — convert via ConvertAPI (ppt → pptx)
        logging.warning("Template is legacy PPT format. Converting via ConvertAPI...")
        with open(CERTIFICATE_TEMPLATE, "rb") as f:
            conv_resp = requests.post(
                f"https://v2.convertapi.com/convert/ppt/to/pptx?Secret={api_secret}",
                files={"File": ("template.ppt", f, "application/vnd.ms-powerpoint")},
                timeout=120,
            )
        if conv_resp.status_code != 200:
            raise RuntimeError(f"ConvertAPI PPT→PPTX failed ({conv_resp.status_code}): {conv_resp.text}")

        conv_result = conv_resp.json()
        conv_files = conv_result.get("Files", [])
        if not conv_files:
            raise RuntimeError(f"ConvertAPI PPT→PPTX returned no files: {conv_result}")

        # Download the converted PPTX
        pptx_url = conv_files[0]["Url"]
        pptx_resp = requests.get(pptx_url, timeout=60)
        pptx_resp.raise_for_status()

        with open(tpl_path, "wb") as dst:
            dst.write(pptx_resp.content)

        if not zipfile.is_zipfile(tpl_path):
            raise RuntimeError("ConvertAPI conversion result is not a valid PPTX")
        logging.info(f"Template converted to PPTX successfully ({len(pptx_resp.content)} bytes).")

    # Make a working copy for this invocation
    work_path = os.path.join(tempfile.gettempdir(), "cert_input.pptx")
    with open(tpl_path, "rb") as src, open(work_path, "wb") as dst:
        dst.write(src.read())

    prs = Presentation(work_path)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            _replace_text_preserving_format(shape, "Full Name", donor_name)
            _replace_text_preserving_format(shape, "Date :  04-03-26", f"Date :  {formatted_date}")

    tmp_path = os.path.join(tempfile.gettempdir(), "cert_output.pptx")
    prs.save(tmp_path)
    logging.info(f"Certificate PPTX created for: {donor_name}")

    try:
        # Step 3: Upload to ConvertAPI and convert PPTX → PNG
        with open(tmp_path, "rb") as f:
            resp = requests.post(
                f"https://v2.convertapi.com/convert/pptx/to/png?Secret={api_secret}",
                files={"File": ("certificate.pptx", f,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
                timeout=60,
            )

        if resp.status_code != 200:
            raise RuntimeError(f"ConvertAPI failed ({resp.status_code}): {resp.text}")

        result = resp.json()
        files = result.get("Files", [])
        if not files:
            raise RuntimeError(f"ConvertAPI returned no files: {result}")

        # Step 3: Download the first slide PNG
        png_url = files[0]["Url"]
        png_resp = requests.get(png_url, timeout=30)
        png_resp.raise_for_status()

        logging.info(f"Certificate PNG downloaded ({len(png_resp.content)} bytes)")
        return png_resp.content

    finally:
        for p in [tmp_path, work_path]:
            if os.path.exists(p):
                os.remove(p)
