"""
Generates personalized certificate PNGs from a PPTX template.

Flow:
  1. Open the PPTX template
  2. Replace "Name" placeholder with the donor's actual name
  3. Upload to CloudConvert API to convert PPTX → PNG
  4. Download and return the PNG bytes

Required environment variables:
  - CLOUDCONVERT_API_KEY
"""

import os
import time
import logging
import tempfile
import requests

from pptx import Presentation

TEMPLATE_DIR = os.path.dirname(os.path.abspath(__file__))
CERTIFICATE_TEMPLATE = os.path.join(
    TEMPLATE_DIR, "Blue and Gold Modern Achievement Certificate.pptx"
)

CLOUDCONVERT_API = "https://api.cloudconvert.com/v2"


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
    Generate a personalized certificate PNG using CloudConvert.

    Args:
        donor_name: The donor's full name to place on the certificate.
        date: The donation date string (unused for this template).

    Returns:
        bytes: The PNG image data of the personalized certificate.
    """
    api_key = os.getenv("CLOUDCONVERT_API_KEY")
    if not api_key:
        raise ValueError("CLOUDCONVERT_API_KEY environment variable is not set")

    headers = {"Authorization": f"Bearer {api_key}"}

    # Step 1: Replace text in PPTX
    prs = Presentation(CERTIFICATE_TEMPLATE)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            _replace_text_preserving_format(shape, "Name", donor_name)

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        tmp_path = tmp.name
    prs.save(tmp_path)
    logging.info(f"Certificate PPTX created for: {donor_name}")

    try:
        # Step 2: Create CloudConvert job (upload + convert + export)
        job_resp = requests.post(
            f"{CLOUDCONVERT_API}/jobs",
            headers={**headers, "Content-Type": "application/json"},
            json={
                "tasks": {
                    "upload": {
                        "operation": "import/upload",
                    },
                    "convert": {
                        "operation": "convert",
                        "input": ["upload"],
                        "output_format": "png",
                        "engine": "office",
                    },
                    "export": {
                        "operation": "export/url",
                        "input": ["convert"],
                    },
                },
            },
        )
        job_resp.raise_for_status()
        job = job_resp.json()["data"]

        # Step 3: Find upload task and upload the PPTX
        upload_task = next(t for t in job["tasks"] if t["name"] == "upload")
        upload_url = upload_task["result"]["form"]["url"]
        upload_params = upload_task["result"]["form"]["parameters"]

        with open(tmp_path, "rb") as f:
            upload_resp = requests.post(
                upload_url,
                data=upload_params,
                files={"file": ("certificate.pptx", f,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
            )
        upload_resp.raise_for_status()
        logging.info("PPTX uploaded to CloudConvert")

        # Step 4: Poll job until complete
        job_id = job["id"]
        for attempt in range(60):
            status_resp = requests.get(
                f"{CLOUDCONVERT_API}/jobs/{job_id}", headers=headers
            )
            status_resp.raise_for_status()
            job_data = status_resp.json()["data"]

            if job_data["status"] == "finished":
                break
            elif job_data["status"] == "error":
                raise RuntimeError(f"CloudConvert job failed: {job_data}")

            time.sleep(2)
        else:
            raise TimeoutError("CloudConvert job did not complete in time")

        # Step 5: Download the PNG
        export_task = next(t for t in job_data["tasks"] if t["name"] == "export")
        download_url = export_task["result"]["files"][0]["url"]

        png_resp = requests.get(download_url)
        png_resp.raise_for_status()

        logging.info(f"Certificate PNG downloaded ({len(png_resp.content)} bytes)")
        return png_resp.content

    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
