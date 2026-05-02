"""
Sends personalized donation emails via SendGrid with PDF receipt attached.
"""

import os
import base64
import logging
import zipfile
import mimetypes

import requests
from docx import Document
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Mail,
    Attachment,
    FileContent,
    FileName,
    FileType,
    Disposition,
    ContentId,
)


def _convert_svg_to_png_convertapi(svg_path: str) -> bytes:
    """Convert an SVG file to PNG bytes via ConvertAPI."""
    api_secret = os.getenv("CONVERTAPI_SECRET")
    if not api_secret:
        raise ValueError("CONVERTAPI_SECRET is required for SVG->PNG conversion")

    with open(svg_path, "rb") as f:
        resp = requests.post(
            f"https://v2.convertapi.com/convert/svg/to/png?Secret={api_secret}&StoreFile=true",
            files={"File": ("header.svg", f, "image/svg+xml")},
            timeout=120,
        )
    if resp.status_code != 200:
        raise RuntimeError(f"ConvertAPI svg->png failed ({resp.status_code}): {resp.text}")

    files = (resp.json().get("Files") or [])
    if not files:
        raise RuntimeError(f"ConvertAPI returned no files: {resp.text}")
    first = files[0]
    url = first.get("Url") or first.get("FileUrl")
    if url:
        r = requests.get(url, timeout=60)
        r.raise_for_status()
        return r.content
    data = first.get("FileData")
    if not data:
        raise RuntimeError(f"ConvertAPI response missing Url/FileData: {list(first.keys())}")
    return base64.b64decode(data)


def _shrink_png_for_email(png_bytes: bytes, max_width: int = 1200) -> bytes:
    """Resize a PNG so its width <= max_width and re-encode for small file size.
    Returns original bytes if Pillow isn't available or anything fails."""
    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(png_bytes))
        if img.width > max_width:
            ratio = max_width / img.width
            new_size = (max_width, int(img.height * ratio))
            img = img.resize(new_size, Image.LANCZOS)
        out = io.BytesIO()
        # Convert to RGB so we can save as JPEG (much smaller for photos/banners)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        img.save(out, format="JPEG", quality=85, optimize=True)
        return out.getvalue()
    except Exception as e:
        logging.warning(f"Could not shrink header image: {e}")
        return png_bytes


def get_header_image():
    """Return (image_bytes, mime, filename) for the email header.

    Prefers an SVG in the working directory (converted to PNG via ConvertAPI
    once, then resized + JPEG-compressed for email-friendly size, cached to
    disk). Falls back to None if no SVG is present.
    """
    here = os.path.dirname(os.path.abspath(__file__))
    svg_candidates = [
        f for f in os.listdir(here)
        if f.lower().endswith(".svg") and "header" in f.lower()
    ]
    if not svg_candidates:
        svg_candidates = [f for f in os.listdir(here) if f.lower().endswith(".svg")]
    if not svg_candidates:
        return None

    svg_path = os.path.join(here, svg_candidates[0])
    cache_path = os.path.join(here, ".header_cache.jpg")

    # Re-build cache if missing or older than the SVG
    if (not os.path.exists(cache_path)
            or os.path.getmtime(cache_path) < os.path.getmtime(svg_path)):
        try:
            logging.info(f"Converting header SVG -> PNG via ConvertAPI: {svg_candidates[0]}")
            png = _convert_svg_to_png_convertapi(svg_path)
            jpg = _shrink_png_for_email(png, max_width=1200)
            with open(cache_path, "wb") as f:
                f.write(jpg)
            logging.info(f"Header cached: {len(jpg)} bytes (from {len(png)} bytes PNG)")
        except Exception as e:
            logging.warning(f"Header SVG->PNG conversion failed: {e}")
            if not os.path.exists(cache_path):
                return None

    with open(cache_path, "rb") as f:
        return f.read(), "image/jpeg", "header.jpg"


def extract_first_image_from_docx(template_path):
    """Return (image_bytes, mime_type, filename) for the first image embedded
    in the given .docx, or None if there is no image."""
    try:
        with zipfile.ZipFile(template_path) as z:
            media = sorted(
                n for n in z.namelist()
                if n.startswith("word/media/")
                and n.lower().rsplit(".", 1)[-1] in ("png", "jpg", "jpeg", "gif")
            )
            if not media:
                return None
            name = media[0]
            data = z.read(name)
            mime, _ = mimetypes.guess_type(name)
            return data, mime or "image/jpeg", os.path.basename(name)
    except Exception as e:
        logging.warning(f"Could not extract header image from {template_path}: {e}")
        return None


def read_email_template(template_path, donor_name):
    """
    Read a .docx email template and return it as plain text + styled HTML,
    preserving bold, sections, and layout from the original document.
    """
    doc = Document(template_path)
    first_name = donor_name.split()[0] if donor_name else "Friend"

    plain_lines = []
    body_parts = []

    for para in doc.paragraphs:
        text = para.text.replace("First Name", first_name)
        plain_lines.append(text)

        if not text.strip():
            body_parts.append("<br>")
            continue

        # Check if paragraph has bold runs
        has_bold = any(r.bold for r in para.runs if r.text.strip())

        # Build run-level HTML preserving bold/normal within a paragraph
        run_html = ""
        for run in para.runs:
            rt = run.text.replace("First Name", first_name)
            if not rt:
                continue
            if run.bold:
                run_html += f"<strong>{rt}</strong>"
            else:
                run_html += rt

        if not run_html:
            run_html = text

        # Section headers (bold standalone lines that are section titles)
        is_section_header = has_bold and all(
            r.bold for r in para.runs if r.text.strip()
        ) and any(
            kw in text for kw in [
                "Why Your Support", "Stay Connected", "FOUNDING MEMBER",
                "Help Us Carry", "From Streets"
            ]
        )

        if is_section_header:
            body_parts.append(
                f'<h3 style="color:#1a3a5c; font-size:16px; margin:24px 0 8px 0; '
                f'border-bottom:2px solid #e8e8e8; padding-bottom:6px;">{run_html}</h3>'
            )
        elif text.startswith("•"):
            body_parts.append(
                f'<li style="margin:4px 0 4px 20px; font-size:15px;">{run_html[1:].strip()}</li>'
            )
        elif "🔗" in text:
            parts = text.split("🔗")
            if len(parts) > 1:
                url = parts[1].strip()
                body_parts.append(
                    f'<p style="font-size:15px; margin:8px 0;">🔗 '
                    f'<a href="{url}" style="color:#2980b9; text-decoration:underline;">{url}</a></p>'
                )
            else:
                body_parts.append(f'<p style="font-size:15px; margin:8px 0;">{run_html}</p>')
        elif text.startswith("With gratitude,"):
            body_parts.append('<!-- CERTIFICATE_PLACEHOLDER -->')
            body_parts.append(f'<p style="font-size:15px; margin:24px 0 4px 0;">{run_html}</p>')
        elif any(kw in text for kw in ["Founder", "Executive Director", "501(c)", "EIN:", "UScension"]):
            body_parts.append(f'<p style="font-size:13px; color:#666; margin:2px 0;">{run_html}</p>')
        else:
            body_parts.append(f'<p style="font-size:15px; line-height:1.6; margin:8px 0;">{run_html}</p>')

    plain_text = "\n".join(plain_lines)

    html_content = f'''
    <!DOCTYPE html>
    <html><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      @media only screen and (max-width:600px) {{
        .email-container {{ width:100% !important; max-width:100% !important; }}
        .email-body, .email-footer {{ padding-left:18px !important; padding-right:18px !important; }}
        .header-img {{ width:100% !important; height:auto !important; }}
      }}
    </style></head>
    <body style="margin:0; padding:0; background-color:#f5f7fa;">
    <div class="email-container" style="max-width:600px; width:100%; margin:0 auto; font-family:Arial, Helvetica, sans-serif; color:#1a1a1a;">

        <!-- Header Banner (image extracted from .docx, or text fallback) -->
        <!-- HEADER_IMAGE_PLACEHOLDER -->

        <!-- Body -->
        <div class="email-body" style="background-color:#ffffff; padding:30px 30px 20px 30px; border-left:1px solid #e8e8e8; border-right:1px solid #e8e8e8;">
            {"".join(body_parts)}
        </div>

        <!-- Footer -->
        <div class="email-footer" style="background-color:#f5f7fa; padding:20px 30px; text-align:center; border-radius:0 0 8px 8px; border:1px solid #e8e8e8; border-top:none;">
            <p style="font-size:12px; color:#888; margin:4px 0;">
                Global Women Foundation &amp; Band of Brothers | 501(c)(3) Nonprofit
            </p>
            <p style="font-size:11px; color:#aaa; margin:4px 0;">
                100 Wilshire Blvd Suite 700 | Santa Monica, CA 90401 |
                <a href="https://www.gwfbob.org" style="color:#2980b9;">www.gwfbob.org</a>
            </p>
        </div>

    </div>
    </body></html>'''

    return plain_text, html_content


def send_donation_email(template_path, to_email, donor_name, amount, pdf_path, amount_dollars=0, date=""):
    """
    Send a donation thank-you email with the receipt PDF attached.
    For donations >= $1000, attaches a personalized certificate PNG from Canva.
    """
    api_key = os.getenv("SENDGRID_API_KEY")
    from_email = os.getenv("FROM_EMAIL", "info@gwfbob.org")
    from_name = os.getenv("FROM_NAME", "Global Women Foundation & Band of Brothers")

    if not api_key:
        raise ValueError("SENDGRID_API_KEY environment variable is not set")

    plain_text, html_content = read_email_template(template_path, donor_name)

    # Use a publicly hosted header image rather than a cid inline attachment.
    # Gmail (especially the mobile app) blocks inline cid images on messages
    # that fail DMARC alignment (sending a gmail.com From: via sendgrid.net),
    # so we serve the banner from a public URL that Gmail's image proxy can
    # fetch. This was the approach confirmed working previously.
    header_url = os.getenv(
        "HEADER_IMAGE_URL",
        "https://raw.githubusercontent.com/afn7081/autmoation/mai/email-header.jpg",
    )
    header_html = f'''
        <div style="width:100%; text-align:center; line-height:0; font-size:0;">
            <a href="https://www.gwfbob.org" style="text-decoration:none;">
              <img src="{header_url}"
                   alt="Global Women Foundation & Band of Brothers"
                   width="600"
                   class="header-img"
                   style="width:100%; max-width:600px; height:auto; display:block; margin:0 auto; border:0; outline:none; text-decoration:none;" />
            </a>
        </div>'''
    html_content = html_content.replace('<!-- HEADER_IMAGE_PLACEHOLDER -->', header_html)

    # Generate certificate for $1000+ donors and embed in email HTML
    cert_png = None
    if amount_dollars >= 1000:
        logging.info(f"Generating certificate for ${amount_dollars} donation to {donor_name}")
        try:
            from certificate_generator import generate_certificate_png

            cert_png = generate_certificate_png(donor_name, date)
            cert_img_html = '''
            <div style="margin:24px 0; text-align:center;">
                <img src="cid:certificate_image" alt="Certificate of Recognition"
                     style="max-width:100%; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.15);" />
            </div>'''
            html_content = html_content.replace('<!-- CERTIFICATE_PLACEHOLDER -->', cert_img_html)
            logging.info("Certificate image placeholder replaced in email HTML")
        except Exception as e:
            logging.error(f"Certificate generation failed for {donor_name}: {e}")
            raise RuntimeError(f"Certificate generation failed: {e}") from e

    first_name = donor_name.split()[0] if donor_name else "Friend"
    subject = f"Thank You for Your Generous Donation, {first_name}!"

    message = Mail(
        from_email=(from_email, from_name),
        to_emails=to_email,
        subject=subject,
        plain_text_content=plain_text,
        html_content=html_content,
    )

    # Attach the PDF receipt
    if pdf_path and os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()

        encoded_pdf = base64.b64encode(pdf_data).decode()
        attachment = Attachment(
            FileContent(encoded_pdf),
            FileName("Donation_Receipt.pdf"),
            FileType("application/pdf"),
            Disposition("attachment"),
        )
        message.add_attachment(attachment)

    # Header image is served via public URL (see header_html above), so no
    # cid attachment is needed.

    # Attach certificate inline + as download for $1000+ donors
    if cert_png:
        encoded_cert = base64.b64encode(cert_png).decode()

        # Inline image referenced by cid:certificate_image
        inline_attachment = Attachment(
            FileContent(encoded_cert),
            FileName("certificate.png"),
            FileType("image/png"),
            Disposition("inline"),
            ContentId("certificate_image"),
        )
        message.add_attachment(inline_attachment)

        # Also attach as downloadable file
        dl_attachment = Attachment(
            FileContent(encoded_cert),
            FileName("Certificate_of_Recognition.png"),
            FileType("image/png"),
            Disposition("attachment"),
        )
        message.add_attachment(dl_attachment)

        logging.info("Certificate embedded and attached successfully")

    sg = SendGridAPIClient(api_key)
    response = sg.send(message)

    return response.status_code
