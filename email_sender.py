"""
Sends personalized donation emails via SendGrid with PDF receipt attached.
"""

import os
import base64
import logging

from docx import Document
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Mail,
    Attachment,
    FileContent,
    FileName,
    FileType,
    Disposition,
)


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
    <div style="max-width:600px; margin:0 auto; font-family:Arial, Helvetica, sans-serif; color:#1a1a1a;">

        <!-- Header Banner -->
        <div style="background-color:#1a3a5c; padding:24px 30px; text-align:center; border-radius:8px 8px 0 0;">
            <h1 style="color:#ffffff; font-size:20px; margin:0; letter-spacing:0.5px;">
                Global Women Foundation &amp; Band of Brothers
            </h1>
            <p style="color:#a8c4dc; font-size:12px; margin:6px 0 0 0;">
                Ending Veteran Homelessness | Building Community
            </p>
        </div>

        <!-- Body -->
        <div style="background-color:#ffffff; padding:30px 30px 20px 30px; border-left:1px solid #e8e8e8; border-right:1px solid #e8e8e8;">
            {"".join(body_parts)}
        </div>

        <!-- Footer -->
        <div style="background-color:#f5f7fa; padding:20px 30px; text-align:center; border-radius:0 0 8px 8px; border:1px solid #e8e8e8; border-top:none;">
            <p style="font-size:12px; color:#888; margin:4px 0;">
                Global Women Foundation &amp; Band of Brothers | 501(c)(3) Nonprofit
            </p>
            <p style="font-size:11px; color:#aaa; margin:4px 0;">
                100 Wilshire Blvd Suite 700 | Santa Monica, CA 90401 |
                <a href="https://www.gwfbob.org" style="color:#2980b9;">www.gwfbob.org</a>
            </p>
        </div>

    </div>'''

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
        message.attachment = attachment

    # Attach certificate inline + as download for $1000+ donors
    if cert_png:
        encoded_cert = base64.b64encode(cert_png).decode()

        # Inline image referenced by cid:certificate_image
        inline_attachment = Attachment(
            FileContent(encoded_cert),
            FileName("certificate.png"),
            FileType("image/png"),
            Disposition("inline"),
        )
        inline_attachment.content_id = "certificate_image"
        message.attachment = inline_attachment

        # Also attach as downloadable file
        dl_attachment = Attachment(
            FileContent(encoded_cert),
            FileName("Certificate_of_Recognition.png"),
            FileType("image/png"),
            Disposition("attachment"),
        )
        message.attachment = dl_attachment

        logging.info("Certificate embedded and attached successfully")

    sg = SendGridAPIClient(api_key)
    response = sg.send(message)

    return response.status_code
