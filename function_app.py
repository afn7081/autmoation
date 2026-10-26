"""
Azure Functions app — Stripe webhook handler for donation emails.
"""

import os
import json
import logging
import tempfile
import threading
from datetime import datetime, timezone

import azure.functions as func
import stripe

from email_sender import send_donation_email
from pdf_generator import generate_receipt_pdf
from asset_store import get_asset

app = func.FunctionApp()

# Blob names in the Azure Storage container.
GENERAL_EMAIL_BLOB = "General Email Template.docx"
DONOR_ABOVE_1000_BLOB = "Email Template for Donors above $1000.docx"
FORMAT_BLOB = "format.docx"

# ── Idempotency: dedupe by charge_id across retries/restarts ─────────────────
_PROCESSED_PI_FILE = os.path.join(tempfile.gettempdir(), "processed_charges.txt")
_PROCESSED_LOCK = threading.Lock()


def _already_processed(pi_id: str) -> bool:
    if not pi_id:
        return False
    with _PROCESSED_LOCK:
        if not os.path.exists(_PROCESSED_PI_FILE):
            return False
        with open(_PROCESSED_PI_FILE, "r") as f:
            return pi_id in {line.strip() for line in f}


def _mark_processed(pi_id: str) -> None:
    if not pi_id:
        return
    with _PROCESSED_LOCK:
        with open(_PROCESSED_PI_FILE, "a") as f:
            f.write(pi_id + "\n")


def format_currency(amount_cents):
    amount = amount_cents / 100
    return f"${amount:,.2f}"


def format_date(timestamp):
    dt = datetime.fromtimestamp(timestamp, tz=timezone.utc)
    return dt.strftime("%B %d, %Y")


def _custom_field_value(custom_fields, key):
    """Return the text value of a checkout custom field by key, if present."""
    for field in custom_fields or []:
        if field.get("key") == key:
            return (field.get("text") or {}).get("value")
    return None


def extract_donor_info(session):
    """Extract donor name, email, amount, and date from a Checkout Session object."""
    custom_fields = session.get("custom_fields", [])
    first_name = _custom_field_value(custom_fields, "firstname")
    last_name = _custom_field_value(custom_fields, "lastname")

    name = " ".join(part for part in [first_name, last_name] if part).strip()

    customer_details = session.get("customer_details") or {}

    # Fallback to the name collected by Stripe if custom fields are empty
    if not name:
        name = customer_details.get("name") or "Valued Donor"

    email = customer_details.get("email") or session.get("customer_email")

    amount_cents = session.get("amount_total", 0)
    created = session.get("created", int(datetime.now(timezone.utc).timestamp()))

    return {
        "name": name,
        "email": email,
        "amount_cents": amount_cents,
        "amount_formatted": format_currency(amount_cents),
        "date_formatted": format_date(created),
    }


def handle_successful_payment(session, event_id="unknown"):
    """Process a completed checkout session: generate PDF receipt and send email."""
    session_id = session.get("id", "unknown")
    donor = extract_donor_info(session)

    log_prefix = f"[evt={event_id}][session={session_id}]"

    if not donor["email"]:
        print(f"{log_prefix} No email found. Skipping.", flush=True)
        logging.warning(f"{log_prefix} No email found. Skipping.")
        return

    amount_dollars = donor["amount_cents"] / 100

    logging.info(
        f"{log_prefix} Processing: donor={donor['name']}, "
        f"email={donor['email']}, amount=${amount_dollars:.2f}"
    )
    print(f"{log_prefix} Processing: donor={donor['name']}, email={donor['email']}, amount=${amount_dollars:.2f}", flush=True)

    if amount_dollars >= 2500:
        email_template_path = get_asset(DONOR_ABOVE_1000_BLOB)
        logging.info(f"{log_prefix} Using $2500+ template with certificate")
    else:
        email_template_path = get_asset(GENERAL_EMAIL_BLOB)

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = generate_receipt_pdf(
            template_path=get_asset(FORMAT_BLOB),
            output_dir=tmpdir,
            donor_name=donor["name"],
            amount=donor["amount_formatted"],
            date=donor["date_formatted"],
        )
        logging.info(f"{log_prefix} PDF receipt generated")

        send_donation_email(
            template_path=email_template_path,
            to_email=donor["email"],
            donor_name=donor["name"],
            amount=donor["amount_formatted"],
            pdf_path=pdf_path,
            amount_dollars=amount_dollars,
            date=donor["date_formatted"],
        )

    logging.info(f"{log_prefix} Email sent to {donor['email']} for ${amount_dollars:.2f}")


# ── Stripe Webhook Endpoint ──────────────────────────────────────────────────

@app.route(route="webhook", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def stripe_webhook(req: func.HttpRequest) -> func.HttpResponse:
    payload = req.get_body().decode("utf-8")
    sig_header = req.headers.get("Stripe-Signature")
    webhook_secret = os.getenv("STRIPE_WEBHOOK_SECRET")
    stripe.api_key = os.getenv("STRIPE_SECRET_KEY")

    try:
        event = stripe.Webhook.construct_event(payload, sig_header, webhook_secret)
    except ValueError:
        logging.error("Invalid webhook payload")
        return func.HttpResponse(json.dumps({"error": "Invalid payload"}), status_code=400)
    except stripe.error.SignatureVerificationError:
        logging.error("Invalid webhook signature")
        return func.HttpResponse(json.dumps({"error": "Invalid signature"}), status_code=400)

    event_id = event.get("id", "unknown")
    event_type = event["type"]
    logging.info(f"[evt={event_id}] Received {event_type}")

    if event_type == "checkout.session.completed":
        session = event["data"]["object"]
        session_id = session.get("id", "")

        # Only process fully paid sessions
        if session.get("payment_status") != "paid":
            logging.info(f"[evt={event_id}][session={session_id}] Not paid (status={session.get('payment_status')}), skipping")
            return func.HttpResponse(json.dumps({"status": "ignored"}), status_code=200)

        if _already_processed(session_id):
            logging.info(f"[evt={event_id}][session={session_id}] Already processed, skipping (dedup)")
            return func.HttpResponse(json.dumps({"status": "duplicate"}), status_code=200)

        # Mark BEFORE processing so concurrent retries also see it
        _mark_processed(session_id)

        # Process synchronously — PDF/certificate generation and email delivery.
        try:
            handle_successful_payment(session, event_id)
            logging.info(f"[evt={event_id}][session={session_id}] Processing completed successfully")
        except Exception as e:
            logging.exception(f"[evt={event_id}][session={session_id}] Processing failed: {e}")
            return func.HttpResponse(json.dumps({"error": str(e)}), status_code=500)

    return func.HttpResponse(json.dumps({"status": "ok"}), status_code=200)

# making a change
# ── Health Check ─────────────────────────────────────────────────────────────

@app.route(route="health", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def health(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse(json.dumps({"status": "healthy"}), status_code=200)
