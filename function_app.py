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

# ── Idempotency: dedupe by payment_intent_id across retries/restarts ─────────
_PROCESSED_PI_FILE = os.path.join(tempfile.gettempdir(), "processed_payment_intents.txt")
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


def extract_donor_info(payment_intent):
    """Extract donor name, email, amount, and date from a Stripe PaymentIntent."""
    charges = payment_intent.get("charges", {}).get("data", [])
    latest_charge = payment_intent.get("latest_charge")

    name = "Valued Donor"
    email = None

    if charges:
        charge = charges[0]
        billing = charge.get("billing_details", {})
        name = billing.get("name") or name
        email = billing.get("email") or charge.get("receipt_email")
    elif isinstance(latest_charge, dict):
        billing = latest_charge.get("billing_details", {})
        name = billing.get("name") or name
        email = billing.get("email") or latest_charge.get("receipt_email")
    elif isinstance(latest_charge, str) and latest_charge.startswith("ch_"):
        try:
            charge = stripe.Charge.retrieve(latest_charge)
            billing = charge.get("billing_details", {})
            name = billing.get("name") or name
            email = billing.get("email") or charge.get("receipt_email")
        except Exception:
            pass

    # Check metadata for donor name (useful for test payments)
    metadata = payment_intent.get("metadata", {})
    if name == "Valued Donor" and metadata.get("donor_name"):
        name = metadata["donor_name"]

    if not email:
        email = payment_intent.get("receipt_email")

    if not email:
        customer_id = payment_intent.get("customer")
        if customer_id:
            try:
                customer = stripe.Customer.retrieve(customer_id)
                email = customer.get("email")
                if name == "Valued Donor":
                    name = customer.get("name") or name
            except Exception:
                pass

    amount_cents = payment_intent.get("amount", 0)
    created = payment_intent.get("created", int(datetime.now(timezone.utc).timestamp()))

    return {
        "name": name,
        "email": email,
        "amount_cents": amount_cents,
        "amount_formatted": format_currency(amount_cents),
        "date_formatted": format_date(created),
    }


def handle_successful_payment(payment_intent, event_id="unknown"):
    """Process a successful payment: generate PDF receipt and send email."""
    pi_id = payment_intent.get("id", "unknown")
    donor = extract_donor_info(payment_intent)

    log_prefix = f"[evt={event_id}][pi={pi_id}]"

    if not donor["email"]:
        logging.warning(f"{log_prefix} No email found. Skipping.")
        return

    amount_dollars = donor["amount_cents"] / 100

    logging.info(
        f"{log_prefix} Processing: donor={donor['name']}, "
        f"email={donor['email']}, amount=${amount_dollars:.2f}"
    )

    if amount_dollars >= 1000:
        email_template_path = get_asset(DONOR_ABOVE_1000_BLOB)
        logging.info(f"{log_prefix} Using $1000+ template with certificate")
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

    if event_type == "payment_intent.succeeded":
        pi = event["data"]["object"]
        pi_id = pi.get("id", "")

        if _already_processed(pi_id):
            logging.info(f"[evt={event_id}][pi={pi_id}] Already processed, skipping (dedup)")
            return func.HttpResponse(json.dumps({"status": "duplicate"}), status_code=200)

        # Mark BEFORE processing so concurrent retries also see it
        _mark_processed(pi_id)

        # Process asynchronously so we can ACK Stripe within its ~30s timeout.
        # The heavy work (PDF + certificate generation + SVG->PNG conversion +
        # SendGrid send) can take longer than Stripe is willing to wait, which
        # caused "context deadline exceeded" timeouts and unnecessary retries.
        def _process_async(pi=pi, event_id=event_id):
            try:
                handle_successful_payment(pi, event_id)
            except Exception as e:
                logging.exception(f"[evt={event_id}] Async processing failed: {e}")

        threading.Thread(
            target=_process_async, name=f"webhook-{pi_id}", daemon=True
        ).start()
        logging.info(f"[evt={event_id}][pi={pi_id}] Dispatched to background worker")

    return func.HttpResponse(json.dumps({"status": "ok"}), status_code=200)

# making a change
# ── Health Check ─────────────────────────────────────────────────────────────

@app.route(route="health", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def health(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse(json.dumps({"status": "healthy"}), status_code=200)
