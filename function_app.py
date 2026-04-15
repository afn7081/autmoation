"""
Azure Functions app — Stripe webhook handler for donation emails.
"""

import os
import json
import logging
import tempfile
from datetime import datetime, timezone

import azure.functions as func
import stripe

from email_sender import send_donation_email
from pdf_generator import generate_receipt_pdf

app = func.FunctionApp()

TEMPLATE_DIR = os.path.dirname(os.path.abspath(__file__))
GENERAL_EMAIL_TEMPLATE = os.path.join(TEMPLATE_DIR, "General Email Template.docx")
DONOR_ABOVE_1000_TEMPLATE = os.path.join(TEMPLATE_DIR, "Email Template for Donors above $1000.docx")
FORMAT_TEMPLATE = os.path.join(TEMPLATE_DIR, "format.docx")


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
        email_template_path = DONOR_ABOVE_1000_TEMPLATE
        logging.info(f"{log_prefix} Using $1000+ template with certificate")
    else:
        email_template_path = GENERAL_EMAIL_TEMPLATE

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = generate_receipt_pdf(
            template_path=FORMAT_TEMPLATE,
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
        payment_intent = event["data"]["object"]
        handle_successful_payment(payment_intent, event_id)
    elif event_type == "charge.succeeded":
        charge = event["data"]["object"]
        pi_id = charge.get("payment_intent")
        if pi_id:
            try:
                payment_intent = stripe.PaymentIntent.retrieve(pi_id)
                handle_successful_payment(payment_intent, event_id)
            except Exception as e:
                logging.error(f"[evt={event_id}] Failed to retrieve PaymentIntent {pi_id}: {e}")
                return func.HttpResponse(
                    json.dumps({"error": "Failed to process"}), status_code=500
                )

    return func.HttpResponse(json.dumps({"status": "ok"}), status_code=200)

# making a change
# ── Health Check ─────────────────────────────────────────────────────────────

@app.route(route="health", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def health(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse(json.dumps({"status": "healthy"}), status_code=200)
