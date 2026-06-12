"""
Fee, Expenses & Cancellation Policy section builder for the Upbuild Services Agreement.

Generates the body text of Section 3 ("Fees, Expenses, and Cancellation Policy")
from structured inputs, so a single template can support many engagement types.
"""

from typing import List, Optional


FEE_ITEM_TYPES = [
    "1:1 Coaching",
    "Executive Coaching",
    "Joint/Group Coaching",
    "Workshop",
    "Leadership Team Session",
    "Full Day Facilitation",
    "Half Day Facilitation",
    "Other Custom Service",
]

PRICING_TYPES = ["Hourly", "Daily", "Per Session", "Fixed Fee", "Custom Text"]

FEE_STRUCTURE_MODES = ["Fixed Rates", "Variable Rates", "Estimated Annual Investment"]

CANCELLATION_TYPES = ["Standard Coaching", "Workshop / Offsite", "Custom"]

INVOICE_FREQUENCIES = ["Quarterly", "Monthly", "Upon Completion", "Custom"]
PAYMENT_TERMS = ["15 Days", "30 Days", "Custom"]

EXPENSE_CATEGORIES = ["Travel", "Lodging", "Meals", "Transportation", "Materials", "Custom"]
TRAVEL_CLASSES = ["Economy", "Premium Economy", "Business", "Custom"]

ENGAGEMENT_TYPES = [
    "Standard Coaching",
    "Executive Coaching",
    "Variable Coaching Rates",
    "Fixed-Fee Workshop / Offsite",
    "Enterprise Engagement",
    "Custom",
]


# ---------------------------------------------------------------------------
# Engagement-type defaults
# ---------------------------------------------------------------------------

def default_state(engagement_type: str) -> dict:
    """Return a dict of sensible defaults for the given engagement type."""

    state = {
        "fee_structure_mode": "Fixed Rates",
        "estimate_amount": "",
        "estimate_disclaimer": True,
        "include_additional_services": True,
        "additional_services_text": (
            "Fees for any additional services will depend on the scope of the "
            "services and will be discussed and agreed upon in advance."
        ),
        "include_expenses": False,
        "expense_categories": ["Travel", "Lodging", "Meals"],
        "expense_preapproval": True,
        "travel_class": "Business",
        "travel_cap": "5,000",
        "expense_approver": "the Client",
        "invoice_frequency": "Quarterly",
        "invoice_frequency_custom": "",
        "payment_terms": "30 Days",
        "payment_terms_custom": "",
    }

    if engagement_type == "Standard Coaching":
        state["fee_items"] = [
            {"name": "1:1 coaching", "pricing_type": "Hourly", "rate": "650", "notes": ""},
        ]
        state["cancellation_policies"] = [
            {"policy_type": "Standard Coaching", "service_type": "session",
             "notice_period": "2", "charge_text": "the full session fee", "additional_text": ""},
        ]

    elif engagement_type == "Executive Coaching":
        state["fee_items"] = [
            {"name": "executive coaching", "pricing_type": "Hourly", "rate": "975", "notes": ""},
        ]
        state["cancellation_policies"] = [
            {"policy_type": "Standard Coaching", "service_type": "session",
             "notice_period": "2", "charge_text": "the full session fee", "additional_text": ""},
        ]

    elif engagement_type == "Variable Coaching Rates":
        state["fee_structure_mode"] = "Variable Rates"
        state["fee_items"] = [
            {"name": "1:1 coaching", "pricing_type": "Hourly", "rate": "", "notes": ""},
            {"name": "executive coaching", "pricing_type": "Hourly", "rate": "", "notes": ""},
        ]
        state["cancellation_policies"] = [
            {"policy_type": "Standard Coaching", "service_type": "session",
             "notice_period": "2", "charge_text": "the full session fee", "additional_text": ""},
        ]

    elif engagement_type == "Fixed-Fee Workshop / Offsite":
        state["fee_items"] = [
            {"name": "the workshop", "pricing_type": "Fixed Fee", "rate": "5,000", "notes": ""},
        ]
        state["include_expenses"] = True
        state["cancellation_policies"] = [
            {"policy_type": "Workshop / Offsite", "service_type": "workshop or in-person session",
             "notice_period": "10", "charge_text": "in full", "additional_text": ""},
        ]

    elif engagement_type == "Enterprise Engagement":
        state["fee_structure_mode"] = "Estimated Annual Investment"
        state["estimate_amount"] = "150,000"
        state["fee_items"] = [
            {"name": "coaching", "pricing_type": "Hourly", "rate": "1,200", "notes": ""},
            {"name": "in-person leadership team sessions", "pricing_type": "Daily", "rate": "10,000", "notes": ""},
            {"name": "group coaching sessions", "pricing_type": "Per Session", "rate": "3,000", "notes": ""},
        ]
        state["include_expenses"] = True
        state["expense_categories"] = ["Travel"]
        state["travel_class"] = "Premium Economy or Business"
        state["expense_approver"] = "Client's CEO"
        state["cancellation_policies"] = [
            {"policy_type": "Standard Coaching", "service_type": "individual coaching session",
             "notice_period": "2", "charge_text": "the full session fee", "additional_text": ""},
            {"policy_type": "Workshop / Offsite", "service_type": "workshop or in-person session",
             "notice_period": "10", "charge_text": "in full", "additional_text": ""},
        ]

    else:  # Custom
        state["fee_items"] = [
            {"name": "", "pricing_type": "Fixed Fee", "rate": "", "notes": ""},
        ]
        state["cancellation_policies"] = [
            {"policy_type": "Custom", "service_type": "", "notice_period": "",
             "charge_text": "", "additional_text": ""},
        ]

    return state


# ---------------------------------------------------------------------------
# Text generation helpers
# ---------------------------------------------------------------------------

def _money(value: str) -> str:
    """Format a numeric string with thousands separators; pass through non-numeric text."""
    cleaned = value.replace(",", "").replace("$", "").strip()
    if cleaned.isdigit():
        return f"{int(cleaned):,}"
    return value.strip()


def _fee_item_clause(item: dict, mode: str) -> Optional[str]:
    """A short clause like '$650 per hour for 1:1 coaching' (no leading 'The fee for...')."""
    name = item.get("name", "").strip()
    if not name:
        return None
    rate = item.get("rate", "").strip()
    ptype = item.get("pricing_type", "Hourly")

    if mode == "Variable Rates":
        return None  # handled separately

    if not rate:
        return None

    if ptype == "Hourly":
        return f"${_money(rate)} per hour for {name}"
    if ptype == "Daily":
        return f"${_money(rate)} per day for {name}"
    if ptype == "Per Session":
        return f"${_money(rate)} per session for {name}"
    if ptype == "Fixed Fee":
        return f"a fixed fee of ${_money(rate)} for {name}"
    # Custom Text
    return rate.rstrip(".")


def build_fees_paragraphs(fee_items: List[dict], mode: str, estimate_amount: str, estimate_disclaimer: bool) -> List[str]:
    """Build the first paragraph(s) of Section 3 covering fee structure."""
    paragraphs = []
    intro = "Fees for Services are based on the scope of work and participation levels agreed upon by the parties."

    if mode == "Variable Rates":
        names = [item.get("name", "").strip() for item in fee_items if item.get("name", "").strip()]
        if names:
            if len(names) == 1:
                names_phrase = names[0]
            else:
                names_phrase = ", ".join(names[:-1]) + f", and {names[-1]}"
            sentence = (
                f"The fee for {names_phrase} will be agreed upon on a "
                f"case-by-case basis based on the Coach and Client Employee."
            )
            paragraphs.append(f"{intro} {sentence}")
        else:
            paragraphs.append(intro)

    else:
        clauses = [c for c in (_fee_item_clause(item, mode) for item in fee_items) if c]
        if clauses:
            if len(clauses) == 1:
                clause_phrase = clauses[0]
            else:
                clause_phrase = ", ".join(clauses[:-1]) + f", and {clauses[-1]}"
            sentence = f"For this engagement, the Company is providing Services at {clause_phrase}."
            paragraphs.append(f"{intro} {sentence}")
        else:
            paragraphs.append(intro)

        if mode == "Estimated Annual Investment" and estimate_amount.strip():
            est = (
                f"Based on the anticipated scope, the total annual investment is "
                f"estimated at approximately ${_money(estimate_amount)}, exclusive of "
                f"travel and related expenses."
            )
            if estimate_disclaimer:
                est += (
                    " Actual fees will vary based on final participation, session "
                    "cadence, and any adjustments to the scope of Services."
                )
            paragraphs[-1] = paragraphs[-1] + " " + est

    return paragraphs


def build_additional_services_text(include: bool, custom_text: str) -> List[str]:
    if not include:
        return []
    text = custom_text.strip() or (
        "Fees for any additional services will depend on the scope of the "
        "services and will be discussed and agreed upon in advance."
    )
    return [text]


def build_expenses_text(config: dict) -> List[str]:
    if not config.get("include_expenses"):
        return []

    categories = config.get("expense_categories", [])
    custom_category = config.get("expense_category_custom", "").strip()
    cat_names = [c.lower() for c in categories if c != "Custom"]
    if "Custom" in categories and custom_category:
        cat_names.append(custom_category.lower())

    if cat_names:
        if len(cat_names) == 1:
            cat_phrase = cat_names[0]
        else:
            cat_phrase = ", ".join(cat_names[:-1]) + f", and {cat_names[-1]}"
        categories_clause = f", including {cat_phrase}"
    else:
        categories_clause = ""

    preapproved = "pre-approved " if config.get("expense_preapproval") else ""

    intro = (
        f"The Client will reimburse the Company for all reasonable and "
        f"{preapproved}out-of-pocket expenses incurred in connection with the "
        f"Services{categories_clause}."
    )

    sentences = [intro]

    if "Travel" in categories:
        travel_class = config.get("travel_class", "Business")
        if travel_class == "Custom":
            travel_class = config.get("travel_class_custom", "").strip() or "business"
        travel_cap = config.get("travel_cap", "").strip()
        approver = config.get("expense_approver", "the Client").strip() or "the Client"

        travel_sentence = f"Class of travel shall be {travel_class.lower()} round trip fare"
        if travel_cap:
            travel_sentence += f" not to exceed ${_money(travel_cap)}"
        travel_sentence += f" unless pre-approved by {approver}."
        sentences.append(travel_sentence)

    return [" ".join(sentences)]


def build_invoicing_text(config: dict) -> List[str]:
    frequency = config.get("invoice_frequency", "Quarterly")
    if frequency == "Quarterly":
        freq_sentence = (
            "Invoices will be issued on a quarterly basis and will reflect actual "
            "Services delivered during the applicable period."
        )
    elif frequency == "Monthly":
        freq_sentence = (
            "Invoices will be issued on a monthly basis and will reflect actual "
            "Services delivered during the applicable period."
        )
    elif frequency == "Upon Completion":
        freq_sentence = "The Company will provide an invoice to the Client upon completion of the Services."
    else:
        custom = config.get("invoice_frequency_custom", "").strip()
        freq_sentence = custom or "Invoices will be issued on a periodic basis agreed upon by the parties."

    terms = config.get("payment_terms", "15 Days")
    if terms == "15 Days":
        terms_phrase = "15 days"
    elif terms == "30 Days":
        terms_phrase = "30 days"
    else:
        terms_phrase = config.get("payment_terms_custom", "").strip() or "30 days"

    return [f"{freq_sentence} Payment is due within {terms_phrase} of the invoice date."]


def build_cancellation_text(policies: List[dict]) -> List[str]:
    """Return the cancellation policy as a lead-in sentence plus bullet lines (prefixed with '• ')."""
    bullets = []

    for policy in policies:
        ptype = policy.get("policy_type")
        service = policy.get("service_type", "").strip()
        if not ptype:
            continue

        if ptype == "Standard Coaching":
            notice = policy.get("notice_period", "2").strip() or "2"
            charge = policy.get("charge_text", "the full session fee").strip() or "the full session fee"
            label = service or "individual coaching session"
            bullets.append(
                f"For {label}s, sessions canceled more than {notice} business days "
                f"in advance will not be charged; sessions canceled within {notice} "
                f"business days of the scheduled time will be charged {charge}."
            )

        elif ptype == "Workshop / Offsite":
            notice = policy.get("notice_period", "10").strip() or "10"
            charge = policy.get("charge_text", "in full").strip() or "in full"
            label = service or "workshop or in-person session"
            bullets.append(
                f"For {label}s, cancellations made more than {notice} business days "
                f"in advance will not be charged. Cancellations made within {notice} "
                f"business days of the scheduled date will be charged {charge}, given "
                f"the time reserved and preparation involved. The Company will make "
                f"reasonable efforts to reschedule where possible."
            )

        else:  # Custom
            notice = policy.get("notice_period", "").strip()
            charge = policy.get("charge_text", "").strip()
            label = service or "Services"
            if notice and charge:
                bullets.append(
                    f"For {label}, cancellations made more than {notice} business "
                    f"days in advance will not be charged; cancellations made within "
                    f"{notice} business days will be charged {charge}."
                )
            additional = policy.get("additional_text", "").strip()
            if additional:
                bullets.append(additional)

    if not bullets:
        return []

    lines = ["The cancellation policy is as follows:"]
    lines.extend(f"• {b}" for b in bullets)
    return lines


def build_fee_section_lines(state: dict) -> List[str]:
    """Build the [Fee Section] body: Fees -> Additional Services -> Expenses."""
    lines: List[str] = []

    lines.extend(build_fees_paragraphs(
        state.get("fee_items", []),
        state.get("fee_structure_mode", "Fixed Rates"),
        state.get("estimate_amount", ""),
        state.get("estimate_disclaimer", True),
    ))

    lines.extend(build_additional_services_text(
        state.get("include_additional_services", True),
        state.get("additional_services_text", ""),
    ))

    lines.extend(build_expenses_text(state))

    return lines


def build_cancellation_section_lines(state: dict) -> List[str]:
    """Build the [Cancellation Policy] body."""
    return build_cancellation_text(state.get("cancellation_policies", []))


def lines_to_text(lines: List[str]) -> str:
    """Join lines into a single editable text blob, one per line."""
    return "\n".join(lines)


def text_to_lines(text: str) -> List[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]
