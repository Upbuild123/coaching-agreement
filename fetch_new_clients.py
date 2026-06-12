"""
Fetches Airtable "New Coaching Client" notification emails from Gmail,
lets you review and correct each client's details, then generates the
agreement document into Contract Automation/agreements/.

Setup:
  1. Enable 2-Step Verification on the Gmail account
  2. Generate an App Password: myaccount.google.com → Security → App Passwords
  3. Add to Contract Automation/.env:
       GMAIL_USER=ea.michael108@gmail.com
       GMAIL_PASSWORD=xxxxxxxxxxxxxxxxxxxx   (16-char app password, no spaces)

Run:
  python3 fetch_new_clients.py
"""

import imaplib
import email
import os
import re
import sys
import shutil
from pathlib import Path
from bs4 import BeautifulSoup
import docx

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

GMAIL_IMAP     = "imap.gmail.com"
SUBJECT_FILTER = "Airtable forms: someone has responded to Upbuild New Coaching Clients"
BASE_DIR       = Path(__file__).parent.parent
TEMPLATE       = BASE_DIR / "Dummy Agreement.docx"
AGREEMENTS_DIR = Path(__file__).parent / "agreements"

COACHES = ["Rasanath Das", "Hari Prasada Das", "Michael Sloyer", "Tzipi Weiss", "Vipin Goyal"]

COACH_MAP = {
    "rasanath": "Rasanath Das",
    "hari":     "Hari Prasada Das",
    "michael":  "Michael Sloyer",
    "tzipi":    "Tzipi Weiss",
    "vipin":    "Vipin Goyal",
}


# ---------------------------------------------------------------------------
# Env
# ---------------------------------------------------------------------------

def load_env():
    env_file = Path(__file__).parent / ".env"
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


# ---------------------------------------------------------------------------
# Gmail
# ---------------------------------------------------------------------------

def connect_gmail(user: str, password: str) -> imaplib.IMAP4_SSL:
    mail = imaplib.IMAP4_SSL(GMAIL_IMAP)
    mail.login(user, password)
    mail.select("inbox")
    return mail


def fetch_matching_emails(mail: imaplib.IMAP4_SSL) -> list:
    _, data = mail.search(None, f'SUBJECT "{SUBJECT_FILTER}"')
    ids = data[0].split()
    if not ids:
        return []
    results = []
    for uid in ids:
        _, msg_data = mail.fetch(uid, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        body = extract_body(msg)
        if body:
            parsed = parse_airtable_email(body)
            if parsed:
                results.append(parsed)
    return results


def extract_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct in ("text/html", "text/plain"):
                return part.get_payload(decode=True).decode(errors="replace")
    else:
        return msg.get_payload(decode=True).decode(errors="replace")
    return None


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

def parse_airtable_email(html: str):
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(separator="\n")

    def field(name: str) -> str:
        m = re.search(rf"{re.escape(name)}\s*\n\s*(.+)", text, re.IGNORECASE)
        return m.group(1).strip() if m else ""

    client_name  = field("Client Name")
    client_email = field("Email")
    coach_raw    = field("Coach")
    rate_raw     = field("Rate")

    if not client_name:
        return None

    parts      = client_name.strip().split()
    first_name = parts[0] if parts else ""
    last_name  = " ".join(parts[1:]) if len(parts) > 1 else ""
    coach      = resolve_coach(coach_raw)
    rate       = re.sub(r"[^\d]", "", rate_raw)

    return {
        "first_name":   first_name,
        "last_name":    last_name,
        "client_email": client_email,
        "coach":        coach,
        "rate":         rate,
        "coach_raw":    coach_raw,
    }


def resolve_coach(raw: str) -> str:
    key = raw.strip().lower().split()[0] if raw.strip() else ""
    return COACH_MAP.get(key, raw.strip())


# ---------------------------------------------------------------------------
# Interactive review & edit
# ---------------------------------------------------------------------------

def prompt_edit(label: str, current: str) -> str:
    value = input(f"  {label} [{current}]: ").strip()
    return value if value else current


def review_client(c: dict):
    print(f"\n  Client Name:  {c['first_name']} {c['last_name']}")
    print(f"  Email:        {c['client_email']}")
    print(f"  Coach:        {c['coach']}")
    print(f"  Rate:         ${c['rate']}/session")

    if c["coach"] not in COACHES:
        print(f"  ⚠  Coach '{c['coach_raw']}' not in known list — please correct below")

    action = input("\n  (e)dit details  /  (s)kip  /  (c)ontinue: ").strip().lower()

    if action == "s":
        return None

    if action == "e":
        c["first_name"]   = prompt_edit("First name",  c["first_name"])
        c["last_name"]    = prompt_edit("Last name",   c["last_name"])
        c["client_email"] = prompt_edit("Email",       c["client_email"])
        c["rate"]         = prompt_edit("Rate ($)",    c["rate"])

        print("\n  Coaches:")
        for i, name in enumerate(COACHES, 1):
            marker = " <--" if name == c["coach"] else ""
            print(f"    {i}. {name}{marker}")
        choice = input(f"  Select coach (1-5) or Enter to keep [{c['coach']}]: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(COACHES):
            c["coach"] = COACHES[int(choice) - 1]

    return c


# ---------------------------------------------------------------------------
# Agreement generation
# ---------------------------------------------------------------------------

def replace_in_paragraph(para, old: str, new: str):
    if old == new or old not in para.text:
        return
    runs = para.runs
    full = "".join(r.text for r in runs)
    # Collect all occurrence positions, then replace right-to-left
    # so earlier positions stay valid after each substitution.
    positions = []
    idx = 0
    while True:
        i = full.find(old, idx)
        if i == -1:
            break
        positions.append(i)
        idx = i + len(old)
    for occ_start in reversed(positions):
        occ_end = occ_start + len(old)
        pos, first = 0, False
        for run in runs:
            rs, re_ = pos, pos + len(run.text)
            if re_ > occ_start and rs < occ_end:
                pre  = run.text[:max(0, occ_start - rs)]
                post = run.text[max(0, occ_end - rs):]
                run.text = (pre + new + post) if not first else (pre + post)
                first = True
            pos = re_


def build_agreement(c: dict) -> Path:
    AGREEMENTS_DIR.mkdir(exist_ok=True)
    client_full = f"{c['first_name']} {c['last_name']}"
    out_path    = AGREEMENTS_DIR / f"{client_full} Upbuild Agreement.docx"
    shutil.copy(TEMPLATE, out_path)
    doc = docx.Document(out_path)
    for para in doc.paragraphs:
        replace_in_paragraph(para, "Alexis Gevorgian", client_full)
        replace_in_paragraph(para, "Alexis",           c["first_name"])
        replace_in_paragraph(para, "Rasanath Das",     c["coach"])
        replace_in_paragraph(para, "$600",             f"${c['rate']}")
    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    load_env()
    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pass = os.environ.get("GMAIL_PASSWORD")

    if not gmail_user or not gmail_pass:
        print("ERROR: Set GMAIL_USER and GMAIL_PASSWORD in Contract Automation/.env")
        sys.exit(1)

    print("\n=== Upbuild Contract Automation ===\n")
    print(f"Connecting to Gmail as {gmail_user}...")
    mail = connect_gmail(gmail_user, gmail_pass)
    print("Searching for new client emails...")
    clients = fetch_matching_emails(mail)
    mail.logout()

    if not clients:
        print("No new client emails found.")
        return

    print(f"\nFound {len(clients)} client(s):\n")
    for i, c in enumerate(clients, 1):
        print(f"  {i}. {c['first_name']} {c['last_name']}  |  {c['client_email']}  |  Coach: {c['coach']}  |  ${c['rate']}/session")

    print("\n" + "-" * 50)

    for i, c in enumerate(clients, 1):
        print(f"\n[{i}/{len(clients)}] Reviewing: {c['first_name']} {c['last_name']}")
        c = review_client(c)
        if c is None:
            print("  Skipped.")
            continue

        print("\n  Generating agreement...")
        doc_path = build_agreement(c)
        print(f"  Saved: agreements/{doc_path.name}")

    print("\nDone.")


if __name__ == "__main__":
    main()
