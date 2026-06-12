"""
Interactive script: generates a client agreement and sends it for signing via Secured Signing.

Run:
    python3 send_agreement.py

Credentials are read from .env in the same directory:
    SS_API_KEY     — your Secured Signing API key
    SS_API_SECRET  — your Secured Signing API secret
"""

import os
import sys
import hmac
import hashlib
import base64
import time
import uuid
import json
import shutil
from pathlib import Path
from datetime import datetime, timedelta

import requests
import docx

BASE_URL  = "https://api.securedsigning.com/web/v1.4"
EMAIL_TEMPLATE = "Upbuild Coaching Agreement"
DUE_DAYS  = 30
TEMPLATE  = Path(__file__).parent / "Dummy Agreement.docx"
COACHES   = ["Rasanath Das", "Hari Prasada Das", "Michael Sloyer", "Tzipi Weiss", "Vipin Goyal"]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_env():
    env_file = Path(__file__).parent / ".env"
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

def ask(prompt: str) -> str:
    value = input(prompt).strip()
    if not value:
        print("  Value cannot be empty.")
        return ask(prompt)
    return value

def replace_in_paragraph(para, old: str, new: str):
    if old == new or old not in para.text:
        return
    runs = para.runs
    full = "".join(r.text for r in runs)
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
            rs, re = pos, pos + len(run.text)
            if re > occ_start and rs < occ_end:
                pre  = run.text[:max(0, occ_start - rs)]
                post = run.text[max(0, occ_end - rs):]
                run.text = (pre + new + post) if not first else (pre + post)
                first = True
            pos = re


# ---------------------------------------------------------------------------
# Document generation
# ---------------------------------------------------------------------------

def build_agreement(first_name: str, last_name: str, coach: str, rate: str) -> Path:
    client_full  = f"{first_name} {last_name}"
    out_path     = Path(__file__).parent / f"{client_full} Upbuild Agreement.docx"

    shutil.copy(TEMPLATE, out_path)
    doc = docx.Document(out_path)

    replacements = [
        ("Alexis Gevorgian", client_full),
        ("Alexis",           first_name),
        ("Rasanath Das",     coach),
        ("$600",             f"${rate}"),
    ]

    for para in doc.paragraphs:
        for old, new in replacements:
            replace_in_paragraph(para, old, new)

    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------------------
# Secured Signing API
# ---------------------------------------------------------------------------

def auth_headers(api_key: str, api_secret: str, referer: str) -> dict:
    timestamp = str(int(time.time()))
    nonce     = uuid.uuid4().hex[:32]
    string_to_sign = f"{api_key}\n{timestamp}\n{nonce}"
    # .NET SDK uses ASCII encoding for both key and message; secret is used as-is
    sig = base64.b64encode(
        hmac.new(api_secret.encode("ascii"), string_to_sign.encode("ascii"), hashlib.sha256).digest()
    ).decode()
    return {
        "X-CUSTOM-API-KEY":   api_key,
        "X-CUSTOM-SIGNATURE": sig,
        "X-CUSTOM-DATE":      timestamp,
        "X-CUSTOM-NONCE":     nonce,
        "Referer":            referer,
    }

def upload_document(api_key: str, api_secret: str, referer: str, doc_path: Path) -> str:
    headers = auth_headers(api_key, api_secret, referer)
    mime    = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    with open(doc_path, "rb") as f:
        resp = requests.post(
            f"{BASE_URL}/Document/Uploader",
            headers=headers,
            files={"file": (doc_path.name, f, mime)},
        )
    if not resp.ok:
        print(f"  ERROR uploading: {resp.status_code} {resp.text}")
        sys.exit(1)
    ref = resp.json().get("Reference")
    if not ref:
        print(f"  ERROR: no Reference in upload response: {resp.text}")
        sys.exit(1)
    return ref

def send_smarttag(api_key: str, api_secret: str, referer: str, doc_ref: str,
                  email: str, first_name: str, last_name: str) -> dict:
    headers = auth_headers(api_key, api_secret, referer)
    headers["Content-Type"] = "application/json"
    due_date = (datetime.now() + timedelta(days=DUE_DAYS)).strftime("%Y-%m-%d")
    body = {
        "DocumentReferences": [doc_ref],
        "DueDate": due_date,
        "Signers": [{
            "Email":     email,
            "FirstName": first_name,
            "LastName":  last_name,
            "InvitationEmailTemplateReference": EMAIL_TEMPLATE,
        }],
    }
    resp = requests.post(f"{BASE_URL}/SmartTag/Send", headers=headers, json=body)
    if not resp.ok:
        print(f"  ERROR sending: {resp.status_code} {resp.text}")
        sys.exit(1)
    return resp.json()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    load_env()
    api_key    = os.environ.get("SS_API_KEY")
    api_secret = os.environ.get("SS_API_SECRET")
    referer    = os.environ.get("SS_REFERER", "https://upbuild.com")
    if not api_key or not api_secret:
        print("ERROR: Set SS_API_KEY and SS_API_SECRET in your .env file.")
        sys.exit(1)

    print("\n=== Upbuild Agreement Sender ===\n")

    # Coach
    print("Coaches:")
    for i, name in enumerate(COACHES, 1):
        print(f"  {i}. {name}")
    while True:
        choice = input("\nSelect coach (1-5): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(COACHES):
            coach = COACHES[int(choice) - 1]
            break
        print("  Please enter a number between 1 and 5.")

    # Client
    first_name = ask("\nClient first name: ")
    last_name  = ask("Client last name:  ")
    email      = ask("Client email:      ")

    # Rate
    rate = ask("Session rate (numbers only, e.g. 600): $")

    # Confirm
    print(f"\nReady to send:")
    print(f"  Coach:  {coach}")
    print(f"  Client: {first_name} {last_name} <{email}>")
    print(f"  Rate:   ${rate}/session")
    confirm = input("\nProceed? (y/n): ").strip().lower()
    if confirm != "y":
        print("Cancelled.")
        sys.exit(0)

    # Build doc
    print("\nGenerating agreement...")
    doc_path = build_agreement(first_name, last_name, coach, rate)
    print(f"  Saved: {doc_path.name}")

    # Upload
    print("Uploading to Secured Signing...")
    doc_ref = upload_document(api_key, api_secret, referer, doc_path)
    print(f"  Uploaded — ref: {doc_ref}")

    # Send
    print(f"Sending to {email}...")
    result = send_smarttag(api_key, api_secret, referer, doc_ref, email, first_name, last_name)
    print("  Sent!\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
