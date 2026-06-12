"""
Upbuild Contract Automation — Streamlit App

Environment variables (set in .env locally, or Streamlit Cloud secrets):
  GMAIL_USER      — sending Gmail address
  GMAIL_PASSWORD  — 16-char Gmail App Password
  NOTIFY_EMAIL    — address to receive the agreement (defaults to GMAIL_USER)
"""

import io
import os
import shutil
import smtplib
import tempfile
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import docx
import streamlit as st

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

COACHES = [
    "Rasanath Das",
    "Hari Prasada Das",
    "Michael Sloyer",
    "Tzipi Weiss",
    "Vipin Goyal",
]

COACH_EMAILS = {
    "Rasanath Das":    "rasanath@upbuild.com",
    "Hari Prasada Das": "hari@upbuild.com",
    "Tzipi Weiss":     "tzipi@upbuild.com",
    "Vipin Goyal":     "vipin@upbuild.com",
}

TEMPLATE = Path(__file__).parent / "Upbuild Coaching Agreement - Sample.docx"


# ---------------------------------------------------------------------------
# Env loader (local dev only; Streamlit Cloud uses st.secrets)
# ---------------------------------------------------------------------------

def load_env():
    env_file = Path(__file__).parent / ".env"
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


def get_secret(key: str, default: str = "") -> str:
    """Read from Streamlit secrets first, fall back to env vars."""
    try:
        return st.secrets[key]
    except (KeyError, FileNotFoundError):
        return os.environ.get(key, default)


# ---------------------------------------------------------------------------
# Agreement generation
# ---------------------------------------------------------------------------

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
            rs, re_ = pos, pos + len(run.text)
            if re_ > occ_start and rs < occ_end:
                pre  = run.text[:max(0, occ_start - rs)]
                post = run.text[max(0, occ_end - rs):]
                run.text = (pre + new + post) if not first else (pre + post)
                first = True
            pos = re_


def build_agreement_bytes(first_name: str, last_name: str, coach: str, rate: str) -> bytes:
    client_full = f"{first_name} {last_name}"
    rate = f"{int(rate):,}"

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    shutil.copy(TEMPLATE, tmp_path)
    document = docx.Document(tmp_path)

    replacements = [
        ("[Client Full Name]",  client_full),
        ("[Client First Name]", first_name),
        ("[Coach Name]",        coach),
        ("[Rate]",              rate),
    ]

    for para in document.paragraphs:
        for old, new in replacements:
            replace_in_paragraph(para, old, new)

    buf = io.BytesIO()
    document.save(buf)
    tmp_path.unlink(missing_ok=True)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

def send_email(first_name: str, last_name: str, client_email: str,
               coach: str, rate: str, doc_bytes: bytes):
    gmail_user = get_secret("GMAIL_USER")
    gmail_pass = get_secret("GMAIL_PASSWORD")
    notify_to  = get_secret("NOTIFY_EMAIL") or gmail_user

    recipients = [notify_to]
    if coach in COACH_EMAILS:
        recipients.append(COACH_EMAILS[coach])

    client_full = f"{first_name} {last_name}"
    filename    = f"{client_full} Upbuild Agreement.docx"

    msg = MIMEMultipart()
    msg["From"]    = gmail_user
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = f"New Coaching Agreement — {client_full}"

    body = (
        f"A new coaching agreement has been generated.\n\n"
        f"  Client:  {client_full}\n"
        f"  Email:   {client_email}\n"
        f"  Coach:   {coach}\n"
        f"  Rate:    ${rate}/session\n\n"
        f"The agreement is attached."
    )
    msg.attach(MIMEText(body, "plain"))

    attachment = MIMEApplication(doc_bytes, Name=filename)
    attachment["Content-Disposition"] = f'attachment; filename="{filename}"'
    msg.attach(attachment)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(gmail_user, gmail_pass)
        server.sendmail(gmail_user, recipients, msg.as_string())


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

load_env()

st.set_page_config(page_title="New Coaching Agreement", page_icon="📄", layout="centered")

st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f5f5f0 0%, #eef2ee 100%);
    }
    .block-container {
        max-width: 640px;
        padding-top: 3rem;
    }
    #upbuild-card {
        background: #ffffff;
        border-radius: 16px;
        padding: 2.5rem 2.5rem 1.5rem;
        box-shadow: 0 4px 24px rgba(0,0,0,0.06);
        border: 1px solid #e8e8e2;
        margin-bottom: 1.5rem;
    }
    #upbuild-card h1 {
        font-size: 1.6rem;
        font-weight: 700;
        color: #2c3e2e;
        margin-bottom: 0.25rem;
    }
    #upbuild-card p {
        color: #6b7268;
        font-size: 0.95rem;
        margin-bottom: 0;
    }
    div[data-testid="stForm"] {
        background: #ffffff;
        border-radius: 16px;
        padding: 2rem 2.5rem 2.5rem;
        box-shadow: 0 4px 24px rgba(0,0,0,0.06);
        border: 1px solid #e8e8e2;
    }
    .stTextInput input, .stSelectbox div[data-baseweb="select"] {
        border-radius: 8px;
    }
    .stButton button, button[kind="primary"], button[kind="formSubmit"] {
        background-color: #4a7c5e !important;
        color: #fff !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 0.6rem 1.5rem !important;
        transition: background-color 0.15s ease !important;
    }
    .stButton button:hover, button[kind="primary"]:hover, button[kind="formSubmit"]:hover {
        background-color: #3a6349 !important;
    }
    label {
        font-weight: 500 !important;
        color: #3a3f3a !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div id="upbuild-card">
    <h1>📄 New Coaching Agreement</h1>
    <p>Fill in the client's details below to generate and send the agreement.</p>
</div>
""", unsafe_allow_html=True)

with st.form("agreement_form"):
    col1, col2 = st.columns(2)
    first_name = col1.text_input("Client first name", placeholder="Jane")
    last_name  = col2.text_input("Client last name",  placeholder="Smith")

    client_email = st.text_input("Client email", placeholder="jane@example.com",
                                  help="Not added to the agreement — sent to you for follow-up.")

    coach = st.selectbox("Coach", options=COACHES)

    rate = st.text_input("Session rate ($)", placeholder="600",
                          help="Numbers only, e.g. 600")

    st.write("")
    submitted = st.form_submit_button("Generate & Send Agreement", type="primary", use_container_width=True)

if submitted:
    errors = []
    if not first_name.strip():
        errors.append("First name is required.")
    if not last_name.strip():
        errors.append("Last name is required.")
    if not client_email.strip() or "@" not in client_email:
        errors.append("A valid client email is required.")
    if not rate.strip().isdigit():
        errors.append("Rate must be a whole number (e.g. 600).")

    if errors:
        for e in errors:
            st.error(e)
    else:
        with st.spinner("Generating agreement and sending email…"):
            try:
                doc_bytes = build_agreement_bytes(
                    first_name.strip(), last_name.strip(), coach, rate.strip()
                )
                send_email(
                    first_name.strip(), last_name.strip(),
                    client_email.strip(), coach, rate.strip(),
                    doc_bytes,
                )
                st.success(
                    f"Agreement for {first_name.strip()} {last_name.strip()} "
                    f"generated and sent to you!"
                )
            except Exception as exc:
                st.error(f"Something went wrong: {exc}")
