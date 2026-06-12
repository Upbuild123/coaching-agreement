"""One-off script to generate 'Upbuild Services Agreement - Template.docx'.

Run once locally: python3 build_services_template.py
"""

import docx
from docx.shared import Pt
from pathlib import Path

OUT = Path(__file__).parent / "Upbuild Services Agreement - Template.docx"

doc = docx.Document()

normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(11)


def h1(text):
    doc.add_heading(text, level=1)


def p(text):
    doc.add_paragraph(text)


def bold_lead(lead, rest):
    para = doc.add_paragraph()
    run = para.add_run(lead)
    run.bold = True
    para.add_run(rest)


# Title
doc.add_heading("Upbuild Services Agreement", level=0)

p(
    "This Agreement is entered into by and between: [Client Company Name] (Client) and "
    "Upbuild (Company) and is effective as of the date the Client signs below. The Company "
    "agrees to provide coaching, consulting, mediation, workshops, 360 Degree Reviews, and "
    "related leadership development services through one or more Upbuild coaches (each, a "
    "“Coach”) for the Client and/or its employees, leaders, and other designated participants. "
    "This Agreement supersedes all prior agreements between the parties."
)

h1("1. Description of Services")

bold_lead(
    "Coaching. ",
    "Coaching is a partnership between the Coach and Client Employees in a "
    "thought-provoking and creative process that inspires Client Employees to maximize "
    "personal and professional potential (Coaching)."
)
bold_lead(
    "360 Degree Reviews. ",
    "A “360” is a structured process facilitated by the Coach for the Client and/or "
    "Client Employees to get feedback from team members, peers, and other stakeholders "
    "(360 Degree Reviews)."
)
bold_lead(
    "Mediation. ",
    "Mediation is a structured and customized process facilitated by the Coach to help "
    "identify gaps in trust and rebuild it between or among Client Employees where it has "
    "eroded."
)
bold_lead(
    "Workshops. ",
    "Workshops are group learning opportunities facilitated by the Coach designed to develop "
    "applied knowledge in a specific area, build trust, foster connection, and unlock "
    "potential in the Client’s organization."
)
bold_lead(
    "Consulting. ",
    "Throughout the term of this Agreement, at times convenient to the Coach and the Client, "
    "and upon the Client’s request, the Coach may provide additional consulting and/or "
    "professional services based on the hourly rate below."
)

h1("2. Coach-Client Relationship")
p(
    "The Coach agrees to maintain the ethics and standards of behavior established by the "
    "International Coach Federation (coachfederation.org/ethics). The Client acknowledges "
    "that the Services may involve discussions relating to leadership, communication, "
    "decision-making, relationships, organizational dynamics, career development, and other "
    "personal and professional matters. The Client and its employees, leaders, and other "
    "participants remain solely responsible for their decisions, actions, and implementation "
    "of any insights, recommendations, or learnings arising from the Services. The Client "
    "acknowledges that the Services are not therapy and do not substitute for therapy, if "
    "needed. The Client also understands that the Services do not involve the diagnosis, "
    "treatment, or cure of mental disorders, medical conditions, or diseases, and are not a "
    "substitute for counseling, psychotherapy, substance abuse treatment, or other medical, "
    "mental health, legal, financial, or professional advice. Participants are encouraged to "
    "seek independent professional guidance, as appropriate."
)

h1("3. Fees, Expenses, and Cancellation Policy")
doc.add_paragraph("[Fee Section]")
doc.add_paragraph("[Cancellation Policy]")
p(
    "Invoices will be issued on a quarterly basis and will reflect actual Services delivered "
    "during the applicable period. Payment is due within 30 days of the invoice date."
)

h1("4. Termination")
p(
    "Either party may terminate this Agreement upon 15 days’ written notice. All fees for "
    "Services rendered will become due and payable within 30 days of termination and the "
    "parties’ rights and obligations described in this Agreement will survive the termination "
    "of the relationship. The Client agrees to compensate the Company for all Services "
    "rendered through the date of termination."
)
p(
    "When a Client Employee is ready to complete coaching, the Company requests that the "
    "Client Employee schedule a final coaching session as a “completion session” for closure."
)

h1("5. Confidentiality and Intellectual Property")
p(
    "The Company and the Client agree to keep confidential any non-public information shared "
    "in connection with the Services and to use such information only for purposes related to "
    "this engagement."
)
p(
    "The Company will treat information shared by the Client, its employees, and other "
    "participants in connection with the Services as confidential and will not disclose such "
    "information except as necessary to provide the Services, as permitted by this Agreement, "
    "or as required by law or applicable ethical obligations."
)
p(
    "For coaching engagements, the parties acknowledge that the effectiveness of coaching "
    "depends upon a relationship of trust and confidentiality. Accordingly, individual "
    "coaching conversations, communications, notes, assessments, and other information shared "
    "by a coaching participant will not be disclosed to the Client without that participant’s "
    "consent, except as required by law or applicable ethical obligations."
)
p(
    "The Company may share with the Client high-level observations, developmental themes, "
    "participation information, attendance information, progress updates, and "
    "recommendations related to the Services, provided that such information does not "
    "disclose confidential communications attributable to a specific individual without that "
    "individual’s consent."
)
p(
    "The Company may discuss Client matters internally with its coaches, employees, "
    "contractors, and advisors as reasonably necessary for training, supervision, quality "
    "assurance, and the delivery of Services, provided such individuals are subject to "
    "appropriate confidentiality obligations."
)
p(
    "The Client acknowledges that the Company may use proprietary methodologies, frameworks, "
    "assessments, exercises, workshop materials, presentations, and other intellectual "
    "property in connection with the Services (“Upbuild Materials”). All Upbuild Materials "
    "shall remain the sole property of the Company. The Client may use such materials "
    "internally for its own business purposes but may not reproduce, distribute, sell, "
    "publish, or provide them to third parties without the Company’s prior written consent."
)
p(
    "Confidential information does not include information that becomes publicly available "
    "through no fault of the receiving party or information independently developed without "
    "use of the other party’s confidential information."
)

h1("6. Independent Contractor")
p(
    "The relationship between the Company and the Client is that of independent contracting "
    "parties. Nothing in this Agreement shall be construed to create a partnership, joint "
    "venture, employment relationship, agency relationship, or other similar arrangement "
    "between the parties."
)
p(
    "The Company and its Coaches shall have sole discretion regarding the manner, method, and "
    "means by which the Services are provided. Nothing in this Agreement shall be interpreted "
    "as giving the Client authority to direct or control the professional judgment, coaching "
    "approach, or delivery of the Services by the Company or its Coaches."
)
p(
    "Neither the Company nor any Coach shall have authority to bind, obligate, or make "
    "representations on behalf of the Client unless expressly authorized in writing."
)
p(
    "The Company shall be solely responsible for all compensation, taxes, insurance, and other "
    "obligations relating to its employees, contractors, Coaches, and other representatives. "
    "The Client shall have no responsibility for any compensation, benefits, taxes, or other "
    "obligations relating to individuals providing Services on behalf of the Company."
)

h1("7. Limitation of Liability")
p(
    "The Services provided under this Agreement are intended to support learning, "
    "development, communication, leadership growth, and organizational effectiveness. While "
    "the Company will use reasonable care and professional judgment in delivering the "
    "Services, the Company does not guarantee any particular outcome, result, business "
    "performance, employment outcome, or organizational change."
)
p(
    "The Client acknowledges that decisions made by the Client, its employees, team members, "
    "and other participants remain their sole responsibility. The Company shall not be "
    "responsible for any actions taken or not taken by the Client or any participant based on "
    "the Services provided."
)
p(
    "To the fullest extent permitted by law, the Company shall not be liable for any indirect, "
    "incidental, special, consequential, punitive, or exemplary damages arising out of or "
    "relating to this Agreement or the Services, including, without limitation, lost profits, "
    "lost revenue, loss of business opportunity, or business interruption."
)
p(
    "Except in cases of the Company’s gross negligence, willful misconduct, or violation of "
    "applicable law, the Company’s total liability arising out of or relating to this "
    "Agreement or the Services shall not exceed the total fees paid by the Client under this "
    "Agreement."
)
p(
    "The Client agrees to indemnify and hold harmless the Company and its officers, employees, "
    "contractors, coaches, and representatives from any third-party claims, damages, "
    "liabilities, costs, or expenses arising from (i) the Client’s breach of this Agreement, "
    "(ii) the Client’s misuse of the Services, or (iii) the actions or omissions of the "
    "Client, its employees, contractors, or representatives, except to the extent resulting "
    "from the Company’s gross negligence or willful misconduct."
)

h1("8. Applicable Law and Attorneys’ Fees")
p(
    "This Agreement shall be governed by the laws of the State of New York. In the event of a "
    "legal dispute arising out of this Agreement, the prevailing party shall be entitled to "
    "recover its reasonable attorneys’ fees and costs from the non-prevailing party."
)

h1("9. Severability.")
p(
    "The parties intend this Agreement to be enforced to the fullest extent permissible under "
    "the laws and public policies. If any particular provision of this Agreement as "
    "adjudicated is invalid, prohibited or unenforceable for any reason, such provision shall "
    "be ineffective, without invalidating the remaining provisions of this Agreement or "
    "affecting the validity or enforceability of this Agreement."
)

doc.add_paragraph("")
p(
    "The Client’s authorized signature below indicates that the Client has read and agrees to "
    "abide by the terms and conditions of this Agreement.  The Client affirms and agrees that "
    "no promises or agreements which are not herein expressed have been made to the Client in "
    "executing this Agreement."
)

doc.add_paragraph("")
p("Client")
p("Printed Name: [Client Signer Name], on behalf of [Client Company Name]")
p("Signature:")
doc.add_paragraph("")
p("[!Sign.1.F,client, ,]")

doc.add_paragraph("")
p("Company")
p("Printed Name: Ariel Weiss, on behalf of Upbuild")
p("Signature:")
doc.add_paragraph("")
p("[!Sign.2.F,Partner, ,]")

doc.save(OUT)
print(f"Saved {OUT}")
