import streamlit as st
import requests
import html
import json

from msal import PublicClientApplication
from datetime import datetime, timedelta, timezone
from groq import Groq

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------
# CONFIG
# ---------------------------------------------------
CLIENT_ID = st.secrets["CLIENT_ID"]
GROQ_API_KEY = st.secrets["GROQ_API_KEY"]

AUTHORITY = "https://login.microsoftonline.com/consumers"

SCOPES = [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Mail.Read"
]

app = PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

groq_client = Groq(api_key=GROQ_API_KEY)

# ---------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------
st.set_page_config(
    page_title="Outlook AI Report Generator",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Outlook AI Report Generator")

st.write(
    "Generate AI-powered DOCX reports from your Outlook emails."
)

# ---------------------------------------------------
# COLOR SYSTEM
# ---------------------------------------------------
NAVY = RGBColor(0x0D, 0x1B, 0x2A)
BLUE = RGBColor(0x1A, 0x56, 0xDB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT = RGBColor(0x1A, 0x20, 0x2C)
MUTED = RGBColor(0x6B, 0x7A, 0x8E)
SURFACE = RGBColor(0xF8, 0xFA, 0xFC)

RED = RGBColor(0xEF, 0x44, 0x44)
GREEN = RGBColor(0x10, 0xB9, 0x81)
AMBER = RGBColor(0xF5, 0x9E, 0x0B)
SLATE = RGBColor(0x94, 0xA3, 0xB8)

RED_BG = RGBColor(0xFE, 0xF2, 0xF2)
GREEN_BG = RGBColor(0xF0, 0xFD, 0xF4)
AMBER_BG = RGBColor(0xFF, 0xFB, 0xEB)
SLATE_BG = RGBColor(0xF8, 0xFA, 0xFC)

RED_TXT = RGBColor(0x99, 0x1B, 0x1B)
GREEN_TXT = RGBColor(0x06, 0x5F, 0x46)
AMBER_TXT = RGBColor(0x92, 0x40, 0x0E)
SLATE_TXT = RGBColor(0x4A, 0x55, 0x68)

TOTAL_DXA = 9026

PRIORITY_THEME = {
    "Critical": (RED_BG, RED, RED_TXT),
    "High": (AMBER_BG, AMBER, AMBER_TXT),
    "Medium": (GREEN_BG, GREEN, GREEN_TXT),
    "Low": (SLATE_BG, SLATE, SLATE_TXT),
}

# ---------------------------------------------------
# DOCX HELPERS
# ---------------------------------------------------
def hex_rgb(rgb):
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def set_cell_bg(cell, rgb):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_rgb(rgb))
    tcPr.append(shd)


def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')

    for m in [('top', top), ('left', left), ('bottom', bottom), ('right', right)]:
        node = OxmlElement(f'w:{m[0]}')
        node.set(qn('w:w'), str(m[1]))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)

    tcPr.append(tcMar)


def set_col_widths(table, widths):
    table.autofit = False

    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width


def add_horizontal_rule(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:color'), 'D1D5DB')
    pBdr.append(bottom)
    pPr.append(pBdr)


def badge_cell(cell, text, bg, txt):
    set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text.upper())
    run.bold = True
    run.font.size = Pt(8)
    run.font.color.rgb = txt


def priority_badge(cell, priority):
    bg, _, txt = PRIORITY_THEME.get(priority, PRIORITY_THEME['Low'])
    badge_cell(cell, priority, bg, txt)

# ---------------------------------------------------
# AI ANALYSIS
# ---------------------------------------------------
def analyze_email(email: dict) -> dict:

    prompt = f"""
You are an executive email analyst.

Analyse the following email and return ONLY valid JSON.

Subject: {email.get("subject", "(no subject)")}

From: {
    email.get("from", {})
    .get("emailAddress", {})
    .get("address", "Unknown")
}

Preview: {email.get("bodyPreview", "")[:500]}

Return JSON in this exact format:

{{
  "priority": "Critical|High|Medium|Low",
  "summary": "Short professional summary",
  "action_item": "One action item or No action required",
  "sentiment": "Positive|Neutral|Negative|Urgent",
  "category": "Meeting|Finance|Support|Project|Legal|Other"
}}
"""

    try:

        response = groq_client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.2
        )

        raw = response.choices[0].message.content.strip()

        raw = raw.replace("```json", "")
        raw = raw.replace("```", "")
        raw = raw.strip()

        return json.loads(raw)

    except Exception:

        return {
            "priority": "Medium",
            "summary": "AI summary could not be generated.",
            "action_item": "Review manually.",
            "sentiment": "Neutral",
            "category": "Other"
        }

# ---------------------------------------------------
# WEEKLY OVERVIEW
# ---------------------------------------------------
def weekly_overview(emails, analyses):

    critical = sum(
        1 for a in analyses
        if a.get("priority") == "Critical"
    )

    high = sum(
        1 for a in analyses
        if a.get("priority") == "High"
    )

    actions = [
        a.get("action_item")
        for a in analyses
        if a.get("action_item") != "No action required"
    ]

    prompt = f"""
Write a professional executive overview paragraph.

Emails analysed: {len(emails)}

Critical emails: {critical}

High priority emails: {high}

Key actions:
{' ; '.join(actions[:5]) if actions else 'None'}

Return plain professional paragraph only.
"""

    try:

        response = groq_client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.3
        )

        return response.choices[0].message.content.strip()

    except Exception:

        return "Overview generation failed."

# ---------------------------------------------------
# PROFESSIONAL DOCX GENERATOR
# ---------------------------------------------------
def generate_docx(emails, analyses, overview_text):

    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(2)

    # ---------------------------------------------------
    # COVER HEADER
    # ---------------------------------------------------
    tbl = doc.add_table(rows=3, cols=1)
    tbl.style = "Table Grid"
    set_col_widths(tbl, [Inches(6.8)])

    c0 = tbl.cell(0, 0)
    set_cell_bg(c0, NAVY)
    r = c0.paragraphs[0].add_run("OUTLOOK AI REPORT")
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = BLUE

    c1 = tbl.cell(1, 0)
    set_cell_bg(c1, NAVY)
    r = c1.paragraphs[0].add_run(
        "Weekly Email Intelligence Report"
    )
    r.bold = True
    r.font.size = Pt(24)
    r.font.color.rgb = WHITE

    c2 = tbl.cell(2, 0)
    set_cell_bg(c2, NAVY)
    r = c2.paragraphs[0].add_run(
        datetime.utcnow().strftime('%d %b %Y')
    )
    r.font.size = Pt(10)
    r.font.color.rgb = WHITE

    doc.add_paragraph()

    # ---------------------------------------------------
    # KPI SECTION
    # ---------------------------------------------------
    critical = sum(
        1 for a in analyses
        if a.get("priority") == "Critical"
    )

    high = sum(
        1 for a in analyses
        if a.get("priority") == "High"
    )

    actions = sum(
        1 for a in analyses
        if a.get("action_item") != "No action required"
    )

    kpi = doc.add_table(rows=2, cols=4)
    kpi.style = "Table Grid"
    set_col_widths(kpi, [Inches(1.7), Inches(1.7), Inches(1.7), Inches(1.7)])

    values = [
        (str(len(emails)), "Emails", BLUE),
        (str(critical), "Critical", RED),
        (str(high), "High Priority", AMBER),
        (str(actions), "Action Items", GREEN)
    ]

    for idx, (num, label, color) in enumerate(values):

        cell1 = kpi.cell(0, idx)
        cell2 = kpi.cell(1, idx)

        set_cell_bg(cell1, WHITE)
        set_cell_bg(cell2, WHITE)

        p1 = cell1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER

        r1 = p1.add_run(num)
        r1.bold = True
        r1.font.size = Pt(22)
        r1.font.color.rgb = color

        p2 = cell2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

        r2 = p2.add_run(label)
        r2.font.size = Pt(9)
        r2.font.color.rgb = MUTED

    doc.add_paragraph()

    # ---------------------------------------------------
    # OVERVIEW
    # ---------------------------------------------------
    h = doc.add_paragraph()
    rh = h.add_run("EXECUTIVE OVERVIEW")
    rh.bold = True
    rh.font.size = Pt(12)
    rh.font.color.rgb = NAVY

    add_horizontal_rule(doc)

    ov = doc.add_table(rows=1, cols=2)
    ov.style = "Table Grid"
    set_col_widths(ov, [Inches(0.15), Inches(6.65)])

    left = ov.cell(0, 0)
    right = ov.cell(0, 1)

    set_cell_bg(left, BLUE)
    set_cell_bg(right, WHITE)

    left.width = Inches(0.1)

    pr = right.paragraphs[0]
    rr = pr.add_run(overview_text)
    rr.font.size = Pt(10)
    rr.font.color.rgb = TEXT

    doc.add_paragraph()

    # ---------------------------------------------------
    # EMAIL ANALYSIS
    # ---------------------------------------------------
    heading = doc.add_paragraph()
    rh = heading.add_run("EMAIL ANALYSIS")
    rh.bold = True
    rh.font.size = Pt(12)
    rh.font.color.rgb = NAVY

    add_horizontal_rule(doc)

    for idx, (email, analysis) in enumerate(
        zip(emails, analyses),
        start=1
    ):

        subject = email.get("subject", "No Subject")

        sender = (
            email.get("from", {})
            .get("emailAddress", {})
            .get("address", "Unknown")
        )

        received = email.get(
            "receivedDateTime",
            ""
        )

        preview = email.get(
            "bodyPreview",
            ""
        )[:300]

        priority = analysis.get("priority", "Medium")

        # Header card
        card = doc.add_table(rows=1, cols=3)
        card.style = "Table Grid"
        set_col_widths(card, [Inches(0.6), Inches(5.2), Inches(1.0)])

        c1 = card.cell(0, 0)
        c2 = card.cell(0, 1)
        c3 = card.cell(0, 2)

        set_cell_bg(c1, SURFACE)
        set_cell_bg(c2, SURFACE)

        r = c1.paragraphs[0].add_run(f"#{idx}")
        r.bold = True
        r.font.color.rgb = MUTED

        r = c2.paragraphs[0].add_run(subject)
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = TEXT

        priority_badge(c3, priority)

        # Meta table
        meta = doc.add_table(rows=2, cols=4)
        meta.style = "Table Grid"
        set_col_widths(meta, [Inches(1.8), Inches(1.8), Inches(1.6), Inches(1.6)])

        labels = [
            "FROM",
            "DATE",
            "CATEGORY",
            "SENTIMENT"
        ]

        values = [
            sender,
            received,
            analysis.get("category"),
            analysis.get("sentiment")
        ]

        for i in range(4):

            top = meta.cell(0, i)
            bottom = meta.cell(1, i)

            set_cell_bg(top, SURFACE)
            set_cell_bg(bottom, WHITE)

            rt = top.paragraphs[0].add_run(labels[i])
            rt.bold = True
            rt.font.size = Pt(8)
            rt.font.color.rgb = MUTED

            rb = bottom.paragraphs[0].add_run(str(values[i]))
            rb.font.size = Pt(9)
            rb.font.color.rgb = TEXT

        # Body content
        body = doc.add_table(rows=3, cols=2)
        body.style = "Table Grid"
        set_col_widths(body, [Inches(1.7), Inches(5.1)])

        rows = [
            (
                "AI SUMMARY",
                analysis.get("summary")
            ),
            (
                "ACTION REQUIRED",
                analysis.get("action_item")
            ),
            (
                "EMAIL PREVIEW",
                preview
            )
        ]

        for r_idx, (label, value) in enumerate(rows):

            lc = body.cell(r_idx, 0)
            vc = body.cell(r_idx, 1)

            set_cell_bg(lc, SURFACE)
            set_cell_bg(vc, WHITE)

            rl = lc.paragraphs[0].add_run(label)
            rl.bold = True
            rl.font.size = Pt(8)
            rl.font.color.rgb = MUTED

            rv = vc.paragraphs[0].add_run(str(value))
            rv.font.size = Pt(9)
            rv.font.color.rgb = TEXT

        doc.add_paragraph()

    # ---------------------------------------------------
    # ACTION ITEMS PAGE
    # ---------------------------------------------------
    doc.add_page_break()

    action_heading = doc.add_paragraph()
    ra = action_heading.add_run("ACTION ITEMS SUMMARY")
    ra.bold = True
    ra.font.size = Pt(14)
    ra.font.color.rgb = NAVY

    add_horizontal_rule(doc)

    actions_table = doc.add_table(rows=1, cols=4)
    actions_table.style = "Table Grid"
    set_col_widths(actions_table, [Inches(2.6), Inches(1.0), Inches(1.2), Inches(2.0)])

    headers = [
        "Subject",
        "Priority",
        "Category",
        "Action"
    ]

    for idx, h in enumerate(headers):

        cell = actions_table.cell(0, idx)

        set_cell_bg(cell, NAVY)

        r = cell.paragraphs[0].add_run(h)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = WHITE

    for email, analysis in zip(emails, analyses):

        if analysis.get("action_item") == "No action required":
            continue

        row = actions_table.add_row().cells

        row[0].text = email.get("subject", "")[:50]
        row[1].text = analysis.get("priority", "")
        row[2].text = analysis.get("category", "")
        row[3].text = analysis.get("action_item", "")

    output_file = "weekly_ai_report.docx"

    doc.save(output_file)

    return output_file

# ---------------------------------------------------
# LOGIN
# ---------------------------------------------------
if "access_token" not in st.session_state:

    st.info(
        "🔓 Login with your Microsoft account to continue"
    )

    if st.button("🔐 Login with Microsoft"):

        try:

            flow = app.initiate_device_flow(
                scopes=SCOPES
            )

            if "user_code" not in flow:

                st.error("❌ Device flow failed")

                st.stop()

            st.markdown("### 👇 Complete Login")

            st.code(flow["user_code"])

            st.markdown(
                f"""
Visit: **{flow['verification_uri']}**

Enter the code above.
"""
            )

            with st.spinner(
                "Waiting for login..."
            ):

                result = app.acquire_token_by_device_flow(
                    flow
                )

            if "access_token" in result:

                st.session_state["access_token"] = (
                    result["access_token"]
                )

                st.success(
                    "✅ Login successful"
                )

                st.rerun()

            else:

                st.error(
                    result.get(
                        "error_description",
                        "Login failed"
                    )
                )

        except Exception as e:

            st.error(f"❌ {str(e)}")

# ---------------------------------------------------
# AFTER LOGIN
# ---------------------------------------------------
else:

    access_token = st.session_state["access_token"]

    col1, col2 = st.columns([3, 1])

    with col1:
        generate_btn = st.button(
            "📄 Generate AI Report"
        )

    with col2:
        logout_btn = st.button(
            "🚪 Logout"
        )

    # Logout
    if logout_btn:

        st.session_state.clear()

        st.rerun()

    # Generate Report
    if generate_btn:

        try:

            with st.spinner(
                "📬 Fetching emails..."
            ):

                last_week = (
                    datetime.now(timezone.utc)
                    - timedelta(days=7)
                ).strftime("%Y-%m-%dT%H:%M:%SZ")

                url = (
                    "https://graph.microsoft.com/v1.0/me/messages"
                    f"?$filter=receivedDateTime ge {last_week}"
                    "&$top=20"
                    "&$orderby=receivedDateTime DESC"
                )

                headers = {
                    "Authorization":
                    f"Bearer {access_token}"
                }

                response = requests.get(
                    url,
                    headers=headers,
                    timeout=20
                )

                if response.status_code == 401:

                    st.error(
                        "❌ Session expired. Please login again."
                    )

                    st.session_state.clear()

                    st.stop()

                if response.status_code != 200:

                    st.error(
                        f"❌ Graph API Error: "
                        f"{response.status_code}"
                    )

                    try:
                        st.json(response.json())
                    except Exception:
                        st.write(response.text)

                    st.stop()

                emails = response.json().get(
                    "value",
                    []
                )

            st.success(
                f"✅ {len(emails)} emails fetched"
            )

            if not emails:

                st.info(
                    "No emails found in last 7 days"
                )

                st.stop()

            analyses = []

            progress = st.progress(0)

            with st.spinner(
                "🤖 Analysing emails using AI..."
            ):

                total = len(emails)

                for idx, email in enumerate(emails):

                    result = analyze_email(email)

                    analyses.append(result)

                    progress.progress(
                        (idx + 1) / total
                    )

            with st.spinner(
                "🧠 Generating executive overview..."
            ):

                overview = weekly_overview(
                    emails,
                    analyses
                )

            with st.spinner(
                "📄 Generating DOCX report..."
            ):

                report_file = generate_docx(
                    emails,
                    analyses,
                    overview
                )

            st.success(
                "✅ AI Report Generated Successfully"
            )

            st.subheader("📩 Email Preview")

            for email, analysis in zip(
                emails[:5],
                analyses[:5]
            ):

                subject = html.escape(
                    email.get(
                        "subject",
                        "No Subject"
                    )
                )

                sender = html.escape(
                    email.get("from", {})
                    .get("emailAddress", {})
                    .get("address", "Unknown")
                )

                summary = analysis.get(
                    "summary",
                    ""
                )

                st.markdown(
                    f"""
### 📩 {subject}

**From:** {sender}

**AI Summary:** {summary}

---
"""
                )

            with open(report_file, "rb") as file:

                st.download_button(
                    label="⬇️ Download DOCX Report",
                    data=file,
                    file_name=report_file,
                    mime=(
                        "application/vnd.openxmlformats-officedocument"
                        ".wordprocessingml.document"
                    )
                )

        except requests.exceptions.Timeout:

            st.error("⏱️ Request timeout")

        except requests.exceptions.ConnectionError:

            st.error("🌐 Network connection error")

        except Exception as e:

            st.error(f"❌ {str(e)}")
