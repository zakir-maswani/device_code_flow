import streamlit as st
import requests
import html
import json
import os
from msal import PublicClientApplication
from datetime import datetime, timedelta, timezone
from groq import Groq

from docx import Document
from docx.shared import Pt

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
st.write("Generate professional AI-powered Outlook email reports.")

# ---------------------------------------------------
# AI ANALYSIS
# ---------------------------------------------------
def analyze_email(email: dict) -> dict:

    prompt = f"""
You are an executive email analyst.

Analyse the email and return ONLY valid JSON.

Subject: {email.get('subject', '(no subject)')}
From: {email.get('from', {}).get('emailAddress', {}).get('address', 'Unknown')}
Preview: {email.get('bodyPreview', '')[:500]}

Return:
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

        raw = response.choices[0].message.content
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
# OVERVIEW
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
Write a professional executive email digest.

Emails analysed: {len(emails)}
Critical emails: {critical}
High priority emails: {high}
Key actions: {'; '.join(actions[:5]) if actions else 'None'}

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
# DOCX GENERATOR
# ---------------------------------------------------
def generate_docx(emails, analyses, overview_text):

    doc = Document()

    title = doc.add_heading(
        "Weekly Outlook AI Report",
        level=1
    )

    title.runs[0].font.size = Pt(24)

    doc.add_paragraph(
        f"Generated: {datetime.utcnow().strftime('%d %b %Y %H:%M UTC')}"
    )

    doc.add_paragraph(
        f"Total Emails Analysed: {len(emails)}"
    )

    doc.add_heading("Executive Overview", level=2)
    doc.add_paragraph(overview_text)

    doc.add_heading("Email Analysis", level=2)

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

        received = email.get("receivedDateTime", "")

        preview = email.get("bodyPreview", "")

        doc.add_heading(f"#{idx} - {subject}", level=3)

        doc.add_paragraph(f"From: {sender}")
        doc.add_paragraph(f"Received: {received}")
        doc.add_paragraph(f"Priority: {analysis.get('priority')}")
        doc.add_paragraph(f"Category: {analysis.get('category')}")
        doc.add_paragraph(f"Sentiment: {analysis.get('sentiment')}")

        doc.add_paragraph(
            f"Summary: {analysis.get('summary')}"
        )

        doc.add_paragraph(
            f"Action Item: {analysis.get('action_item')}"
        )

        doc.add_paragraph(
            f"Preview: {preview}"
        )

        doc.add_paragraph("-" * 80)

    output_file = "weekly_ai_report.docx"

    doc.save(output_file)

    return output_file

# ---------------------------------------------------
# LOGIN
# ---------------------------------------------------
if "access_token" not in st.session_state:

    st.info("🔐 Login with Microsoft")

    if st.button("Login with Microsoft"):

        try:

            flow = app.initiate_device_flow(scopes=SCOPES)

            if "user_code" not in flow:
                st.error("Device flow failed")
                st.stop()

            st.code(flow["user_code"])

            st.write(
                f"Visit: {flow['verification_uri']}"
            )

            with st.spinner("Waiting for authentication..."):

                result = app.acquire_token_by_device_flow(flow)

            if "access_token" in result:

                st.session_state["access_token"] = result[
                    "access_token"
                ]

                st.success("Login successful")

                st.rerun()

            else:

                st.error(
                    result.get(
                        "error_description",
                        "Login failed"
                    )
                )

        except Exception as e:
            st.error(str(e))

# ---------------------------------------------------
# AFTER LOGIN
# ---------------------------------------------------
else:

    access_token = st.session_state["access_token"]

    col1, col2 = st.columns(2)

    with col1:
        generate_btn = st.button(
            "📄 Generate AI Report"
        )

    with col2:
        logout_btn = st.button("🚪 Logout")

    if logout_btn:

        st.session_state.clear()
        st.rerun()

    if generate_btn:

        try:

            with st.spinner("Fetching emails..."):

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
                    "Authorization": f"Bearer {access_token}"
                }

                response = requests.get(
                    url,
                    headers=headers,
                    timeout=20
                )

                if response.status_code != 200:

                    st.error(
                        f"Graph API Error: {response.status_code}"
                    )

                    st.stop()

                emails = response.json().get("value", [])

            st.success(f"{len(emails)} emails fetched")

            analyses = []

            progress = st.progress(0)

            with st.spinner("Analysing emails using AI..."):

                for idx, email in enumerate(emails):

                    result = analyze_email(email)

                    analyses.append(result)

                    progress.progress(
                        (idx + 1) / len(emails)
                    )

            with st.spinner("Generating DOCX report..."):

                overview = weekly_overview(
                    emails,
                    analyses
                )

                report_file = generate_docx(
                    emails,
                    analyses,
                    overview
                )

            st.success("AI Report Generated Successfully")

            # Preview emails
            st.subheader("Email Preview")

            for email, analysis in zip(emails[:5], analyses[:5]):

                subject = html.escape(
                    email.get("subject", "No Subject")
                )

                sender = html.escape(
                    email.get("from", {})
                    .get("emailAddress", {})
                    .get("address", "Unknown")
                )

                summary = analysis.get("summary")

                st.markdown(f"""
### 📩 {subject}
**From:** {sender}

**AI Summary:** {summary}

---
""")

            # Download button
            with open(report_file, "rb") as file:

                st.download_button(
                    label="⬇️ Download DOCX Report",
                    data=file,
                    file_name=report_file,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except requests.exceptions.Timeout:
            st.error("Request timeout")

        except requests.exceptions.ConnectionError:
            st.error("Network error")

        except Exception as e:
            st.error(str(e))
