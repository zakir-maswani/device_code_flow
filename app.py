import streamlit as st
import requests
import json
from msal import PublicClientApplication
from datetime import datetime, timedelta, timezone
import os

# Try importing groq, handle if not available
try:
    from groq import Groq
    GROQ_AVAILABLE = True
except ImportError:
    GROQ_AVAILABLE = False

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------
# PAGE CONFIG (MUST BE FIRST)
# ---------------------------------------------------
st.set_page_config(
    page_title="Outlook Weekly Emails",
    page_icon="📧",
    layout="wide"
)

# ---------------------------------------------------
# LOAD SECRETS SAFELY
# ---------------------------------------------------
try:
    CLIENT_ID = st.secrets["CLIENT_ID"]
except KeyError:
    st.error("❌ CLIENT_ID not found in secrets!")
    st.info("Add CLIENT_ID to .streamlit/secrets.toml")
    st.stop()

# Load Groq API Key safely
GROQ_API_KEY = None
groq_client = None

try:
    GROQ_API_KEY = st.secrets.get("GROQ_API_KEY")
    if GROQ_API_KEY and GROQ_AVAILABLE:
        groq_client = Groq(api_key=GROQ_API_KEY)
except Exception as e:
    st.warning(f"⚠️ Groq not configured: {str(e)}")

# ---------------------------------------------------
# CONFIG
# ---------------------------------------------------
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Mail.Read"
]

app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# ---------------------------------------------------
# DOCX HELPER FUNCTIONS
# ---------------------------------------------------
def hex_rgb(rgb: RGBColor) -> str:
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"

def set_cell_bg(cell, rgb: RGBColor):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_rgb(rgb))
    tcPr.append(shd)

def paragraph(doc_or_cell, text, bold=False, italic=False,
              size_pt=10, color=RGBColor(26, 32, 44), align=WD_ALIGN_PARAGRAPH.LEFT):
    if hasattr(doc_or_cell, "paragraphs") and not hasattr(doc_or_cell, "add_paragraph"):
        if doc_or_cell.paragraphs:
            p = doc_or_cell.paragraphs[0]
            p.clear()
        else:
            p = doc_or_cell.add_paragraph()
    else:
        p = doc_or_cell.add_paragraph()
    p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = "Calibri"
    run.font.size = Pt(size_pt)
    run.font.color.rgb = color
    return p

# ---------------------------------------------------
# AI ANALYSIS
# ---------------------------------------------------
def analyze_email(email: dict) -> dict:
    """Analyze email using Groq Llama"""
    if not groq_client:
        return {
            "priority": "Medium",
            "summary": "AI analysis not available",
            "action_item": "Review manually.",
            "sentiment": "Neutral",
            "category": "Other"
        }
    
    try:
        subject = email.get("subject", "(no subject)")[:100]
        sender = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
        preview = email.get("bodyPreview", "")[:400]
        
        prompt = f"""Analyze this email and return ONLY valid JSON, no markdown:

Subject: {subject}
From: {sender}
Preview: {preview}

Return exactly:
{{
  "priority": "Critical|High|Medium|Low",
  "summary": "2-3 sentences about what this email is about",
  "action_item": "One next step, or 'No action required'",
  "sentiment": "Positive|Neutral|Negative|Urgent",
  "category": "Meeting|Finance|Support|Project|Legal|Other"
}}"""
        
        res = groq_client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
        )
        
        raw = res.choices[0].message.content.strip()
        raw = raw.replace("```json", "").replace("```", "").strip()
        return json.loads(raw)
    
    except Exception as e:
        return {
            "priority": "Medium",
            "summary": "Could not analyze",
            "action_item": "Review manually.",
            "sentiment": "Neutral",
            "category": "Other"
        }

# ---------------------------------------------------
# DOCX GENERATION
# ---------------------------------------------------
def generate_docx(emails):
    """Generate professional docx report"""
    try:
        doc = Document()
        
        # Setup page
        section = doc.sections[0]
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        
        # Title
        title = doc.add_paragraph()
        title_run = title.add_run("📧 Weekly Email Intelligence Report")
        title_run.font.size = Pt(24)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(13, 27, 42)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Date
        date_para = doc.add_paragraph()
        date_run = date_para.add_run(f"Generated: {datetime.utcnow().strftime('%d %b %Y')}")
        date_run.font.size = Pt(10)
        date_run.font.color.rgb = RGBColor(107, 114, 128)
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph()
        
        # Summary
        summary_heading = doc.add_paragraph()
        summary_run = summary_heading.add_run("SUMMARY")
        summary_run.font.size = Pt(12)
        summary_run.font.bold = True
        summary_run.font.color.rgb = RGBColor(75, 85, 99)
        
        doc.add_paragraph(f"Total emails analyzed: {len(emails)}")
        
        critical = sum(1 for e in emails if e.get("analysis", {}).get("priority") == "Critical")
        high = sum(1 for e in emails if e.get("analysis", {}).get("priority") == "High")
        
        doc.add_paragraph(f"Critical priority: {critical}")
        doc.add_paragraph(f"High priority: {high}")
        
        doc.add_paragraph()
        
        # Email Details
        details_heading = doc.add_paragraph()
        details_run = details_heading.add_run("EMAIL DETAILS")
        details_run.font.size = Pt(12)
        details_run.font.bold = True
        details_run.font.color.rgb = RGBColor(75, 85, 99)
        
        for idx, email in enumerate(emails, 1):
            email_para = doc.add_paragraph()
            email_para.paragraph_format.left_indent = Inches(0.25)
            
            # Email header
            subject = email.get('subject', 'No Subject')[:80]
            subj_run = email_para.add_run(f"{idx}. {subject}\n")
            subj_run.font.bold = True
            subj_run.font.size = Pt(11)
            
            analysis = email.get("analysis", {})
            
            # From
            from_addr = email.get('from', {}).get('emailAddress', {}).get('address', 'Unknown')
            from_run = email_para.add_run(f"From: {from_addr}\n")
            from_run.font.size = Pt(9)
            
            # Priority
            priority = analysis.get("priority", "Medium")
            priority_color_map = {
                "Critical": RGBColor(239, 68, 68),
                "High": RGBColor(245, 158, 11),
                "Medium": RGBColor(16, 185, 129),
                "Low": RGBColor(148, 163, 184),
            }
            priority_run = email_para.add_run(f"Priority: {priority} | ")
            priority_run.font.size = Pt(9)
            priority_run.font.color.rgb = priority_color_map.get(priority, RGBColor(75, 85, 99))
            priority_run.font.bold = True
            
            # Sentiment
            sentiment = analysis.get("sentiment", "Neutral")
            sentiment_run = email_para.add_run(f"Sentiment: {sentiment}\n")
            sentiment_run.font.size = Pt(9)
            
            # Category
            category = analysis.get("category", "Other")
            category_run = email_para.add_run(f"Category: {category}\n")
            category_run.font.size = Pt(9)
            
            # Summary
            summary = analysis.get("summary", "No summary available")
            summary_run = email_para.add_run(f"Summary: {summary}\n")
            summary_run.font.size = Pt(9)
            summary_run.italic = True
            
            # Action Item
            action = analysis.get("action_item", "No action required")
            action_run = email_para.add_run(f"Action: {action}\n")
            action_run.font.size = Pt(9)
            action_run.font.color.rgb = RGBColor(26, 86, 219)
            action_run.font.bold = True
            
            # Preview
            preview = email.get("bodyPreview", "")[:200]
            preview_run = email_para.add_run(f"Preview: {preview}...\n")
            preview_run.font.size = Pt(9)
            preview_run.font.color.rgb = RGBColor(107, 114, 128)
        
        # Save
        output_path = "/tmp/weekly_email_report.docx"
        doc.save(output_path)
        return output_path
    
    except Exception as e:
        st.error(f"Error generating document: {str(e)}")
        return None

# ---------------------------------------------------
# UI
# ---------------------------------------------------
st.title("📧 Outlook Weekly Email Dashboard")

with st.sidebar:
    st.title("⚙️ Dashboard")
    st.markdown("---")
    st.markdown("### Features")
    st.markdown("""
    - Secure Microsoft login
    - Fetch last 7 days emails
    - AI analysis (Groq/Llama)
    - Download as Word document
    """)

# Login
if "access_token" not in st.session_state:
    st.info("🔓 Login with your Microsoft account to continue")
    
    if st.button("🔐 Login with Microsoft"):
        try:
            flow = app.initiate_device_flow(scopes=SCOPES)
            
            if "user_code" not in flow:
                st.error("❌ Device flow failed")
                st.stop()
            
            st.markdown("### 👇 Complete Login")
            st.code(flow["user_code"])
            st.markdown(f"Visit: **{flow['verification_uri']}** and enter the code above.")
            
            with st.spinner("Waiting for login..."):
                result = app.acquire_token_by_device_flow(flow)
            
            if "access_token" in result:
                st.session_state["access_token"] = result["access_token"]
                st.success("✅ Login successful")
                st.rerun()
            else:
                st.error(f"❌ {result.get('error_description', 'Login failed')}")
        
        except Exception as e:
            st.error(f"❌ {str(e)}")

# After Login
else:
    access_token = st.session_state["access_token"]
    
    col1, col2 = st.columns([3, 1])
    with col1:
        fetch_btn = st.button("📥 Fetch Last 7 Days & Generate Report")
    with col2:
        logout_btn = st.button("🚪 Logout")
    
    if logout_btn:
        st.session_state.clear()
        st.rerun()
    
    if fetch_btn:
        with st.spinner("📬 Fetching emails..."):
            try:
                last_week = (
                    datetime.now(timezone.utc) - timedelta(days=7)
                ).strftime("%Y-%m-%dT%H:%M:%SZ")
                
                url = (
                    "https://graph.microsoft.com/v1.0/me/messages"
                    f"?$filter=receivedDateTime ge {last_week}"
                    "&$top=50"
                    "&$orderby=receivedDateTime DESC"
                )
                
                headers = {"Authorization": f"Bearer {access_token}"}
                response = requests.get(url, headers=headers, timeout=20)
                
                if response.status_code == 200:
                    emails_raw = response.json().get("value", [])
                    st.success(f"✅ Fetched {len(emails_raw)} emails")
                    
                    if not emails_raw:
                        st.info("No emails found in last 7 days")
                    else:
                        # Analyze emails
                        with st.spinner("🤖 Analyzing emails with AI..."):
                            emails_with_analysis = []
                            progress_bar = st.progress(0)
                            
                            for idx, email in enumerate(emails_raw):
                                analysis = analyze_email(email)
                                emails_with_analysis.append({
                                    "subject": email.get("subject", "No Subject"),
                                    "from": email.get("from", {}),
                                    "receivedDateTime": email.get("receivedDateTime", ""),
                                    "bodyPreview": email.get("bodyPreview", ""),
                                    "analysis": analysis
                                })
                                progress = (idx + 1) / len(emails_raw)
                                progress_bar.progress(progress)
                        
                        st.success(f"✅ Analyzed {len(emails_with_analysis)} emails")
                        
                        # Display emails
                        st.markdown("---")
                        st.subheader("📧 Email Analysis")
                        
                        for email in emails_with_analysis[:10]:
                            with st.expander(f"📩 {email['subject'][:60]}"):
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.write(f"**From:** {email['from'].get('emailAddress', {}).get('address', 'Unknown')}")
                                    st.write(f"**Date:** {email['receivedDateTime'][:16]}")
                                with col2:
                                    analysis = email['analysis']
                                    st.write(f"**Priority:** {analysis.get('priority', 'N/A')}")
                                    st.write(f"**Sentiment:** {analysis.get('sentiment', 'N/A')}")
                                
                                st.write(f"**Category:** {analysis.get('category', 'N/A')}")
                                st.write(f"**Summary:** {analysis.get('summary', 'N/A')}")
                                st.write(f"**Action:** {analysis.get('action_item', 'No action required')}")
                                st.write(f"**Preview:** {email['bodyPreview'][:150]}...")
                        
                        # Generate DOCX
                        st.markdown("---")
                        with st.spinner("📄 Generating Word document..."):
                            docx_path = generate_docx(emails_with_analysis)
                        
                        if docx_path and os.path.exists(docx_path):
                            with open(docx_path, "rb") as f:
                                st.download_button(
                                    label="📥 Download Report (DOCX)",
                                    data=f.read(),
                                    file_name=f"email_report_{datetime.now().strftime('%Y%m%d')}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                            st.success("✅ Document generated successfully!")
                        else:
                            st.error("❌ Failed to generate document")
                
                elif response.status_code == 401:
                    st.error("❌ Session expired. Please login again.")
                    st.session_state.clear()
                else:
                    st.error(f"❌ Error: {response.status_code}")
            
            except requests.exceptions.Timeout:
                st.error("⏱️ Request timeout")
            except requests.exceptions.ConnectionError:
                st.error("🌐 Network connection error")
            except Exception as e:
                st.error(f"❌ {str(e)}")
