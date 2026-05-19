import streamlit as st
import requests
import html
import json

from msal import PublicClientApplication
from datetime import datetime, timedelta, timezone
from groq import Groq

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------
# CONFIG
# ---------------------------------------------------
CLIENT_ID = st.secrets.get("CLIENT_ID", "test_client")
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", "test_key")

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
    "🚀 Generate professional AI-powered DOCX reports from your Outlook emails with automatic formatting."
)

# ---------------------------------------------------
# COLOR SYSTEM
# ---------------------------------------------------
NAVY      = RGBColor(0x0D, 0x1B, 0x2A)
BLUE      = RGBColor(0x1A, 0x56, 0xDB)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
TEXT      = RGBColor(0x1A, 0x20, 0x2C)
MUTED     = RGBColor(0x6B, 0x7A, 0x8E)
SURFACE   = RGBColor(0xF8, 0xFA, 0xFC)
LIGHT_BLU = RGBColor(0xDB, 0xEA, 0xFE)

RED       = RGBColor(0xEF, 0x44, 0x44)
GREEN     = RGBColor(0x10, 0xB9, 0x81)
AMBER     = RGBColor(0xF5, 0x9E, 0x0B)
SLATE     = RGBColor(0x94, 0xA3, 0xB8)
PURPLE    = RGBColor(0x8B, 0x5C, 0xF6)
TEAL      = RGBColor(0x06, 0xB6, 0xD4)

RED_BG    = RGBColor(0xFE, 0xF2, 0xF2)
GREEN_BG  = RGBColor(0xF0, 0xFD, 0xF4)
AMBER_BG  = RGBColor(0xFF, 0xFB, 0xEB)
SLATE_BG  = RGBColor(0xF8, 0xFA, 0xFC)

RED_TXT   = RGBColor(0x99, 0x1B, 0x1B)
GREEN_TXT = RGBColor(0x06, 0x5F, 0x46)
AMBER_TXT = RGBColor(0x92, 0x40, 0x0E)
SLATE_TXT = RGBColor(0x4A, 0x55, 0x68)
BLUE_TXT  = RGBColor(0x1E, 0x40, 0xAF)

PRIORITY_THEME = {
    "Critical": (RED_BG,   RED,   RED_TXT),
    "High":     (AMBER_BG, AMBER, AMBER_TXT),
    "Medium":   (GREEN_BG, GREEN, GREEN_TXT),
    "Low":      (SLATE_BG, SLATE, SLATE_TXT),
}

TOTAL_DXA = 9026   # A4 with 1.5 cm margins each side

# ---------------------------------------------------
# DOCX HELPERS
# ---------------------------------------------------
def hex_rgb(rgb):
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def set_cell_bg(cell, rgb):
    tcPr = cell._tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),  "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_rgb(rgb))
    tcPr.append(shd)


def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for name, val in [("top", top), ("left", left), ("bottom", bottom), ("right", right)]:
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_col_widths(table, widths):
    table.autofit = False
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width


def add_horizontal_rule(doc, color="1A56DB", size=8):
    p   = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    str(size))
    bot.set(qn("w:color"), color)
    bot.set(qn("w:space"), "1")
    pBdr.append(bot)
    pPr.append(pBdr)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(6)


def badge_cell(cell, text, bg, txt):
    set_cell_bg(cell, bg)
    p   = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text.upper())
    run.bold           = True
    run.font.size      = Pt(8)
    run.font.color.rgb = txt


def priority_badge(cell, priority):
    bg, _, txt = PRIORITY_THEME.get(priority, PRIORITY_THEME["Low"])
    badge_cell(cell, priority, bg, txt)


def add_kv_row(table, label, value, label_bg=SURFACE, value_bg=WHITE, font_size=9):
    """Add a label/value pair row to an existing table."""
    row = table.add_row().cells
    set_cell_bg(row[0], label_bg)
    set_cell_bg(row[1], value_bg)
    set_cell_margins(row[0], top=60, bottom=60, left=100, right=80)
    set_cell_margins(row[1], top=60, bottom=60, left=100, right=100)

    rl = row[0].paragraphs[0].add_run(label)
    rl.bold           = True
    rl.font.size      = Pt(8)
    rl.font.color.rgb = MUTED

    rv = row[1].paragraphs[0]
    rv.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    rv.word_wrap = True
    r  = rv.add_run(str(value))
    r.font.size      = Pt(font_size)
    r.font.color.rgb = TEXT


# ---------------------------------------------------
# AI ANALYSIS  (enhanced prompt, robust parsing)
# ---------------------------------------------------
def analyze_email(email: dict) -> dict:

    subject  = email.get("subject", "(no subject)")
    sender   = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
    preview  = email.get("bodyPreview", "")[:600]
    is_read  = email.get("isRead", True)
    importance = email.get("importance", "normal")
    has_attachments = email.get("hasAttachments", False)

    prompt = f"""
You are a senior executive email analyst. Analyse this email and return ONLY a valid JSON object — no markdown, no explanation.

Email Details:
- Subject: {subject}
- From: {sender}
- Read status: {"Read" if is_read else "Unread"}
- Importance flag: {importance}
- Has attachments: {has_attachments}
- Body preview: {preview}

Return this exact JSON structure:
{{
  "priority": "Critical|High|Medium|Low",
  "summary": "2-3 sentence professional summary of the email",
  "action_item": "Specific next step required, or 'No action required'",
  "sentiment": "Positive|Neutral|Negative|Urgent",
  "category": "Meeting|Finance|Support|Project|Legal|HR|Marketing|IT|Operations|Other",
  "deadline_hint": "Any date/deadline mentioned or 'None detected'",
  "key_people": "Comma-separated names/emails mentioned or 'None'",
  "risk_level": "High|Medium|Low|None",
  "response_needed": "Yes|No",
  "response_timeframe": "Immediate|Within 24h|Within a week|Not required"
}}
"""

    for model in ["llama-3.3-70b-versatile", "llama-3.1-70b-versatile", "llama-3.1-8b-instant"]:
        try:
            response = groq_client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=500,
            )
            raw = response.choices[0].message.content.strip()
            # Strip any accidental markdown fences
            raw = raw.replace("```json", "").replace("```", "").strip()
            # Extract first JSON object found
            start = raw.find("{")
            end   = raw.rfind("}") + 1
            if start != -1 and end > start:
                return json.loads(raw[start:end])
        except Exception:
            continue

    # Fallback
    return {
        "priority": "Medium",
        "summary": "AI summary could not be generated. Please review manually.",
        "action_item": "Review manually.",
        "sentiment": "Neutral",
        "category": "Other",
        "deadline_hint": "None detected",
        "key_people": "None",
        "risk_level": "None",
        "response_needed": "No",
        "response_timeframe": "Not required",
    }


# ---------------------------------------------------
# WEEKLY OVERVIEW  (robust with fallback models)
# ---------------------------------------------------
def weekly_overview(emails, analyses):

    critical  = sum(1 for a in analyses if a.get("priority") == "Critical")
    high      = sum(1 for a in analyses if a.get("priority") == "High")
    unread    = sum(1 for e in emails   if not e.get("isRead", True))
    responses = sum(1 for a in analyses if a.get("response_needed") == "Yes")
    actions   = [a.get("action_item", "") for a in analyses
                 if a.get("action_item") not in ("No action required", "", None)]

    # Category breakdown
    categories = {}
    for a in analyses:
        cat = a.get("category", "Other")
        categories[cat] = categories.get(cat, 0) + 1
    top_cats = ", ".join(
        f"{k} ({v})" for k, v in sorted(categories.items(), key=lambda x: -x[1])[:3]
    )

    prompt = f"""
You are a C-suite executive assistant. Write a concise professional executive overview paragraph (3-5 sentences) for an email intelligence report.

Data:
- Total emails analysed: {len(emails)}
- Unread emails: {unread}
- Critical priority: {critical}
- High priority: {high}
- Responses needed: {responses}
- Top categories: {top_cats}
- Key action items: {'; '.join(actions[:5]) if actions else 'None outstanding'}

Write only the paragraph text. No headers, no bullet points, no preamble.
"""

    for model in ["llama-3.3-70b-versatile", "llama-3.1-70b-versatile", "llama-3.1-8b-instant"]:
        try:
            response = groq_client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=300,
            )
            text = response.choices[0].message.content.strip()
            if text:
                return text
        except Exception:
            continue

    # Deterministic fallback — always succeeds
    return (
        f"This report covers {len(emails)} emails received over the past 7 days, "
        f"of which {unread} remain unread. "
        f"There are {critical} critical and {high} high-priority items requiring attention, "
        f"with {responses} emails awaiting a response. "
        f"Key topics include {top_cats if top_cats else 'general correspondence'}. "
        f"Immediate action is recommended on the items flagged Critical or High priority."
    )


# ---------------------------------------------------
# PROFESSIONAL DOCX GENERATOR
# ---------------------------------------------------
def generate_docx(emails, analyses, overview_text):

    doc = Document()

    # Page setup — A4
    section = doc.sections[0]
    section.page_width    = Cm(21)
    section.page_height   = Cm(29.7)
    section.left_margin   = Cm(1.5)
    section.right_margin  = Cm(1.5)
    section.top_margin    = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    # ── COVER HEADER ──────────────────────────────────────────────────────────
    tbl = doc.add_table(rows=4, cols=1)
    tbl.style = "Table Grid"
    set_col_widths(tbl, [Inches(7.0)])

    # Badge row
    c0 = tbl.cell(0, 0)
    set_cell_bg(c0, NAVY)
    set_cell_margins(c0, top=100, bottom=40, left=150, right=150)
    r = c0.paragraphs[0].add_run("OUTLOOK  ·  AI EMAIL INTELLIGENCE")
    r.bold = True; r.font.size = Pt(9); r.font.color.rgb = BLUE

    # Title row
    c1 = tbl.cell(1, 0)
    set_cell_bg(c1, NAVY)
    set_cell_margins(c1, top=100, bottom=80, left=150, right=150)
    p1 = c1.paragraphs[0]
    p1.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p1.add_run("Weekly Email Intelligence Report")
    r.bold = True; r.font.size = Pt(26); r.font.color.rgb = WHITE

    # Subtitle row
    c2 = tbl.cell(2, 0)
    set_cell_bg(c2, NAVY)
    set_cell_margins(c2, top=40, bottom=60, left=150, right=150)
    r = c2.paragraphs[0].add_run("Powered by AI — Confidential")
    r.font.size = Pt(10); r.font.color.rgb = MUTED; r.italic = True

    # Date row
    c3 = tbl.cell(3, 0)
    set_cell_bg(c3, BLUE)
    set_cell_margins(c3, top=70, bottom=70, left=150, right=150)
    date_str = datetime.utcnow().strftime("%d %B %Y  |  %H:%M UTC")
    r = c3.paragraphs[0].add_run(f"Generated: {date_str}")
    r.bold = True; r.font.size = Pt(10); r.font.color.rgb = WHITE

    doc.add_paragraph()

    # ── KPI DASHBOARD ─────────────────────────────────────────────────────────
    critical_count  = sum(1 for a in analyses if a.get("priority") == "Critical")
    high_count      = sum(1 for a in analyses if a.get("priority") == "High")
    action_count    = sum(1 for a in analyses if a.get("action_item") not in ("No action required", "", None))
    unread_count    = sum(1 for e in emails   if not e.get("isRead", True))
    response_count  = sum(1 for a in analyses if a.get("response_needed") == "Yes")
    attachment_count = sum(1 for e in emails  if e.get("hasAttachments", False))

    kpi = doc.add_table(rows=2, cols=6)
    kpi.style = "Table Grid"
    col_w = Inches(1.167)
    set_col_widths(kpi, [col_w] * 6)

    kpi_data = [
        (str(len(emails)),    "Total Emails",    BLUE),
        (str(unread_count),   "Unread",          PURPLE),
        (str(critical_count), "Critical",        RED),
        (str(high_count),     "High Priority",   AMBER),
        (str(action_count),   "Action Items",    GREEN),
        (str(response_count), "Need Response",   TEAL),
    ]

    for idx, (num, label, color) in enumerate(kpi_data):
        top_c = kpi.cell(0, idx)
        bot_c = kpi.cell(1, idx)
        set_cell_bg(top_c, WHITE); set_cell_bg(bot_c, SURFACE)
        set_cell_margins(top_c, top=90, bottom=40, left=60, right=60)
        set_cell_margins(bot_c, top=40, bottom=80, left=60, right=60)

        p = top_c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(num)
        r.bold = True; r.font.size = Pt(22); r.font.color.rgb = color

        p2 = bot_c.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(label)
        r2.font.size = Pt(8); r2.font.color.rgb = MUTED

    doc.add_paragraph()

    # ── REPORT METADATA ───────────────────────────────────────────────────────
    # Category breakdown table
    categories = {}
    sentiments  = {}
    for a in analyses:
        cat = a.get("category", "Other")
        sent = a.get("sentiment", "Neutral")
        categories[cat]  = categories.get(cat, 0)  + 1
        sentiments[sent] = sentiments.get(sent, 0) + 1

    ph = doc.add_paragraph()
    ph.paragraph_format.space_before = Pt(6)
    r = ph.add_run("REPORT METADATA")
    r.bold = True; r.font.size = Pt(11); r.font.color.rgb = NAVY
    add_horizontal_rule(doc, color="1A56DB", size=6)

    meta_tbl = doc.add_table(rows=1, cols=2)
    meta_tbl.style = "Table Grid"
    set_col_widths(meta_tbl, [Inches(3.5), Inches(3.5)])

    # Left: category breakdown
    lc = meta_tbl.cell(0, 0)
    rc = meta_tbl.cell(0, 1)
    set_cell_bg(lc, SURFACE); set_cell_bg(rc, SURFACE)
    set_cell_margins(lc, top=80, bottom=80, left=100, right=60)
    set_cell_margins(rc, top=80, bottom=80, left=60, right=100)

    lp = lc.paragraphs[0]
    lr = lp.add_run("Email Categories\n")
    lr.bold = True; lr.font.size = Pt(9); lr.font.color.rgb = NAVY

    for cat, cnt in sorted(categories.items(), key=lambda x: -x[1]):
        pct = int(cnt / len(emails) * 100) if emails else 0
        lc.add_paragraph().add_run(f"  {cat}: {cnt}  ({pct}%)").font.size = Pt(8)

    rp = rc.paragraphs[0]
    rr = rp.add_run("Sentiment Distribution\n")
    rr.bold = True; rr.font.size = Pt(9); rr.font.color.rgb = NAVY

    for sent, cnt in sorted(sentiments.items(), key=lambda x: -x[1]):
        pct = int(cnt / len(emails) * 100) if emails else 0
        rc.add_paragraph().add_run(f"  {sent}: {cnt}  ({pct}%)").font.size = Pt(8)

    doc.add_paragraph()

    # ── EXECUTIVE OVERVIEW ────────────────────────────────────────────────────
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(6)
    rh = h.add_run("EXECUTIVE OVERVIEW")
    rh.bold = True; rh.font.size = Pt(13); rh.font.color.rgb = NAVY
    add_horizontal_rule(doc, color="1A56DB", size=8)

    ov = doc.add_table(rows=1, cols=2)
    ov.style = "Table Grid"
    set_col_widths(ov, [Inches(0.15), Inches(6.85)])

    left  = ov.cell(0, 0); right = ov.cell(0, 1)
    set_cell_bg(left, BLUE); set_cell_bg(right, WHITE)
    set_cell_margins(right, top=100, bottom=100, left=120, right=120)
    left.width = Inches(0.1)

    pr = right.paragraphs[0]
    pr.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    rr = pr.add_run(overview_text)
    rr.font.size = Pt(11); rr.font.color.rgb = TEXT

    doc.add_paragraph()

    # ── EMAIL ANALYSIS ────────────────────────────────────────────────────────
    heading = doc.add_paragraph()
    heading.paragraph_format.space_before = Pt(12)
    rh = heading.add_run("DETAILED EMAIL ANALYSIS")
    rh.bold = True; rh.font.size = Pt(13); rh.font.color.rgb = NAVY
    add_horizontal_rule(doc, color="1A56DB", size=8)

    for idx, (email, analysis) in enumerate(zip(emails, analyses), start=1):

        subject   = email.get("subject", "No Subject")[:120]
        sender_addr = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
        sender_name = email.get("from", {}).get("emailAddress", {}).get("name", "")
        display_sender = f"{sender_name} <{sender_addr}>" if sender_name else sender_addr
        received  = email.get("receivedDateTime", "")
        preview   = email.get("bodyPreview", "")[:350]
        priority  = analysis.get("priority", "Medium")

        # Parse date nicely
        try:
            dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
            received_display = dt.strftime("%d %b %Y  %H:%M UTC")
            day_of_week      = dt.strftime("%A")
        except Exception:
            received_display = received[:19]
            day_of_week      = ""

        is_read         = email.get("isRead", True)
        importance      = email.get("importance", "normal").capitalize()
        has_attachments = email.get("hasAttachments", False)
        flag_status     = email.get("flag", {}).get("flagStatus", "notFlagged")
        conversation_id = email.get("conversationId", "")[:20]
        categories_tags = ", ".join(email.get("categories", [])) or "None"
        web_link        = email.get("webLink", "")

        # ── Card header (index + subject + priority badge) ──
        card = doc.add_table(rows=1, cols=3)
        card.style = "Table Grid"
        set_col_widths(card, [Inches(0.55), Inches(5.55), Inches(0.9)])

        c_idx = card.cell(0, 0); c_sub = card.cell(0, 1); c_pri = card.cell(0, 2)

        # Index badge background matches priority
        bg_color, _, _ = PRIORITY_THEME.get(priority, PRIORITY_THEME["Low"])
        set_cell_bg(c_idx, bg_color)
        set_cell_bg(c_sub, NAVY)
        set_cell_margins(c_idx, top=80, bottom=80, left=60, right=60)
        set_cell_margins(c_sub, top=80, bottom=80, left=120, right=100)

        ri = c_idx.paragraphs[0]
        ri.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r  = ri.add_run(f"#{idx}")
        r.bold = True; r.font.size = Pt(12); r.font.color.rgb = NAVY

        ps = c_sub.paragraphs[0]
        ps.word_wrap = True
        rs = ps.add_run(subject)
        rs.bold = True; rs.font.size = Pt(11); rs.font.color.rgb = WHITE

        priority_badge(c_pri, priority)

        # ── Status strip (read/unread, importance, attachments, flag) ──
        status_bg = RED_BG if not is_read else GREEN_BG
        status_strip = doc.add_table(rows=1, cols=4)
        status_strip.style = "Table Grid"
        set_col_widths(status_strip, [Inches(1.75), Inches(1.75), Inches(1.75), Inches(1.75)])

        status_items = [
            ("READ STATUS",   "🔴 UNREAD" if not is_read else "✅ Read",   RED if not is_read else GREEN),
            ("IMPORTANCE",    importance,                                   RED if importance == "High" else MUTED),
            ("ATTACHMENTS",   "📎 Yes" if has_attachments else "No",       BLUE if has_attachments else MUTED),
            ("FLAGGED",       "🚩 Flagged" if flag_status == "flagged" else "Not flagged", AMBER if flag_status == "flagged" else MUTED),
        ]

        for si, (slabel, sval, scolor) in enumerate(status_items):
            sc = status_strip.cell(0, si)
            set_cell_bg(sc, SURFACE)
            set_cell_margins(sc, top=50, bottom=50, left=80, right=80)
            sp = sc.paragraphs[0]
            rl = sp.add_run(f"{slabel}\n")
            rl.bold = True; rl.font.size = Pt(7); rl.font.color.rgb = MUTED
            rv = sp.add_run(sval)
            rv.bold = True; rv.font.size = Pt(9); rv.font.color.rgb = scolor

        # ── Sender / Date / Category / Sentiment meta ──
        meta = doc.add_table(rows=2, cols=4)
        meta.style = "Table Grid"
        set_col_widths(meta, [Inches(2.1), Inches(1.7), Inches(1.6), Inches(1.6)])

        meta_labels  = ["FROM",     "RECEIVED",         "CATEGORY",                      "SENTIMENT"]
        meta_values  = [
            display_sender[:55],
            f"{received_display}\n({day_of_week})",
            analysis.get("category", "Other"),
            analysis.get("sentiment", "Neutral"),
        ]

        for i in range(4):
            top = meta.cell(0, i); bot = meta.cell(1, i)
            set_cell_bg(top, SURFACE); set_cell_bg(bot, WHITE)
            set_cell_margins(top, top=50, bottom=30, left=80, right=80)
            set_cell_margins(bot, top=50, bottom=60, left=80, right=80)

            rt = top.paragraphs[0].add_run(meta_labels[i])
            rt.bold = True; rt.font.size = Pt(7); rt.font.color.rgb = MUTED

            rb = bot.paragraphs[0]
            rb.word_wrap = True
            r  = rb.add_run(str(meta_values[i]))
            r.font.size = Pt(9); r.font.color.rgb = TEXT

        # ── AI Content body ──
        body_rows = [
            ("AI SUMMARY",       analysis.get("summary", "")[:220]),
            ("ACTION REQUIRED",  analysis.get("action_item", "")[:150]),
            ("EMAIL PREVIEW",    preview),
        ]

        body = doc.add_table(rows=len(body_rows), cols=2)
        body.style = "Table Grid"
        set_col_widths(body, [Inches(1.6), Inches(5.4)])

        for r_idx, (lbl, val) in enumerate(body_rows):
            lc2 = body.cell(r_idx, 0); vc2 = body.cell(r_idx, 1)
            set_cell_bg(lc2, SURFACE); set_cell_bg(vc2, WHITE)
            set_cell_margins(lc2, top=70, bottom=70, left=100, right=80)
            set_cell_margins(vc2, top=70, bottom=70, left=100, right=100)

            rl2 = lc2.paragraphs[0].add_run(lbl)
            rl2.bold = True; rl2.font.size = Pt(8); rl2.font.color.rgb = MUTED

            rv2 = vc2.paragraphs[0]
            rv2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            rv2.word_wrap = True
            r3 = rv2.add_run(str(val))
            r3.font.size = Pt(9); r3.font.color.rgb = TEXT

        # ── Extended intelligence row ──
        intel_rows = [
            ("DEADLINE HINT",     analysis.get("deadline_hint", "None detected")),
            ("KEY PEOPLE",        analysis.get("key_people", "None")),
            ("RISK LEVEL",        analysis.get("risk_level", "None")),
            ("RESPONSE NEEDED",   f"{analysis.get('response_needed','No')}  —  {analysis.get('response_timeframe','Not required')}"),
            ("OUTLOOK CATEGORIES", categories_tags),
        ]

        intel = doc.add_table(rows=len(intel_rows), cols=2)
        intel.style = "Table Grid"
        set_col_widths(intel, [Inches(1.6), Inches(5.4)])

        for r_idx, (lbl, val) in enumerate(intel_rows):
            lc3 = intel.cell(r_idx, 0); vc3 = intel.cell(r_idx, 1)
            alt_bg = LIGHT_BLU if r_idx % 2 == 0 else WHITE
            set_cell_bg(lc3, SURFACE); set_cell_bg(vc3, alt_bg)
            set_cell_margins(lc3, top=55, bottom=55, left=100, right=80)
            set_cell_margins(vc3, top=55, bottom=55, left=100, right=100)

            rl3 = lc3.paragraphs[0].add_run(lbl)
            rl3.bold = True; rl3.font.size = Pt(8); rl3.font.color.rgb = MUTED

            rv3 = vc3.paragraphs[0]
            rv3.word_wrap = True
            r4  = rv3.add_run(str(val))
            r4.font.size = Pt(9); r4.font.color.rgb = TEXT

        doc.add_paragraph()

    # ── ACTION ITEMS SUMMARY ──────────────────────────────────────────────────
    doc.add_page_break()

    ah = doc.add_paragraph()
    ah.paragraph_format.space_before = Pt(12)
    ra = ah.add_run("ACTION ITEMS SUMMARY")
    ra.bold = True; ra.font.size = Pt(14); ra.font.color.rgb = NAVY
    add_horizontal_rule(doc, color="1A56DB", size=8)

    act_tbl = doc.add_table(rows=1, cols=5)
    act_tbl.style = "Table Grid"
    set_col_widths(act_tbl, [Inches(2.3), Inches(0.9), Inches(1.0), Inches(1.0), Inches(1.8)])

    for i, h_txt in enumerate(["Subject", "Priority", "Category", "Response", "Action Required"]):
        cell = act_tbl.cell(0, i)
        set_cell_bg(cell, NAVY)
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        r = cell.paragraphs[0].add_run(h_txt)
        r.bold = True; r.font.size = Pt(9); r.font.color.rgb = WHITE

    for email, analysis in zip(emails, analyses):
        if analysis.get("action_item") in ("No action required", "", None):
            continue
        row = act_tbl.add_row().cells
        for cell in row:
            set_cell_margins(cell, top=65, bottom=65, left=100, right=100)

        vals = [
            email.get("subject", "")[:55],
            analysis.get("priority", ""),
            analysis.get("category", ""),
            analysis.get("response_timeframe", ""),
            analysis.get("action_item", "")[:70],
        ]
        colors = [TEXT, RED if vals[1] == "Critical" else (AMBER if vals[1] == "High" else TEXT), TEXT, TEXT, TEXT]

        for i, (val, col) in enumerate(zip(vals, colors)):
            p = row[i].paragraphs[0]
            p.word_wrap = True
            if i in (1, 2, 3):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(9); r.font.color.rgb = col
            if i == 1 and vals[1] in ("Critical", "High"):
                r.bold = True

    # ── UNREAD EMAILS PAGE ────────────────────────────────────────────────────
    unread_emails = [(e, a) for e, a in zip(emails, analyses) if not e.get("isRead", True)]

    if unread_emails:
        doc.add_page_break()

        uh = doc.add_paragraph()
        uh.paragraph_format.space_before = Pt(12)
        ru = uh.add_run(f"UNREAD EMAILS  ({len(unread_emails)} items)")
        ru.bold = True; ru.font.size = Pt(14); ru.font.color.rgb = RED
        add_horizontal_rule(doc, color="EF4444", size=8)

        un_tbl = doc.add_table(rows=1, cols=4)
        un_tbl.style = "Table Grid"
        set_col_widths(un_tbl, [Inches(2.5), Inches(1.5), Inches(0.9), Inches(2.1)])

        for i, h_txt in enumerate(["Subject", "From", "Priority", "AI Summary"]):
            cell = un_tbl.cell(0, i)
            set_cell_bg(cell, RED)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            r = cell.paragraphs[0].add_run(h_txt)
            r.bold = True; r.font.size = Pt(9); r.font.color.rgb = WHITE

        for email, analysis in unread_emails:
            row = un_tbl.add_row().cells
            for cell in row:
                set_cell_bg(cell, RED_BG)
                set_cell_margins(cell, top=65, bottom=65, left=100, right=100)

            un_vals = [
                email.get("subject", "")[:60],
                email.get("from", {}).get("emailAddress", {}).get("address", "")[:35],
                analysis.get("priority", ""),
                analysis.get("summary", "")[:100],
            ]
            for i, val in enumerate(un_vals):
                p = row[i].paragraphs[0]; p.word_wrap = True
                r = p.add_run(val)
                r.font.size = Pt(9); r.font.color.rgb = RED_TXT
                if i == 2: r.bold = True

    # ── REPORT FOOTER ─────────────────────────────────────────────────────────
    doc.add_page_break()

    footer_h = doc.add_paragraph()
    rf = footer_h.add_run("REPORT NOTES & DISCLAIMER")
    rf.bold = True; rf.font.size = Pt(12); rf.font.color.rgb = NAVY
    add_horizontal_rule(doc, color="1A56DB", size=6)

    notes = [
        "This report was generated automatically using AI-powered analysis of your Outlook inbox.",
        "Email summaries and action items are AI-generated and should be reviewed before acting upon them.",
        "Priority classifications are based on AI analysis of email content, sender, and context clues.",
        "Sentiment analysis reflects the perceived tone of email content and may not be 100% accurate.",
        f"Report generated: {datetime.utcnow().strftime('%d %B %Y at %H:%M UTC')}",
        f"Emails analysed: {len(emails)}  |  Period: Last 7 days",
    ]
    for note in notes:
        np = doc.add_paragraph()
        np.paragraph_format.space_before = Pt(2)
        np.paragraph_format.space_after  = Pt(2)
        nr = np.add_run(f"• {note}")
        nr.font.size = Pt(9); nr.font.color.rgb = MUTED

    output_file = "weekly_ai_report.docx"
    doc.save(output_file)
    return output_file


# ---------------------------------------------------
# LOGIN
# ---------------------------------------------------
if "access_token" not in st.session_state:

    st.info("🔓 Login with your Microsoft account to continue")

    if st.button("🔐 Login with Microsoft", use_container_width=True):
        try:
            flow = app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                st.error("❌ Device flow failed")
                st.stop()

            st.markdown("### 👇 Complete Login")
            st.code(flow["user_code"])
            st.markdown(f"Visit: **{flow['verification_uri']}**\n\nEnter the code above.")

            with st.spinner("Waiting for login..."):
                result = app.acquire_token_by_device_flow(flow)

            if "access_token" in result:
                st.session_state["access_token"] = result["access_token"]
                st.success("✅ Login successful")
                st.rerun()
            else:
                st.error(result.get("error_description", "Login failed"))
        except Exception as e:
            st.error(f"❌ {str(e)}")

# ---------------------------------------------------
# AFTER LOGIN
# ---------------------------------------------------
else:
    access_token = st.session_state["access_token"]

    col1, col2 = st.columns([3, 1])
    with col1:
        generate_btn = st.button("📄 Generate AI Report", use_container_width=True)
    with col2:
        logout_btn = st.button("🚪 Logout", use_container_width=True)

    if logout_btn:
        st.session_state.clear()
        st.rerun()

    if generate_btn:
        try:
            with st.spinner("📬 Fetching emails..."):
                last_week = (datetime.now(timezone.utc) - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
                url = (
                    "https://graph.microsoft.com/v1.0/me/messages"
                    f"?$filter=receivedDateTime ge {last_week}"
                    "&$top=20"
                    "&$orderby=receivedDateTime DESC"
                    "&$select=subject,from,receivedDateTime,bodyPreview,isRead,importance,"
                    "hasAttachments,flag,categories,conversationId,webLink"
                )
                headers = {"Authorization": f"Bearer {access_token}"}
                response = requests.get(url, headers=headers, timeout=20)

                if response.status_code == 401:
                    st.error("❌ Session expired. Please login again.")
                    st.session_state.clear()
                    st.stop()

                if response.status_code != 200:
                    st.error(f"❌ Graph API Error: {response.status_code}")
                    try:
                        st.json(response.json())
                    except Exception:
                        st.write(response.text)
                    st.stop()

                emails = response.json().get("value", [])

            st.success(f"✅ {len(emails)} emails fetched")

            if not emails:
                st.info("No emails found in last 7 days")
                st.stop()

            analyses = []
            progress = st.progress(0)

            with st.spinner("🤖 Analysing emails using AI..."):
                for idx, email in enumerate(emails):
                    result = analyze_email(email)
                    analyses.append(result)
                    progress.progress((idx + 1) / len(emails))

            with st.spinner("🧠 Generating executive overview..."):
                overview = weekly_overview(emails, analyses)

            with st.spinner("📄 Generating professional DOCX report..."):
                report_file = generate_docx(emails, analyses, overview)

            st.success("✅ AI Report Generated Successfully — Download below.")

            with open(report_file, "rb") as file:
                st.download_button(
                    label="⬇️ Download DOCX Report",
                    data=file,
                    file_name=report_file,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        except requests.exceptions.Timeout:
            st.error("⏱️ Request timeout")
        except requests.exceptions.ConnectionError:
            st.error("🌐 Network connection error")
        except Exception as e:
            st.error(f"❌ {str(e)}")
