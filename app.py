import streamlit as st
import requests
from msal import PublicClientApplication, SerializableTokenCache
from datetime import datetime

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Outlook Mail",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

    section[data-testid="stSidebar"] { background: #0f1117; border-right: 1px solid #1e2130; }
    section[data-testid="stSidebar"] * { color: #c9d1e0 !important; }

    .email-card {
        background: #16192a; border: 1px solid #1e2540; border-radius: 12px;
        padding: 18px 22px; margin-bottom: 12px;
        transition: border-color 0.2s, box-shadow 0.2s; position: relative;
    }
    .email-card:hover { border-color: #3b82f6; box-shadow: 0 0 0 1px #3b82f620; }
    .email-card.unread { border-left: 3px solid #3b82f6; background: #141829; }
    .email-card.urgent { border-left: 3px solid #ef4444; }

    .email-subject { font-size: 15px; font-weight: 600; color: #e8eaf0; margin: 0 0 4px 0; line-height: 1.4; }
    .email-subject.read { font-weight: 400; color: #8b92a5; }

    .email-meta {
        display: flex; gap: 16px; font-size: 12px; color: #5a6478;
        font-family: 'DM Mono', monospace; margin-bottom: 10px; flex-wrap: wrap;
    }
    .email-preview {
        font-size: 13px; color: #6b7485; line-height: 1.6;
        border-top: 1px solid #1e2540; padding-top: 10px; margin-top: 6px;
    }

    .badge { display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 11px; font-weight: 500; font-family: 'DM Mono', monospace; }
    .badge-unread { background: #1e3a5f; color: #60a5fa; }
    .badge-urgent { background: #3b1010; color: #f87171; }
    .badge-attach { background: #1a2e1a; color: #4ade80; }

    .open-link a { font-size: 12px; color: #3b82f6 !important; text-decoration: none; font-family: 'DM Mono', monospace; }
    .open-link a:hover { text-decoration: underline; }

    .stat-card { background: #16192a; border: 1px solid #1e2540; border-radius: 10px; padding: 16px 20px; text-align: center; }
    .stat-number { font-size: 28px; font-weight: 600; color: #e8eaf0; }
    .stat-label  { font-size: 12px; color: #5a6478; margin-top: 2px; font-family: 'DM Mono', monospace; }

    .stButton > button {
        background: #2563eb; color: white; border: none; border-radius: 8px;
        padding: 10px 20px; font-family: 'DM Sans', sans-serif; font-weight: 500; transition: background 0.2s;
    }
    .stButton > button:hover { background: #1d4ed8; color: white; }

    #MainMenu { visibility: hidden; }
    footer     { visibility: hidden; }
    header     { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
CLIENT_ID = st.secrets["CLIENT_ID"]
TENANT_ID = st.secrets["TENANT_ID"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES    = ["Mail.Read", "User.Read"]

# ─────────────────────────────────────────────
# MSAL — persistent token cache in session
# ─────────────────────────────────────────────
def get_msal_app():
    """Build MSAL app using a serializable cache stored in session_state."""
    cache = SerializableTokenCache()
    if "token_cache" in st.session_state:
        cache.deserialize(st.session_state["token_cache"])

    msal_app = PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )
    return msal_app, cache


def save_cache(cache):
    """Persist cache back to session_state after any token operation."""
    if cache.has_state_changed:
        st.session_state["token_cache"] = cache.serialize()


def get_valid_token():
    """
    Return a fresh access token.
    1. Try silent acquire (uses refresh token automatically).
    2. If that fails, clear session and ask user to re-login.
    """
    msal_app, cache = get_msal_app()
    accounts = msal_app.get_accounts()

    if accounts:
        result = msal_app.acquire_token_silent(SCOPES, account=accounts[0])
        save_cache(cache)
        if result and "access_token" in result:
            # Update stored token
            st.session_state["access_token"] = result["access_token"]
            return result["access_token"], None

    # Silent failed — force re-login
    for k in ["access_token", "user_info", "emails", "stats", "next_link"]:
        st.session_state.pop(k, None)
    return None, "Session expired. Please sign in again."

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def get_user_info(token):
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {token}"}
    )
    return (r.json(), None) if r.status_code == 200 else (None, r.text)


def get_emails(token, top=10, folder="all", search="", only_unread=False):
    folder_map = {
        "all":    "me/messages",
        "inbox":  "me/mailFolders/inbox/messages",
        "sent":   "me/mailFolders/sentitems/messages",
        "drafts": "me/mailFolders/drafts/messages",
        "junk":   "me/mailFolders/junkemail/messages",
    }
    path = folder_map.get(folder, "me/messages")
    filters = []
    if only_unread:
        filters.append("isRead eq false")

    url = f"https://graph.microsoft.com/v1.0/{path}?$top={top}&$orderby=receivedDateTime desc"
    if filters:
        url += f"&$filter={' and '.join(filters)}"
    if search:
        url += f"&$search=\"{search}\""

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        **({"ConsistencyLevel": "eventual"} if search else {})
    }
    r = requests.get(url, headers=headers)

    if r.status_code == 200:
        data = r.json()
        return data.get("value", []), data.get("@odata.nextLink"), None
    elif r.status_code == 401:
        # Token invalid — try silent refresh then retry once
        new_token, err = get_valid_token()
        if err:
            return [], None, err
        # Retry with fresh token
        headers["Authorization"] = f"Bearer {new_token}"
        r2 = requests.get(url, headers=headers)
        if r2.status_code == 200:
            data = r2.json()
            return data.get("value", []), data.get("@odata.nextLink"), None
        return [], None, f"Error {r2.status_code}: {r2.text}"
    elif r.status_code == 403:
        return [], None, "403: Mail.Read permission not granted."
    else:
        try:
            msg = r.json().get("error", {}).get("message", r.text)
        except Exception:
            msg = r.text
        return [], None, f"Error {r.status_code}: {msg}"


def format_date(dt_str):
    if not dt_str:
        return ""
    try:
        dt  = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        now = datetime.now(dt.tzinfo)
        diff = now - dt
        if diff.days == 0:   return dt.strftime("Today %I:%M %p")
        if diff.days == 1:   return dt.strftime("Yesterday %I:%M %p")
        if diff.days < 7:    return dt.strftime("%A %I:%M %p")
        return dt.strftime("%b %d, %Y")
    except Exception:
        return dt_str[:10]


def render_email_card(email):
    subject     = email.get("subject") or "(No Subject)"
    sender_obj  = email.get("from", {}).get("emailAddress", {})
    sender_name = sender_obj.get("name", "")
    sender_addr = sender_obj.get("address", "Unknown")
    preview     = email.get("bodyPreview", "").replace("\r\n", " ").strip()
    time_str    = format_date(email.get("receivedDateTime", ""))
    is_read     = email.get("isRead", True)
    has_attach  = email.get("hasAttachments", False)
    web_link    = email.get("webLink", "")
    importance  = email.get("importance", "normal")

    card_class = "email-card" + (" unread" if not is_read else "") + (" urgent" if importance == "high" else "")
    subj_class = "email-subject" + ("" if not is_read else " read")
    display_name = sender_name if sender_name else sender_addr

    badges = ""
    if not is_read:          badges += '<span class="badge badge-unread">● UNREAD</span> '
    if importance == "high": badges += '<span class="badge badge-urgent">↑ URGENT</span> '
    if has_attach:           badges += '<span class="badge badge-attach">⊕ ATTACHMENT</span> '

    open_btn     = f'<div class="open-link"><a href="{web_link}" target="_blank">Open in Outlook ↗</a></div>' if web_link else ""
    preview_html = f'<div class="email-preview">{preview[:300]}{"…" if len(preview) > 300 else ""}</div>' if preview else ""

    st.markdown(f"""
    <div class="{card_class}">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px;">
            <div style="flex:1; min-width:0;">
                <div style="margin-bottom:6px;">{badges}</div>
                <div class="{subj_class}">{subject}</div>
                <div class="email-meta">
                    <span>👤 {display_name}</span>
                    <span>✉ {sender_addr}</span>
                    <span>🕐 {time_str}</span>
                </div>
            </div>
            <div style="flex-shrink:0;">{open_btn}</div>
        </div>
        {preview_html}
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# LOGIN PAGE
# ─────────────────────────────────────────────
if "access_token" not in st.session_state:
    st.markdown("""
    <div style="max-width:420px; margin:80px auto; background:#16192a; border:1px solid #1e2540;
                border-radius:16px; padding:48px 40px; text-align:center;">
        <div style="font-size:48px; margin-bottom:16px;">📧</div>
        <div style="font-size:26px; font-weight:600; color:#e8eaf0; margin-bottom:8px;">Outlook Mail</div>
        <div style="font-size:14px; color:#5a6478; margin-bottom:32px;">
            Sign in with your Microsoft account to access your inbox
        </div>
    </div>
    """, unsafe_allow_html=True)

    col = st.columns([1, 2, 1])[1]
    with col:
        if st.button("🔐  Sign in with Microsoft", use_container_width=True):
            msal_app, cache = get_msal_app()
            flow = msal_app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                st.error("Device flow failed. Check your Azure App Registration.")
                st.stop()

            st.info("**Step 1** — Copy this code:")
            st.code(flow["user_code"], language=None)
            st.markdown(f"**Step 2** — [Open Microsoft sign-in page ↗]({flow['verification_uri']})")
            st.caption("Complete sign-in in the browser, then wait here.")

            with st.spinner("Waiting for authentication…"):
                result = msal_app.acquire_token_by_device_flow(flow)

            save_cache(cache)

            if "access_token" in result:
                st.session_state["access_token"] = result["access_token"]
                st.success("✅ Signed in successfully!")
                st.rerun()
            else:
                st.error(f"Sign-in failed: {result.get('error_description', 'Unknown error')}")

# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
else:
    # ── Ensure token is still valid on every page load ──
    token, err = get_valid_token()
    if err:
        st.warning(err)
        st.stop()

    # ── Fetch user info once ──
    if "user_info" not in st.session_state:
        user, err = get_user_info(token)
        if user:
            st.session_state["user_info"] = user
        else:
            st.error(f"Session error: {err}")
            del st.session_state["access_token"]
            st.rerun()

    user         = st.session_state.get("user_info", {})
    display_name = user.get("displayName", "User")
    email_addr   = user.get("mail") or user.get("userPrincipalName", "")

    # ── SIDEBAR ──
    with st.sidebar:
        st.markdown(f"""
        <div style="padding:16px 0 24px;">
            <div style="font-size:13px; color:#5a6478; margin-bottom:4px;">Signed in as</div>
            <div style="font-size:15px; font-weight:600; color:#e8eaf0;">{display_name}</div>
            <div style="font-size:12px; color:#3b82f6; font-family:'DM Mono',monospace;">{email_addr}</div>
        </div>
        """, unsafe_allow_html=True)

        st.divider()

        st.markdown("**📁 Folder**")
        folder = st.selectbox("folder", label_visibility="collapsed",
            options=["all", "inbox", "sent", "drafts", "junk"],
            format_func=lambda x: {
                "all":    "📬  All Messages",
                "inbox":  "📥  Inbox",
                "sent":   "📤  Sent",
                "drafts": "📝  Drafts",
                "junk":   "🗑️  Junk",
            }[x]
        )

        st.markdown("**🔢 Count**")
        top_n = st.selectbox("count", label_visibility="collapsed",
                             options=[5, 10, 20, 50], index=1)

        st.markdown("**🔍 Search**")
        search_query = st.text_input("search", label_visibility="collapsed",
                                     placeholder="Search subject, sender…")

        only_unread = st.checkbox("Unread only")

        st.markdown("")
        fetch_btn = st.button("📥  Fetch Emails", use_container_width=True)

        st.divider()
        if st.button("🚪  Logout", use_container_width=True):
            for k in ["access_token", "token_cache", "emails", "user_info",
                      "stats", "next_link", "folder"]:
                st.session_state.pop(k, None)
            st.rerun()

    # ── FETCH ──
    if fetch_btn:
        # Always get a fresh/valid token before fetching
        token, err = get_valid_token()
        if err:
            st.error(err)
            st.stop()

        with st.spinner("Fetching emails…"):
            emails, next_link, err = get_emails(
                token, top=top_n, folder=folder,
                search=search_query, only_unread=only_unread,
            )
        if err:
            st.error(err)
        else:
            st.session_state["emails"]    = emails
            st.session_state["next_link"] = next_link
            st.session_state["folder"]    = folder
            st.session_state["stats"] = {
                "total":  len(emails),
                "unread": sum(1 for e in emails if not e.get("isRead", True)),
                "urgent": sum(1 for e in emails if e.get("importance") == "high"),
                "attach": sum(1 for e in emails if e.get("hasAttachments", False)),
            }

    # ── HEADER ──
    st.markdown(f"""
    <div style="display:flex; align-items:center; gap:12px; margin-bottom:24px;">
        <span style="font-size:28px;">📧</span>
        <div>
            <div style="font-size:22px; font-weight:600; color:#e8eaf0;">Outlook Mail</div>
            <div style="font-size:13px; color:#5a6478; font-family:'DM Mono',monospace;">{email_addr}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── STATS ──
    if "stats" in st.session_state:
        s = st.session_state["stats"]
        c1, c2, c3, c4 = st.columns(4)
        for col, num, label, color in [
            (c1, s["total"],  "TOTAL",     "#3b82f6"),
            (c2, s["unread"], "UNREAD",    "#60a5fa"),
            (c3, s["urgent"], "URGENT",    "#ef4444"),
            (c4, s["attach"], "W/ ATTACH", "#4ade80"),
        ]:
            with col:
                st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-number" style="color:{color};">{num}</div>
                    <div class="stat-label">{label}</div>
                </div>
                """, unsafe_allow_html=True)
        st.markdown("<div style='margin-bottom:20px;'></div>", unsafe_allow_html=True)

    # ── EMAIL LIST ──
    if "emails" not in st.session_state:
        st.markdown("""
        <div style="text-align:center; padding:80px 0; color:#3a4155;">
            <div style="font-size:48px; margin-bottom:16px;">📭</div>
            <div style="font-size:16px;">Select a folder and click <b>Fetch Emails</b></div>
        </div>
        """, unsafe_allow_html=True)
    else:
        emails = st.session_state["emails"]
        if not emails:
            st.markdown("""
            <div style="text-align:center; padding:60px 0; color:#3a4155;">
                <div style="font-size:40px; margin-bottom:12px;">📭</div>
                <div style="font-size:15px;">No emails found.</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            folder_label = st.session_state.get("folder", "all")
            st.markdown(
                f'<div style="font-size:13px; color:#5a6478; font-family:\'DM Mono\',monospace; '
                f'margin-bottom:16px;">{len(emails)} email(s) · {folder_label.upper()}</div>',
                unsafe_allow_html=True
            )
            for email in emails:
                render_email_card(email)

            if st.session_state.get("next_link"):
                st.caption("More emails available — increase the count to load more.")
