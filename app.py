import streamlit as st
import requests
from msal import PublicClientApplication

# ----------------------------
# Load secrets (Streamlit Cloud)
# ----------------------------
CLIENT_ID = st.secrets["CLIENT_ID"]
TENANT_ID = st.secrets["TENANT_ID"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Read", "User.Read"]

# ----------------------------
# MSAL App
# ----------------------------
app = PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

# ----------------------------
# Functions
# ----------------------------
def get_user_info(access_token):
    """Verify token works by fetching user profile first."""
    url = "https://graph.microsoft.com/v1.0/me"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json(), None
    else:
        return None, response.text

def get_emails(access_token, top=10):
    url = (
        "https://graph.microsoft.com/v1.0/me/messages"
        f"?$top={top}&$select=subject,from,receivedDateTime,bodyPreview"
        f"&$orderby=receivedDateTime desc"
    )
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)

    # ── Debug block (remove once working) ──────────────────────────────────
    with st.expander("🔍 API Debug Info", expanded=False):
        st.write(f"**Status Code:** {response.status_code}")
        st.write(f"**Request URL:** {url}")
        try:
            st.json(response.json())
        except Exception:
            st.text(response.text)
    # ───────────────────────────────────────────────────────────────────────

    if response.status_code == 200:
        emails = response.json().get("value", [])
        if not emails:
            st.warning("⚠️ API returned 200 but 0 emails. Your mailbox may be empty, "
                       "or the account has no messages in the default folder.")
        return emails
    elif response.status_code == 401:
        st.error("❌ 401 Unauthorized — your token has expired or lacks Mail.Read permission. "
                 "Please logout and login again.")
        del st.session_state["access_token"]
        st.rerun()
    elif response.status_code == 403:
        st.error("❌ 403 Forbidden — Mail.Read permission was not granted by your admin. "
                 "Check API permissions in Azure Portal and ensure admin consent is given.")
    else:
        try:
            err = response.json()
            st.error(f"❌ Error {response.status_code}: {err.get('error', {}).get('message', response.text)}")
        except Exception:
            st.error(f"❌ Unexpected error {response.status_code}: {response.text}")
    return []

# ----------------------------
# UI
# ----------------------------
st.set_page_config(page_title="Outlook Dashboard", page_icon="📧", layout="wide")
st.title("📧 Outlook Email Dashboard (MS Graph)")

# ----------------------------
# LOGIN SECTION
# ----------------------------
if "access_token" not in st.session_state:
    st.info("Please login with Microsoft to continue.")
    if st.button("🔐 Login with Microsoft"):
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            st.error("Device flow failed. Check your Azure App Registration.")
            st.stop()

        # Show the device code message prominently
        st.info(flow["message"])
        st.code(flow["user_code"], language=None)  # make code easy to copy
        st.markdown(f"👉 [Click here to authenticate]({flow['verification_uri']})",
                    unsafe_allow_html=True)

        with st.spinner("Waiting for you to complete authentication in the browser…"):
            result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            st.session_state["access_token"] = result["access_token"]
            # Store token metadata for debugging
            st.session_state["token_scopes"] = result.get("scope", "")
            st.success("✅ Login successful!")
            st.rerun()
        else:
            st.error(f"Login failed: {result.get('error_description', 'Unknown error')}")

# ----------------------------
# AFTER LOGIN
# ----------------------------
else:
    # ── Verify token and show logged-in user ────────────────────────────
    if "user_info" not in st.session_state:
        user, err = get_user_info(st.session_state["access_token"])
        if user:
            st.session_state["user_info"] = user
        else:
            st.error(f"Token validation failed: {err}")
            del st.session_state["access_token"]
            st.rerun()

    user = st.session_state.get("user_info", {})
    display_name = user.get("displayName", "User")
    email_addr   = user.get("mail") or user.get("userPrincipalName", "")

    st.success(f"🎉 Logged in as **{display_name}** ({email_addr})")

    # Show granted scopes (useful to verify Mail.Read is included)
    scopes_granted = st.session_state.get("token_scopes", "")
    if scopes_granted:
        with st.expander("🔑 Token Scopes Granted"):
            st.write(scopes_granted)
        if "Mail.Read" not in scopes_granted and "mail.read" not in scopes_granted.lower():
            st.warning("⚠️ **Mail.Read** is NOT in the granted scopes! "
                       "Go to Azure Portal → App Registration → API Permissions "
                       "and add Mail.Read, then grant Admin Consent.")

    col1, col2 = st.columns(2)
    with col1:
        top_n = st.selectbox("How many emails?", [5, 10, 20, 50], index=1)
        if st.button("📥 Fetch Emails"):
            with st.spinner("Fetching emails…"):
                st.session_state["emails"] = get_emails(
                    st.session_state["access_token"],
                    top=top_n
                )
    with col2:
        if st.button("🚪 Logout"):
            for key in ["access_token", "emails", "user_info", "token_scopes"]:
                st.session_state.pop(key, None)
            st.rerun()

    # ----------------------------
    # DISPLAY EMAILS
    # ----------------------------
    if "emails" in st.session_state:
        emails = st.session_state["emails"]
        st.subheader(f"📬 Latest {len(emails)} Email(s)")

        if not emails:
            st.info("No emails found. Check the debug info above for the raw API response.")
        else:
            for email in emails:
                with st.container(border=True):
                    subject  = email.get("subject") or "(No Subject)"
                    sender   = (email.get("from", {})
                                     .get("emailAddress", {})
                                     .get("address", "Unknown"))
                    preview  = email.get("bodyPreview", "")
                    time_str = email.get("receivedDateTime", "")

                    col_a, col_b = st.columns([3, 1])
                    with col_a:
                        st.markdown(f"**📌 {subject}**")
                        st.caption(f"👤 {sender}")
                    with col_b:
                        st.caption(f"⏰ {time_str[:10] if time_str else ''}")

                    if preview:
                        st.write(preview)
