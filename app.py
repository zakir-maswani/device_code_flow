import streamlit as st
import requests
from msal import PublicClientApplication

# ----------------------------
# Config
# ----------------------------
CLIENT_ID = st.secrets["CLIENT_ID"]
AUTHORITY = "https://login.microsoftonline.com/consumers"  # personal accounts
SCOPES = ["Mail.Read"]

app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# ----------------------------
# Functions
# ----------------------------
def get_emails(access_token):
    url = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=10"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        st.error(f"Error {response.status_code}: {response.text}")
        return []

def try_silent_login():
    """Reuse cached token if available (avoids re-login on rerun)."""
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    return None

# ----------------------------
# UI
# ----------------------------
st.title("📧 Outlook Email Dashboard")

# Try silent login first
if "access_token" not in st.session_state:
    token = try_silent_login()
    if token:
        st.session_state["access_token"] = token

# ----------------------------
# LOGIN
# ----------------------------
if "access_token" not in st.session_state:
    st.info("Login with your Microsoft account to read emails.")
    if st.button("🔐 Login with Microsoft"):
        with st.spinner("Starting device flow..."):
            flow = app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                st.error("Device flow failed. Check your Azure app registration.")
                st.stop()

        # Show the code to the user clearly
        st.markdown("### 👇 Complete login in your browser")
        st.code(flow["user_code"], language=None)
        st.markdown(f"Go to **[microsoft.com/devicelogin]({flow['verification_uri']})** and enter the code above.")

        with st.spinner("Waiting for you to complete login..."):
            result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            st.session_state["access_token"] = result["access_token"]
            st.success("✅ Logged in!")
            st.rerun()
        else:
            err = result.get("error_description", "Unknown error")
            st.error(f"Login failed: {err}")

# ----------------------------
# AFTER LOGIN
# ----------------------------
else:
    st.success("🎉 You are logged in!")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Fetch Emails"):
            with st.spinner("Fetching emails..."):
                st.session_state["emails"] = get_emails(st.session_state["access_token"])
    with col2:
        if st.button("🚪 Logout"):
            for key in ["access_token", "emails"]:
                st.session_state.pop(key, None)
            st.rerun()

    # ----------------------------
    # DISPLAY
    # ----------------------------
    if "emails" in st.session_state:
        emails = st.session_state["emails"]
        st.subheader(f"📬 {len(emails)} Latest Emails")
        for email in emails:
            with st.expander(email.get("subject", "No Subject")):
                sender = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
                st.write(f"**👤 From:** {sender}")
                st.write(f"**⏰ Time:** {email.get('receivedDateTime', '')}")
                st.write(f"**📝 Preview:** {email.get('bodyPreview', '')}")
