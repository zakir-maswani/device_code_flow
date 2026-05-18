import streamlit as st
import requests
from msal import PublicClientApplication
from datetime import datetime, timedelta, timezone

# ----------------------------
# CONFIG
# ----------------------------
CLIENT_ID = st.secrets["CLIENT_ID"]

AUTHORITY = "https://login.microsoftonline.com/consumers"

SCOPES = [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Mail.Read"
]

app = PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

# ----------------------------
# PAGE CONFIG
# ----------------------------
st.set_page_config(
    page_title="Outlook Weekly Emails",
    page_icon="📧",
    layout="wide"
)

# ----------------------------
# CUSTOM CSS
# ----------------------------
st.markdown("""
<style>

.main {
    padding-top: 1rem;
}

.stButton > button {
    width: 100%;
    border-radius: 10px;
    height: 3em;
    font-weight: 600;
    border: none;
}

.email-card {
    background-color: #ffffff10;
    border: 1px solid #ffffff20;
    padding: 18px;
    border-radius: 14px;
    margin-bottom: 15px;
}

.security-box {
    background: linear-gradient(90deg, #0f172a, #111827);
    padding: 16px;
    border-radius: 14px;
    border: 1px solid #334155;
    margin-top: 10px;
    margin-bottom: 20px;
}

.header-box {
    padding: 20px;
    border-radius: 18px;
    background: linear-gradient(135deg, #2563eb, #1d4ed8);
    color: white;
    margin-bottom: 25px;
}

.small-text {
    font-size: 14px;
    opacity: 0.85;
}

</style>
""", unsafe_allow_html=True)

# ----------------------------
# HEADER
# ----------------------------
st.markdown("""
<div class="header-box">
    <h1>📧 Outlook Weekly Email Dashboard</h1>
    <p class="small-text">
        Securely access and view your last 7 days emails using Microsoft Graph API.
    </p>
</div>
""", unsafe_allow_html=True)

# ----------------------------
# SIDEBAR
# ----------------------------
with st.sidebar:

    st.title("⚙️ Dashboard")

    st.markdown("---")

    st.markdown("### 🔐 Privacy & Security")

    st.markdown("""
    <div class="security-box">
        ✅ Your emails are accessed securely using Microsoft authentication.<br><br>
        
        ✅ Your data is <b>NOT stored</b> anywhere.<br><br>
        
        ✅ No database is used.<br><br>
        
        ✅ Emails are fetched temporarily during your session only.<br><br>
        
        ✅ Authentication is handled directly by Microsoft.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    st.markdown("### 📌 Features")

    st.markdown("""
    - View last 7 days emails  
    - Microsoft secure login  
    - Fast email fetching  
    - Clean dashboard UI  
    - Secure session-based access  
    """)

# ----------------------------
# LOGIN
# ----------------------------
if "access_token" not in st.session_state:

    st.info("🔓 Login with your Microsoft account to continue")

    if st.button("🔐 Login with Microsoft"):

        flow = app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            st.error("❌ Device flow failed")
            st.stop()

        st.markdown("### 👇 Complete Login")

        st.code(flow["user_code"])

        st.markdown(
            f"""
            Visit: **{flow['verification_uri']}**  
            and enter the code above.
            """
        )

        with st.spinner("Waiting for login..."):

            result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:

            st.session_state["access_token"] = result["access_token"]

            st.success("✅ Login successful")

            st.rerun()

        else:

            st.error("❌ Login failed")

            st.write(result)

# ----------------------------
# AFTER LOGIN
# ----------------------------
else:

    access_token = st.session_state["access_token"]

    col1, col2 = st.columns([3, 1])

    with col1:
        fetch_btn = st.button("📥 Fetch Last 7 Days Emails")

    with col2:
        logout_btn = st.button("🚪 Logout")

    if logout_btn:
        st.session_state.clear()
        st.rerun()

    if fetch_btn:

        with st.spinner("📬 Fetching emails..."):

            last_week = (
                datetime.now(timezone.utc) - timedelta(days=7)
            ).strftime("%Y-%m-%dT%H:%M:%SZ")

            url = (
                "https://graph.microsoft.com/v1.0/me/messages"
                f"?$filter=receivedDateTime ge {last_week}"
                "&$top=100"
                "&$orderby=receivedDateTime DESC"
            )

            headers = {
                "Authorization": f"Bearer {access_token}"
            }

            response = requests.get(url, headers=headers)

            if response.status_code == 200:

                emails = response.json().get("value", [])

                st.success(f"✅ {len(emails)} emails found")

                if not emails:
                    st.info("No emails found in last 7 days")

                for email in emails:

                    subject = email.get("subject", "No Subject")

                    sender = (
                        email.get("from", {})
                        .get("emailAddress", {})
                        .get("address", "Unknown")
                    )

                    received = email.get("receivedDateTime", "")

                    preview = email.get("bodyPreview", "")

                    st.markdown(f"""
                    <div class="email-card">
                        <h4>📩 {subject}</h4>
                        <p><b>From:</b> {sender}</p>
                        <p><b>Received:</b> {received}</p>
                        <p>{preview}</p>
                    </div>
                    """, unsafe_allow_html=True)

            else:

                st.error(f"❌ Error: {response.status_code}")

                try:
                    st.json(response.json())
                except:
                    st.write(response.text)
