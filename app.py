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
# PAGE
# ----------------------------
st.set_page_config(
    page_title="Outlook Weekly Emails",
    layout="wide"
)

st.title("📧 Outlook Weekly Emails")

# ----------------------------
# LOGIN
# ----------------------------
if "access_token" not in st.session_state:

    st.info("Login with your Microsoft account")

    if st.button("🔐 Login with Microsoft"):

        flow = app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            st.error("Device flow failed")
            st.stop()

        st.code(flow["user_code"])

        st.markdown(
            f"Go to: {flow['verification_uri']} and enter the code above."
        )

        result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            st.session_state["access_token"] = result["access_token"]
            st.success("✅ Login successful")
            st.rerun()

        else:
            st.error("Login failed")
            st.write(result)

# ----------------------------
# FETCH EMAILS
# ----------------------------
else:

    access_token = st.session_state["access_token"]

    col1, col2 = st.columns([1, 1])

    with col1:
        fetch_btn = st.button("📥 Fetch Last 1 Week Emails")

    with col2:
        if st.button("🚪 Logout"):
            st.session_state.clear()
            st.rerun()

    if fetch_btn:

        with st.spinner("Fetching emails..."):

            # Last 7 days
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

                    with st.expander(subject):

                        st.write(f"**From:** {sender}")
                        st.write(f"**Received:** {received}")
                        st.write(f"**Preview:** {preview}")

            else:

                st.error(f"Error: {response.status_code}")

                try:
                    st.json(response.json())
                except:
                    st.write(response.text)
