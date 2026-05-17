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
def get_access_token():
    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        return None, "Device flow failed"

    return flow, None


def get_emails(access_token, top=10):
    url = (
        #"https://graph.microsoft.com/v1.0/me/messages"
        #f"?$top={top}&$select=subject,from,receivedDateTime,bodyPreview"
        "https://graph.microsoft.com/v1.0/me/messages?%24top=5&24skip=5"
    )

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        st.error(response.text)
        return []


# ----------------------------
# UI
# ----------------------------
st.title("📧 Outlook Email Dashboard (MS Graph)")

# ----------------------------
# LOGIN SECTION
# ----------------------------
if "access_token" not in st.session_state:

    st.info("Please login with Microsoft to continue.")

    if st.button("🔐 Login with Microsoft"):

        flow = app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            st.error("Device flow failed. Check Azure setup.")
        else:
            st.info(flow["message"])

            result = app.acquire_token_by_device_flow(flow)

            if "access_token" in result:
                st.session_state["access_token"] = result["access_token"]
                st.success("✅ Login successful!")
                st.rerun()
            else:
                st.error(result.get("error_description"))

# ----------------------------
# AFTER LOGIN
# ----------------------------
else:
    st.success("🎉 You are logged in!")

    col1, col2 = st.columns(2)

    with col1:
        if st.button("📥 Fetch Emails"):
            st.session_state["emails"] = get_emails(
                st.session_state["access_token"],
                top=10
            )

    with col2:
        if st.button("🚪 Logout"):
            del st.session_state["access_token"]
            if "emails" in st.session_state:
                del st.session_state["emails"]
            st.rerun()

    # ----------------------------
    # DISPLAY EMAILS
    # ----------------------------
    if "emails" in st.session_state:

        st.subheader("📬 Latest Emails")

        for email in st.session_state["emails"]:
            st.markdown("---")

            subject = email.get("subject", "No Subject")
            sender = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
            preview = email.get("bodyPreview", "")
            time = email.get("receivedDateTime", "")

            st.write(f"**📌 Subject:** {subject}")
            st.write(f"**👤 From:** {sender}")
            st.write(f"**⏰ Time:** {time}")
            st.write(f"**📝 Preview:** {preview}")
