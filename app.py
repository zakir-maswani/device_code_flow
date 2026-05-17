import streamlit as st
from msal import PublicClientApplication

# ----------------------------
# Azure App Config
# ----------------------------

CLIENT_ID = st.secrets["CLIENT_ID"]
TENANT_ID = st.secrets["TENANT_ID"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Read", "Mail.ReadWrite", "User.Read"]

# ----------------------------
# MSAL App
# ----------------------------
app = PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

st.title("📧 Outlook Automation LOgin")

# ----------------------------
# Login Button
# ----------------------------
if "access_token" not in st.session_state:

    if st.button("🔐 Login with Microsoft"):

        flow = app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            st.error("Device flow failed. Check Azure config.")
        else:
            st.info(flow["message"])  # shows login instructions

            result = app.acquire_token_by_device_flow(flow)

            if "access_token" in result:
                st.session_state["access_token"] = result["access_token"]
                st.success("✅ Login successful!")

                st.rerun()
            else:
                st.error(result.get("error_description"))

# ----------------------------
# After Login UI
# ----------------------------
else:
    st.success("🎉 You are logged in with Microsoft Graph!")

    st.write("Token preview:")
    st.code(st.session_state["access_token"][:50] + "...")

    if st.button("Logout"):
        del st.session_state["access_token"]
        st.rerun()
