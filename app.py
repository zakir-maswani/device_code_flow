import streamlit as st
import requests
from msal import PublicClientApplication
import jwt
import socket
import time
import json
from datetime import datetime
from typing import Optional, Dict, Any

# ----------------------------
# Config
# ----------------------------
CLIENT_ID = st.secrets.get("CLIENT_ID", "NOT_SET")
TENANT_ID = st.secrets.get("TENANT_ID", "NOT_SET")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Read"]

app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# ----------------------------
# Debug Utilities
# ----------------------------

def check_system_time() -> Dict[str, Any]:
    """Check server time vs token expiry"""
    now = datetime.utcnow()
    timestamp = int(time.time())
    return {
        "current_time": now.isoformat(),
        "timestamp": timestamp,
        "timezone": "UTC"
    }

def check_network() -> Dict[str, Any]:
    """Test network connectivity to Microsoft services"""
    services = {
        "graph.microsoft.com": None,
        "login.microsoftonline.com": None
    }
    
    for service in services:
        try:
            socket.gethostbyname(service)
            services[service] = "✅ Reachable"
        except socket.gaierror as e:
            services[service] = f"❌ Error: {e}"
    
    return services

def decode_token(token: str) -> Optional[Dict]:
    """Safely decode JWT token without verification"""
    try:
        decoded = jwt.decode(token, options={"verify_signature": False})
        return decoded
    except Exception as e:
        return {"error": str(e)}

def validate_token(token: str) -> Dict[str, Any]:
    """Validate token and check claims"""
    decoded = decode_token(token)
    
    if "error" in decoded:
        return {
            "valid": False, 
            "error": decoded["error"],
            "checks": {},
            "full_token": {}
        }
    
    now = int(time.time())
    exp = decoded.get("exp", 0)
    
    checks = {
        "has_aud": "aud" in decoded,
        "aud_value": decoded.get("aud"),
        "is_graph": decoded.get("aud") == "00000003-0000-0000-c000-000000000000",
        "has_mail_scope": "Mail.Read" in decoded.get("scp", ""),
        "scopes": decoded.get("scp", "").split(),
        "account_type": "Personal" if decoded.get("idp") == "live.com" else "Work/School",
        "idp": decoded.get("idp"),
        "expired": exp < now,
        "expires_in_seconds": max(0, exp - now),
        "expires_at": datetime.utcfromtimestamp(exp).isoformat() if exp > 0 else "Invalid",
        "issued_at": datetime.utcfromtimestamp(decoded.get("iat", 0)).isoformat() if decoded.get("iat", 0) > 0 else "Invalid",
    }
    
    return {
        "valid": exp > now and "Mail.Read" in decoded.get("scp", ""),
        "checks": checks,
        "full_token": decoded if decoded else {}
    }

# ----------------------------
# Functions
# ----------------------------

def get_emails(access_token: str) -> tuple[list, Dict[str, Any]]:
    """Fetch emails with full debugging"""
    debug_info = {
        "endpoint": "https://graph.microsoft.com/v1.0/me/messages",
        "status_code": None,
        "response_time": None,
        "error": None,
        "raw_response": None
    }
    
    url = "https://graph.microsoft.com/v1.0/me/messages?$top=10&$orderby=receivedDateTime desc"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    try:
        start = time.time()
        response = requests.get(url, headers=headers, timeout=10, verify=True)
        debug_info["response_time"] = round(time.time() - start, 2)
        debug_info["status_code"] = response.status_code
        
        st.write(f"📊 **Status:** {response.status_code}")
        st.write(f"⏱️ **Response Time:** {debug_info['response_time']}s")
        
        if response.status_code == 200:
            data = response.json()
            debug_info["raw_response"] = data
            return data.get("value", []), debug_info
        else:
            debug_info["error"] = response.text
            try:
                error_json = response.json()
                debug_info["error_details"] = error_json
                st.error(f"❌ Error {response.status_code}")
                st.code(json.dumps(error_json, indent=2, default=str), language="json")
            except:
                st.error(f"❌ Error {response.status_code}")
                st.code(response.text, language="text")
            return [], debug_info
    
    except requests.exceptions.Timeout:
        debug_info["error"] = "Request timeout (10s)"
        st.error("⏱️ Request timed out - Microsoft Graph not responding")
        return [], debug_info
    except requests.exceptions.SSLError as e:
        debug_info["error"] = f"SSL Error: {e}"
        st.error(f"🔒 SSL Certificate Error: {e}")
        return [], debug_info
    except requests.exceptions.ConnectionError as e:
        debug_info["error"] = f"Connection Error: {e}"
        st.error(f"🌐 Cannot connect to Microsoft Graph: {e}")
        return [], debug_info
    except Exception as e:
        debug_info["error"] = str(e)
        st.error(f"❌ Unexpected error: {e}")
        return [], debug_info

def try_silent_login() -> Optional[str]:
    """Reuse cached token if available (avoids re-login on rerun)."""
    try:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                return result["access_token"]
    except Exception as e:
        st.warning(f"Silent login failed: {e}")
    return None

# ----------------------------
# UI
# ----------------------------
st.set_page_config(page_title="📧 Outlook Email Dashboard", layout="wide")
st.title("📧 Outlook Email Dashboard")

# Sidebar for diagnostics
with st.sidebar:
    st.header("🔧 Diagnostics")
    if st.checkbox("Show System Info", value=False):
        sys_time = check_system_time()
        st.write("**System Time:**")
        st.code(json.dumps(sys_time, indent=2, default=str), language="json")
    
    if st.checkbox("Show Network Status", value=False):
        network = check_network()
        st.write("**Network Connectivity:**")
        for service, status in network.items():
            st.write(f"- {service}: {status}")
    
    if st.checkbox("Show Secrets", value=False):
        st.write("**Configuration:**")
        st.write(f"- CLIENT_ID: {CLIENT_ID[:20]}..." if len(CLIENT_ID) > 20 else f"- CLIENT_ID: {CLIENT_ID}")
        st.write(f"- TENANT_ID: {TENANT_ID[:20]}..." if len(TENANT_ID) > 20 else f"- TENANT_ID: {TENANT_ID}")
        st.write(f"- AUTHORITY: {AUTHORITY}")

# Try silent login first
if "access_token" not in st.session_state:
    token = try_silent_login()
    if token:
        st.session_state["access_token"] = token

# ----------------------------
# LOGIN
# ----------------------------
if "access_token" not in st.session_state:
    st.info("🔓 Login with your Microsoft account to read emails.")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔐 Login with Microsoft", use_container_width=True):
            try:
                with st.spinner("🔄 Starting device flow..."):
                    flow = app.initiate_device_flow(scopes=SCOPES)
                    
                    if "user_code" not in flow:
                        st.error("❌ Device flow failed. Check your Azure app registration.")
                        st.code(json.dumps(flow, indent=2, default=str), language="json")
                        st.stop()

                st.markdown("### 👇 Complete login in your browser")
                st.code(flow["user_code"], language=None)
                st.markdown(f"Go to **[microsoft.com/devicelogin]({flow['verification_uri']})** and enter the code above.")
                st.markdown(f"**Expires in:** {flow.get('expires_in', 'N/A')} seconds")

                with st.spinner("⏳ Waiting for you to complete login..."):
                    result = app.acquire_token_by_device_flow(flow)

                if "access_token" in result:
                    st.session_state["access_token"] = result["access_token"]
                    st.success("✅ Logged in!")
                    st.rerun()
                else:
                    err = result.get("error_description", result.get("error", "Unknown error"))
                    st.error(f"❌ Login failed: {err}")
                    st.code(json.dumps(result, indent=2, default=str), language="json")
            except Exception as e:
                st.error(f"❌ Login error: {e}")
    
    with col2:
        if st.button("📋 View Debug Info", use_container_width=True):
            with st.expander("🔍 Configuration Debug", expanded=True):
                st.write("**Secrets Status:**")
                st.write(f"- CLIENT_ID set: {CLIENT_ID != 'NOT_SET'}")
                st.write(f"- TENANT_ID set: {TENANT_ID != 'NOT_SET'}")
                st.write(f"- SCOPES: {SCOPES}")

# ----------------------------
# AFTER LOGIN
# ----------------------------
else:
    st.success("🎉 You are logged in!")
    
    token = st.session_state["access_token"]
    
    # Create tabs for organization
    tab_main, tab_debug, tab_token = st.tabs(["📧 Emails", "🔧 Full Debug", "🔐 Token Info"])
    
    # ----------------------------
    # MAIN TAB - Emails
    # ----------------------------
    with tab_main:
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📥 Fetch Emails", use_container_width=True):
                with st.spinner("📬 Fetching emails..."):
                    emails, debug_info = get_emails(token)
                    st.session_state["emails"] = emails
                    st.session_state["last_debug"] = debug_info
        
        with col2:
            if st.button("🚪 Logout", use_container_width=True):
                for key in ["access_token", "emails", "last_debug"]:
                    st.session_state.pop(key, None)
                st.rerun()
        
        # Display emails
        if "emails" in st.session_state:
            emails = st.session_state["emails"]
            st.subheader(f"📬 {len(emails)} Latest Emails")
            
            if len(emails) == 0:
                st.info("No emails found.")
            else:
                for i, email in enumerate(emails, 1):
                    with st.expander(f"{i}. {email.get('subject', 'No Subject')}"):
                        sender = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
                        st.write(f"**👤 From:** {sender}")
                        st.write(f"**⏰ Received:** {email.get('receivedDateTime', 'N/A')}")
                        st.write(f"**📝 Preview:** {email.get('bodyPreview', 'N/A')}")
                        st.write(f"**🔗 ID:** `{email.get('id', 'N/A')}`")
    
    # ----------------------------
    # TOKEN DEBUG TAB
    # ----------------------------
    with tab_token:
        st.subheader("🔐 Token Claims")
        
        token_validation = validate_token(token)
        
        # Status indicator
        if token_validation.get("valid", False):
            st.success("✅ Token is valid")
        else:
            st.error("❌ Token is invalid or expired")
        
        # Show error if present
        if "error" in token_validation:
            st.error(f"Error decoding token: {token_validation['error']}")
        
        # Token checks
        checks = token_validation.get("checks", {})
        
        if checks:
            st.write("**Token Validation Checks:**")
            
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"✅ Has AUD claim: {checks.get('has_aud', 'N/A')}")
                st.write(f"✅ Is Graph audience: {checks.get('is_graph', 'N/A')}")
                st.write(f"✅ Has Mail.Read scope: {checks.get('has_mail_scope', 'N/A')}")
                st.write(f"✅ Account type: {checks.get('account_type', 'N/A')}")
            
            with col2:
                st.write(f"✅ Expired: {checks.get('expired', 'N/A')}")
                st.write(f"✅ Expires in: {checks.get('expires_in_seconds', 'N/A')} seconds")
                st.write(f"✅ Issued at: {checks.get('issued_at', 'N/A')}")
                st.write(f"✅ Expires at: {checks.get('expires_at', 'N/A')}")
            
            # All claims
            full_token = token_validation.get("full_token", {})
            if full_token:
                with st.expander("📋 All Token Claims"):
                    try:
                        token_str = json.dumps(full_token, indent=2, default=str)
                        st.code(token_str, language="json")
                    except Exception as e:
                        st.warning(f"Could not serialize token: {e}")
                        st.write(f"Token data: {type(full_token)}")
        else:
            st.info("No token checks available")
    
    # ----------------------------
    # FULL DEBUG TAB
    # ----------------------------
    with tab_debug:
        st.subheader("🔧 Full Diagnostics")
        
        # System time
        with st.expander("⏰ System Time Check", expanded=True):
            sys_time = check_system_time()
            st.code(json.dumps(sys_time, indent=2, default=str), language="json")
            
            token_validation = validate_token(token)
            checks = token_validation.get("checks", {})
            
            if checks:
                exp_time = checks.get("expires_at", "Unknown")
                expires_in = checks.get("expires_in_seconds", 0)
                st.write(f"**Token expires at:** {exp_time}")
                st.write(f"**Seconds remaining:** {expires_in}")
                
                if expires_in < 60:
                    st.warning("⚠️ Token will expire soon!")
            else:
                st.warning("Could not validate token")
        
        # Network connectivity
        with st.expander("🌐 Network Connectivity", expanded=True):
            network = check_network()
            for service, status in network.items():
                st.write(f"**{service}:** {status}")
            
            # Test Graph endpoint
            st.write("**Testing Graph endpoint...**")
            try:
                test_response = requests.head("https://graph.microsoft.com/v1.0/me", timeout=5)
                st.write(f"✅ Graph endpoint reachable (HTTP {test_response.status_code})")
            except requests.exceptions.Timeout:
                st.error("❌ Timeout - slow network or endpoint down")
            except requests.exceptions.ConnectionError as e:
                st.error(f"❌ Cannot connect: {e}")
        
        # Configuration
        with st.expander("⚙️ Configuration", expanded=False):
            st.write("**Secrets:**")
            st.write(f"- CLIENT_ID: {CLIENT_ID[:30]}..." if len(CLIENT_ID) > 30 else f"- CLIENT_ID: {CLIENT_ID}")
            st.write(f"- TENANT_ID: {TENANT_ID}")
            st.write(f"- AUTHORITY: {AUTHORITY}")
            st.write(f"- SCOPES: {SCOPES}")
            
            st.write("**MSAL Settings:**")
            st.write(f"- App instance: {str(app)[:100]}")
            
            accounts = app.get_accounts()
            st.write(f"**Cached accounts: {len(accounts)}**")
            for acc in accounts:
                st.write(f"- Username: {acc.get('username')}")
                st.write(f"- Home Account ID: {acc.get('home_account_id')}")
        
        # Last API call debug
        if "last_debug" in st.session_state:
            with st.expander("📊 Last API Call Debug", expanded=True):
                try:
                    debug_str = json.dumps(st.session_state["last_debug"], indent=2, default=str)
                    st.code(debug_str, language="json")
                except Exception as e:
                    st.warning(f"Could not serialize debug info: {e}")
                    st.write(str(st.session_state["last_debug"]))
        
        # Environment info
        with st.expander("🖥️ Environment Info", expanded=False):
            import platform
            import sys
            
            st.write(f"**Python Version:** {sys.version}")
            st.write(f"**Platform:** {platform.platform()}")
            st.write(f"**Streamlit Version:** {st.__version__}")

# ----------------------------
# Footer
# ----------------------------
st.divider()
st.caption("🔐 This app uses OAuth2 device flow for secure authentication | No credentials are stored")
