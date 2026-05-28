# 𝗢𝘂𝘁𝗹𝗼𝗼𝗸 𝗔𝗜 𝗥𝗲𝗽𝗼𝗿𝘁 𝗚𝗲𝗻𝗲𝗿𝗮𝘁𝗼𝗿 ⭕

> Transform your email inbox into executive intelligence reports with AI-powered analysis and professional formatting.

[![Python 3.8+](https://img.shields.io/badge/Python-3.8+-3776ab?logo=python&logoColor=white)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Built%20with-Streamlit-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io/)
[![Microsoft Graph API](https://img.shields.io/badge/Microsoft%20Graph%20API-Green?logo=microsoft&logoColor=white)](https://learn.microsoft.com/en-us/graph/)
[![Groq AI](https://img.shields.io/badge/AI%20Powered%20by-Groq-333?logo=groq)](https://groq.com/)

---

## ✨ What This Does

This application automatically analyzes your Outlook emails from the past 7 days using advanced AI and generates a professional, beautifully formatted Word document report. Every email is:

- 🤖 **Analyzed** by AI for priority, summary, and action items
- 🎯 **Categorized** by type (Meeting, Finance, Support, Legal, etc.)
- 🚨 **Prioritized** as Critical, High, Medium, or Low
- 📊 **Summarized** with executive insights
- 📝 **Documented** in a polished DOCX format with styling and metrics

---

## 🎯 Key Features

### Email Intelligence
- ✅ Fetches last 7 days of emails from Microsoft Outlook
- ✅ AI-powered priority classification (Critical → Low)
- ✅ Executive summary generation (max 2 sentences)
- ✅ Action item extraction and tracking
- ✅ Sentiment analysis (Positive, Neutral, Negative, Urgent)
- ✅ Email categorization (8+ categories)
- ✅ Deadline and approval detection
- ✅ Business risk assessment

### Report Generation
- ✅ **Professional DOCX export** with custom styling
- ✅ **Dynamic KPI dashboard** (Total, Critical, High Priority, Actions)
- ✅ **Executive overview** powered by AI
- ✅ **Detailed email cards** with metadata and analysis
- ✅ **Action items page** for quick reference
- ✅ **Responsive design** with color-coded priorities
- ✅ **Auto-fitted text** based on content length

### Security & Authentication
- ✅ Secure Microsoft OAuth 2.0 authentication
- ✅ Device flow login (no password required)
- ✅ Session management with token handling
- ✅ API error handling and graceful fallbacks

---

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- Microsoft Account (Outlook)
- Groq API Key (for AI analysis)
- Microsoft Azure App Registration

### Installation

#### 1. Clone & Setup
```bash
# Clone the repository
git clone https://github.com/yourusername/outlook-ai-report-generator.git
cd outlook-ai-report-generator

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

#### 2. Configure Secrets

Create a `.streamlit/secrets.toml` file:

```toml
# Microsoft Azure App Registration
CLIENT_ID = "your-azure-client-id-here"

# Groq API Key
GROQ_API_KEY = "your-groq-api-key-here"
```

#### 3. Register Azure App

1. Go to [Azure Portal](https://portal.azure.com/)
2. Create a new **App Registration**
3. Add platform: **Mobile and desktop applications**
4. Set redirect URI: `https://localhost`
5. Add API permissions:
   - `Mail.Read` (read user mailbox)
   - `User.Read` (read user profile)
6. Copy your **Client ID**

#### 4. Get Groq API Key

1. Visit [Groq Console](https://console.groq.com/)
2. Sign up or log in
3. Create API key
4. Copy to `secrets.toml`

#### 5. Run the App

```bash
streamlit run app.py
```

Visit `http://localhost:8501` in your browser.

---

## 📖 How to Use

### Step 1: Login
```
🔐 Click "Login with Microsoft"
→ Enter device code on Microsoft login page
→ Wait for authentication
```

### Step 2: Generate Report
```
📊 Click "Generate AI Report"
→ System fetches your last 7 days of emails
→ AI analyzes each email (priority, summary, action)
→ Executive overview is generated
→ Professional DOCX is created
```

### Step 3: Download & Review
```
📥 Download the DOCX file
→ Open in Microsoft Word
→ Review KPI metrics, overview, and detailed analysis
→ Share with stakeholders
```

---

## 🏗️ Architecture

### Component Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    Streamlit UI (Frontend)                  │
│  • Login Interface  • Report Generation  • Download Button   │
└────────────────────────┬────────────────────────────────────┘
                         │
        ┌────────────────┼────────────────┐
        │                │                │
        ▼                ▼                ▼
   ┌─────────┐   ┌──────────────┐   ┌──────────┐
   │Microsoft│   │ Groq AI      │   │ Python   │
   │Graph API│   │ (Llama 3.1)  │   │ Libraries│
   │  OAuth  │   │ Analysis     │   │ (python- │
   │         │   │              │   │docx, etc)│
   └─────────┘   └──────────────┘   └──────────┘
        │                │                │
        └────────────────┼────────────────┘
                         │
                         ▼
              ┌─────────────────────────┐
              │  Professional DOCX      │
              │  Report Generation      │
              │  • Styling              │
              │  • Tables & Formatting  │
              │  • Color Themes         │
              └─────────────────────────┘
```

### Data Flow

1. **Authentication**: Device flow OAuth with Microsoft
2. **Fetch**: Graph API retrieves emails (last 7 days)
3. **Analyze**: Each email processed through Groq AI
4. **Aggregate**: Data collected into analyses list
5. **Generate**: Professional DOCX created with styling
6. **Download**: User downloads final report

---

## 📊 Email Analysis Details

### Priority Classification

| Level | Triggers | Color |
|-------|----------|-------|
| 🔴 **Critical** | Security incidents, legal escalations, production outages, executive escalations, deadlines <24h | Red |
| 🟠 **High** | Customer issues, approval requests, urgent meetings, important decisions | Amber |
| 🟢 **Medium** | Standard business communications requiring attention | Green |
| ⚪ **Low** | Informational emails, newsletters, promotions, automated notifications | Slate |

### Email Categories

```
Meeting       Finance       Support       Project       Legal
HR            Sales         Security      Operations    Other
```

### Analysis Output

Each email analysis includes:

```json
{
  "priority": "High",
  "summary": "Concise 2-sentence executive summary",
  "action_item": "Specific next step or 'No action required'",
  "sentiment": "Positive/Neutral/Negative/Urgent",
  "category": "Meeting",
  "deadline": "2024-05-30",
  "approval_required": true,
  "follow_up_required": false,
  "business_risk": "Medium"
}
```

---

## 📄 Report Structure

The generated DOCX includes:

### Page 1
```
┌─────────────────────────────────────┐
│      OUTLOOK AI REPORT              │
│  Weekly Email Intelligence Report   │
│         30 May 2024                 │
├─────────────────────────────────────┤
│  KPI Dashboard (4 Metrics)          │
│  • Total Emails    • Critical       │
│  • High Priority   • Action Items   │
├─────────────────────────────────────┤
│  Executive Overview (AI Generated)  │
├─────────────────────────────────────┤
│  Email Analysis Cards (Per Email)   │
│  • Subject & Sender                 │
│  • Priority Badge                   │
│  • AI Summary                       │
│  • Action Items                     │
│  • Email Preview                    │
└─────────────────────────────────────┘
```

### Page 2
```
┌─────────────────────────────────────┐
│   ACTION ITEMS SUMMARY              │
├─────────────────────────────────────┤
│  Sortable Table:                    │
│  • Subject | Priority | Category    │
│  • Action Item (Truncated)          │
│                                     │
│  (Only includes emails with actions)│
└─────────────────────────────────────┘
```

---

## 🔧 Configuration

### Customization Options

#### Colors (in `app.py`)
```python
NAVY = RGBColor(0x0D, 0x1B, 0x2A)      # Header background
BLUE = RGBColor(0x1A, 0x56, 0xDB)      # Accent color
RED = RGBColor(0xEF, 0x44, 0x44)       # Critical priority
AMBER = RGBColor(0xF5, 0x9E, 0x0B)     # High priority
GREEN = RGBColor(0x10, 0xB9, 0x81)     # Medium priority
```

#### Email Fetch Limit
```python
# In the report generation section:
"&$top=20"  # Change 20 to fetch more/fewer emails
```

#### AI Model Temperature
```python
# In analyze_email() function:
temperature=0.1  # 0.0 = deterministic, 1.0 = creative
```

---

## 🛠️ Troubleshooting

### Common Issues

#### ❌ "Device flow failed"
- **Cause**: Azure App Registration not configured correctly
- **Solution**: Verify CLIENT_ID in secrets.toml matches Azure portal

#### ❌ "Graph API Error: 401"
- **Cause**: Access token expired or invalid
- **Solution**: Click Logout and login again

#### ❌ "No emails found in last 7 days"
- **Cause**: No recent emails in mailbox
- **Solution**: Check email filters, change date range in code (line ~380)

#### ❌ "AI summary could not be generated"
- **Cause**: Groq API error or invalid API key
- **Solution**: Verify GROQ_API_KEY in secrets.toml

#### ❌ Request timeout
- **Cause**: Network or API slowness
- **Solution**: Check internet connection, try again

### Debug Mode

Add to `app.py` to see detailed logs:
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

---

## 📦 Dependencies

```
streamlit==1.28+           # Web UI framework
msal==1.24+                # Microsoft authentication
groq==0.4+                 # Groq AI API
python-docx==0.8+          # DOCX generation
requests==2.31+            # HTTP requests
```

Full requirements in `requirements.txt`

---

## 🔐 Security Best Practices

### ✅ Do
- ✅ Store API keys in `.streamlit/secrets.toml` (not in code)
- ✅ Use environment variables in production
- ✅ Rotate API keys regularly
- ✅ Restrict Azure app permissions to minimum needed
- ✅ Clear session on logout

### ❌ Don't
- ❌ Hardcode API keys in code
- ❌ Share `secrets.toml` in version control
- ❌ Use production keys in development
- ❌ Share generated reports with sensitive data publicly

---

## 🚀 Deployment

### Streamlit Cloud

1. Push code to GitHub (without secrets.toml)
2. Visit [Streamlit Cloud](https://streamlit.io/cloud)
3. Create new app and select repository
4. Add secrets in dashboard:
   ```
   CLIENT_ID = "xxx"
   GROQ_API_KEY = "xxx"
   ```
5. Deploy!

### Docker

```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["streamlit", "run", "app.py", "--server.port=8501"]
```

```bash
docker build -t outlook-ai-report .
docker run -p 8501:8501 outlook-ai-report
```

---

## 📈 Performance Notes

| Metric | Value | Notes |
|--------|-------|-------|
| Emails Fetched | 20 max | Configurable via `$top` parameter |
| AI Analysis Time | ~2-3s per email | Depends on Groq API |
| Report Generation | ~5-10s | DOCX creation and formatting |
| Total Pipeline | ~1-2 min | For 20 emails |

**Optimization Tips:**
- Reduce email count for faster processing
- Batch process emails in background jobs
- Cache AI results if re-running

---

## 🤝 Contributing

### Development Workflow

```bash
# 1. Fork & clone
git clone https://github.com/yourusername/outlook-ai-report-generator.git

# 2. Create feature branch
git checkout -b feature/amazing-feature

# 3. Make changes & test
streamlit run app.py

# 4. Commit & push
git commit -m "Add amazing feature"
git push origin feature/amazing-feature

# 5. Open Pull Request
```

### Ideas for Contributions
- [ ] Add more email categories
- [ ] Support for other email providers (Gmail, etc.)
- [ ] Custom report templates
- [ ] Email scheduling
- [ ] Team/group report generation
- [ ] Real-time email monitoring
- [ ] Export to PDF, HTML
- [ ] Mobile app version

---

## 📝 License

This project is licensed under the MIT License - see LICENSE file for details.

---

## 🙏 Acknowledgments

- **Microsoft Graph API** - Email data access
- **Groq AI** - Fast LLM inference (Llama 3.1)
- **Streamlit** - Web framework
- **python-docx** - DOCX generation

---

## 📞 Support & Contact

- 📧 Email: support@yourapp.com
- 🐛 Bug Reports: [GitHub Issues](https://github.com/yourusername/outlook-ai-report-generator/issues)
- 💡 Feature Requests: [GitHub Discussions](https://github.com/yourusername/outlook-ai-report-generator/discussions)
- 🐦 Twitter: [@yourhandle](https://twitter.com/yourhandle)

---

## 🎓 Learning Resources

### Understanding the Code

1. **Microsoft Graph API**: [Official Docs](https://learn.microsoft.com/en-us/graph/)
2. **OAuth 2.0**: [Auth0 Guide](https://auth0.com/intro-to-iam/what-is-oauth-2)
3. **Streamlit**: [Tutorial](https://docs.streamlit.io/library/get-started)
4. **python-docx**: [Documentation](https://python-docx.readthedocs.io/)
5. **Groq API**: [API Reference](https://console.groq.com/docs)

### Video Tutorials
- Building Streamlit apps
- OAuth authentication flows
- Word document generation in Python
- AI email analysis workflows

---

## 📊 Stats & Metrics

```
📧 Supported: Outlook emails
🤖 AI Model: Llama 3.1 (8B) via Groq
📄 Export: DOCX format
🔐 Auth: OAuth 2.0
⚡ Avg Speed: ~1-2 minutes
📈 Emails/Run: Up to 20
💾 Report Size: ~500KB
```

---

## 🎯 Roadmap

### Version 2.0 (Planned)
- [ ] Multiple email provider support
- [ ] Custom report templates
- [ ] Scheduled report generation
- [ ] Team collaboration features
- [ ] Advanced filtering options
- [ ] Export to PDF format

### Version 3.0 (Future)
- [ ] Real-time email monitoring
- [ ] Mobile application
- [ ] Browser extension
- [ ] Integration with Slack/Teams
- [ ] Advanced analytics dashboard

---

<div align="center">

### ⭐ If you found this helpful, please star the repository!

Made with ❤️ using Streamlit, Groq AI, and Microsoft Graph API

[⬆ back to top](#outlook-ai-report-generator-)

</div>
