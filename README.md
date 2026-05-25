# 🚀 Job Application Automation System

An intelligent, AI-powered system to automate job applications using web scraping, email extraction, and personalized email generation.

## ✨ Features

1. **🔍 Web Scraping** - Scrape Google search results using Apify SERP scraper
2. **📧 Email Extraction** - Automatically extract email addresses from scraped data
3. **📄 Resume Upload** - Support for TXT, PDF, and DOCX formats
4. **🤖 AI Email Generation** - Generate personalized job application emails using Groq AI
5. **✏️ Email Editing** - Review and edit generated emails before sending
6. **📤 Bulk Email Sending** - Send emails to multiple recipients with SMTP
7. **📊 Progress Tracking** - Real-time progress updates and success tracking

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────┐
│                   User Interface                     │
│               (Streamlit Web App)                    │
└──────────────────┬──────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────┐
│              Job Automation Backend                  │
│           (job_automation.py)                        │
├─────────────────────────────────────────────────────┤
│  • Apify SERP Scraper Integration                   │
│  • Email Regex Extraction                           │
│  • Resume Parser (PDF/DOCX/TXT)                     │
│  • Groq API Integration                             │
│  • SMTP Email Sender                                │
└──────────────────┬──────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────┐
│              External Services                       │
│  • Apify (Web Scraping)                             │
│  • Groq AI (Email Generation)                       │
│  • SMTP Server (Email Sending)                      │
└─────────────────────────────────────────────────────┘
```

## 📋 Prerequisites

1. **Python 3.8+**
2. **API Keys:**
   - Apify API Key ([Get it here](https://apify.com))
   - Groq API Key ([Get it here](https://console.groq.com))
3. **SMTP Credentials** (Gmail, Outlook, etc.)

## 🚀 Installation

### Step 1: Clone or Download

```bash
# If you have git
git clone <your-repo-url>
cd job-automation

# Or download and extract the files
```

### Step 2: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 3: Run the Application

```bash
streamlit run streamlit_app.py
```

The app will open in your default browser at `http://localhost:8501`

## 📖 Usage Guide

### 1. Configure API Keys

- Open the sidebar in the Streamlit app
- Enter your **Apify API Key**
- Enter your **Groq API Key**

### 2. Get Data (Two Options)

**Option A: Scrape with Apify**
- Enter search keywords (e.g., "hotel, email, jobs")
- Select number of pages to scrape
- Click "Run Scraper"
- Wait 1-2 minutes for results

**Option B: Upload Existing Data**
- Upload a CSV or Excel file with scraped data
- File must have a "Results" column with description text

### 3. Extract Emails

- Click "Extract Emails" button
- System will find all email addresses in the data
- Review the extracted emails list

### 4. Select Recipients & Upload Resume

- **Select Email Recipients:**
  - Click on individual emails to select/deselect
  - Use "Select All" or "Deselect All" buttons
  - Add manual emails if needed

- **Upload Resume:**
  - Upload your resume (TXT, PDF, or DOCX)
  - System will extract text for AI processing

### 5. Generate AI Emails

- Click "Generate Personalized Emails"
- AI will create custom emails for each recipient
- Wait while emails are generated (2-3 seconds each)

### 6. Review & Edit

- Review each generated email
- Edit any email content as needed
- Check subject lines and body text

### 7. Send Emails

- Configure SMTP settings:
  - **Gmail:** smtp.gmail.com, port 587
  - **Outlook:** smtp.office365.com, port 587
- Enter your email credentials
- Choose whether to attach resume
- Click "Send All Emails"

## 🔧 Configuration

### Gmail SMTP Setup

1. Enable 2-factor authentication on your Gmail account
2. Generate an [App Password](https://myaccount.google.com/apppasswords)
3. Use the app password in the SMTP configuration

### Outlook SMTP Setup

- Host: `smtp.office365.com`
- Port: `587`
- Use your regular Outlook password

## 📊 Example Data Format

Your uploaded CSV/Excel should have this structure:

```csv
Results,Search_term,Page_number
"Title: Hotel Manager Position
Description: We're hiring! Email: hr@hotelexample.com
Contact us for details","hotel email jobs",1
```

## 🎯 Use Cases

1. **Job Hunting** - Apply to multiple companies automatically
2. **Freelancing** - Reach out to potential clients
3. **Business Development** - Contact prospects with personalized emails
4. **Networking** - Connect with industry professionals

## 🛠️ Troubleshooting

### Scraper Times Out
- Reduce number of pages to scrape
- Check your Apify API key and credits
- Try uploading pre-scraped data instead

### No Emails Found
- Check if "Results" column contains email addresses
- Verify data format is correct
- Try different column name in extraction

### Email Sending Fails
- Verify SMTP credentials are correct
- For Gmail, use App Password not regular password
- Check firewall/antivirus isn't blocking SMTP
- Verify "Less secure app access" settings

### AI Generation Fails
- Check Groq API key is valid
- Verify you have sufficient API credits
- Ensure resume text was uploaded correctly

## 📦 File Structure

```
job-automation/
├── job_automation.py      # Backend automation class
├── streamlit_app.py        # Streamlit UI
├── requirements.txt        # Python dependencies
└── README.md              # This file
```

## 🔐 Security Notes

- Never commit API keys to version control
- Use environment variables for production
- Keep SMTP credentials secure
- Be careful with email sending limits to avoid spam flags

## 📈 Best Practices

1. **Test First** - Send to yourself first to verify
2. **Personalize** - Review and edit AI-generated emails
3. **Respect Limits** - Don't send too many emails at once
4. **Follow Up** - Track responses and follow up appropriately
5. **Be Professional** - Ensure all emails are polite and professional

## ⚡ Performance Tips

- **Batch Processing** - Process 10-20 emails at a time
- **Rate Limiting** - Add delays between sends (built-in: 1 sec)
- **API Limits** - Monitor your Apify and Groq API usage
- **Resume Size** - Keep resume under 5 pages for best AI results

## 🚨 Known Limitations

1. Email extraction works best with public contact information
2. AI generation quality depends on resume and job description quality
3. SMTP servers may have daily sending limits
4. Apify scraping is subject to Google's rate limits

## 🤝 Contributing

Feel free to:
- Report bugs
- Suggest features
- Submit pull requests
- Improve documentation

## 📄 License

This project is for educational and personal use. Please respect:
- Website terms of service when scraping
- Email marketing laws (CAN-SPAM, GDPR)
- Platform usage policies

## 💡 Tips for Success

1. **Quality over Quantity** - Better to send 10 great emails than 100 generic ones
2. **Customize Keywords** - Use specific search terms for better targeting
3. **Update Resume** - Keep your resume current for best AI results
4. **Test Thoroughly** - Always test with your own email first
5. **Track Results** - Keep notes on which approaches work best

## 🆘 Support

If you encounter issues:
1. Check the troubleshooting section
2. Verify all API keys are correct
3. Ensure dependencies are installed
4. Check console for error messages

## 🎓 Learning Resources

- [Streamlit Documentation](https://docs.streamlit.io)
- [Apify Documentation](https://docs.apify.com)
- [Groq API Documentation](https://console.groq.com/docs)
- [SMTP Email Basics](https://developers.google.com/gmail/api/guides)

---

**Built with ❤️ using Python, Streamlit, and AI**

**Version:** 1.0.0  
**Last Updated:** January 2026