"""
Job Application Automation - Complete Single File Version
All-in-one solution with manual email entry and AI-powered email generation
"""

import streamlit as st
import pandas as pd
import json
from pathlib import Path
import tempfile
import os
import re
import time
import requests
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import List, Dict, Optional, Tuple

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

try:
    import docx
except ImportError:
    docx = None


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def validate_email(email: str) -> bool:
    """Validate email format"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


def parse_email_content(email_content: str) -> Tuple[str, str]:
    """Parse generated email to extract subject and body"""
    lines = email_content.strip().split('\n')
    subject = ""
    body_lines = []
    
    for i, line in enumerate(lines):
        if line.lower().startswith('subject:'):
            subject = line.split(':', 1)[1].strip()
        elif i > 0 and (subject or i > 2):
            body_lines.append(line)
    
    body = '\n'.join(body_lines).strip()
    
    if not subject:
        subject = "Job Application - Opportunity Inquiry"
    
    return subject, body


def parse_email_list(email_string: str) -> List[str]:
    """Parse comma or newline separated email list"""
    separators = [',', ';', '\n', ' ']
    emails = [email_string]
    
    for sep in separators:
        new_emails = []
        for email in emails:
            new_emails.extend(email.split(sep))
        emails = new_emails
    
    valid_emails = []
    for email in emails:
        email = email.strip()
        if email and validate_email(email):
            valid_emails.append(email.lower())
    
    return list(set(valid_emails))


# =============================================================================
# AUTOMATION CLASS
# =============================================================================

class JobApplicationAutomation:
    """Main automation class"""
    
    def __init__(self, groq_api_key: str = None):
        self.groq_api_key = groq_api_key
        self.email_regex = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    
    def read_resume(self, file_path: str) -> str:
        """Read resume from various file formats"""
        try:
            if file_path.endswith('.txt'):
                with open(file_path, 'r', encoding='utf-8') as f:
                    text = f.read()
                    
            elif file_path.endswith('.pdf') and PyPDF2:
                text = ""
                with open(file_path, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    for page in pdf_reader.pages:
                        text += page.extract_text()
                        
            elif file_path.endswith('.docx') and docx:
                doc = docx.Document(file_path)
                text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
                
            else:
                raise ValueError("❌ Resume must be .txt, .pdf, or .docx format")
            
            return text.strip()
            
        except Exception as e:
            raise Exception(f"❌ Error reading resume: {str(e)}")
    
    def generate_email_with_groq(self, resume_text: str, job_description: str, 
                                  recipient_info: Dict, custom_instructions: str = "") -> str:
        """Generate personalized job application email using Groq API"""
        if not self.groq_api_key:
            raise ValueError("❌ Groq API key is required")
        
        company = recipient_info.get('company', recipient_info.get('title', 'Hiring Team'))
        
        prompt = f"""You are an expert job application email writer. Generate a professional, compelling email for a job application.

RESUME SUMMARY:
{resume_text[:3000]}

COMPANY/POSITION: {company}
JOB DESCRIPTION:
{job_description[:2000]}

{f"CUSTOM INSTRUCTIONS: {custom_instructions}" if custom_instructions else ""}

INSTRUCTIONS:
1. Write a professional subject line that mentions the job/position
2. Start with appropriate greeting (Dear Hiring Manager/Sir/Madam)
3. Brief introduction (1-2 sentences) - who you are and why you're writing
4. Highlight 2-3 most relevant skills/experiences from the resume that match this opportunity
5. Show genuine enthusiasm and interest in the role/company
6. Include a clear call to action (request for interview/discussion)
7. Professional closing with full name

FORMATTING REQUIREMENTS:
- Keep email between 150-250 words
- Use professional but friendly tone
- Be specific and avoid generic statements
- Format as a complete email ready to send
- Include "Subject:" line at the top
- Do NOT include email signature or contact details at the end

Generate the complete email now:"""

        try:
            url = "https://api.groq.com/openai/v1/chat/completions"
            headers = {
                "Authorization": f"Bearer {self.groq_api_key}",
                "Content-Type": "application/json"
            }
            
            payload = {
                "model": "mixtral-8x7b-32768",
                "messages": [
                    {
                        "role": "system",
                        "content": "You are a professional job application email writer. Write clear, compelling, and personalized emails."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                "temperature": 0.7,
                "max_tokens": 800,
                "top_p": 0.9
            }
            
            response = requests.post(url, json=payload, headers=headers, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            email_content = data["choices"][0]["message"]["content"].strip()
            
            return email_content
            
        except requests.exceptions.Timeout:
            raise Exception("⏱️ Groq API request timed out")
        except requests.exceptions.RequestException as e:
            raise Exception(f"❌ Groq API error: {str(e)}")
        except Exception as e:
            raise Exception(f"❌ Error generating email: {str(e)}")
    
    def send_email_smtp(self, to_emails: List[str], subject: str, body: str,
                       smtp_config: Dict, resume_path: str = None) -> Dict[str, bool]:
        """Send emails using SMTP"""
        results = {}
        
        for to_email in to_emails:
            try:
                msg = MIMEMultipart()
                msg['From'] = f"{smtp_config.get('from_name', '')} <{smtp_config['from_email']}>"
                msg['To'] = to_email
                msg['Subject'] = subject
                msg['Date'] = datetime.now().strftime('%a, %d %b %Y %H:%M:%S %z')
                
                msg.attach(MIMEText(body, 'plain', 'utf-8'))
                
                if resume_path and os.path.exists(resume_path):
                    with open(resume_path, 'rb') as f:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(f.read())
                    
                    encoders.encode_base64(part)
                    filename = os.path.basename(resume_path)
                    part.add_header('Content-Disposition', f'attachment; filename= {filename}')
                    msg.attach(part)
                
                with smtplib.SMTP(smtp_config['host'], smtp_config['port']) as server:
                    server.starttls()
                    server.login(smtp_config['username'], smtp_config['password'])
                    server.send_message(msg)
                
                results[to_email] = True
                time.sleep(1)
                
            except Exception as e:
                results[to_email] = False
                print(f"  ❌ Failed to send to {to_email}: {str(e)}")
        
        return results


# =============================================================================
# STREAMLIT UI
# =============================================================================

# Page configuration
st.set_page_config(
    page_title="Job Application Automation",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: 600;
        border: none;
        padding: 0.75rem;
        border-radius: 10px;
    }
    h1, h2, h3 {
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'automation' not in st.session_state:
    st.session_state.automation = None
if 'manual_email_list' not in st.session_state:
    st.session_state.manual_email_list = []
if 'resume_text' not in st.session_state:
    st.session_state.resume_text = ""
if 'resume_path' not in st.session_state:
    st.session_state.resume_path = None
if 'generated_emails' not in st.session_state:
    st.session_state.generated_emails = {}
if 'current_step' not in st.session_state:
    st.session_state.current_step = 1

# Title
st.markdown("<h1 style='text-align: center; font-size: 3rem;'>🚀 Job Application Automation</h1>", 
            unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #ddd;'>AI-Powered Email Generation</p>", 
            unsafe_allow_html=True)
st.markdown("---")

# Sidebar
st.sidebar.title("⚙️ Configuration")
groq_key = st.sidebar.text_input("Groq API Key (Required)", type="password", 
                                  help="Get from console.groq.com")

if groq_key:
    st.session_state.automation = JobApplicationAutomation(groq_key)
    st.sidebar.success("✅ API Key configured!")
else:
    st.sidebar.warning("⚠️ Please enter Groq API key")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Status")
if st.session_state.manual_email_list:
    st.sidebar.info(f"📧 Emails: {len(st.session_state.manual_email_list)}")
if st.session_state.resume_text:
    st.sidebar.success(f"📄 Resume: Uploaded")

# Main content
if not groq_key:
    st.info("👈 Please enter your Groq API key in the sidebar")
    st.info("💡 Get free API key from: https://console.groq.com")
    st.stop()

# Step 1: Email Entry & Resume
if st.session_state.current_step == 1:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📧 Step 1: Add Email Addresses")
        
        # Bulk entry
        st.subheader("Bulk Entry (paste multiple emails)")
        bulk_emails = st.text_area(
            "Paste emails here (one per line or comma-separated):",
            height=150,
            placeholder="hr@company1.com\njobs@company2.com\ncareers@company3.com"
        )
        
        if st.button("➕ Add Bulk Emails"):
            if bulk_emails:
                parsed_emails = parse_email_list(bulk_emails)
                if parsed_emails:
                    for email in parsed_emails:
                        if email not in [e['email'] for e in st.session_state.manual_email_list]:
                            st.session_state.manual_email_list.append({
                                'email': email,
                                'company': '',
                                'description': 'Manual Entry'
                            })
                    st.success(f"✅ Added {len(parsed_emails)} emails!")
                    st.rerun()
        
        st.markdown("---")
        st.subheader("Single Entry with Details")
        
        col_a, col_b = st.columns(2)
        with col_a:
            single_email = st.text_input("Email Address")
            company_name = st.text_input("Company Name (Optional)")
        with col_b:
            job_title = st.text_input("Job Title (Optional)")
        
        if st.button("➕ Add Email with Details"):
            if single_email and validate_email(single_email):
                description = "Manual Entry"
                if company_name:
                    description += f" - {company_name}"
                if job_title:
                    description += f" ({job_title})"
                
                st.session_state.manual_email_list.append({
                    'email': single_email,
                    'company': company_name,
                    'description': description
                })
                st.success(f"✅ Added {single_email}")
                st.rerun()
            else:
                st.error("Invalid email address")
        
        # Display added emails
        if st.session_state.manual_email_list:
            st.markdown("---")
            st.subheader(f"📋 Email List ({len(st.session_state.manual_email_list)} emails)")
            
            for i, email_data in enumerate(st.session_state.manual_email_list):
                col_x, col_y = st.columns([5, 1])
                with col_x:
                    st.text(f"✉️ {email_data['email']} - {email_data.get('company', 'N/A')}")
                with col_y:
                    if st.button("🗑️", key=f"del_{i}"):
                        st.session_state.manual_email_list.pop(i)
                        st.rerun()
            
            if st.button("🗑️ Clear All"):
                st.session_state.manual_email_list = []
                st.rerun()
    
    with col2:
        st.header("📄 Step 2: Upload Resume")
        resume_file = st.file_uploader("Upload your resume", type=['txt', 'pdf', 'docx'])
        
        if resume_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(resume_file.name).suffix) as tmp:
                tmp.write(resume_file.getvalue())
                tmp_path = tmp.name
            
            try:
                resume_text = st.session_state.automation.read_resume(tmp_path)
                st.session_state.resume_text = resume_text
                st.session_state.resume_path = tmp_path
                st.success("✅ Resume uploaded!")
                st.text_area("Preview", resume_text[:300] + "...", height=150, disabled=True)
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
        
        st.markdown("---")
        st.subheader("🎯 Custom Instructions")
        custom_instructions = st.text_area(
            "Optional instructions for email generation:",
            placeholder="E.g., 'Emphasize my 5 years of hotel management experience'",
            height=100,
            key="custom_inst"
        )
        
        if st.session_state.resume_text and st.session_state.manual_email_list:
            if st.button("✨ Generate Emails", type="primary"):
                st.session_state.custom_instructions = custom_instructions
                st.session_state.current_step = 2
                st.rerun()
        else:
            if not st.session_state.resume_text:
                st.warning("⚠️ Upload resume")
            if not st.session_state.manual_email_list:
                st.warning("⚠️ Add emails")

# Step 2: Generate Emails
elif st.session_state.current_step == 2:
    st.header("🤖 Generating AI-Powered Emails")
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total = len(st.session_state.manual_email_list)
    
    try:
        for i, email_data in enumerate(st.session_state.manual_email_list):
            status_text.text(f"Generating {i+1}/{total} for {email_data['email']}...")
            progress_bar.progress((i + 1) / total)
            
            email_content = st.session_state.automation.generate_email_with_groq(
                resume_text=st.session_state.resume_text,
                job_description=email_data.get('description', ''),
                recipient_info=email_data,
                custom_instructions=st.session_state.get('custom_instructions', '')
            )
            
            st.session_state.generated_emails[email_data['email']] = email_content
            time.sleep(2)  # Rate limiting
        
        st.session_state.current_step = 3
        st.success(f"✅ Generated {total} emails!")
        status_text.empty()
        progress_bar.empty()
        time.sleep(1)
        st.rerun()
        
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        status_text.empty()
        progress_bar.empty()

# Step 3: Review & Edit
elif st.session_state.current_step == 3:
    st.header("✏️ Review & Edit Emails")
    
    for email, content in st.session_state.generated_emails.items():
        with st.expander(f"📧 {email}", expanded=False):
            edited = st.text_area(
                "Email Content (editable):",
                value=content,
                height=300,
                key=f"edit_{email}"
            )
            st.session_state.generated_emails[email] = edited
            
            subject, body = parse_email_content(edited)
            st.markdown(f"**Subject:** {subject}")
    
    if st.button("✅ Continue to Send", type="primary"):
        st.session_state.current_step = 4
        st.rerun()
    
    if st.button("⬅️ Back"):
        st.session_state.current_step = 1
        st.rerun()

# Step 4: Send Emails
elif st.session_state.current_step == 4:
    st.header("📤 Send Emails")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        smtp_host = st.text_input("SMTP Host", value="smtp.gmail.com")
        smtp_port = st.number_input("SMTP Port", value=587)
    
    with col2:
        smtp_username = st.text_input("Email Username")
        smtp_password = st.text_input("Password", type="password")
    
    with col3:
        from_email = st.text_input("From Email", value=smtp_username)
        from_name = st.text_input("Your Name")
    
    attach_resume = st.checkbox("Attach Resume", value=True)
    
    st.info(f"📊 Ready to send {len(st.session_state.generated_emails)} emails")
    st.warning("⚠️ Gmail users: Use App Password (not regular password)")
    
    if st.button("🚀 Send All", type="primary"):
        if not all([smtp_host, smtp_username, smtp_password, from_email]):
            st.error("❌ Fill in all SMTP fields")
        else:
            smtp_config = {
                'host': smtp_host,
                'port': int(smtp_port),
                'username': smtp_username,
                'password': smtp_password,
                'from_email': from_email,
                'from_name': from_name
            }
            
            progress = st.progress(0)
            status = st.empty()
            
            try:
                all_results = {}
                emails_list = list(st.session_state.generated_emails.keys())
                
                for i, email in enumerate(emails_list):
                    status.text(f"Sending {i+1}/{len(emails_list)} to {email}...")
                    progress.progress((i + 1) / len(emails_list))
                    
                    content = st.session_state.generated_emails[email]
                    subject, body = parse_email_content(content)
                    
                    result = st.session_state.automation.send_email_smtp(
                        to_emails=[email],
                        subject=subject,
                        body=body,
                        smtp_config=smtp_config,
                        resume_path=st.session_state.resume_path if attach_resume else None
                    )
                    
                    all_results.update(result)
                
                status.empty()
                progress.empty()
                
                success_count = sum(all_results.values())
                st.success(f"✅ Sent {success_count}/{len(all_results)} emails!")
                
                for email, success in all_results.items():
                    if success:
                        st.success(f"✅ {email}")
                    else:
                        st.error(f"❌ {email}")
                
                if st.button("🔄 Start New Campaign"):
                    for key in ['manual_email_list', 'resume_text', 'resume_path', 
                               'generated_emails', 'custom_instructions']:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.session_state.current_step = 1
                    st.rerun()
                
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                status.empty()
                progress.empty()
    
    if st.button("⬅️ Back"):
        st.session_state.current_step = 3
        st.rerun()

# Footer
st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: #ddd;'>
        <p>🚀 Job Application Automation • Powered by Groq AI</p>
    </div>
""", unsafe_allow_html=True)