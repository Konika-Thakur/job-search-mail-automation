"""
Complete Job Application Automation System - Production Ready (PROFESSIONAL VERSION)
All Features: Google Jobs, Freelancing, Profile Management, Multiple Attachments, Excel Processing
Fixed: File naming preservation, Professional styling, Emoji removal
"""

import streamlit as st
import pandas as pd
import json
from pathlib import Path
import tempfile
import os
import time
from datetime import datetime
import re
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import List, Dict, Tuple

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

try:
    import docx
except ImportError:
    docx = None


# =============================================================================
# CSS LOADER
# =============================================================================

def load_css(css_file_path: str) -> None:
    """Load and inject external CSS file into Streamlit app"""
    css_path = Path(css_file_path)
    if css_path.exists():
        with open(css_path, "r", encoding="utf-8") as f:
            css_content = f.read()
        st.markdown(f"<style>{css_content}</style>", unsafe_allow_html=True)
    else:
        st.warning(f"CSS file not found: {css_file_path}")


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


def read_document(file_path: str) -> str:
    """Read document content from various formats"""
    try:
        ext = Path(file_path).suffix.lower()
        
        if ext == '.txt':
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                if content.strip():
                    return content
                else:
                    raise Exception("Text file is empty")
            except Exception as e:
                raise Exception(f"Error reading TXT file: {str(e)}")
                
        elif ext == '.pdf':
            if PyPDF2 is None:
                raise Exception("PyPDF2 not installed. Install with: pip install PyPDF2")
            
            try:
                text = ""
                with open(file_path, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    
                    if len(pdf_reader.pages) == 0:
                        raise Exception("PDF has no pages")
                    
                    for page_num, page in enumerate(pdf_reader.pages):
                        try:
                            page_text = page.extract_text()
                            if page_text:
                                text += page_text + "\n"
                        except Exception as page_error:
                            st.warning(f"Could not extract text from page {page_num + 1}: {str(page_error)}")
                            continue
                
                if text.strip():
                    return text.strip()
                else:
                    raise Exception("Could not extract any text from PDF. The PDF might be scanned/image-based.")
                    
            except Exception as e:
                raise Exception(f"Error reading PDF: {str(e)}")
                
        elif ext == '.docx':
            if docx is None:
                raise Exception("python-docx not installed. Install with: pip install python-docx")
            
            try:
                doc = docx.Document(file_path)
                paragraphs = [paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()]
                
                if paragraphs:
                    return "\n".join(paragraphs)
                else:
                    raise Exception("DOCX file appears to be empty")
                    
            except Exception as e:
                raise Exception(f"Error reading DOCX file: {str(e)}")
        else:
            raise Exception(f"Unsupported file format: {ext}. Please use TXT, PDF, or DOCX")
            
    except Exception as e:
        st.error(f"Document Reading Error: {str(e)}")
        return ""


# =============================================================================
# JOB SEARCH CLASS
# =============================================================================

class JobSearchEngine:
    """Handles job searching from multiple sources using SerpAPI"""
    
    def __init__(self, serp_api_key: str):
        self.serp_api_key = serp_api_key.strip()
        self.base_url = "https://serpapi.com/search"
    
    def search_google_jobs(self, keywords: str, location: str = "India") -> pd.DataFrame:
        """Search Google Jobs using SerpAPI"""
        try:
            params = {
                "engine": "google_jobs",
                "q": keywords,
                "location": location,
                "api_key": self.serp_api_key
            }
            
            response = requests.get(self.base_url, params=params, timeout=30)
            response.raise_for_status()
            results = response.json()
            
            jobs = results.get("jobs_results", [])
            job_data = []
            
            for job in jobs:
                job_data.append({
                    "Title": job.get("title", ""),
                    "Company": job.get("company_name", ""),
                    "Location": job.get("location", ""),
                    "Type": job.get("detected_extensions", {}).get("schedule_type", ""),
                    "Posted": job.get("detected_extensions", {}).get("posted_at", ""),
                    "Description": job.get("description", "")[:500],
                    "Apply_Link": job.get("apply_options", [{}])[0].get("link", "") if job.get("apply_options") else ""
                })
            
            return pd.DataFrame(job_data)
        except Exception as e:
            st.error(f"Google Jobs search error: {str(e)}")
            return pd.DataFrame()
    
    def search_freelance_platforms(self, keywords: str) -> pd.DataFrame:
        """Search freelancing platforms using Google search"""
        all_results = []
        platforms = [
            ("upwork.com", "Upwork"),
            ("freelancer.com", "Freelancer"),
            ("fiverr.com", "Fiverr"),
            ("toptal.com", "Toptal"),
            ("guru.com", "Guru")
        ]
        
        for site, platform_name in platforms:
            try:
                query = f"site:{site} {keywords}"
                
                params = {
                    "engine": "google",
                    "q": query,
                    "num": 10,
                    "api_key": self.serp_api_key
                }
                
                response = requests.get(self.base_url, params=params, timeout=30)
                response.raise_for_status()
                results = response.json()
                
                for result in results.get("organic_results", []):
                    all_results.append({
                        "Platform": platform_name,
                        "Title": result.get("title", ""),
                        "Link": result.get("link", ""),
                        "Description": result.get("snippet", "")[:300]
                    })
                
                time.sleep(1)
            except Exception as e:
                st.warning(f"{platform_name} search error: {str(e)}")
                continue
        
        return pd.DataFrame(all_results)
    
    def extract_emails_from_results(self, df: pd.DataFrame) -> List[str]:
        """Extract email addresses from search results"""
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        emails = []
        
        for _, row in df.iterrows():
            content = " ".join([str(v) for v in row.values if pd.notna(v)])
            found_emails = re.findall(email_pattern, content.lower())
            
            for email in found_emails:
                if validate_email(email):
                    emails.append(email)
        
        return list(set(emails))


# =============================================================================
# EMAIL AUTOMATION CLASS
# =============================================================================

class EmailAutomation:
    """Handles email generation and sending with profile support"""
    
    def __init__(self, groq_api_key: str, sender_profile: Dict):
        self.groq_api_key = groq_api_key.strip()
        self.sender_profile = sender_profile
    
    def generate_email(self, resume_text: str, job_desc: str, 
                      recipient_info: Dict, custom_inst: str = "") -> str:
        """Generate personalized email with sender profile"""
        
        def clean_text(text: str) -> str:
            """Remove emojis and non-ASCII characters that cause encoding issues"""
            if not text:
                return ""
            text = str(text)
            emoji_pattern = re.compile(
                "["
                "\U0001F600-\U0001F64F"
                "\U0001F300-\U0001F5FF"
                "\U0001F680-\U0001F6FF"
                "\U0001F1E0-\U0001F1FF"
                "\U00002500-\U00002BEF"
                "\U00002702-\U000027B0"
                "\U000024C2-\U0001F251"
                "\U0001f926-\U0001f937"
                "\U00010000-\U0010ffff"
                "\u2640-\u2642"
                "\u2600-\u2B55"
                "\u200d"
                "\u23cf"
                "\u23e9"
                "\u231a"
                "\ufe0f"
                "\u3030"
                "]+", 
                flags=re.UNICODE
            )
            cleaned = emoji_pattern.sub(r'', text)
            cleaned = ''.join(char for char in cleaned if 32 <= ord(char) <= 126 or char in '\n\r\t')
            return cleaned.strip()
        
        # Build profile context
        profile_context = []
        if self.sender_profile.get('name'):
            name = clean_text(str(self.sender_profile['name']))
            if name:
                profile_context.append(f"Name: {name}")
        if self.sender_profile.get('current_role'):
            role = clean_text(str(self.sender_profile['current_role']))
            if role:
                profile_context.append(f"Current Role: {role}")
        if self.sender_profile.get('location'):
            loc = clean_text(str(self.sender_profile['location']))
            if loc:
                profile_context.append(f"Location: {loc}")
        
        if self.sender_profile.get('linkedin'):
            profile_context.append(f"LinkedIn: {str(self.sender_profile['linkedin'])}")
        if self.sender_profile.get('github'):
            profile_context.append(f"GitHub: {str(self.sender_profile['github'])}")
        if self.sender_profile.get('portfolio'):
            profile_context.append(f"Portfolio: {str(self.sender_profile['portfolio'])}")
        
        profile_str = "\n".join(profile_context) if profile_context else "Professional applicant"
        
        company = clean_text(str(recipient_info.get('company', recipient_info.get('title', 'Hiring Team'))))
        if not company:
            company = "Hiring Team"
        
        resume_text = str(resume_text)[:2000] if resume_text else ""
        job_desc = str(job_desc)[:1200] if job_desc else ""
        custom_inst = str(custom_inst)[:300] if custom_inst else ""
        
        clean_resume = clean_text(resume_text)
        clean_job_desc = clean_text(job_desc)
        clean_custom_inst = clean_text(custom_inst)
        
        prompt = f"""Generate a professional job application email.

APPLICANT INFO:
{profile_str}

RESUME HIGHLIGHTS:
{clean_resume[:1500]}

TARGET: {company}

JOB INFO:
{clean_job_desc[:1000]}"""

        if clean_custom_inst:
            prompt += f"\n\nSPECIAL INSTRUCTIONS:\n{clean_custom_inst[:250]}"

        prompt += """

REQUIREMENTS:
- Start with "Subject: [Job Title] Application - [Your Name]"
- Professional greeting
- 2-3 paragraphs highlighting relevant skills
- Call to action
- Sign off with "Best regards,"

Write the email now (150-200 words):"""

        try:
            url = "https://api.groq.com/openai/v1/chat/completions"
            headers = {
                "Authorization": f"Bearer {self.groq_api_key}",
                "Content-Type": "application/json"
            }
            
            payload = {
                "model": "llama-3.3-70b-versatile",
                "messages": [
                    {"role": "system", "content": "You are a professional email writer. Write clear, concise job application emails."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 600
            }
            
            if len(prompt) > 4000:
                raise Exception(f"Prompt too long ({len(prompt)} chars). Please shorten your resume or job description.")
            
            response = requests.post(url, json=payload, headers=headers, timeout=30)
            
            if response.status_code == 401:
                raise Exception("Invalid Groq API key. Please check your API key and try again.")
            elif response.status_code == 400:
                error_detail = response.text
                raise Exception(f"Bad request to Groq API. The prompt may be too long or contain invalid characters. Details: {error_detail[:200]}")
            
            response.raise_for_status()
            
            data = response.json()
            email_content = data["choices"][0]["message"]["content"].strip()
            
            signature = self._build_signature()
            return email_content + "\n" + signature
            
        except requests.exceptions.RequestException as e:
            if "401" in str(e):
                raise Exception("Invalid Groq API key. Please verify your API key is correct.")
            elif "400" in str(e):
                raise Exception(f"Bad request - your resume or job description may be too long or contain invalid characters. Try shortening them. Error: {str(e)[:150]}")
            else:
                raise Exception(f"API request failed: {str(e)}")
        except Exception as e:
            raise Exception(f"Email generation error: {str(e)}")
    
    def _build_signature(self) -> str:
        """Build professional email signature"""
        sig = ["\n\nBest regards,"]
        if self.sender_profile.get('name'):
            sig.append(self.sender_profile['name'])
        if self.sender_profile.get('current_role'):
            sig.append(self.sender_profile['current_role'])
        if self.sender_profile.get('email'):
            sig.append(f"Email: {self.sender_profile['email']}")
        if self.sender_profile.get('phone'):
            sig.append(f"Phone: {self.sender_profile['phone']}")
        if self.sender_profile.get('linkedin'):
            sig.append(f"LinkedIn: {self.sender_profile['linkedin']}")
        if self.sender_profile.get('github'):
            sig.append(f"GitHub: {self.sender_profile['github']}")
        if self.sender_profile.get('portfolio'):
            sig.append(f"Portfolio: {self.sender_profile['portfolio']}")
        if self.sender_profile.get('location'):
            sig.append(f"Location: {self.sender_profile['location']}")
        
        return "\n".join(sig)
    
    @staticmethod
    def send_mail(to_email, subject, body, smtp_config, attachments_with_names=None):
        """Send email with attachments preserving original filenames"""
        msg = MIMEMultipart()
        msg["From"] = smtp_config["from_email"]
        msg["To"] = to_email
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        if attachments_with_names:
            for file_path, original_name in attachments_with_names:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename={original_name}"
                    )
                    msg.attach(part)

        try:
            with smtplib.SMTP(smtp_config["host"], smtp_config["port"]) as server:
                server.starttls()
                server.login(smtp_config["username"], smtp_config["password"])
                server.sendmail(smtp_config["from_email"], to_email, msg.as_string())
            return True
        except Exception as e:
            st.error(f"Failed to send to {to_email}: {e}")
            return False


# =============================================================================
# EXCEL PROCESSOR
# =============================================================================

def process_excel_for_emails(file_path: str) -> List[str]:
    """Extract emails from Excel file"""
    try:
        df = pd.read_excel(file_path)
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        emails = []
        
        for col in df.columns:
            for val in df[col].dropna():
                found_emails = re.findall(email_pattern, str(val).lower())
                emails.extend([e for e in found_emails if validate_email(e)])
        
        return list(set(emails))
    except Exception as e:
        st.error(f"Excel processing error: {str(e)}")
        return []


# =============================================================================
# STREAMLIT UI
# =============================================================================

st.set_page_config(
    page_title="AI Job Application System",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load external CSS
load_css(Path(__file__).parent / "styles.css")

# Initialize session state
if 'sender_profile' not in st.session_state:
    st.session_state.sender_profile = {}
if 'email_list' not in st.session_state:
    st.session_state.email_list = []
if 'resume_text' not in st.session_state:
    st.session_state.resume_text = ""
if 'resume_path' not in st.session_state:
    st.session_state.resume_path = None
if 'resume_original_name' not in st.session_state:
    st.session_state.resume_original_name = None
if 'additional_docs' not in st.session_state:
    st.session_state.additional_docs = []
if 'generated_emails' not in st.session_state:
    st.session_state.generated_emails = {}
if 'current_step' not in st.session_state:
    st.session_state.current_step = 1
if 'job_search_results' not in st.session_state:
    st.session_state.job_search_results = {}
if 'custom_instructions' not in st.session_state:
    st.session_state.custom_instructions = ""
# Manual template state
if 'manual_subject' not in st.session_state:
    st.session_state.manual_subject = ""
if 'manual_body' not in st.session_state:
    st.session_state.manual_body = ""
if 'email_mode' not in st.session_state:
    st.session_state.email_mode = "ai"  # "ai" | "manual"

# Header
st.markdown("""
    <h1 style='text-align: center; margin-bottom: 0;'>
        AI Job Application System
    </h1>
    <p style='text-align: center; color: var(--text-secondary); font-size: 1.1rem; margin-top: 0.5rem;'>
    </p>
""", unsafe_allow_html=True)
st.markdown("---")

# Sidebar - Configuration
with st.sidebar:
    st.title("Configuration")
    
    with st.expander("API Keys", expanded=True):
        groq_key = st.text_input("Groq API Key*", type="password",
                                  help="Required for AI email generation")
        serp_key = st.text_input("SerpAPI Key", type="password",
                                  help="Optional: For job search features")
    
    st.markdown("---")
    
    with st.expander("Your Profile", expanded=True):
        st.markdown("**Personal Information**")
        st.session_state.sender_profile['name'] = st.text_input(
            "Full Name*", value=st.session_state.sender_profile.get('name', ''))
        st.session_state.sender_profile['email'] = st.text_input(
            "Email Address*", value=st.session_state.sender_profile.get('email', ''))
        st.session_state.sender_profile['email_app_key'] = st.text_input(
            "Email App Key*", type="password", value=st.session_state.sender_profile.get('email_app_key', ''))
        st.session_state.sender_profile['phone'] = st.text_input(
            "Phone Number", value=st.session_state.sender_profile.get('phone', ''))
        
        st.markdown("**Professional Details**")
        st.session_state.sender_profile['current_role'] = st.text_input(
            "Current Role/Title", value=st.session_state.sender_profile.get('current_role', ''))
        st.session_state.sender_profile['location'] = st.text_input(
            "Location", value=st.session_state.sender_profile.get('location', ''))
        
        st.markdown("**Online Presence**")
        st.session_state.sender_profile['linkedin'] = st.text_input(
            "LinkedIn URL", value=st.session_state.sender_profile.get('linkedin', ''))
        st.session_state.sender_profile['github'] = st.text_input(
            "GitHub URL", value=st.session_state.sender_profile.get('github', ''))
        st.session_state.sender_profile['portfolio'] = st.text_input(
            "Portfolio URL", value=st.session_state.sender_profile.get('portfolio', ''))
    
    st.markdown("---")
    
    st.markdown("### Status")
    if st.session_state.sender_profile.get('name') and st.session_state.sender_profile.get('email'):
        st.success(f"Profile: {st.session_state.sender_profile['name']}")
    else:
        st.warning("Complete your profile")
    
    if st.session_state.email_list:
        st.info(f"{len(st.session_state.email_list)} recipients")
    if st.session_state.resume_text:
        st.success("Resume uploaded")
    if st.session_state.additional_docs:
        st.info(f"{len(st.session_state.additional_docs)} additional docs")
    # Show active email mode
    if st.session_state.email_mode == "manual" and st.session_state.manual_subject:
        st.info("Mode: Manual Template")
    elif st.session_state.generated_emails:
        st.info(f"{len(st.session_state.generated_emails)} AI emails ready")

# Check requirements
if not groq_key:
    st.warning("Please enter Groq API Key in the sidebar")
    st.info("Get free API key from: https://console.groq.com/keys")
    st.markdown("""
        ### How to get your Groq API Key:
        1. Go to https://console.groq.com
        2. Sign up or log in
        3. Navigate to 'API Keys' section
        4. Click 'Create API Key'
        5. Copy and paste it in the sidebar
    """)
    st.stop()

if not (st.session_state.sender_profile.get('name') and st.session_state.sender_profile.get('email')):
    st.warning("Please complete your profile in the sidebar")
    st.stop()

# Main Content Tabs
tab1, tab2, tab3, tab4, tab_manual, tab5 = st.tabs([
    "Job Search",
    "Email Management",
    "Documents",
    "Generate Emails",
    "✍️ Manual Template",
    "Send Emails"
])

# TAB 1: Job Search
with tab1:
    st.header("Search Jobs & Extract Emails")
    
    if not serp_key:
        st.warning("SerpAPI key required for job search. Add it in the sidebar.")
        st.info("Get free API key from: https://serpapi.com/")
    else:
        job_searcher = JobSearchEngine(serp_key)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Google Jobs Search")
            job_keywords = st.text_input(
                "Search Keywords",
                placeholder="e.g., 'python developer', 'hotel manager'",
                key="job_keywords"
            )
            job_location = st.text_input("Location", value="India", key="job_location")
            
            if st.button("Search Google Jobs", key="search_google"):
                if job_keywords:
                    with st.spinner("Searching Google Jobs..."):
                        results = job_searcher.search_google_jobs(job_keywords, job_location)
                        if not results.empty:
                            st.session_state.job_search_results['google'] = results
                            st.success(f"Found {len(results)} jobs!")
                            st.dataframe(results, use_container_width=True)
                            
                            if st.button("Extract Emails from Results"):
                                emails = job_searcher.extract_emails_from_results(results)
                                if emails:
                                    for email in emails:
                                        if email not in [e['email'] for e in st.session_state.email_list]:
                                            st.session_state.email_list.append({
                                                'email': email,
                                                'company': 'From Job Search',
                                                'description': 'Google Jobs Search Result'
                                            })
                                    st.success(f"Added {len(emails)} emails!")
                                else:
                                    st.info("No emails found in search results")
                        else:
                            st.warning("No results found")
        
        with col2:
            st.subheader("Freelance Platforms")
            freelance_keywords = st.text_input(
                "Skills/Keywords",
                placeholder="e.g., 'python', 'web design', 'Gen AI'",
                key="freelance_keywords"
            )
            
            if st.button("Search Freelance Sites", key="search_freelance"):
                if freelance_keywords:
                    with st.spinner("Searching Upwork, Fiverr, Freelancer, Toptal, Guru..."):
                        results = job_searcher.search_freelance_platforms(freelance_keywords)
                        if not results.empty:
                            st.session_state.job_search_results['freelance'] = results
                            st.success(f"Found {len(results)} opportunities!")
                            st.dataframe(results, use_container_width=True)
                        else:
                            st.warning("No results found")
    
    st.markdown("---")
    st.subheader("Extract Emails from Excel")
    excel_file = st.file_uploader(
        "Upload Excel file (from Apify scraper or any source)",
        type=['xlsx', 'xls', 'csv'],
        key="excel_upload"
    )
    
    if excel_file:
        if st.button("Extract Emails from Excel"):
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', mode='wb') as tmp:
                tmp.write(excel_file.getvalue())
                tmp.flush()
                tmp_path = tmp.name
            
            with st.spinner("Processing Excel file..."):
                emails = process_excel_for_emails(tmp_path)
                if emails:
                    for email in emails:
                        if email not in [e['email'] for e in st.session_state.email_list]:
                            st.session_state.email_list.append({
                                'email': email,
                                'company': 'From Excel',
                                'description': 'Extracted from uploaded file'
                            })
                    st.success(f"Extracted and added {len(emails)} emails!")
                else:
                    st.warning("No valid emails found in Excel file")
            
            os.unlink(tmp_path)

# TAB 2: Email Management
with tab2:
    st.header("Manage Recipient Emails")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("Bulk Email Entry")
        bulk_emails = st.text_area(
            "Paste multiple emails (one per line or comma-separated):",
            height=150,
            placeholder="hr@company1.com\njobs@company2.com, careers@company3.com"
        )
        
        if st.button("Add Bulk Emails"):
            if bulk_emails:
                parsed = parse_email_list(bulk_emails)
                added = 0
                for email in parsed:
                    if email not in [e['email'] for e in st.session_state.email_list]:
                        st.session_state.email_list.append({
                            'email': email,
                            'company': '',
                            'description': 'Manual Entry'
                        })
                        added += 1
                st.success(f"Added {added} new emails!")
                st.rerun()
        
        st.markdown("---")
        st.subheader("Single Email with Details")
        
        col_a, col_b = st.columns(2)
        with col_a:
            single_email = st.text_input("Email Address")
            company = st.text_input("Company Name (Optional)")
        with col_b:
            job_title = st.text_input("Job Title (Optional)")
            job_desc = st.text_area("Job Description (Optional)", height=100)
        
        if st.button("Add Email with Details"):
            if single_email and validate_email(single_email):
                desc = "Manual Entry"
                if company:
                    desc += f" - {company}"
                if job_title:
                    desc += f" ({job_title})"
                if job_desc:
                    desc += f"\n{job_desc}"
                
                st.session_state.email_list.append({
                    'email': single_email,
                    'company': company,
                    'description': desc
                })
                st.success(f"Added {single_email}")
                st.rerun()
            else:
                st.error("Invalid email address")
    
    with col2:
        st.subheader("Email List Preview")
        st.metric("Total Recipients", len(st.session_state.email_list))
        
        if st.session_state.email_list:
            if st.button("Clear All Emails"):
                st.session_state.email_list = []
                st.rerun()
    
    if st.session_state.email_list:
        st.markdown("---")
        st.subheader(f"Current Email List ({len(st.session_state.email_list)} emails)")
        
        for i, email_data in enumerate(st.session_state.email_list):
            col_x, col_y = st.columns([6, 1])
            with col_x:
                st.text(f"{email_data['email']} - {email_data.get('company', 'N/A')}")
            with col_y:
                if st.button("Delete", key=f"del_email_{i}"):
                    st.session_state.email_list.pop(i)
                    st.rerun()

# TAB 3: Documents
with tab3:
    st.header("Upload Documents")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Resume (Required)")
        st.info("Supported formats: TXT, PDF, DOCX")
        
        resume_file = st.file_uploader(
            "Upload your resume",
            type=['txt', 'pdf', 'docx'],
            key="resume_upload",
            help="Upload your resume in TXT, PDF, or DOCX format"
        )
        
        if resume_file:
            try:
                original_filename = resume_file.name
                file_size_mb = len(resume_file.getvalue()) / (1024 * 1024)
                st.info(f"File: {original_filename} ({file_size_mb:.2f} MB)")
                
                suffix = Path(original_filename).suffix
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, mode='wb') as tmp:
                    tmp.write(resume_file.getvalue())
                    tmp.flush()
                    tmp_path = tmp.name
                
                if not os.path.exists(tmp_path):
                    st.error("Failed to save uploaded file")
                    st.stop()
                
                with st.spinner("Reading resume..."):
                    resume_text = read_document(tmp_path)
                
                if resume_text and len(resume_text.strip()) > 50:
                    st.session_state.resume_text = resume_text
                    st.session_state.resume_path = tmp_path
                    st.session_state.resume_original_name = original_filename
                    st.success(f"Resume uploaded successfully! ({len(resume_text)} characters)")
                    
                    with st.expander("Preview Resume Content", expanded=False):
                        preview_text = resume_text[:800] + ("..." if len(resume_text) > 800 else "")
                        st.text_area("Resume Preview", preview_text, height=250, disabled=True)
                        st.caption(f"Total length: {len(resume_text)} characters")
                
                elif resume_text:
                    st.error("Resume appears too short. Please upload a complete resume.")
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)
                else:
                    st.error("Failed to extract text from resume. Please check the file format.")
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)
                    
            except Exception as e:
                st.error(f"Error processing resume: {str(e)}")
                st.info("Tips:")
                st.markdown("""
                - **For PDFs**: Ensure it's a text-based PDF (not a scanned image)
                - **For scanned PDFs**: Convert to text using an OCR tool first
                - **Alternative**: Save your resume as a TXT file for best compatibility
                """)
    
    with col2:
        st.subheader("Additional Documents")
        st.markdown("*Upload cover letter, certifications, portfolio, etc.*")
        
        additional_files = st.file_uploader(
            "Upload additional documents (optional)",
            type=['pdf', 'docx', 'txt'],
            accept_multiple_files=True,
            key="additional_docs_upload"
        )
        
        if additional_files:
            st.session_state.additional_docs = []
            success_count = 0
            
            for file in additional_files:
                try:
                    original_name = file.name
                    suffix = Path(original_name).suffix
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, mode='wb') as tmp:
                        tmp.write(file.getvalue())
                        tmp.flush()
                        st.session_state.additional_docs.append({
                            'name': original_name,
                            'path': tmp.name
                        })
                        success_count += 1
                except Exception as e:
                    st.error(f"Failed to upload {file.name}: {str(e)}")
            
            if success_count > 0:
                st.success(f"Uploaded {success_count} additional document(s)")
                for doc in st.session_state.additional_docs:
                    st.text(f"Attached: {doc['name']}")
    
    if st.session_state.resume_text:
        st.markdown("---")
        st.success("Resume is loaded and ready!")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.metric("Characters", len(st.session_state.resume_text))
        with col_b:
            st.metric("Words", len(st.session_state.resume_text.split()))
        with col_c:
            st.metric("Additional Docs", len(st.session_state.additional_docs))
    
    st.markdown("---")
    st.subheader("Custom Instructions for Email Generation")
    custom_instructions = st.text_area(
        "Provide specific instructions for tailoring your emails:",
        placeholder="e.g., 'Emphasize my 5 years of hotel management experience'\n'Mention my MBA degree prominently'\n'Highlight my Python and AI skills'",
        height=150,
        value=st.session_state.custom_instructions,
        key="custom_instructions_input"
    )
    
    if st.button("Save Instructions"):
        st.session_state.custom_instructions = custom_instructions
        st.success("Instructions saved!")

# TAB 4: Generate Emails
with tab4:
    st.header("Generate AI-Powered Emails")
    
    if not st.session_state.resume_text:
        st.warning("Please upload your resume in the Documents tab first")
    elif not st.session_state.email_list:
        st.warning("Please add recipient emails in the Email Management tab")
    else:
        st.info(f"Ready to generate emails for {len(st.session_state.email_list)} recipients")
        
        email_automation = EmailAutomation(groq_key, st.session_state.sender_profile)
        
        if st.button("Generate All Emails", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            total = len(st.session_state.email_list)
            custom_inst = st.session_state.custom_instructions
            
            try:
                for i, email_data in enumerate(st.session_state.email_list):
                    status_text.text(f"Generating {i+1}/{total} for {email_data['email']}...")
                    progress_bar.progress((i + 1) / total)
                    
                    email_content = email_automation.generate_email(
                        resume_text=st.session_state.resume_text,
                        job_desc=email_data.get('description', ''),
                        recipient_info=email_data,
                        custom_inst=custom_inst
                    )
                    
                    st.session_state.generated_emails[email_data['email']] = email_content
                    st.session_state.email_mode = "ai"
                    time.sleep(2)
                
                status_text.empty()
                progress_bar.empty()
                st.success(f"Generated {total} emails successfully!")
                
            except Exception as e:
                st.error(f"Error: {str(e)}")
                if "Invalid Groq API key" in str(e):
                    st.error("Please verify your Groq API key is correct. Get one from https://console.groq.com/keys")
                status_text.empty()
                progress_bar.empty()
        
        if st.session_state.generated_emails:
            st.markdown("---")
            st.subheader("Review & Edit Generated Emails")
            
            for email, content in st.session_state.generated_emails.items():
                with st.expander(f"Email: {email}", expanded=False):
                    edited = st.text_area(
                        "Email Content (editable):",
                        value=content,
                        height=350,
                        key=f"edit_{email}"
                    )
                    st.session_state.generated_emails[email] = edited
                    subject, body = parse_email_content(edited)
                    st.markdown(f"**Subject:** `{subject}`")

# =============================================================================
# TAB MANUAL: Write Once, Send to All — No Groq needed
# =============================================================================
with tab_manual:
    st.header("✍️ Manual Email Template")

    st.info(
        "**Write your email once — it is sent to every recipient as-is, "
        "only the To: address changes.** Zero Groq API calls, "
        "instant for any number of recipients."
    )

    if not st.session_state.email_list:
        st.warning("Add recipient emails in the **Email Management** tab first.")

    st.markdown("---")

    # ── Subject line ──────────────────────────────────────────────────────────
    manual_subject = st.text_input(
        "Email Subject *",
        value=st.session_state.manual_subject,
        placeholder="e.g., Application for Python Developer Position – John Smith",
        key="manual_subject_input"
    )

    st.markdown("---")

    # ── Body ──────────────────────────────────────────────────────────────────
    st.markdown(
        "**Email Body** — use `{name}` anywhere to insert the recipient's "
        "company / name automatically."
    )

    manual_body = st.text_area(
        "Email Body *",
        value=st.session_state.manual_body,
        height=380,
        placeholder=(
            "Dear Hiring Manager,\n\n"
            "I am writing to express my interest in opportunities at {name}.\n\n"
            "[Your pitch here — 2-3 paragraphs]\n\n"
            "I have attached my resume for your consideration and would welcome "
            "the opportunity to discuss how my background aligns with your needs.\n\n"
            "Best regards,\n"
            "{your_name}\n"
            "{your_phone} | {your_linkedin}"
        ),
        key="manual_body_input"
    )

    # ── Placeholder reference ─────────────────────────────────────────────────
    with st.expander("📌 Available placeholders", expanded=False):
        st.markdown("""
| Placeholder | Replaced with |
|---|---|
| `{name}` | Recipient's company / name (from Email Management) |
| `{your_name}` | Your full name (from sidebar profile) |
| `{your_role}` | Your current role (from sidebar profile) |
| `{your_email}` | Your email address |
| `{your_phone}` | Your phone number |
| `{your_linkedin}` | Your LinkedIn URL |
| `{your_github}` | Your GitHub URL |
| `{your_portfolio}` | Your portfolio URL |
        """)

    st.markdown("---")

    # ── Live preview ──────────────────────────────────────────────────────────
    profile = st.session_state.sender_profile
    placeholder_map = {
        "{your_name}":      profile.get("name", ""),
        "{your_role}":      profile.get("current_role", ""),
        "{your_email}":     profile.get("email", ""),
        "{your_phone}":     profile.get("phone", ""),
        "{your_linkedin}":  profile.get("linkedin", ""),
        "{your_github}":    profile.get("github", ""),
        "{your_portfolio}": profile.get("portfolio", ""),
        "{name}":           "(Recipient Name / Company)",
    }

    preview_subject = manual_subject
    preview_body    = manual_body
    for ph, val in placeholder_map.items():
        preview_subject = preview_subject.replace(ph, val)
        preview_body    = preview_body.replace(ph, val)

    with st.expander("👁️ Live Preview", expanded=False):
        st.markdown(f"**Subject:** `{preview_subject or '(no subject yet)'}`")
        st.text_area(
            "Body preview:",
            value=preview_body or "(start typing your email body above)",
            height=300,
            disabled=True,
            key="manual_preview_area"
        )

    st.markdown("---")

    # ── Action buttons ────────────────────────────────────────────────────────
    col_save, col_apply = st.columns(2)

    with col_save:
        if st.button("💾 Save Template", use_container_width=True):
            if not manual_subject.strip():
                st.error("Please enter a subject line.")
            elif not manual_body.strip():
                st.error("Please enter the email body.")
            else:
                st.session_state.manual_subject = manual_subject
                st.session_state.manual_body    = manual_body
                st.session_state.email_mode     = "manual"
                st.success("Template saved! Click **Apply** when ready, then go to **Send Emails**.")

    with col_apply:
        apply_label = (
            f"🚀 Apply to All {len(st.session_state.email_list)} Recipients"
            if st.session_state.email_list else "🚀 Apply Template to All Recipients"
        )
        if st.button(apply_label, type="primary", use_container_width=True):
            if not manual_subject.strip():
                st.error("Please enter a subject line.")
            elif not manual_body.strip():
                st.error("Please enter the email body.")
            elif not st.session_state.email_list:
                st.error("No recipients found. Add emails in the Email Management tab first.")
            else:
                # Warn if overwriting AI-generated emails
                if st.session_state.email_mode == "ai" and st.session_state.generated_emails:
                    st.warning(
                        "This will replace your previously AI-generated emails "
                        "with the manual template."
                    )

                st.session_state.manual_subject = manual_subject
                st.session_state.manual_body    = manual_body
                st.session_state.email_mode     = "manual"

                count = 0
                for email_data in st.session_state.email_list:
                    recipient_name = (
                        email_data.get("company") or
                        email_data.get("name") or
                        "Hiring Manager"
                    )
                    body = manual_body
                    subj = manual_subject

                    # Substitute all placeholders per recipient
                    per_recipient_map = {
                        "{name}":           recipient_name,
                        "{your_name}":      profile.get("name", ""),
                        "{your_role}":      profile.get("current_role", ""),
                        "{your_email}":     profile.get("email", ""),
                        "{your_phone}":     profile.get("phone", ""),
                        "{your_linkedin}":  profile.get("linkedin", ""),
                        "{your_github}":    profile.get("github", ""),
                        "{your_portfolio}": profile.get("portfolio", ""),
                    }
                    for ph, val in per_recipient_map.items():
                        body = body.replace(ph, val)
                        subj = subj.replace(ph, val)

                    # Store in same format as AI-generated so Send tab works unchanged
                    st.session_state.generated_emails[email_data["email"]] = (
                        f"Subject: {subj}\n\n{body}"
                    )
                    count += 1

                st.success(
                    f"Template applied to **{count} recipients**! "
                    f"Go to the **Send Emails** tab to send them all."
                )
                st.balloons()

    # ── Status strip ─────────────────────────────────────────────────────────
    if st.session_state.email_mode == "manual" and st.session_state.manual_subject:
        n = len(st.session_state.generated_emails)
        if n:
            st.success(
                f"**Manual template active** — {n} emails queued in Send Emails tab."
            )
        else:
            st.info("Template saved. Click **Apply** above to queue emails for sending.")

    if st.session_state.email_mode == "ai" and st.session_state.generated_emails:
        st.warning(
            f"{len(st.session_state.generated_emails)} AI-generated emails are currently "
            "queued. Applying this template will replace them."
        )

# TAB 5: Send Emails
with tab5:
    st.header("Send Emails")

    if not st.session_state.generated_emails:
        st.warning(
            "No emails queued yet. Either:\n"
            "- Generate AI emails in the **Generate Emails** tab, or\n"
            "- Write and apply a template in the **✍️ Manual Template** tab."
        )
    else:
        mode_label = "Manual Template" if st.session_state.email_mode == "manual" else "AI Generated"
        st.success(
            f"{len(st.session_state.generated_emails)} emails ready to send  |  Mode: **{mode_label}**"
        )

        # Attachments
        attachments_with_names = []
        if st.session_state.resume_path and st.session_state.resume_original_name:
            attachments_with_names.append((
                st.session_state.resume_path,
                st.session_state.resume_original_name
            ))
        if st.session_state.additional_docs:
            for doc in st.session_state.additional_docs:
                attachments_with_names.append((doc['path'], doc['name']))

        if attachments_with_names:
            st.info(f"Will attach {len(attachments_with_names)} document(s) with original filenames:")
            for _, name in attachments_with_names:
                st.text(f"  • {name}")

        if not st.session_state.sender_profile.get("email_app_key"):
            st.error("Email App Key is missing! Please add it in the sidebar.")
            st.info("Get Gmail App Password: https://support.google.com/accounts/answer/185833")
            st.stop()

        smtp_config = {
            "host": "smtp.gmail.com",
            "port": 587,
            "username": st.session_state.sender_profile.get("email", ""),
            "password": st.session_state.sender_profile.get("email_app_key", ""),
            "from_email": st.session_state.sender_profile.get("email", ""),
            "from_name": st.session_state.sender_profile.get("name", "")
        }

        if st.button("Send All Emails", type="primary"):
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

                    result = EmailAutomation.send_mail(
                        to_email=email,
                        subject=subject,
                        body=body,
                        smtp_config=smtp_config,
                        attachments_with_names=attachments_with_names
                    )

                    all_results[email] = result
                    time.sleep(1)

                status.empty()
                progress.empty()

                success_count = sum(all_results.values())
                st.success(f"Successfully sent {success_count}/{len(all_results)} emails!")

                col_success, col_failed = st.columns(2)
                with col_success:
                    st.subheader("Successful")
                    for email, success in all_results.items():
                        if success:
                            st.success(f"{email}")
                with col_failed:
                    st.subheader("Failed")
                    for email, success in all_results.items():
                        if not success:
                            st.error(f"{email}")

                if st.button("Start New Campaign"):
                    for key in ['email_list', 'resume_text', 'resume_path', 'resume_original_name',
                               'generated_emails', 'additional_docs', 'custom_instructions',
                               'job_search_results']:
                        if key in st.session_state:
                            if key == 'custom_instructions':
                                st.session_state[key] = ""
                            elif key in ['email_list', 'additional_docs']:
                                st.session_state[key] = []
                            else:
                                st.session_state[key] = {}
                    st.session_state.manual_subject = ""
                    st.session_state.manual_body    = ""
                    st.session_state.email_mode     = "ai"
                    st.session_state.current_step   = 1
                    st.rerun()

            except Exception as e:
                st.error(f"Error: {str(e)}")
                status.empty()
                progress.empty()


# Footer
st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: var(--text-secondary); padding: 2rem;'>
        <h3 style='color: var(--text-primary);'></h3>
        <p></p>
        <p style='font-size: 0.9rem; opacity: 0.8;'>
            Features: Google Jobs Search • Freelance Platform Search • 
            Multi-Document Attachments • AI-Powered Email Generation • Manual Template • Bulk Email Management
        </p>
        <p style='font-size: 0.8rem; opacity: 0.6;'>
            Dark Theme Edition: Original Filenames Preserved • Theme-Independent Styling • Professional Interface
        </p>
    </div>
""", unsafe_allow_html=True)