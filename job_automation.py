"""
Complete Job Application Automation System - Production Ready (PROFESSIONAL VERSION)
All Features: Google Jobs, Freelancing, Profile Management, Multiple Attachments, Excel Processing
Fixed: PyPDF2 → pypdf, File naming preservation, Professional styling, Emoji removal
Added: Anti-spam delayed sending (30s+ between emails) for 500+ mail campaigns
"""

import streamlit as st
import pandas as pd
import json
from pathlib import Path
import tempfile
import os
import time
import random
from datetime import datetime, timedelta
import re
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import List, Dict, Tuple

# ── PDF library: try pypdf first (modern), fall back to PyPDF2 (legacy) ──
try:
    from pypdf import PdfReader as PdfReaderLib
    PDF_LIB = "pypdf"
except ImportError:
    try:
        from PyPDF2 import PdfReader as PdfReaderLib
        PDF_LIB = "PyPDF2"
    except ImportError:
        PdfReaderLib = None
        PDF_LIB = None

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
            if PdfReaderLib is None:
                raise Exception(
                    "No PDF library found. Run:  pip install pypdf\n"
                    "(pypdf is the modern replacement for the deprecated PyPDF2)"
                )
            
            try:
                text = ""
                with open(file_path, 'rb') as f:
                    pdf_reader = PdfReaderLib(f)
                    
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
                    raise Exception(
                        "Could not extract any text from PDF. "
                        "The PDF might be scanned/image-based. "
                        "Try converting it to DOCX or TXT first."
                    )
                    
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

        if len(prompt) > 4000:
            raise Exception(f"Prompt too long ({len(prompt)} chars). Please shorten your resume or job description.")

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

        # ── Retry with exponential backoff for 429 rate-limit errors ──
        MAX_RETRIES = 6
        base_wait   = 15  # seconds — Groq free tier resets quickly

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                response = requests.post(url, json=payload, headers=headers, timeout=30)

                if response.status_code == 401:
                    raise Exception("Invalid Groq API key. Please check your API key and try again.")

                if response.status_code == 400:
                    raise Exception(
                        f"Bad request to Groq API. Prompt may be too long or contain invalid characters. "
                        f"Details: {response.text[:200]}"
                    )

                if response.status_code == 429:
                    # Parse retry-after header if present, otherwise use exponential backoff
                    retry_after = response.headers.get("retry-after") or response.headers.get("x-ratelimit-reset-requests")
                    if retry_after:
                        try:
                            wait_secs = int(float(retry_after))
                        except (ValueError, TypeError):
                            wait_secs = base_wait * (2 ** (attempt - 1))
                    else:
                        wait_secs = base_wait * (2 ** (attempt - 1))  # 15, 30, 60, 120, 240, 480

                    wait_secs = min(wait_secs, 480)  # cap at 8 minutes

                    if attempt < MAX_RETRIES:
                        # Show live countdown in the UI
                        rate_slot = st.empty()
                        for s in range(int(wait_secs), 0, -1):
                            mm, ss = divmod(s, 60)
                            rate_slot.warning(
                                f"⏳ Groq rate limit hit (attempt {attempt}/{MAX_RETRIES}). "
                                f"Retrying in **{mm:02d}:{ss:02d}** — "
                                f"this is normal for free-tier accounts with 100+ emails."
                            )
                            time.sleep(1)
                        rate_slot.empty()
                        continue  # retry
                    else:
                        raise Exception(
                            f"Groq API rate limit exceeded after {MAX_RETRIES} retries. "
                            f"Tip: upgrade to Groq paid tier, or reduce batch size and try again later."
                        )

                response.raise_for_status()

                data = response.json()
                email_content = data["choices"][0]["message"]["content"].strip()
                signature = self._build_signature()
                return email_content + "\n" + signature

            except requests.exceptions.RequestException as e:
                err_str = str(e)
                if "401" in err_str:
                    raise Exception("Invalid Groq API key. Please verify your API key is correct.")
                elif "429" in err_str:
                    # Handled above via status code check; shouldn't reach here, but guard anyway
                    if attempt < MAX_RETRIES:
                        wait_secs = base_wait * (2 ** (attempt - 1))
                        time.sleep(min(wait_secs, 480))
                        continue
                    raise Exception("Groq API rate limit exceeded. Please wait a few minutes and try again.")
                elif "400" in err_str:
                    raise Exception(f"Bad request — try shortening your resume or job description. Error: {err_str[:150]}")
                else:
                    raise Exception(f"API request failed: {err_str}")
            except Exception as e:
                if "rate limit" in str(e).lower() or "429" in str(e):
                    raise  # already a clean message
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

        # Attach files with original names
        if attachments_with_names:
            for file_path, original_name in attachments_with_names:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    # Use original filename instead of temp filename
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
            return True, None          # (success, error_message)
        except Exception as e:
            return False, str(e)       # (success, error_message)


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
# STREAMLIT UI - PROFESSIONAL STYLING
# =============================================================================

st.set_page_config(
    page_title="AI Job Application System",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Professional Dark Theme CSS - Theme-independent
st.markdown("""
    <style>
    :root {
        --primary-color: #3b82f6;
        --secondary-color: #2563eb;
        --background-color: #0f172a;
        --card-background: #1e293b;
        --text-primary: #f1f5f9;
        --text-secondary: #94a3b8;
        --border-color: #334155;
        --success-color: #10b981;
        --warning-color: #f59e0b;
        --error-color: #ef4444;
        --hover-bg: #293548;
    }
    .main { background-color: var(--background-color) !important; color: var(--text-primary) !important; }
    [data-testid="stSidebar"] { background-color: #1a2332 !important; border-right: 1px solid var(--border-color); }
    [data-testid="stSidebar"] * { color: var(--text-primary) !important; }
    h1, h2, h3, h4, h5, h6 { color: var(--text-primary) !important; font-weight: 600 !important; }
    h1 { font-size: 2.5rem !important; margin-bottom: 0.5rem !important; }
    .stButton>button {
        width: 100%; background-color: var(--primary-color) !important;
        color: white !important; font-weight: 500 !important; border: none !important;
        padding: 0.75rem 1.5rem !important; border-radius: 0.5rem !important;
        transition: all 0.2s ease !important; box-shadow: 0 2px 4px rgba(0,0,0,0.3) !important;
    }
    .stButton>button:hover {
        background-color: var(--secondary-color) !important;
        box-shadow: 0 4px 8px rgba(59, 130, 246, 0.4) !important; transform: translateY(-1px);
    }
    .stTextInput>div>div>input, .stTextArea>div>div>textarea, .stSelectbox>div>div>select {
        background-color: var(--card-background) !important; border: 1px solid var(--border-color) !important;
        border-radius: 0.5rem !important; color: var(--text-primary) !important; padding: 0.5rem !important;
    }
    .stTextInput>div>div>input:focus, .stTextArea>div>div>textarea:focus {
        border-color: var(--primary-color) !important;
        box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.2) !important; outline: none !important;
    }
    .stTextInput label, .stTextArea label, .stSelectbox label { color: var(--text-secondary) !important; }
    .info-box {
        padding: 1.5rem; background-color: var(--card-background); border-radius: 0.75rem;
        border: 1px solid var(--border-color); margin: 1rem 0; box-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }
    .streamlit-expanderHeader {
        background-color: var(--card-background) !important; border: 1px solid var(--border-color) !important;
        border-radius: 0.5rem !important; color: var(--text-primary) !important; font-weight: 500 !important;
    }
    .streamlit-expanderHeader:hover { background-color: var(--hover-bg) !important; }
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem; background-color: var(--card-background); border-radius: 0.5rem;
        padding: 0.5rem; border: 1px solid var(--border-color);
    }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent; border-radius: 0.375rem;
        color: var(--text-secondary); font-weight: 500; padding: 0.5rem 1rem;
    }
    .stTabs [data-baseweb="tab"]:hover { background-color: var(--hover-bg); }
    .stTabs [aria-selected="true"] { background-color: var(--primary-color) !important; color: white !important; }
    .stSuccess { background-color: rgba(16,185,129,0.15) !important; color: #6ee7b7 !important; border-left: 4px solid var(--success-color) !important; border-radius: 0.5rem !important; }
    .stWarning { background-color: rgba(245,158,11,0.15) !important; color: #fbbf24 !important; border-left: 4px solid var(--warning-color) !important; border-radius: 0.5rem !important; }
    .stError { background-color: rgba(239,68,68,0.15) !important; color: #fca5a5 !important; border-left: 4px solid var(--error-color) !important; border-radius: 0.5rem !important; }
    .stInfo { background-color: rgba(59,130,246,0.15) !important; color: #93c5fd !important; border-left: 4px solid var(--primary-color) !important; border-radius: 0.5rem !important; }
    .stDataFrame { border: 1px solid var(--border-color) !important; border-radius: 0.5rem !important; background-color: var(--card-background) !important; }
    [data-testid="stMetricValue"] { color: var(--primary-color) !important; font-size: 2rem !important; font-weight: 600 !important; }
    [data-testid="stMetricLabel"] { color: var(--text-secondary) !important; }
    .uploadedFile { background-color: var(--card-background) !important; border: 1px solid var(--border-color) !important; border-radius: 0.5rem !important; }
    [data-testid="stFileUploader"] { background-color: var(--card-background) !important; border: 1px dashed var(--border-color) !important; border-radius: 0.5rem !important; }
    .stProgress > div > div { background-color: var(--primary-color) !important; }
    hr { border-color: var(--border-color) !important; opacity: 0.3 !important; }
    header[data-testid="stHeader"] { background-color: #1a2332 !important; }
    .stMarkdown { color: var(--text-primary) !important; }
    .caption { color: var(--text-secondary) !important; }
    code { background-color: var(--card-background) !important; color: #93c5fd !important; border: 1px solid var(--border-color) !important; border-radius: 0.25rem !important; padding: 0.2rem 0.4rem !important; }
    ::placeholder { color: var(--text-secondary) !important; opacity: 0.6 !important; }

    /* ── Anti-spam countdown & log styles ── */
    .countdown-box {
        background: linear-gradient(135deg, #1e293b, #0f172a);
        border: 1px solid #3b82f6; border-radius: 12px;
        padding: 1.25rem 2rem; text-align: center; margin: 0.75rem 0;
        box-shadow: 0 0 20px rgba(59,130,246,0.15);
    }
    .countdown-number { font-size: 3rem; font-weight: 700; color: #3b82f6; line-height: 1; letter-spacing: 2px; }
    .countdown-label { font-size: 0.85rem; color: #94a3b8; margin-top: 0.3rem; }
    .log-success { background: rgba(16,185,129,0.08); border-left: 3px solid #10b981; padding: 0.4rem 0.9rem; border-radius: 0 6px 6px 0; margin: 0.2rem 0; font-size: 0.88rem; color: #6ee7b7; }
    .log-fail { background: rgba(239,68,68,0.08); border-left: 3px solid #ef4444; padding: 0.4rem 0.9rem; border-radius: 0 6px 6px 0; margin: 0.2rem 0; font-size: 0.88rem; color: #fca5a5; }
    .log-wait { background: rgba(245,158,11,0.08); border-left: 3px solid #f59e0b; padding: 0.4rem 0.9rem; border-radius: 0 6px 6px 0; margin: 0.2rem 0; font-size: 0.88rem; color: #fbbf24; }
    .batch-header { background: linear-gradient(90deg, #1e40af22, transparent); border-left: 4px solid #3b82f6; border-radius: 0 8px 8px 0; padding: 0.6rem 1rem; color: #93c5fd; font-weight: 600; font-size: 0.95rem; margin: 1rem 0 0.4rem 0; }
    </style>
""", unsafe_allow_html=True)

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
if 'send_results' not in st.session_state:
    st.session_state.send_results = {}
if 'manual_subject' not in st.session_state:
    st.session_state.manual_subject = ""
if 'manual_body' not in st.session_state:
    st.session_state.manual_body = ""
if 'email_mode' not in st.session_state:
    st.session_state.email_mode = "ai"

# Header
st.markdown("""
    <h1 style='text-align: center; margin-bottom: 0;'>
        AI Job Application System
    </h1>
    <p style='text-align: center; color: #94a3b8; font-size: 1.1rem; margin-top: 0.5rem;'>
        Automated Job Search • AI Email Generation • Anti-Spam Smart Sending
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
    if st.session_state.email_mode == "manual" and st.session_state.manual_subject:
        st.info("Mode: Manual Template")
    elif st.session_state.generated_emails:
        st.info(f"{len(st.session_state.generated_emails)} AI emails ready")
    if st.session_state.send_results:
        sent_ok = sum(1 for v in st.session_state.send_results.values() if v)
        st.info(f"Sent: {sent_ok}/{len(st.session_state.send_results)}")

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

    # ── Show install hint if no PDF library available ──
    if PdfReaderLib is None:
        st.error(
            "**PDF library not installed.** Run this command in your terminal, then restart the app:\n\n"
            "```\npip install pypdf\n```"
        )
    else:
        st.success(f"PDF library ready: `{PDF_LIB}` ✓")
    
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
                - **Install pypdf**: Run `pip install pypdf` then restart the app
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
        total_recip = len(st.session_state.email_list)
        st.info(f"Ready to generate emails for {total_recip} recipients")

        # ── Rate-limit guidance ──
        if total_recip > 30:
            st.warning(
                f"**Large batch ({total_recip} emails detected).**  "
                f"Groq free tier allows ~30 requests/minute. "
                f"A built-in delay is added between calls and the app will auto-pause & retry if rate-limited. "
                f"Estimated generation time: **~{max(1, total_recip // 25)} min**. Leave this tab open."
            )

        # Configurable generation delay
        gen_delay = st.slider(
            "Delay between API calls (seconds)",
            min_value=2, max_value=30, value=5, step=1,
            help="Higher = fewer rate-limit hits. 5s works for most accounts. "
                 "Increase to 10-15s if you keep hitting 429 errors."
        )

        email_automation = EmailAutomation(groq_key, st.session_state.sender_profile)
        
        if st.button("Generate All Emails", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            gen_stats   = st.empty()
            
            total      = len(st.session_state.email_list)
            custom_inst = st.session_state.custom_instructions
            succeeded  = 0
            failed_gen = 0
            
            try:
                for i, email_data in enumerate(st.session_state.email_list):
                    status_text.markdown(
                        f"⚙️ Generating **{i+1}/{total}** for `{email_data['email']}`..."
                    )
                    progress_bar.progress((i + 1) / total)
                    
                    try:
                        email_content = email_automation.generate_email(
                            resume_text=st.session_state.resume_text,
                            job_desc=email_data.get('description', ''),
                            recipient_info=email_data,
                            custom_inst=custom_inst
                        )
                        st.session_state.generated_emails[email_data['email']] = email_content
                        succeeded += 1
                    except Exception as per_email_err:
                        failed_gen += 1
                        st.warning(f"Skipped `{email_data['email']}`: {str(per_email_err)[:120]}")

                    gen_stats.info(f"✅ Generated: {succeeded}  |  ⚠️ Skipped: {failed_gen}  |  Remaining: {total - i - 1}")

                    # Polite delay between Groq API calls to avoid rate limits
                    if i < total - 1:
                        time.sleep(gen_delay)
                
                status_text.empty()
                progress_bar.empty()
                gen_stats.empty()

                if succeeded > 0:
                    st.success(f"Generated {succeeded}/{total} emails successfully!")
                if failed_gen > 0:
                    st.warning(
                        f"{failed_gen} email(s) were skipped due to rate limits or errors. "
                        f"You can re-run generation — already-generated emails are preserved."
                    )
                
            except Exception as e:
                st.error(f"Error: {str(e)}")
                if "Invalid Groq API key" in str(e):
                    st.error("Please verify your Groq API key is correct. Get one from https://console.groq.com/keys")
                status_text.empty()
                progress_bar.empty()
                gen_stats.empty()
        
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

# =============================================================================
# TAB MANUAL: Write Once, Send to All — No Groq needed
# =============================================================================
with tab_manual:
    st.header("✍️ Manual Email Template")
    st.info(
        "**Write your email once — it is sent to every recipient as-is, "
        "only the To address changes.** No Groq API calls needed. "
        "Perfect for 500+ emails where you want full control over the content."
    )

    if not st.session_state.email_list:
        st.warning("Add recipient emails in the **Email Management** tab first.")

    st.markdown("---")

    # ── Subject ──
    manual_subject = st.text_input(
        "Email Subject *",
        value=st.session_state.manual_subject,
        placeholder="e.g., Application for Python Developer Position – Konika Sharma",
        key="manual_subject_input"
    )

    st.markdown("---")

    # ── Body ──
    st.markdown("**Email Body** — use `{name}` as an optional placeholder for the recipient's name "
                "(pulled from company field if available, otherwise left blank)")

    manual_body = st.text_area(
        "Email Body *",
        value=st.session_state.manual_body,
        height=350,
        placeholder=(
            "Dear Hiring Manager,\n\n"
            "I am writing to express my interest in the [Job Title] position at your company.\n\n"
            "[Your 2-3 paragraph pitch here]\n\n"
            "I have attached my resume for your consideration and would welcome the opportunity "
            "to discuss how my background aligns with your needs.\n\n"
            "Best regards,\n"
            "[Your Name]\n"
            "[Phone] | [LinkedIn]"
        ),
        key="manual_body_input"
    )

    # ── Variable reference ──
    with st.expander("📌 Available placeholders", expanded=False):
        st.markdown("""
        | Placeholder | Replaced with |
        |---|---|
        | `{name}` | Recipient company/name (from Email Management) |
        | `{your_name}` | Your full name (from sidebar profile) |
        | `{your_role}` | Your current role (from sidebar profile) |
        | `{your_email}` | Your email (from sidebar profile) |
        | `{your_phone}` | Your phone (from sidebar profile) |
        | `{your_linkedin}` | Your LinkedIn URL |
        """)

    st.markdown("---")

    # ── Preview ──
    preview_subject = manual_subject
    preview_body    = manual_body

    # Substitute preview placeholders with profile values
    profile = st.session_state.sender_profile
    for ph, val in [
        ("{your_name}",     profile.get("name", "")),
        ("{your_role}",     profile.get("current_role", "")),
        ("{your_email}",    profile.get("email", "")),
        ("{your_phone}",    profile.get("phone", "")),
        ("{your_linkedin}", profile.get("linkedin", "")),
        ("{name}",          "(Recipient Name)"),
    ]:
        preview_body    = preview_body.replace(ph, val)
        preview_subject = preview_subject.replace(ph, val)

    with st.expander("👁️ Preview (first recipient substitution)", expanded=False):
        st.markdown(f"**Subject:** `{preview_subject}`")
        st.text_area("Body preview:", value=preview_body, height=300, disabled=True, key="preview_area")

    st.markdown("---")

    # ── Action buttons ──
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
                st.success("Template saved! Go to **Send Emails** tab to send.")

    with col_apply:
        if st.button("🚀 Apply Template to All Recipients", type="primary", use_container_width=True):
            if not manual_subject.strip():
                st.error("Please enter a subject line.")
            elif not manual_body.strip():
                st.error("Please enter the email body.")
            elif not st.session_state.email_list:
                st.error("No recipients found. Add emails in the Email Management tab.")
            else:
                # Save template
                st.session_state.manual_subject = manual_subject
                st.session_state.manual_body    = manual_body
                st.session_state.email_mode     = "manual"

                # Build the final body for each recipient by substituting placeholders
                profile = st.session_state.sender_profile
                count = 0
                for email_data in st.session_state.email_list:
                    recipient_name = (
                        email_data.get("company") or
                        email_data.get("name") or
                        "Hiring Manager"
                    )
                    body = manual_body
                    subj = manual_subject
                    for ph, val in [
                        ("{name}",          recipient_name),
                        ("{your_name}",     profile.get("name", "")),
                        ("{your_role}",     profile.get("current_role", "")),
                        ("{your_email}",    profile.get("email", "")),
                        ("{your_phone}",    profile.get("phone", "")),
                        ("{your_linkedin}", profile.get("linkedin", "")),
                    ]:
                        body = body.replace(ph, val)
                        subj = subj.replace(ph, val)

                    # Store in generated_emails dict using same format as AI mode
                    # Format: "Subject: <subj>\n\n<body>" — parse_email_content() handles this
                    st.session_state.generated_emails[email_data["email"]] = f"Subject: {subj}\n\n{body}"
                    count += 1

                st.success(
                    f"✅ Template applied to **{count} recipients**! "
                    f"Go to the **Send Emails** tab to send them all."
                )
                st.balloons()

    # ── Status ──
    if st.session_state.email_mode == "manual" and st.session_state.manual_subject:
        st.markdown("---")
        n_applied = len(st.session_state.generated_emails)
        if n_applied:
            st.success(f"**Template active** — ready to send to {n_applied} recipients via Send Emails tab.")
        else:
            st.info("Template saved but not yet applied. Click **Apply Template to All Recipients**.")

    if st.session_state.email_mode == "ai" and st.session_state.generated_emails:
        st.markdown("---")
        st.warning(
            "AI-generated emails are currently queued in the Send tab. "
            "Clicking **Apply Template** above will **replace** them with your manual template."
        )


# TAB 5: Send Emails — ANTI-SPAM DELAYED SENDING
# All original logic preserved. Added: delay controls, live countdown,
# per-email status log, batch cooldown, resume support, results export.
# =============================================================================
with tab5:
    st.header("Send Emails")

    if not st.session_state.generated_emails:
        st.warning("Please generate emails first in the 'Generate Emails' tab")
    else:
        total_generated = len(st.session_state.generated_emails)
        st.success(f"{total_generated} emails ready to send")

        # ── Attachments (original logic preserved) ──
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

        # ── SMTP check (original logic preserved) ──
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

        # ── Anti-spam delay settings ──
        st.markdown("---")
        st.markdown("### Anti-Spam Sending Settings")
        st.info(
            "**Why delay?** Sending 500+ emails instantly flags your account as spam. "
            "Spacing them with randomness mimics human behaviour and protects your domain reputation."
        )

        col_d1, col_d2, col_d3 = st.columns(3)
        with col_d1:
            delay_between = st.slider(
                "Delay between emails (seconds)",
                min_value=15, max_value=300, value=30, step=5,
                help="30s minimum recommended. Use 45-60s for 500+ emails."
            )
        with col_d2:
            jitter = st.slider(
                "Random jitter ± (seconds)",
                min_value=0, max_value=30, value=10, step=5,
                help="Adds randomness so emails don't fire at a perfectly regular interval."
            )
        with col_d3:
            batch_size = st.slider(
                "Batch size",
                min_value=5, max_value=50, value=20, step=5,
                help="After each batch a longer pause is taken."
            )

        batch_pause_min = st.slider(
            "Pause between batches (minutes)",
            min_value=1, max_value=20, value=5,
            help="Cool-down after each batch. Gmail limits ~500 emails/day."
        )

        # Estimated time
        num_batches = max(1, (total_generated + batch_size - 1) // batch_size)
        est_secs = (total_generated * delay_between) + (num_batches * batch_pause_min * 60)
        est_label = (
            f"{est_secs // 3600}h {(est_secs % 3600) // 60}m"
            if est_secs >= 3600 else f"{est_secs // 60}m {est_secs % 60}s"
        )

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Emails", total_generated)
        c2.metric("Batches", num_batches)
        c3.metric("Est. Duration", est_label)
        c4.metric("Delay/Email", f"{delay_between}s ±{jitter}s")

        st.markdown("---")

        # ── Resume support: detect already-sent emails ──
        already_sent_set = {k for k, v in st.session_state.send_results.items() if v}
        remaining_emails = [e for e in st.session_state.generated_emails.keys() if e not in already_sent_set]

        if already_sent_set:
            st.info(f"**Resume detected:** {len(already_sent_set)} already sent. {len(remaining_emails)} remaining.")
            col_rb1, col_rb2 = st.columns(2)
            with col_rb1:
                do_send = st.button("Resume — Send Remaining", type="primary")
            with col_rb2:
                if st.button("Restart from Beginning"):
                    st.session_state.send_results = {}
                    remaining_emails = list(st.session_state.generated_emails.keys())
                    st.rerun()
        else:
            do_send = st.button("Send All Emails", type="primary")

        # ── MAIN SEND LOOP ──
        if do_send:
            emails_to_process = remaining_emails if already_sent_set else list(st.session_state.generated_emails.keys())
            total_to_send = len(emails_to_process)

            progress = st.progress(0)
            status = st.empty()
            countdown_slot = st.empty()

            sc1, sc2, sc3, sc4 = st.columns(4)
            stat_sent   = sc1.empty()
            stat_failed = sc2.empty()
            stat_left   = sc3.empty()
            stat_eta    = sc4.empty()

            log_area = st.container()

            sent_count   = len(already_sent_set)
            failed_count = sum(1 for v in st.session_state.send_results.values() if not v)
            start_time   = datetime.now()

            try:
                for batch_start in range(0, total_to_send, batch_size):
                    batch         = emails_to_process[batch_start: batch_start + batch_size]
                    batch_num     = batch_start // batch_size + 1
                    total_batches = (total_to_send + batch_size - 1) // batch_size

                    log_area.markdown(
                        f"<div class='batch-header'>Batch {batch_num}/{total_batches} — "
                        f"Emails {batch_start + 1}–{min(batch_start + batch_size, total_to_send)}</div>",
                        unsafe_allow_html=True
                    )

                    for idx, email in enumerate(batch):
                        global_idx = batch_start + idx + 1
                        status.markdown(f"📤 **Sending** `{email}` — {global_idx}/{total_to_send}")
                        progress.progress(global_idx / total_to_send)

                        content = st.session_state.generated_emails[email]
                        subject, body = parse_email_content(content)

                        # ── Original send_mail call — now returns (bool, err_str) ──
                        result, err_msg = EmailAutomation.send_mail(
                            to_email=email,
                            subject=subject,
                            body=body,
                            smtp_config=smtp_config,
                            attachments_with_names=attachments_with_names
                        )

                        st.session_state.send_results[email] = result
                        ts = datetime.now().strftime("%H:%M:%S")

                        if result:
                            sent_count += 1
                            log_area.markdown(
                                f"<div class='log-success'>✅ [{ts}] <b>{email}</b> — Sent</div>",
                                unsafe_allow_html=True
                            )
                        else:
                            failed_count += 1
                            log_area.markdown(
                                f"<div class='log-fail'>❌ [{ts}] <b>{email}</b> — Failed: {err_msg}</div>",
                                unsafe_allow_html=True
                            )

                        stat_sent.metric("Sent ✅", sent_count)
                        stat_failed.metric("Failed ❌", failed_count)
                        stat_left.metric("Remaining", total_to_send - global_idx)
                        elapsed = (datetime.now() - start_time).total_seconds()
                        if global_idx > 0:
                            eta_s = (elapsed / global_idx) * (total_to_send - global_idx)
                            stat_eta.metric("ETA", str(timedelta(seconds=int(eta_s))))

                        # ── Countdown between emails (skip after last email overall) ──
                        is_last_email    = (global_idx == total_to_send)
                        is_last_in_batch = (idx == len(batch) - 1)

                        if not is_last_email:
                            if is_last_in_batch and batch_start + batch_size < total_to_send:
                                pause_secs = batch_pause_min * 60
                                wait_label = f"Batch {batch_num} done — cooling down for {batch_pause_min} min..."
                            else:
                                actual_jitter = random.uniform(-jitter, jitter)
                                pause_secs = max(5, delay_between + actual_jitter)
                                wait_label = f"Waiting {pause_secs:.0f}s before next email..."

                            log_area.markdown(
                                f"<div class='log-wait'>⏳ {wait_label}</div>",
                                unsafe_allow_html=True
                            )

                            for remaining_secs in range(int(pause_secs), 0, -1):
                                mm, ss = divmod(remaining_secs, 60)
                                countdown_slot.markdown(
                                    f"<div class='countdown-box'>"
                                    f"<div class='countdown-number'>{mm:02d}:{ss:02d}</div>"
                                    f"<div class='countdown-label'>{wait_label}</div>"
                                    f"</div>",
                                    unsafe_allow_html=True
                                )
                                time.sleep(1)

                            countdown_slot.empty()

            except Exception as e:
                st.error(f"Error during sending: {str(e)}")
            finally:
                status.empty()
                countdown_slot.empty()

            # ── Final summary ──
            st.markdown("---")
            final_sent   = sum(1 for v in st.session_state.send_results.values() if v)
            final_failed = sum(1 for v in st.session_state.send_results.values() if not v)

            if final_failed == 0:
                st.success(f"All {final_sent} emails sent successfully!")
            else:
                st.warning(f"Done: {final_sent} sent, {final_failed} failed")

            col_success, col_failed_col = st.columns(2)
            with col_success:
                st.subheader("Successful")
                for email, ok in st.session_state.send_results.items():
                    if ok:
                        st.success(f"{email}")
            with col_failed_col:
                st.subheader("Failed")
                for email, ok in st.session_state.send_results.items():
                    if not ok:
                        st.error(f"{email}")

            # Download results CSV
            result_df = pd.DataFrame([
                {"Email": e, "Status": "Sent" if ok else "Failed"}
                for e, ok in st.session_state.send_results.items()
            ])
            st.download_button(
                "Download Results CSV",
                result_df.to_csv(index=False).encode("utf-8"),
                "send_results.csv", "text/csv"
            )

            # Reset button (original logic preserved)
            if st.button("Start New Campaign"):
                for key in ['email_list', 'resume_text', 'resume_path', 'resume_original_name',
                           'generated_emails', 'additional_docs', 'custom_instructions',
                           'job_search_results', 'send_results']:
                    if key in st.session_state:
                        if key == 'custom_instructions':
                            st.session_state[key] = ""
                        elif key in ['email_list', 'additional_docs']:
                            st.session_state[key] = []
                        else:
                            st.session_state[key] = {}
                # Reset manual template state too
                st.session_state.manual_subject = ""
                st.session_state.manual_body    = ""
                st.session_state.email_mode     = "ai"
                st.session_state.current_step = 1
                st.rerun()


# Footer
st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: #94a3b8; padding: 2rem;'>
        <h3 style='color: #f1f5f9;'>AI Job Application System</h3>
        <p>Powered by Groq AI • SerpAPI Integration • Professional Email Automation</p>
        <p style='font-size: 0.9rem; opacity: 0.8;'>
            Features: Google Jobs Search • Freelance Platform Search • 
            Multi-Document Attachments • AI-Powered Email Generation • Bulk Email Management
        </p>
        <p style='font-size: 0.8rem; opacity: 0.6;'>
            Dark Theme Edition • pypdf PDF Support • Original Filenames Preserved •
            Anti-Spam Delayed Sending • Live Countdown • Batch Cooldowns • Resume/Retry
        </p>
    </div>
""", unsafe_allow_html=True)