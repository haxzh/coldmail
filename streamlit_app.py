import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from email_automation import (
    DEFAULT_EXCEL_PATH,
    DEFAULT_RESUME_PATH,
    DEFAULT_SUBJECT_TEMPLATE,
    DEFAULT_TEMPLATE_PATH,
    EmailJobConfig,
    load_template,
    read_contacts,
    render_template,
    safe_text,
    send_emails,
    should_process_status,
)


SETTINGS_FILE = Path("mail_settings.json")
UPLOAD_DIR = Path(".runtime_uploads")
st: Optional[object] = None
THEME_FILE = Path("theme_settings.json")


def _has_streamlit_context() -> bool:
    """Return True when running via `streamlit run` script runner."""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx

        return get_script_run_ctx() is not None
    except Exception:
        return False


def load_theme_settings() -> Dict[str, Any]:
    """Load theme settings (dark/light mode)."""
    default_theme = {"dark_mode": False}
    if THEME_FILE.exists():
        try:
            loaded = json.loads(THEME_FILE.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                return loaded
        except Exception:
            pass
    return default_theme


def save_theme_settings(theme: Dict[str, Any]) -> None:
    """Save theme settings to file."""
    THEME_FILE.write_text(json.dumps(theme, indent=2), encoding="utf-8")


def _auto_launch_streamlit_if_needed() -> None:
    """
    If user runs `python streamlit_app.py`, relaunch with streamlit so the UI
    opens correctly in a browser without ScriptRunContext warnings.
    """
    if _has_streamlit_context():
        return

    cmd = [sys.executable, "-m", "streamlit", "run", str(Path(__file__).name)]
    if len(sys.argv) > 1:
        cmd.extend(sys.argv[1:])
    subprocess.run(cmd, check=False)
    raise SystemExit(0)


def default_settings() -> Dict[str, Any]:
    return {
        "excel_path": DEFAULT_EXCEL_PATH,
        "resume_path": DEFAULT_RESUME_PATH,
        "template_path": DEFAULT_TEMPLATE_PATH,
        "sender_email": "",
        "app_password": "",
        "remember_password": True,
        "subject_template": DEFAULT_SUBJECT_TEMPLATE,
        "delay_min": 5,
        "delay_max": 10,
        "max_retries": 1,
        "max_emails_per_run": 0,
        "use_custom_template_text": False,
        "template_body_text": "",
    }


def load_settings() -> Dict[str, Any]:
    settings = default_settings()
    if SETTINGS_FILE.exists():
        try:
            loaded = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                settings.update(loaded)
        except Exception:
            pass
    return settings


def save_settings(settings: Dict[str, Any]) -> None:
    SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")


def ensure_upload_dir() -> Path:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    return UPLOAD_DIR


def save_uploaded_file(uploaded_file: Any, fallback_name: str) -> str:
    upload_dir = ensure_upload_dir()
    file_name = Path(getattr(uploaded_file, "name", fallback_name) or fallback_name).name
    target = upload_dir / file_name
    target.write_bytes(uploaded_file.getbuffer())
    return str(target.resolve())


def inject_styles(dark_mode: bool = False) -> None:
    """Inject advanced CSS styles with dark/light mode support."""
    if dark_mode:
        # Dark Mode Styles
        st.markdown(
            """
            <style>
            @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

            :root {
                --primary: #3b82f6;
                --primary-light: #60a5fa;
                --primary-dark: #1e40af;
                --accent: #ec4899;
                --success: #10b981;
                --warning: #f59e0b;
                --danger: #ef4444;
                --bg-primary: #0f172a;
                --bg-secondary: #1e293b;
                --bg-tertiary: #334155;
                --text-primary: #f1f5f9;
                --text-secondary: #cbd5e1;
                --border: #475569;
            }

            html, body, [class*="css"], .stMarkdown, .stTextInput, .stButton, .stNumberInput {
                font-family: 'Inter', 'Manrope', sans-serif;
            }

            .stApp {
                background: linear-gradient(135deg, #0f172a 0%, #1a1f35 50%, #16213e 100%);
                color: var(--text-primary);
            }

            .stApp, .stApp p, .stApp li, .stApp span, .stApp div, .stApp label {
                color: var(--text-primary);
            }

            .stCaption,
            .stHelp,
            .stMarkdown small,
            .stMarkdown p,
            [data-testid="stMetricLabel"],
            [data-testid="stMetricValue"],
            [data-testid="stMetricDelta"] {
                color: var(--text-secondary);
            }

            [data-testid="stWidgetLabel"] p,
            [data-testid="stFileUploader"] label,
            [data-testid="stFileUploaderDropzone"] p,
            [data-testid="stFileUploaderDropzone"] span,
            .stCheckbox label,
            .stRadio label,
            .stSelectbox label,
            .stTextInput label,
            .stNumberInput label,
            .stTextArea label {
                color: var(--text-secondary) !important;
            }

            .stMarkdown h1,
            .stMarkdown h2,
            .stMarkdown h3,
            .stMarkdown h4,
            .stMarkdown h5,
            .stMarkdown h6,
            .stSubheader,
            .stHeader {
                color: var(--text-primary) !important;
            }

            .stTextInput > div > div,
            .stNumberInput > div > div,
            .stPasswordInput > div > div,
            .stTextArea > div > div,
            .stSelectbox > div > div {
                background: rgba(15, 23, 42, 0.72);
                border: 1px solid rgba(148, 163, 184, 0.28);
                border-radius: 12px;
            }

            .stTextInput > div > div > input,
            .stNumberInput > div > div > input,
            .stPasswordInput > div > div > input,
            .stTextArea > div > div > textarea {
                background: transparent;
                color: var(--text-primary);
                caret-color: var(--text-primary);
            }

            .stTextInput > div > div > input::placeholder,
            .stNumberInput > div > div > input::placeholder,
            .stPasswordInput > div > div > input::placeholder,
            .stTextArea > div > div > textarea::placeholder {
                color: rgba(203, 213, 225, 0.75);
            }

            .stTextInput > div > div > input:focus,
            .stNumberInput > div > div > input:focus,
            .stPasswordInput > div > div > input:focus,
            .stTextArea > div > div > textarea:focus {
                border-color: var(--primary-light);
                box-shadow: 0 0 0 3px rgba(96, 165, 250, 0.18);
            }

            .stFileUploader {
                background: rgba(15, 23, 42, 0.55);
                border-radius: 14px;
                padding: 4px;
            }

            [data-testid="stFileUploaderDropzone"] {
                background: rgba(30, 41, 59, 0.95);
                border: 1px dashed rgba(148, 163, 184, 0.35);
                border-radius: 14px;
                color: var(--text-primary);
            }

            [data-testid="stFileUploaderDropzone"] button {
                background: rgba(59, 130, 246, 0.14);
                color: var(--text-primary);
                border: 1px solid rgba(96, 165, 250, 0.35);
            }

            [data-testid="stExpander"] {
                background: rgba(15, 23, 42, 0.55);
                border: 1px solid rgba(148, 163, 184, 0.22);
                border-radius: 14px;
            }

            [data-testid="stExpander"] summary,
            [data-testid="stExpander"] button {
                color: var(--text-primary) !important;
            }

            .stDownloadButton button,
            [data-testid="stDownloadButton"] button {
                background: linear-gradient(135deg, #0ea5e9 0%, #2563eb 100%);
                color: #ffffff !important;
                border: none;
            }

            /* Hero Section */
            .hero {
                padding: 32px 28px;
                border-radius: 20px;
                background: linear-gradient(135deg, rgba(30, 41, 59, 0.96) 0%, rgba(51, 65, 85, 0.96) 50%, rgba(30, 41, 59, 0.96) 100%);
                border: 1px solid rgba(148, 163, 184, 0.25);
                color: white;
                margin-bottom: 24px;
                box-shadow: 0 20px 50px rgba(0, 0, 0, 0.5);
                position: relative;
                overflow: hidden;
            }

            .hero::before {
                content: '';
                position: absolute;
                top: -50%;
                right: -50%;
                width: 200%;
                height: 200%;
                background: radial-gradient(circle, rgba(59, 130, 246, 0.1) 0%, transparent 70%);
                animation: heroGlow 8s ease-in-out infinite;
            }

            @keyframes heroGlow {
                0%, 100% { transform: translate(0, 0); }
                50% { transform: translate(30px, -30px); }
            }

            .hero h2 {
                position: relative;
                z-index: 1;
                margin: 0;
                font-size: 2rem;
                font-weight: 800;
                letter-spacing: -0.5px;
                background: linear-gradient(135deg, #60a5fa, #a78bfa);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                background-clip: text;
            }

            .hero p {
                position: relative;
                z-index: 1;
                margin: 8px 0 0 0;
                opacity: 0.85;
                font-size: 1rem;
                color: var(--text-secondary);
            }

            /* Sidebar */
            .stSidebar {
                background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%);
                border-right: 1px solid var(--border);
            }

            .stSidebar .stCaption,
            .stSidebar .stMarkdown p,
            .stSidebar label,
            .stSidebar [data-testid="stWidgetLabel"] p {
                color: rgba(226, 232, 240, 0.88) !important;
            }

            /* Metrics Cards */
            [data-testid="metric-container"] {
                background: linear-gradient(135deg, rgba(30, 41, 59, 0.94) 0%, rgba(51, 65, 85, 0.94) 100%);
                border: 1px solid var(--border);
                border-radius: 16px;
                padding: 20px;
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
                transition: all 0.3s ease;
            }

            [data-testid="metric-container"]:hover {
                border-color: var(--primary);
                box-shadow: 0 15px 40px rgba(59, 130, 246, 0.2);
                transform: translateY(-2px);
            }

            /* Buttons */
            .stButton > button {
                border-radius: 12px;
                font-weight: 600;
                transition: all 0.3s ease;
                border: none;
            }

            .stButton > button[kind="primary"] {
                background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
                color: white;
            }

            .stButton > button[kind="primary"]:hover {
                background: linear-gradient(135deg, #60a5fa 0%, #3b82f6 100%);
                box-shadow: 0 10px 30px rgba(59, 130, 246, 0.4);
            }

            .stButton > button[kind="secondary"] {
                background: rgba(30, 41, 59, 0.92);
                border: 1px solid rgba(148, 163, 184, 0.3);
                color: var(--text-primary);
            }

            .stButton > button[kind="secondary"]:hover {
                background: rgba(51, 65, 85, 0.98);
                border-color: var(--primary);
            }

            /* Input Fields */
            .stTextInput > div > div > input,
            .stNumberInput > div > div > input,
            .stPasswordInput > div > div > input,
            .stTextArea > div > div > textarea {
                background: var(--bg-secondary);
                border: 1px solid var(--border);
                color: var(--text-primary);
                border-radius: 10px;
                padding: 10px 12px;
                transition: all 0.3s ease;
            }

            .stTextInput > div > div > input:focus,
            .stNumberInput > div > div > input:focus,
            .stPasswordInput > div > div > input:focus,
            .stTextArea > div > div > textarea:focus {
                border-color: var(--primary);
                box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
                outline: none;
            }

            /* Checkboxes & Radio */
            .stCheckbox > label > div > div,
            .stRadio > label > div > div {
                border-radius: 6px;
                border: 1px solid var(--border);
                background: var(--bg-secondary);
            }

            .stCheckbox > label > div > div:hover,
            .stRadio > label > div > div:hover {
                border-color: var(--primary);
            }

            /* Info/Success/Warning/Error Boxes */
            .stSuccess {
                background: rgba(16, 185, 129, 0.1);
                border: 1px solid rgba(16, 185, 129, 0.3);
                border-radius: 12px;
                padding: 16px;
                color: #6ee7b7;
            }

            .stError {
                background: rgba(239, 68, 68, 0.1);
                border: 1px solid rgba(239, 68, 68, 0.3);
                border-radius: 12px;
                padding: 16px;
                color: #fca5a5;
            }

            .stWarning {
                background: rgba(245, 158, 11, 0.1);
                border: 1px solid rgba(245, 158, 11, 0.3);
                border-radius: 12px;
                padding: 16px;
                color: #fcd34d;
            }

            .stInfo {
                background: rgba(59, 130, 246, 0.1);
                border: 1px solid rgba(59, 130, 246, 0.3);
                border-radius: 12px;
                padding: 16px;
                color: #93c5fd;
            }

            /* Hints & Cards */
            .hint {
                padding: 16px 18px;
                border-radius: 14px;
                background: rgba(59, 130, 246, 0.1);
                border: 1px solid rgba(59, 130, 246, 0.3);
                color: #93c5fd;
                font-weight: 500;
                margin-bottom: 12px;
            }

            .active-file {
                padding: 14px 16px;
                border-radius: 12px;
                background: rgba(59, 130, 246, 0.05);
                border: 1px solid rgba(59, 130, 246, 0.2);
                color: #93c5fd;
                font-size: 0.92rem;
                margin-bottom: 10px;
                font-weight: 500;
            }

            /* Divider */
            .stDivider {
                border-color: var(--border);
            }

            /* Sections */
            .stSubheader {
                color: var(--text-primary);
                font-weight: 700;
            }

            /* Text Area */
            .stTextArea > label {
                color: var(--text-secondary);
            }

            /* Caption */
            .stCaption {
                color: var(--text-secondary);
            }

            /* Columns spacing */
            .stColumns {
                gap: 20px;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )
    else:
        # Light Mode Styles
        st.markdown(
            """
            <style>
            @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

            :root {
                --primary: #2563eb;
                --primary-light: #3b82f6;
                --primary-dark: #1e40af;
                --accent: #db2777;
                --success: #059669;
                --warning: #d97706;
                --danger: #dc2626;
                --bg-primary: #ffffff;
                --bg-secondary: #f8fafc;
                --bg-tertiary: #e2e8f0;
                --text-primary: #0f172a;
                --text-secondary: #64748b;
                --border: #cbd5e1;
            }

            html, body, [class*="css"], .stMarkdown, .stTextInput, .stButton, .stNumberInput {
                font-family: 'Inter', 'Manrope', sans-serif;
            }

            .stApp {
                background: linear-gradient(135deg, #f8fafc 0%, #f0f4f8 50%, #e8ecf1 100%);
                color: var(--text-primary);
            }

            /* Hero Section */
            .hero {
                padding: 32px 28px;
                border-radius: 20px;
                background: linear-gradient(135deg, #ffffff 0%, #f8fafc 50%, #f3f4f6 100%);
                border: 1px solid rgba(226, 232, 240, 0.6);
                color: var(--text-primary);
                margin-bottom: 24px;
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.08);
                position: relative;
                overflow: hidden;
            }

            .hero::before {
                content: '';
                position: absolute;
                top: -50%;
                right: -50%;
                width: 200%;
                height: 200%;
                background: radial-gradient(circle, rgba(37, 99, 235, 0.08) 0%, transparent 70%);
                animation: heroGlow 8s ease-in-out infinite;
            }

            @keyframes heroGlow {
                0%, 100% { transform: translate(0, 0); }
                50% { transform: translate(30px, -30px); }
            }

            .hero h2 {
                position: relative;
                z-index: 1;
                margin: 0;
                font-size: 2rem;
                font-weight: 800;
                letter-spacing: -0.5px;
                background: linear-gradient(135deg, #2563eb, #7c3aed);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                background-clip: text;
            }

            .hero p {
                position: relative;
                z-index: 1;
                margin: 8px 0 0 0;
                opacity: 0.8;
                font-size: 1rem;
                color: var(--text-secondary);
            }

            /* Sidebar */
            .stSidebar {
                background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
                border-right: 1px solid var(--border);
            }

            /* Metrics Cards */
            [data-testid="metric-container"] {
                background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
                border: 1px solid var(--border);
                border-radius: 16px;
                padding: 20px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
                transition: all 0.3s ease;
            }

            [data-testid="metric-container"]:hover {
                border-color: var(--primary);
                box-shadow: 0 8px 20px rgba(37, 99, 235, 0.15);
                transform: translateY(-2px);
            }

            /* Buttons */
            .stButton > button {
                border-radius: 12px;
                font-weight: 600;
                transition: all 0.3s ease;
                border: none;
            }

            .stButton > button[kind="primary"] {
                background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
                color: white;
            }

            .stButton > button[kind="primary"]:hover {
                background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
                box-shadow: 0 8px 20px rgba(37, 99, 235, 0.3);
                transform: translateY(-1px);
            }

            .stButton > button[kind="secondary"] {
                background: var(--bg-secondary);
                border: 1px solid var(--border);
                color: var(--text-primary);
            }

            .stButton > button[kind="secondary"]:hover {
                background: white;
                border-color: var(--primary);
                color: var(--primary);
            }

            /* Input Fields */
            .stTextInput > div > div > input,
            .stNumberInput > div > div > input,
            .stPasswordInput > div > div > input,
            .stTextArea > div > div > textarea {
                background: white;
                border: 1px solid var(--border);
                color: var(--text-primary);
                border-radius: 10px;
                padding: 10px 12px;
                transition: all 0.3s ease;
            }

            .stTextInput > div > div > input:focus,
            .stNumberInput > div > div > input:focus,
            .stPasswordInput > div > div > input:focus,
            .stTextArea > div > div > textarea:focus {
                border-color: var(--primary);
                box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
                outline: none;
            }

            /* Checkboxes & Radio */
            .stCheckbox > label > div > div,
            .stRadio > label > div > div {
                border-radius: 6px;
                border: 1px solid var(--border);
                background: white;
            }

            .stCheckbox > label > div > div:hover,
            .stRadio > label > div > div:hover {
                border-color: var(--primary);
            }

            /* Info/Success/Warning/Error Boxes */
            .stSuccess {
                background: rgba(5, 150, 105, 0.1);
                border: 1px solid rgba(5, 150, 105, 0.3);
                border-radius: 12px;
                padding: 16px;
                color: #059669;
            }

            .stError {
                background: rgba(220, 38, 38, 0.1);
                border: 1px solid rgba(220, 38, 38, 0.3);
                border-radius: 12px;
                padding: 16px;
                color: #dc2626;
            }

            .stWarning {
                background: rgba(217, 119, 6, 0.1);
                border: 1px solid rgba(217, 119, 6, 0.3);
                border-radius: 12px;
                padding: 16px;
                color: #b45309;
            }

            .stInfo {
                background: rgba(37, 99, 235, 0.1);
                border: 1px solid rgba(37, 99, 235, 0.3);
                border-radius: 12px;
                padding: 16px;
                color: #2563eb;
            }

            /* Hints & Cards */
            .hint {
                padding: 16px 18px;
                border-radius: 14px;
                background: rgba(37, 99, 235, 0.08);
                border: 1px solid rgba(37, 99, 235, 0.3);
                color: #1e40af;
                font-weight: 500;
                margin-bottom: 12px;
            }

            .active-file {
                padding: 14px 16px;
                border-radius: 12px;
                background: rgba(37, 99, 235, 0.08);
                border: 1px solid rgba(37, 99, 235, 0.3);
                color: #1e40af;
                font-size: 0.92rem;
                margin-bottom: 10px;
                font-weight: 500;
            }

            /* Divider */
            .stDivider {
                border-color: var(--border);
            }

            /* Sections */
            .stSubheader {
                color: var(--text-primary);
                font-weight: 700;
            }

            /* Text Area */
            .stTextArea > label {
                color: var(--text-secondary);
            }

            /* Caption */
            .stCaption {
                color: var(--text-secondary);
            }

            /* Columns spacing */
            .stColumns {
                gap: 20px;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )


def get_status_counts(excel_path: str) -> Dict[str, int]:
    counts = {"Total": 0, "Pending": 0, "Sent": 0, "Failed": 0}
    path = Path(excel_path)
    if not path.exists():
        return counts

    df = read_contacts(excel_path)
    counts["Total"] = len(df)
    counts["Pending"] = int(df["Status"].apply(should_process_status).sum())
    counts["Sent"] = int((df["Status"].astype(str).str.strip().str.lower() == "sent").sum())
    counts["Failed"] = int((df["Status"].astype(str).str.strip().str.lower() == "failed").sum())
    return counts


def render_preview(
    excel_path: str,
    template_path: str,
    subject_template: str,
    template_text: Optional[str] = None,
) -> None:
    df = read_contacts(excel_path)
    pending = df[df["Status"].apply(should_process_status)]
    if pending.empty:
        st.info("No pending rows found (Status empty/Not Sent).")
        return

    first = pending.iloc[0]
    name = safe_text(first.get("Name")) or "Hiring Manager"
    company = safe_text(first.get("Company Name")) or "your company"
    title = safe_text(first.get("Title")) or "the role"
    email = safe_text(first.get("Email"))

    body_template = template_text if safe_text(template_text) else load_template(template_path)
    subject = render_template(subject_template, name=name, company=company, title=title)
    body = render_template(body_template, name=name, company=company, title=title)

    st.subheader("Preview First Pending Email")
    st.write(f"To: {email}")
    st.write(f"Subject: {subject}")
    st.text_area("Body", value=body, height=260)


def main() -> None:
    global st
    import streamlit as st_module

    st = st_module

    st.set_page_config(page_title="Cold Email Dashboard", page_icon="📨", layout="wide")
    
    # Initialize theme
    if "dark_mode" not in st.session_state:
        theme = load_theme_settings()
        st.session_state.dark_mode = theme.get("dark_mode", False)
    
    inject_styles(st.session_state.dark_mode)

    # Top right theme toggle
    col1, col2, col3 = st.columns([1, 0.15, 0.15])
    with col3:
        if st.button("🌙" if not st.session_state.dark_mode else "☀️", key="theme_toggle", help="Toggle Theme"):
            st.session_state.dark_mode = not st.session_state.dark_mode
            save_theme_settings({"dark_mode": st.session_state.dark_mode})
            st.rerun()

    st.markdown(
        """
        <div class="hero">
            <h2>📨 Cold Email Automation Dashboard</h2>
            <p>Advanced email automation with status tracking, preview support, and batch sending capabilities.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    settings = load_settings()
    if "last_excel_download_path" not in st.session_state:
        st.session_state.last_excel_download_path = ""

    with st.sidebar:
        st.header("⚙️ Account & Configuration")
        
        with st.expander("📧 Email Settings", expanded=True):
            sender_email = st.text_input("Sender Gmail", value=settings.get("sender_email", ""), placeholder="your-email@gmail.com")
            app_password = st.text_input(
                "Gmail App Password",
                type="password",
                value=settings.get("app_password", ""),
                placeholder="Enter 16-character app password",
            )
            remember_password = st.checkbox(
                "Remember password on this device",
                value=bool(settings.get("remember_password", True)),
                help="Uncheck for better security on shared computers"
            )
        
        st.divider()
        
        with st.expander("📁 Quick File Select", expanded=True):
            st.caption("📤 Upload Files")
            excel_upload = st.file_uploader(
                "Select Excel File",
                type=["xlsx", "xls"],
                help="Best for deployment: upload directly instead of typing path.",
                key="excel_upload",
            )
            resume_upload = st.file_uploader(
                "Select Resume PDF",
                type=["pdf"],
                help="Optional: PDF to attach to emails",
                key="resume_upload",
            )
            template_upload = st.file_uploader(
                "Select Template File",
                type=["txt"],
                help="Email body template with {name}, {company}, {title} placeholders",
                key="template_upload",
            )

            st.caption("📍 Or Use Local Paths")
            excel_path = st.text_input("Excel Path", value=settings.get("excel_path", DEFAULT_EXCEL_PATH))
            resume_path = st.text_input("Resume PDF Path", value=settings.get("resume_path", DEFAULT_RESUME_PATH))
            template_path = st.text_input(
                "Template Path",
                value=settings.get("template_path", DEFAULT_TEMPLATE_PATH),
            )

            active_excel_path = save_uploaded_file(excel_upload, DEFAULT_EXCEL_PATH) if excel_upload else excel_path
            active_resume_path = save_uploaded_file(resume_upload, DEFAULT_RESUME_PATH) if resume_upload else resume_path
            active_template_path = (
                save_uploaded_file(template_upload, DEFAULT_TEMPLATE_PATH) if template_upload else template_path
            )

            st.markdown(
                f"<div class='active-file'><strong>✅ Excel:</strong> {Path(active_excel_path).name}</div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<div class='active-file'><strong>✅ Resume:</strong> {Path(active_resume_path).name}</div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<div class='active-file'><strong>✅ Template:</strong> {Path(active_template_path).name}</div>",
                unsafe_allow_html=True,
            )
        
        st.divider()
        
        with st.expander("📝 Email Template", expanded=False):
            subject_template = st.text_input(
                "Subject Template",
                value=settings.get("subject_template", DEFAULT_SUBJECT_TEMPLATE),
                help="Use {name}, {company}, {title} for personalization"
            )
            use_custom_template_text = st.checkbox(
                "✏️ Use Email Text from Dashboard",
                value=bool(settings.get("use_custom_template_text", False)),
                help="If enabled, this text is used instead of template file.",
            )
            template_body_text = st.text_area(
                "Email Body Text",
                value=str(settings.get("template_body_text", "")),
                height=180,
                placeholder=(
                    "Write your email body here.\nYou can use:\n{name} - Recipient name\n{company} - Company name\n{title} - Job title"
                ),
            )
        
        st.divider()
        
        with st.expander("⏱️ Timing & Limits", expanded=False):
            delay_min = st.number_input(
                "Delay Min (seconds)",
                min_value=0,
                max_value=300,
                value=int(settings.get("delay_min", 5)),
                step=1,
                help="Minimum delay between emails"
            )
            delay_max = st.number_input(
                "Delay Max (seconds)",
                min_value=0,
                max_value=300,
                value=int(settings.get("delay_max", 10)),
                step=1,
                help="Maximum delay between emails (random between min-max)"
            )
            max_retries = st.number_input(
                "Max Retries on Failure",
                min_value=0,
                max_value=10,
                value=int(settings.get("max_retries", 1)),
                step=1,
                help="Number of times to retry failed emails"
            )
            max_emails_per_run = st.number_input(
                "Max Emails Per Run (0 = no limit)",
                min_value=0,
                max_value=10000,
                value=int(settings.get("max_emails_per_run", 0)),
                step=1,
                help="Limit emails sent in one run (useful for testing)"
            )

        st.divider()
        
        if st.button("💾 Save Settings", use_container_width=True):
            saved_settings = {
                "sender_email": sender_email,
                "app_password": app_password if remember_password else "",
                "remember_password": remember_password,
                "excel_path": excel_path,
                "resume_path": resume_path,
                "template_path": template_path,
                "subject_template": subject_template,
                "delay_min": int(delay_min),
                "delay_max": int(delay_max),
                "max_retries": int(max_retries),
                "max_emails_per_run": int(max_emails_per_run),
                "use_custom_template_text": bool(use_custom_template_text),
                "template_body_text": template_body_text,
            }
            save_settings(saved_settings)
            st.success("✅ Settings saved successfully!")

    left, right = st.columns([1.2, 1])

    with left:
        st.markdown('<div class="hint">💡 Use placeholders: <strong>{name}</strong>, <strong>{company}</strong>, <strong>{title}</strong></div>', unsafe_allow_html=True)
        st.write("")

        # Status Metrics
        st.subheader("📊 Campaign Status")
        try:
            counts = get_status_counts(active_excel_path)
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("📋 Total", counts["Total"], help="Total contacts in spreadsheet")
            with c2:
                st.metric("⏳ Pending", counts["Pending"], help="Ready to send")
            with c3:
                st.metric("✅ Sent", counts["Sent"], help="Successfully sent")
            with c4:
                st.metric("❌ Failed", counts["Failed"], help="Send failures")
        except Exception as err:
            st.warning(f"⚠️ Could not read Excel: {err}")

        st.divider()
        
        # Preview Section
        st.subheader("👁️ Email Preview")
        col_preview_btn, col_spacer = st.columns([1, 2])
        with col_preview_btn:
            preview_btn = st.button("🔍 Preview Email", use_container_width=True)
        
        if preview_btn:
            try:
                render_preview(
                    active_excel_path,
                    active_template_path,
                    subject_template,
                    template_body_text if use_custom_template_text else None,
                )
            except Exception as err:
                st.error(f"❌ Preview failed: {err}")

    with right:
        st.subheader("🚀 Send Campaign")
        
        # Initialize campaign state
        if "campaign_running" not in st.session_state:
            st.session_state.campaign_running = False
        if "campaign_stopped" not in st.session_state:
            st.session_state.campaign_stopped = False
        
        # Preview mode toggle
        preview_only = st.checkbox(
            "🔒 Preview mode only (simulate, don't send)",
            value=False,
            help="Test without actually sending emails"
        )
        
        # Launch button
        col_start, col_stop = st.columns([1, 1])
        with col_start:
            run_now = st.button("▶️ Start Automation", type="primary", use_container_width=True, key="start_btn")
        with col_stop:
            if st.session_state.campaign_running:
                stop_now = st.button("⏹️ Stop Campaign", type="secondary", use_container_width=True, key="stop_btn")
                if stop_now:
                    st.session_state.campaign_stopped = True
                    st.warning("⚠️ Campaign stopped by user. Download the updated Excel below.")
            else:
                st.button("⏹️ Stop Campaign", type="secondary", use_container_width=True, disabled=True, key="stop_btn_disabled")

        # Progress tracking
        if run_now:
            st.session_state.campaign_running = True
            st.session_state.campaign_stopped = False
            
            if delay_min > delay_max:
                st.session_state.campaign_running = False
                st.error("❌ Delay Min cannot be greater than Delay Max!")
                st.stop()

            if not preview_only and (not sender_email.strip() or not app_password.strip()):
                st.session_state.campaign_running = False
                st.error("❌ Sender Gmail and App Password are required for sending!")
                st.stop()

            # Create placeholders for live updates
            progress_placeholder = st.empty()
            status_placeholder = st.empty()
            log_placeholder = st.empty()
            result_placeholder = st.empty()
            summary_placeholder = st.empty()

            logs: List[str] = []
            processed_rows = set()
            failed_rows = set()

            pending_count = 0
            try:
                df = read_contacts(active_excel_path)
                pending_count = int(df["Status"].apply(should_process_status).sum())
                if int(max_emails_per_run) > 0:
                    pending_count = min(pending_count, int(max_emails_per_run))
            except Exception:
                pending_count = 0

            def streamlit_logger(message: str) -> None:
                if st.session_state.campaign_stopped:
                    return
                    
                logs.append(message)
                
                # Track failed emails
                if "failed" in message.lower() or "error" in message.lower():
                    match = re.search(r"Row\s+(\d+)", message)
                    if match:
                        failed_rows.add(int(match.group(1)))
                
                # Track processed rows
                match = re.search(r"Row\s+(\d+)", message)
                if match:
                    processed_rows.add(int(match.group(1)))

                if pending_count > 0:
                    ratio = min(len(processed_rows) / pending_count, 1.0)
                    with progress_placeholder.container():
                        st.progress(ratio)
                    with status_placeholder.container():
                        st.caption(
                            f"📤 Processed: {len(processed_rows)}/{pending_count} emails | ❌ Failed: {len(failed_rows)}"
                        )
                
                with log_placeholder.container():
                    st.text_area("📋 Live Logs", value="\n".join(logs[-150:]), height=280, disabled=True)

            config = EmailJobConfig(
                excel_path=active_excel_path,
                resume_pdf_path=active_resume_path,
                sender_email=sender_email.strip(),
                app_password=app_password.strip(),
                template_path=active_template_path.strip() or None,
                template_text=template_body_text if use_custom_template_text else None,
                subject_template=subject_template.strip() or DEFAULT_SUBJECT_TEMPLATE,
                delay_min=int(delay_min),
                delay_max=int(delay_max),
                max_retries=int(max_retries),
                max_emails_per_run=int(max_emails_per_run),
                preview_only=preview_only,
            )

            try:
                send_emails(config=config, logger=streamlit_logger)
                
                # Campaign completed
                st.session_state.campaign_running = False
                
                with result_placeholder.container():
                    st.success("✅ Campaign completed successfully!")

                # Show summary
                with summary_placeholder.container():
                    col_s1, col_s2, col_s3 = st.columns(3)
                    with col_s1:
                        st.metric("📧 Total Processed", len(processed_rows))
                    with col_s2:
                        st.metric("✅ Successful", len(processed_rows) - len(failed_rows))
                    with col_s3:
                        st.metric("❌ Failed", len(failed_rows))

                updated_excel = Path(active_excel_path)
                if updated_excel.exists():
                    st.session_state.last_excel_download_path = str(updated_excel)
                    
            except Exception as err:
                st.session_state.campaign_running = False
                with result_placeholder.container():
                    st.error(f"❌ Campaign stopped: {err}")

        # Show summary if campaign was stopped
        if st.session_state.campaign_stopped and not st.session_state.campaign_running:
            with summary_placeholder.container():
                st.info("📊 Campaign was stopped. Here's the summary:")
                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    st.metric("📧 Total Processed", len(processed_rows) if 'processed_rows' in locals() else 0)
                with col_s2:
                    st.metric("❌ Failed", len(failed_rows) if 'failed_rows' in locals() else 0)

        # Download section
        st.divider()
        st.subheader("📥 Download Results")
        saved_excel_path = st.session_state.get("last_excel_download_path", "")
        if saved_excel_path:
            updated_excel = Path(saved_excel_path)
            if updated_excel.exists():
                # Read the file and show download button
                excel_bytes = updated_excel.read_bytes()
                
                # Show file info
                st.caption(f"📄 File: {updated_excel.name} | Size: {len(excel_bytes) / 1024:.1f} KB")
                
                st.download_button(
                    "⬇️ Download Updated Excel",
                    data=excel_bytes,
                    file_name=updated_excel.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.success("✅ Excel ready for download - it contains all current statuses including just-sent emails!")


if __name__ == "__main__":
    _auto_launch_streamlit_if_needed()
    main()