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


def _has_streamlit_context() -> bool:
    """Return True when running via `streamlit run` script runner."""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx

        return get_script_run_ctx() is not None
    except Exception:
        return False


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


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&display=swap');

        html, body, [class*="css"], .stMarkdown, .stTextInput, .stButton, .stNumberInput {
            font-family: 'Manrope', sans-serif;
        }

        .stApp {
            background:
                radial-gradient(circle at 15% 15%, rgba(20, 184, 166, 0.08) 0%, rgba(20, 184, 166, 0) 30%),
                radial-gradient(circle at 90% 10%, rgba(37, 99, 235, 0.12) 0%, rgba(37, 99, 235, 0) 35%),
                linear-gradient(180deg, #f7fbff 0%, #eef4ff 100%);
        }
        .hero {
            padding: 22px 24px;
            border-radius: 18px;
            background: linear-gradient(115deg, #0d9488 0%, #0f766e 35%, #2563eb 100%);
            color: white;
            margin-bottom: 14px;
            box-shadow: 0 18px 36px rgba(15, 118, 110, 0.28);
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
        .hero h2 {
            margin: 0;
            font-size: 1.72rem;
            letter-spacing: 0.2px;
        }
        .hero p {
            margin: 6px 0 0 0;
            opacity: 0.95;
        }
        .hint {
            padding: 11px 13px;
            border-radius: 12px;
            background: #e6fffb;
            border: 1px solid #99f6e4;
            color: #134e4a;
        }

        .stSidebar {
            background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
        }

        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #ef4444 0%, #f97316 100%);
            border: none;
            color: #fff;
            font-weight: 700;
            border-radius: 10px;
        }

        .active-file {
            padding: 10px 12px;
            border-radius: 10px;
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            color: #1e3a8a;
            font-size: 0.92rem;
            margin-bottom: 8px;
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


def render_preview(excel_path: str, template_path: str, subject_template: str) -> None:
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

    body_template = load_template(template_path)
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
    inject_styles()

    st.markdown(
        """
        <div class="hero">
            <h2>Cold Email Automation Dashboard</h2>
            <p>One-time setup, reusable settings, preview support, and status-aware sending.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    settings = load_settings()

    with st.sidebar:
        st.header("Account & Config")
        sender_email = st.text_input("Sender Gmail", value=settings.get("sender_email", ""))
        app_password = st.text_input(
            "Gmail App Password",
            type="password",
            value=settings.get("app_password", ""),
        )
        remember_password = st.checkbox(
            "Remember App Password on this machine",
            value=bool(settings.get("remember_password", True)),
        )
        st.caption("For better security, keep this off and enter password only when needed.")

        st.divider()
        st.subheader("Quick File Select")
        excel_upload = st.file_uploader(
            "Select Excel File",
            type=["xlsx", "xls"],
            help="Best for deployment: upload directly instead of typing path.",
            key="excel_upload",
        )
        resume_upload = st.file_uploader(
            "Select Resume PDF",
            type=["pdf"],
            key="resume_upload",
        )
        template_upload = st.file_uploader(
            "Select Template File",
            type=["txt"],
            key="template_upload",
        )

        st.caption("Or keep local paths below if running on your machine.")
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
            f"<div class='active-file'><strong>Active Excel:</strong> {Path(active_excel_path).name}</div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div class='active-file'><strong>Active Resume:</strong> {Path(active_resume_path).name}</div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div class='active-file'><strong>Active Template:</strong> {Path(active_template_path).name}</div>",
            unsafe_allow_html=True,
        )

        st.divider()
        subject_template = st.text_input(
            "Subject Template",
            value=settings.get("subject_template", DEFAULT_SUBJECT_TEMPLATE),
        )
        delay_min = st.number_input(
            "Delay Min (seconds)",
            min_value=0,
            max_value=300,
            value=int(settings.get("delay_min", 5)),
            step=1,
        )
        delay_max = st.number_input(
            "Delay Max (seconds)",
            min_value=0,
            max_value=300,
            value=int(settings.get("delay_max", 10)),
            step=1,
        )
        max_retries = st.number_input(
            "Max Retries",
            min_value=0,
            max_value=10,
            value=int(settings.get("max_retries", 1)),
            step=1,
        )
        max_emails_per_run = st.number_input(
            "Max Emails Per Run (0 = no limit)",
            min_value=0,
            max_value=10000,
            value=int(settings.get("max_emails_per_run", 0)),
            step=1,
        )

        if st.button("Save Settings", use_container_width=True):
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
            }
            save_settings(saved_settings)
            st.success("Settings saved. Next time fields will auto-fill.")

    left, right = st.columns([1.2, 1])

    with left:
        st.markdown('<div class="hint">Use placeholders in template and subject: {name}, {company}, {title}</div>', unsafe_allow_html=True)
        st.write("")

        try:
            counts = get_status_counts(active_excel_path)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total", counts["Total"])
            c2.metric("Pending", counts["Pending"])
            c3.metric("Sent", counts["Sent"])
            c4.metric("Failed", counts["Failed"])
        except Exception as err:
            st.warning(f"Could not read Excel status: {err}")

        st.subheader("Email Preview")
        if st.button("Preview First Pending Email"):
            try:
                render_preview(active_excel_path, active_template_path, subject_template)
            except Exception as err:
                st.error(f"Preview failed: {err}")

    with right:
        st.subheader("Automation")
        preview_only = st.checkbox("Preview mode only (do not send)", value=False)
        run_now = st.button("Start Automation", type="primary", use_container_width=True)

        log_box = st.empty()
        progress_bar = st.progress(0)
        status_label = st.empty()

        if run_now:
            if delay_min > delay_max:
                st.error("Delay Min cannot be greater than Delay Max.")
                st.stop()

            if not preview_only and (not sender_email.strip() or not app_password.strip()):
                st.error("Sender Gmail and App Password are required for sending.")
                st.stop()

            logs: List[str] = []
            processed_rows = set()

            pending_count = 0
            try:
                df = read_contacts(active_excel_path)
                pending_count = int(df["Status"].apply(should_process_status).sum())
                if int(max_emails_per_run) > 0:
                    pending_count = min(pending_count, int(max_emails_per_run))
            except Exception:
                pending_count = 0

            def streamlit_logger(message: str) -> None:
                logs.append(message)
                match = re.search(r"Row\s+(\d+)", message)
                if match:
                    processed_rows.add(int(match.group(1)))

                if pending_count > 0:
                    ratio = min(len(processed_rows) / pending_count, 1.0)
                    progress_bar.progress(ratio)
                    status_label.caption(
                        f"Processed {len(processed_rows)} of {pending_count} pending rows"
                    )
                log_box.text_area("Live Logs", value="\n".join(logs[-200:]), height=320)

            config = EmailJobConfig(
                excel_path=active_excel_path,
                resume_pdf_path=active_resume_path,
                sender_email=sender_email.strip(),
                app_password=app_password.strip(),
                template_path=active_template_path.strip() or None,
                subject_template=subject_template.strip() or DEFAULT_SUBJECT_TEMPLATE,
                delay_min=int(delay_min),
                delay_max=int(delay_max),
                max_retries=int(max_retries),
                max_emails_per_run=int(max_emails_per_run),
                preview_only=preview_only,
            )

            try:
                send_emails(config=config, logger=streamlit_logger)
                progress_bar.progress(1.0)
                st.success("Automation completed successfully.")

                updated_excel = Path(active_excel_path)
                if updated_excel.exists():
                    st.download_button(
                        "Download Updated Excel",
                        data=updated_excel.read_bytes(),
                        file_name=updated_excel.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            except Exception as err:
                st.error(f"Automation stopped: {err}")


if __name__ == "__main__":
    _auto_launch_streamlit_if_needed()
    main()