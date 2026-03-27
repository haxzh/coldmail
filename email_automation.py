import argparse
import random
import re
import smtplib
import ssl
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Callable, Optional

import pandas as pd


REQUIRED_COLUMNS = ["Name", "Email", "Company Name", "Title", "Status"]
DATE_COLUMN = "Date"
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL_PATH = "contacts.xlsx"
DEFAULT_RESUME_PATH = "resume.pdf"
DEFAULT_TEMPLATE_PATH = "email_template.txt"
DEFAULT_SUBJECT_TEMPLATE = "Application for opportunities at {company}"
DEFAULT_BODY_TEMPLATE = (
    "Dear {name},\n\n"
    "I hope you are doing well.\n\n"
    "My name is [Your Name], and I am writing to express my interest in opportunities at "
    "{company}. I am interested in the {title} role and would be grateful if you could "
    "consider my profile for relevant openings.\n\n"
    "I have attached my resume for your review. I would appreciate the opportunity to "
    "connect and discuss how my background can contribute to your team.\n\n"
    "Thank you for your time and consideration.\n\n"
    "Sincerely,\n"
    "[Your Name]\n"
    "[Phone Number]\n"
    "[LinkedIn URL]"
)


@dataclass
class EmailJobConfig:
    excel_path: str
    resume_pdf_path: str
    sender_email: str
    app_password: str
    template_path: Optional[str] = None
    subject_template: str = DEFAULT_SUBJECT_TEMPLATE
    delay_min: int = 5
    delay_max: int = 10
    max_retries: int = 1
    max_emails_per_run: int = 0
    preview_only: bool = False


def safe_text(value: object) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_valid_email(email: str) -> bool:
    pattern = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    return bool(re.match(pattern, email.strip()))


def should_process_status(value: object) -> bool:
    status = safe_text(value).lower()
    return status in {"", "not sent"}


def resolve_path(path_str: str) -> Path:
    """Resolve input path against CWD first, then script directory."""
    path = Path(path_str).expanduser()
    if path.is_absolute():
        return path

    cwd_candidate = Path.cwd() / path
    if cwd_candidate.exists():
        return cwd_candidate

    return BASE_DIR / path


def load_template(template_path: Optional[str]) -> str:
    if not template_path:
        return DEFAULT_BODY_TEMPLATE

    path = resolve_path(template_path)
    if not path.exists():
        raise FileNotFoundError(
            f"Template file not found: {template_path} (resolved: {path})"
        )
    return path.read_text(encoding="utf-8")


def render_template(text: str, *, name: str, company: str, title: str) -> str:
    try:
        return text.format(name=name, company=company, title=title)
    except KeyError as err:
        raise ValueError(
            "Template contains an unknown placeholder. "
            "Use only: {name}, {company}, {title}."
        ) from err


def read_contacts(excel_path: str) -> pd.DataFrame:
    path = resolve_path(excel_path)
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path} (resolved: {path})")

    df = pd.read_excel(path)

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    # Keep writable text columns as object dtype to avoid pandas FutureWarning
    # when writing strings into columns inferred as float (all-NaN in Excel).
    df["Status"] = df["Status"].astype("object")

    if DATE_COLUMN not in df.columns:
        df[DATE_COLUMN] = ""
    else:
        df[DATE_COLUMN] = df[DATE_COLUMN].astype("object")

    return df


def save_excel(df: pd.DataFrame, excel_path: str) -> None:
    path = resolve_path(excel_path)
    try:
        df.to_excel(path, index=False)
    except (PermissionError, OSError) as err:
        raise PermissionError(
            "Cannot update Excel file because it is open in another program. "
            f"Please close the Excel file and run again. Path: {path}"
        ) from err


def build_message(
    sender_email: str,
    receiver_email: str,
    subject: str,
    body: str,
    resume_pdf_path: str,
) -> MIMEMultipart:
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    resolved_resume_path = resolve_path(resume_pdf_path)
    with open(resolved_resume_path, "rb") as file_obj:
        attachment = MIMEApplication(file_obj.read(), _subtype="pdf")
        attachment.add_header(
            "Content-Disposition",
            "attachment",
            filename=resolved_resume_path.name,
        )
        msg.attach(attachment)

    return msg


def send_with_retry(
    smtp_server: smtplib.SMTP_SSL,
    sender_email: str,
    receiver_email: str,
    message: MIMEMultipart,
    max_retries: int,
    logger: Callable[[str], None],
) -> None:
    last_error: Optional[Exception] = None

    for attempt in range(max_retries + 1):
        try:
            smtp_server.sendmail(sender_email, receiver_email, message.as_string())
            return
        except Exception as err:  # pylint: disable=broad-except
            last_error = err
            if attempt < max_retries:
                logger(
                    f"Retry {attempt + 1}/{max_retries} for {receiver_email} after error: {err}"
                )
                time.sleep(2)

    if last_error is not None:
        raise last_error


def send_emails(config: EmailJobConfig, logger: Callable[[str], None]) -> None:
    if config.delay_min < 0 or config.delay_max < 0:
        raise ValueError("Delay values must be non-negative")
    if config.delay_min > config.delay_max:
        raise ValueError("delay-min cannot be greater than delay-max")
    if config.max_retries < 0:
        raise ValueError("max-retries cannot be negative")
    if config.max_emails_per_run < 0:
        raise ValueError("max-emails-per-run cannot be negative")

    resume_path = resolve_path(config.resume_pdf_path)
    if not resume_path.exists():
        raise FileNotFoundError(
            f"Resume file not found: {config.resume_pdf_path} (resolved: {resume_path})"
        )

    df = read_contacts(config.excel_path)
    body_template = load_template(config.template_path)

    smtp_server: Optional[smtplib.SMTP_SSL] = None
    if not config.preview_only:
        context = ssl.create_default_context()
        smtp_server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context)
        smtp_server.login(config.sender_email, config.app_password)
        logger("Authenticated with Gmail SMTP successfully.")
    else:
        logger("Preview mode enabled. No emails will be sent and Excel status will not be updated.")

    sent_count = 0
    failed_count = 0
    skipped_count = 0
    processed_pending_count = 0

    try:
        for index in df.index:
            row = df.loc[index]
            current_status = row.get("Status", "")

            if not should_process_status(current_status):
                skipped_count += 1
                logger(
                    f"SKIPPED Row {index + 2}: Status='{safe_text(current_status)}' "
                    "(already processed)"
                )
                continue

            if (
                config.max_emails_per_run > 0
                and processed_pending_count >= config.max_emails_per_run
            ):
                logger(
                    "Reached configured per-run limit "
                    f"({config.max_emails_per_run}). Stopping this run."
                )
                break

            processed_pending_count += 1

            name = safe_text(row.get("Name")) or "Hiring Manager"
            email = safe_text(row.get("Email"))
            company = safe_text(row.get("Company Name")) or "your company"
            title = safe_text(row.get("Title")) or "the role"
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            subject = render_template(
                config.subject_template,
                name=name,
                company=company,
                title=title,
            )
            body = render_template(
                body_template,
                name=name,
                company=company,
                title=title,
            )

            if not is_valid_email(email):
                logger(f"FAILED Row {index + 2}: Invalid email -> {email}")
                if not config.preview_only:
                    df.loc[index, "Status"] = "Failed"
                    save_excel(df, config.excel_path)
                failed_count += 1
                continue

            if config.preview_only:
                logger(
                    f"PREVIEW Row {index + 2}: To={email} | Subject={subject} | Company={company}"
                )
                continue

            try:
                if smtp_server is None:
                    raise RuntimeError("SMTP server not initialized")

                message_obj = build_message(
                    sender_email=config.sender_email,
                    receiver_email=email,
                    subject=subject,
                    body=body,
                    resume_pdf_path=config.resume_pdf_path,
                )

                send_with_retry(
                    smtp_server=smtp_server,
                    sender_email=config.sender_email,
                    receiver_email=email,
                    message=message_obj,
                    max_retries=config.max_retries,
                    logger=logger,
                )

                df.loc[index, "Status"] = "Sent"
                df.loc[index, DATE_COLUMN] = now_str
                save_excel(df, config.excel_path)

                sent_count += 1
                logger(f"SENT Row {index + 2}: {email}")

            except Exception as send_error:  # pylint: disable=broad-except
                df.loc[index, "Status"] = "Failed"
                save_excel(df, config.excel_path)

                failed_count += 1
                logger(f"FAILED Row {index + 2}: {email} | Error: {send_error}")

            delay_seconds = random.uniform(config.delay_min, config.delay_max)
            logger(f"Sleeping {delay_seconds:.2f} seconds before next email...")
            time.sleep(delay_seconds)

    finally:
        # Final save at the end as a safety checkpoint.
        if not config.preview_only:
            save_excel(df, config.excel_path)
        if smtp_server is not None:
            smtp_server.quit()

    logger(
        "Completed. "
        f"Sent={sent_count}, Failed={failed_count}, Skipped={skipped_count}, "
        f"ProcessedPending={processed_pending_count}."
    )


def create_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Excel-based cold email automation tool (Status-aware)"
    )
    parser.add_argument(
        "--excel",
        default=DEFAULT_EXCEL_PATH,
        help="Path to Excel file (.xlsx). Default: contacts.xlsx",
    )
    parser.add_argument(
        "--resume",
        default=DEFAULT_RESUME_PATH,
        help="Path to PDF resume file. Default: resume.pdf",
    )
    parser.add_argument(
        "--sender",
        default="",
        help="Sender Gmail address",
    )
    parser.add_argument(
        "--app-password",
        dest="app_password",
        default="",
        help="Gmail app password",
    )
    parser.add_argument(
        "--template",
        default=DEFAULT_TEMPLATE_PATH,
        help="Path to email template text file. Default: email_template.txt",
    )
    parser.add_argument(
        "--subject-template",
        default=DEFAULT_SUBJECT_TEMPLATE,
        help="Subject template using placeholders: {name}, {company}, {title}",
    )
    parser.add_argument("--delay-min", type=int, default=5, help="Minimum delay in seconds")
    parser.add_argument("--delay-max", type=int, default=10, help="Maximum delay in seconds")
    parser.add_argument(
        "--max-retries",
        type=int,
        default=1,
        help="Retries for failed sends (default: 1)",
    )
    parser.add_argument(
        "--max-emails-per-run",
        type=int,
        default=0,
        help="Maximum pending rows to process in one run (0 means no limit)",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Preview only. Do not send and do not update Excel status.",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch Tkinter GUI",
    )
    return parser


def run_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("Cold Email Automation")
    root.geometry("900x720")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack(fill="both", expand=True)

    labels = [
        "Excel file",
        "Resume PDF",
        "Sender Gmail",
        "Gmail App Password",
        "Template file (optional)",
        "Subject template",
        "Delay min (sec)",
        "Delay max (sec)",
        "Max retries",
    ]

    defaults = [
        "",
        "",
        "",
        "",
        "",
        DEFAULT_SUBJECT_TEMPLATE,
        "5",
        "10",
        "1",
    ]

    entries = []
    for i, label_text in enumerate(labels):
        tk.Label(frame, text=label_text, anchor="w").grid(row=i, column=0, sticky="w", pady=4)
        entry = tk.Entry(frame, width=72, show="*" if "Password" in label_text else "")
        entry.insert(0, defaults[i])
        entry.grid(row=i, column=1, sticky="we", pady=4)
        entries.append(entry)

    def choose_excel() -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            entries[0].delete(0, tk.END)
            entries[0].insert(0, path)

    def choose_resume() -> None:
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            entries[1].delete(0, tk.END)
            entries[1].insert(0, path)

    def choose_template() -> None:
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            entries[4].delete(0, tk.END)
            entries[4].insert(0, path)

    tk.Button(frame, text="Browse", command=choose_excel).grid(row=0, column=2, padx=6)
    tk.Button(frame, text="Browse", command=choose_resume).grid(row=1, column=2, padx=6)
    tk.Button(frame, text="Browse", command=choose_template).grid(row=4, column=2, padx=6)

    preview_var = tk.BooleanVar(value=False)
    tk.Checkbutton(frame, text="Preview only (no send, no status update)", variable=preview_var).grid(
        row=9, column=1, sticky="w", pady=8
    )

    frame.grid_columnconfigure(1, weight=1)

    log_box = scrolledtext.ScrolledText(frame, width=105, height=22)
    log_box.grid(row=11, column=0, columnspan=3, pady=10, sticky="nsew")
    frame.grid_rowconfigure(11, weight=1)

    def log(message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_box.insert(tk.END, f"[{timestamp}] {message}\n")
        log_box.see(tk.END)
        root.update_idletasks()

    def preview_first_email() -> None:
        try:
            excel_path = entries[0].get().strip()
            template_path = entries[4].get().strip() or None
            subject_template = entries[5].get().strip() or DEFAULT_SUBJECT_TEMPLATE

            df = read_contacts(excel_path)
            processable = df[df["Status"].apply(should_process_status)]
            if processable.empty:
                messagebox.showinfo("Preview", "No rows with empty/Not Sent status.")
                return

            first = processable.iloc[0]
            name = safe_text(first.get("Name")) or "Hiring Manager"
            company = safe_text(first.get("Company Name")) or "your company"
            title = safe_text(first.get("Title")) or "the role"

            body_template = load_template(template_path)
            subject = render_template(
                subject_template,
                name=name,
                company=company,
                title=title,
            )
            body = render_template(
                body_template,
                name=name,
                company=company,
                title=title,
            )

            preview_window = tk.Toplevel(root)
            preview_window.title("Email Preview")
            preview_window.geometry("780x560")

            preview_text = scrolledtext.ScrolledText(preview_window, width=95, height=32)
            preview_text.pack(fill="both", expand=True, padx=10, pady=10)
            preview_text.insert(tk.END, f"Subject: {subject}\n\n{body}")
            preview_text.configure(state="disabled")

        except Exception as err:  # pylint: disable=broad-except
            messagebox.showerror("Preview Error", str(err))

    def start_sending() -> None:
        try:
            config = EmailJobConfig(
                excel_path=entries[0].get().strip(),
                resume_pdf_path=entries[1].get().strip(),
                sender_email=entries[2].get().strip(),
                app_password=entries[3].get().strip(),
                template_path=entries[4].get().strip() or None,
                subject_template=entries[5].get().strip() or DEFAULT_SUBJECT_TEMPLATE,
                delay_min=int(entries[6].get().strip()),
                delay_max=int(entries[7].get().strip()),
                max_retries=int(entries[8].get().strip()),
                max_emails_per_run=0,
                preview_only=preview_var.get(),
            )
        except ValueError:
            messagebox.showerror("Validation Error", "Delay/Retry values must be integers")
            return

        def worker() -> None:
            try:
                send_emails(config=config, logger=log)
                messagebox.showinfo("Completed", "Processing finished. Excel Status column is updated.")
            except Exception as err:  # pylint: disable=broad-except
                messagebox.showerror("Error", str(err))

        threading.Thread(target=worker, daemon=True).start()

    button_row = tk.Frame(frame)
    button_row.grid(row=10, column=0, columnspan=3, pady=5, sticky="w")

    tk.Button(button_row, text="Preview First Pending Email", command=preview_first_email).pack(
        side="left", padx=5
    )
    tk.Button(button_row, text="Start Sending", command=start_sending).pack(side="left", padx=5)

    root.mainloop()


def main() -> None:
    parser = create_arg_parser()
    args = parser.parse_args()

    if args.gui:
        run_gui()
        return

    # Allow running without CLI credentials by prompting in terminal.
    sender_email = args.sender.strip() if args.sender else ""
    app_password = args.app_password.strip() if args.app_password else ""

    if not sender_email:
        sender_email = input("Enter sender Gmail: ").strip()
    if not app_password:
        app_password = input("Enter Gmail App Password: ").strip()

    if not sender_email or not app_password:
        parser.error(
            "Sender and App Password are required. "
            "Provide --sender and --app-password, or enter them when prompted."
        )

    config = EmailJobConfig(
        excel_path=args.excel,
        resume_pdf_path=args.resume,
        sender_email=sender_email,
        app_password=app_password,
        template_path=args.template,
        subject_template=args.subject_template,
        delay_min=args.delay_min,
        delay_max=args.delay_max,
        max_retries=args.max_retries,
        max_emails_per_run=args.max_emails_per_run,
        preview_only=args.preview,
    )

    def cli_log(message: str) -> None:
        print(message)

    send_emails(config=config, logger=cli_log)


if __name__ == "__main__":
    main()
