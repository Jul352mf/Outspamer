import base64
import logging
import os
from datetime import datetime, timedelta
from pathlib import Path
from python_http_client.exceptions import HTTPError
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Attachment,
    Disposition,
    FileContent,
    FileName,
    FileType,
    Mail,
)

log = logging.getLogger(__name__)


def send_with_sendgrid(
    *,
    row,
    html_body: str,
    subject: str,
    attachments_dir: str,
    send_time: datetime | None,
    delay_seconds: float,
    index: int,
    account_name: str | None = None,
    api_key: str | None = None,
    dry_run: bool = False,
    send_now_mode: bool = False,
    cc: str | None = None,
) -> None:
    """Send a single email using SendGrid."""
    if dry_run:
        log.info(
            "DRY-RUN sendgrid %s <%s> cc=%s",
            row.get("vorname", ""),
            row.get("email", ""),
            cc or "",
        )
        return

    api_key = api_key or os.getenv("SENDGRID_API_KEY")
    if not api_key:
        log.error("No SendGrid API key provided")
        return

    message = Mail(
        from_email=account_name or os.getenv("SENDGRID_FROM_EMAIL", ""),
        to_emails=row["email"],
        subject=subject,
        html_content=html_body,
    )

    if cc:
        message.cc = [email.strip() for email in cc.split(";") if email.strip()]

    attach_path = Path(attachments_dir)
    if attach_path.exists():
        for path in attach_path.iterdir():
            if path.is_file():
                data = path.read_bytes()
                encoded = base64.b64encode(data).decode()
                attachment = Attachment(
                    FileContent(encoded),
                    FileName(path.name),
                    FileType("application/octet-stream"),
                    Disposition("attachment"),
                )
                message.add_attachment(attachment)
    else:
        log.warning("Attachments directory '%s' not found", attachments_dir)

    schedule_time = None
    if not send_now_mode and send_time is not None:
        schedule_time = send_time + timedelta(seconds=index * delay_seconds)
        message.send_at = int(schedule_time.timestamp())

    try:
        sg = SendGridAPIClient(api_key)
        sg.send(message)
        if schedule_time:
            log.info(
                "scheduled %s - %s <%s>",
                schedule_time.strftime("%Y-%m-%d %H:%M:%S"),
                row.get("vorname", ""),
                row["email"],
            )
        else:
            log.info("sent (now) -> %s <%s>", row.get("vorname", ""), row["email"])
    except HTTPError as exc:
        log.error("SendGrid HTTP %s: %s", exc.status_code, exc.body.decode(), row.get("vorname", ""), row["email"])
    except Exception as exc:
        log.error("SendGrid error: %s", exc, row.get("vorname", ""), row["email"])

