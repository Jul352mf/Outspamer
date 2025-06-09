import os
import logging
import pathlib
import base64
import mimetypes
from datetime import datetime, timedelta

from postmarker.core import PostmarkClient

log = logging.getLogger(__name__)


def _collect_attachments(directory: str) -> list[dict]:
    path = pathlib.Path(directory)
    if not path.exists():
        log.warning("Attachments directory '%s' not found", directory)
        return []
    items = []
    for item in path.iterdir():
        if item.is_file():
            with item.open("rb") as f:
                content = base64.b64encode(f.read()).decode()
            items.append(
                {
                    "Name": item.name,
                    "Content": content,
                    "ContentType": mimetypes.guess_type(item.name)[0]
                    or "application/octet-stream",
                }
            )
    return items


def send_with_postmark(
    *,
    row,
    html_body: str,
    subject: str,
    attachments_dir: str,
    send_time: datetime | None,
    delay_seconds: float,
    index: int,
    account_name: str | None = None,
    dry_run: bool = False,
    send_now_mode: bool = False,
    cc: str | None = None,
    token: str | None = None,
) -> None:
    """Send a single e-mail using Postmark."""
    if dry_run:
        log.info(
            "DRY-RUN  %s <%s> cc=%s",
            row.get("vorname", ""),
            row.get("email", ""),
            cc or "",
        )
        return

    token = token or os.getenv("POSTMARK_TOKEN")
    if not token:
        log.error("Postmark token not configured")
        return



    data = {
        "From": account_name or os.getenv("POSTMARK_FROM", ""),
        "To": row["email"],
        "Subject": subject,
        "HtmlBody": html_body,
    }
    if cc:
        data["Cc"] = cc
        
    if send_time is not None and not send_now_mode:
        schedule = send_time + timedelta(seconds=index * delay_seconds)
        # Postmark expects an RFC3339 string; keep the timezone info
        data["DeliverAt"] = schedule.isoformat()
        log.info("Scheduled delivery at %s (Zurich time)", schedule.isoformat())

    attachments = _collect_attachments(attachments_dir)
    if attachments:
        data["Attachments"] = attachments

    client = PostmarkClient(server_token=token)

    try:
        client.emails.send(**data)
        log.info("sent -> %s <%s>", row.get("vorname", ""), row["email"])
    except Exception as exc:
        log.error("Postmark send error: %s", exc)
