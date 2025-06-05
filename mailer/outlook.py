import pythoncom, win32com.client as win32
import logging, pathlib
from datetime import datetime

log = logging.getLogger(__name__)

def _select_account(outlook, account_name: str | None):
    if not account_name:
        return None
    for acct in outlook.Session.Accounts:
        if acct.DisplayName == account_name or acct.SmtpAddress.lower() == account_name.lower():
            return acct
    log.error("Specified Outlook account '%s' not found; default account will be used", account_name)
    return None

def send_with_outlook(
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
):
    """Handles a single e‑mail send.

    Parameters
    ----------
    send_time : datetime | None
        Initial campaign start time (localized).  If provided, each e‑mail
        will be scheduled as `send_time + delay_seconds * index`.
    index : int
        Zero‑based order within campaign.
    send_now_mode : bool
        When True, mail is sent immediately and we **sleep** `delay_seconds`
        afterwards to maintain pacing.
    """
    if dry_run:
        log.info("DRY-RUN  %s <%s>", row.get("vorname", ""), row.get("email", ""))
        return

    pythoncom.CoInitialize()
    outlook = win32.Dispatch("Outlook.Application")
    mail    = outlook.CreateItem(0)

    mail.To = row["email"]
    mail.Subject = subject
    mail.HTMLBody = html_body

    # attachments
    for path in pathlib.Path(attachments_dir).iterdir():
        if path.is_file():
            mail.Attachments.Add(str(path))

    # account selection
    account_obj = _select_account(outlook, account_name)
    if account_obj is not None:
        mail.SendUsingAccount = account_obj

    if send_now_mode or send_time is None:
        # direct send + sleep for pacing
        mail.Send()
        log.info("sent (now) -> %s <%s>", row.get("vorname", ""), row["email"])
        # pace in caller (worker)
        return

    # Schedule deferred time
    schedule_time = send_time + (index * delay_seconds)
    # Outlook expects naive local datetime
    mail.DeferredDeliveryTime = schedule_time
    mail.Send()
    log.info("scheduled %s - %s <%s>", schedule_time.strftime("%Y-%m-%d %H:%M:%S"), row.get("Vorname", ""), row["email"])
