import logging
import queue
import threading
import time
from .settings import load
from .outlook import send_with_outlook
import pandas as pd
from pathlib import Path
from datetime import datetime
from jinja2 import Environment, FileSystemLoader, StrictUndefined
from zoneinfo import ZoneInfo

log = logging.getLogger(__name__)
cfg = load()

# synonyms mapping for column names
_COLUMN_MAP = {
    "email": "email",
    "email adresse": "email",
    "vorname": "vorname",
    "nachname": "nachname",
    "company": "company",
    "firma": "company",
    "title": "title",
    "sprache": "language",
    "language": "language",
    "template": "template",
}


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    new_cols = []
    for col in df.columns:
        key = col.strip().lower()
        key = _COLUMN_MAP.get(key, key)
        new_cols.append(key)
    df.columns = new_cols
    return df


def _resolve_leads_path(user_path: str | Path | None) -> Path:
    default_path = cfg["defaults"]["default_leads_file"]
    if user_path is None or user_path == "":
        if not default_path:
            raise FileNotFoundError("No leads file provided and no default set.")
        user_path = default_path
    p = Path(user_path)
    if p.is_absolute():
        if p.exists():
            return p
        raise FileNotFoundError(p)
    # relative
    if p.exists():
        return p.resolve()
    candidate = Path(cfg["paths"]["leads"]) / p
    if candidate.exists():
        return candidate.resolve()
    raise FileNotFoundError(f"{user_path} (looked in cwd and {cfg['paths']['leads']})")


def send_campaign(
    *,
    excel_path: str | None = None,
    subject_line: str,
    template_base: str | None = None,
    sheet_name: str | None = None,
    send_at: str | None = "now",
    account: str | None = None,
    template_column: str | None = "template",
    language_column: str = "language",
    dry_run: bool = False,
):
    paths = cfg["paths"]
    defaults = cfg["defaults"]
    template_base = template_base or defaults["template_base"]

    xls = _resolve_leads_path(excel_path)
    sheet = sheet_name or defaults["sheet_name"]

    leads = pd.read_excel(xls, sheet_name=sheet)
    leads = _normalize_columns(leads)

    required = {"email", "vorname"}
    missing = required - set(leads.columns)
    if missing:
        raise SystemExit(f"Excel missing columns: {missing}")

    # prepare Jinja2 env
    env = Environment(
        loader=FileSystemLoader(paths["templates"]),
        undefined=StrictUndefined,
        autoescape=True,
    )

    # time & pacing
    tz = ZoneInfo(defaults["timezone"])
    if send_at and send_at != "now":
        start = datetime.fromisoformat(send_at)
        if start.tzinfo is None:
            start = start.replace(tzinfo=tz)
        campaign_start = start.astimezone(tz).replace(tzinfo=None)
        send_now_mode = False
    else:
        campaign_start = None
        send_now_mode = True
    delay = float(defaults["delay_seconds"])

    # worker with queue
    q = queue.Queue()

    def worker():
        idx = 0
        while True:
            item = q.get()
            if item is None:
                q.task_done()
                break
            row = item
            # pick template name
            tpl_name = None
            if (
                template_column
                and template_column in leads.columns
                and pd.notna(row.get(template_column, None))
            ):
                tpl_name = str(row[template_column]).strip()
            else:
                lang = (
                    str(row.get(language_column, "")).lower().strip()
                    if language_column in leads.columns
                    else ""
                )
                tpl_name = (
                    f"{template_base}_{lang}.html" if lang else f"{template_base}.html"
                )
            tpl_path = Path(paths["templates"]) / tpl_name
            if not tpl_path.exists():
                log.error(
                    "Template %s not found for %s; skipping", tpl_name, row["email"]
                )
                q.task_done()
                idx += 1
                continue
            template = env.get_template(tpl_name)
            try:
                html = template.render(
                    vorname=row["vorname"],
                    nachname=row.get("nachname", ""),
                    company=row.get("company", ""),
                    title=row.get("title", ""),
                    language=row.get(language_column, ""),
                )
            except Exception as e:
                log.error("Render error for %s: %s", row["email"], e)
                q.task_done()
                idx += 1
                continue

            send_with_outlook(
                row=row,
                html_body=html,
                subject=subject_line,
                attachments_dir=paths["attachments"],
                send_time=campaign_start,
                delay_seconds=delay,
                index=idx,
                account_name=account,
                dry_run=dry_run,
                send_now_mode=send_now_mode,
            )
            if send_now_mode and not dry_run:
                time.sleep(delay)
            q.task_done()
            idx += 1

    t = threading.Thread(target=worker, daemon=True)
    t.start()

    for _, row in leads.iterrows():
        q.put(row.copy())
    q.put(None)
    q.join()
    t.join()
    log.info("Campaign finished.")
