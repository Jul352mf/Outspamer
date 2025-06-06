import logging
import queue
import threading
import time
from .settings import load
from .outlook import send_with_outlook
from .template_utils import process_template_file, extract_subject_and_body
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

    cc_threshold = int(defaults.get("cc_threshold", 3))

    template_cache: dict[str, tuple] = {}

    def get_template(name: str):
        if name not in template_cache:
            path = Path(paths["templates"]) / name
            if not path.exists():
                return None
            process_template_file(path)
            subj, name_sal, generic_sal, body = extract_subject_and_body(path)
            template_cache[name] = (
                env.from_string(body),
                subj,
                name_sal,
                generic_sal,
            )
        return template_cache[name]

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
            template_data = get_template(tpl_name)
            if not template_data:
                log.error(
                    "Template %s not found for %s; skipping", tpl_name, row["email"]
                )
                q.task_done()
                idx += 1
                continue
            template, tpl_subject, tpl_name_sal, tpl_generic_sal = template_data
            try:
                context = {
                    "vorname": row.get("vorname", ""),
                    "nachname": row.get("nachname", ""),
                    "company": row.get("company", ""),
                    "title": row.get("title", ""),
                    "language": row.get(language_column, ""),
                }

                sal_tpl = tpl_name_sal if row.get("use_named_salutation") else tpl_generic_sal
                if sal_tpl:
                    salutation = env.from_string(sal_tpl).render(**context)
                else:
                    salutation = ""
                context["salutation"] = salutation

                html = template.render(**context)
            except Exception as e:
                log.error("Render error for %s: %s", row["email"], e)
                q.task_done()
                idx += 1
                continue

            send_with_outlook(
                row=row,
                html_body=html,
                subject=tpl_subject or subject_line,
                attachments_dir=paths["attachments"],
                send_time=campaign_start,
                delay_seconds=delay,
                index=idx,
                account_name=account,
                dry_run=dry_run,
                send_now_mode=send_now_mode,
                cc=row.get("cc"),
            )
            if send_now_mode and not dry_run:
                time.sleep(delay)
            q.task_done()
            idx += 1

    t = threading.Thread(target=worker, daemon=True)
    t.start()

    # expand semi-colon separated e-mail cells
    if "email" in leads.columns:
        leads["email_list"] = leads["email"].fillna("").apply(
            lambda v: [e.strip() for e in str(v).split(";") if e.strip()]
        )
    else:
        raise SystemExit("Excel missing 'email' column")

    def first_name(rows) -> str | None:
        for _, r in rows.iterrows():
            v = r.get("vorname")
            if pd.notna(v) and str(v).strip():
                return str(v).strip()
        return None

    if "company" in leads.columns:
        for company, group in leads.groupby("company"):
            all_emails: list[str] = []
            for _, r in group.iterrows():
                all_emails.extend(r["email_list"])
            if not all_emails:
                continue

            fname = first_name(group)
            base = group.iloc[0].to_dict()
            name_row = None
            if fname:
                for _, r in group.iterrows():
                    v = r.get("vorname")
                    if pd.notna(v) and str(v).strip():
                        name_row = r.to_dict()
                        break

            if len(all_emails) > cc_threshold:
                first = True
                for email in all_emails:
                    rec = (name_row if first and name_row else base).copy()
                    rec["email"] = email
                    rec["use_named_salutation"] = bool(fname) and first
                    if not rec.get("vorname"):
                        rec["vorname"] = fname if first and fname else ""
                    q.put(rec)
                    first = False
            else:
                rec = (name_row or base).copy()
                rec["email"] = all_emails[0]
                rec["cc"] = ";".join(all_emails[1:]) if len(all_emails) > 1 else None
                rec["use_named_salutation"] = bool(fname)
                if not rec.get("vorname"):
                    rec["vorname"] = fname or ""
                q.put(rec)
    else:
        for _, row in leads.iterrows():
            emails = row["email_list"]
            if not emails:
                continue

            fname = row.get("vorname")
            fname = str(fname).strip() if pd.notna(fname) and str(fname).strip() else None
            if len(emails) > cc_threshold:
                first = True
                for email in emails:
                    rec = row.to_dict()
                    rec["email"] = email
                    rec["use_named_salutation"] = bool(fname) and first
                    rec["vorname"] = fname if first and fname else ""
                    q.put(rec)
                    first = False
            else:
                rec = row.to_dict()
                rec["email"] = emails[0]
                rec["cc"] = ";".join(emails[1:]) if len(emails) > 1 else None
                rec["use_named_salutation"] = bool(fname)
                rec["vorname"] = fname or ""
                q.put(rec)
    q.put(None)
    q.join()
    t.join()
    log.info("Campaign finished.")
