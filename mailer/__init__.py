import logging
import queue
import threading
import time
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd
from jinja2 import Environment, FileSystemLoader, StrictUndefined
from zoneinfo import ZoneInfo

from .settings import Config, load
from .outlook import send_with_outlook
from .sendgrid import send_with_sendgrid
from .template_utils import process_template_file, extract_subject_and_body
from .utils import normalize_columns, resolve_leads_path, extract_subject

log = logging.getLogger(__name__)
cfg: Config = load()


def send_campaign(
    *,
    excel_path: str | None = None,
    subject_line: str | None = None,
    template_base: str | None = None,
    sheet_name: str | None = None,
    send_at: str | None = "now",
    account: str | None = None,
    template_column: str | None = None,
    language_column: str | None = None,
    cc_column: str | None = None,
    dry_run: bool = False,
    provider: str = "sendgrid",
):
    paths = cfg.paths
    defaults = cfg.defaults
    template_base = template_base or defaults.template_base
    account = account or (defaults.account or None)
    template_column = (
        template_column if template_column is not None else defaults.template_column
    )
    language_column = language_column or defaults.language_column
    cc_column = cc_column if cc_column is not None else defaults.cc_column

    if provider.lower() == "sendgrid":
        send_func = send_with_sendgrid
    elif provider.lower() == "outlook":
        send_func = send_with_outlook
    else:
        raise ValueError(f"Unknown provider: {provider}")

    xls = resolve_leads_path(cfg, excel_path)
    sheet = sheet_name or defaults.sheet_name

    leads = pd.read_excel(xls, sheet_name=sheet)
    leads = normalize_columns(leads)

    required = {"email", "vorname"}
    missing = required - set(leads.columns)
    if missing:
        raise SystemExit(f"Excel missing columns: {missing}")

    # prepare Jinja2 env
    env = Environment(
        loader=FileSystemLoader(paths.templates),
        undefined=StrictUndefined,
        autoescape=True,
    )

    cc_threshold = int(defaults.cc_threshold)

    template_cache: dict[str, tuple] = {}

    def get_template(name: str):
        if name not in template_cache:
            path = Path(paths.templates) / name
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
    tz = ZoneInfo(defaults.timezone)
    if send_at and send_at != "now":
        start = datetime.fromisoformat(send_at)
        if start.tzinfo is None:
            start = start.replace(tzinfo=tz)
        campaign_start = start.astimezone(tz).replace(tzinfo=None)
        send_now_mode = False
    else:
        campaign_start = None
        send_now_mode = True
    delay = float(defaults.delay_seconds)

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

            lang = (
                str(row.get(language_column, "")).lower().strip()
                if language_column in leads.columns
                else ""
            )

            # determine languages and hour offsets for this row
            if lang == "ch":
                plan = [("de", 0), ("it", 1), ("fr", 2), ("en", 3)]
                follow_up = False
            else:
                plan = [(lang, 0)] if lang else [("", 0)]
                follow_up = lang not in ("en", "de", "")

            first_context = None
            for lang_code, hour_offset in plan:
                if (
                    template_column
                    and template_column in leads.columns
                    and pd.notna(row.get(template_column, None))
                ):
                    tpl_name = str(row[template_column]).strip()
                else:
                    tpl_name = (
                        f"{template_base}_{lang_code}.html" if lang_code else f"{template_base}.html"
                    )
                template_data = get_template(tpl_name)
                lang_used = lang_code
                if not template_data and lang_code != "en":
                    template_data = get_template(f"{template_base}_en.html")
                    lang_used = "en"
                    follow_up = False
                if not template_data:
                    log.error(
                        "Template %s not found for %s; skipping", tpl_name, row["email"]
                    )
                    continue

                template, stored_subject, tpl_name_sal, tpl_generic_sal = template_data
                try:
                    context = {
                        "vorname": row.get("vorname", ""),
                        "nachname": row.get("nachname", ""),
                        "company": row.get("company", ""),
                        "title": row.get("title", ""),
                        "language": lang_used or row.get(language_column, ""),
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
                    continue

                if stored_subject:
                    rendered_subject = env.from_string(stored_subject).render(**context)
                else:
                    rendered_subject = None
                extracted_subject = extract_subject(html)
                final_subject = (
                    extracted_subject
                    or rendered_subject
                    or subject_line
                    or defaults.subject_line
                )
                cc_value = None
                if cc_column:
                    val = row.get(cc_column)
                    if pd.notna(val) and str(val).strip():
                        cc_value = str(val).strip()

                # compute time
                if campaign_start is None:
                    base_time = datetime.now(tz)
                else:
                    base_time = campaign_start
                if hour_offset == 0 and campaign_start is None:
                    send_time = None
                    send_mode = True
                else:
                    send_time = base_time + timedelta(hours=hour_offset)
                    send_mode = False

                send_func(
                    row=row,
                    html_body=html,
                    subject=final_subject,
                    cc=cc_value,
                    attachments_dir=paths.attachments,
                    send_time=send_time,
                    delay_seconds=delay,
                    index=idx,
                    account_name=account,
                    dry_run=dry_run,
                    send_now_mode=send_mode,
                )

                if hour_offset == 0 and first_context is None:
                    first_context = context

            if follow_up:
                tpl_name = f"{template_base}_en.html"
                template_data = get_template(tpl_name)
                if template_data:
                    template, stored_subject, tpl_name_sal, tpl_generic_sal = template_data
                    ctx = (first_context or {}).copy()
                    ctx["language"] = "en"
                    html = template.render(**ctx)
                    if stored_subject:
                        rendered_subject = env.from_string(stored_subject).render(**ctx)
                    else:
                        rendered_subject = None
                    extracted_subject = extract_subject(html)
                    final_subject = (
                        extracted_subject
                        or rendered_subject
                        or subject_line
                        or defaults.subject_line
                    )
                    if campaign_start is None:
                        base_time = datetime.now(tz)
                    else:
                        base_time = campaign_start
                    send_time = base_time + timedelta(hours=1)
                    send_func(
                        row=row,
                        html_body=html,
                        subject=final_subject,
                        cc=cc_value,
                        attachments_dir=paths.attachments,
                        send_time=send_time,
                        delay_seconds=delay,
                        index=idx,
                        account_name=account,
                        dry_run=dry_run,
                        send_now_mode=False,
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
                extra_cc = ";".join(all_emails[1:]) if len(all_emails) > 1 else None
                if extra_cc:
                    if rec.get("cc"):
                        rec["cc"] = f"{rec['cc']};{extra_cc}"
                    else:
                        rec["cc"] = extra_cc
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
                extra_cc = ";".join(emails[1:]) if len(emails) > 1 else None
                if extra_cc:
                    if rec.get("cc"):
                        rec["cc"] = f"{rec['cc']};{extra_cc}"
                    else:
                        rec["cc"] = extra_cc
                # keep existing cc if no additional recipients
                rec["use_named_salutation"] = bool(fname)
                rec["vorname"] = fname or ""
                q.put(rec)
    q.put(None)
    q.join()
    t.join()
    log.info("Campaign finished.")
