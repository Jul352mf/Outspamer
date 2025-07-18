import logging
import logging.handlers
import sys

import typer
from mailer import send_campaign

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.handlers.RotatingFileHandler(
            "email.log", maxBytes=1_000_000, backupCount=3
        ),
    ],
)

app = typer.Typer(help="Cold outreach mailer – Phase‑0 v4 (all features)")


@app.command()
def run(
    subject: str | None = typer.Option(None, "--subject", "-s", help="E‑mail subject line"),
    leads: str | None = typer.Option(
        None, "--leads", "-l", help="Excel file (defaults leads/ dir)"
    ),
    template_base: str | None = typer.Option(
        None, help="Base name used to build template file like <base>_<lang>.html"
    ),
    sheet: str | None = typer.Option(None, help="Excel sheet name"),
    send_at: str = typer.Option("now", help="'now' or ISO 'YYYY‑MM‑DD HH:MM'"),
    account: str | None = typer.Option(
        None, "--account", "-a", help="Outlook account to send from"
    ),
    cc_column: str = typer.Option("cc", help="Column holding CC addresses"),
    language_column: str = typer.Option(
        "language", help="Column holding language abbreviation (de, en…)"
    ),
    dry_run: bool = typer.Option(False, "--dry-run", "-n", help="Render but not send"),
    provider: str = typer.Option(
        "sendgrid",
        "--provider",
        help="Mail provider ('sendgrid' or 'outlook')",
        show_default=True,
    ),
):
    send_campaign(
        excel_path=leads,
        subject_line=subject,
        template_base=template_base,
        sheet_name=sheet,
        send_at=send_at,
        account=account,
        language_column=language_column,
        cc_column=cc_column,
        dry_run=dry_run,
        provider=provider,
    )


if __name__ == "__main__":
    app()
