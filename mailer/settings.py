from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os
import tomllib
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
TOML = ROOT / "settings.toml"
load_dotenv(ROOT / ".env", override=False)


@dataclass
class Paths:
    attachments: str
    templates: str
    leads: str


@dataclass
class Defaults:
    delay_seconds: float = 2.5
    sheet_name: str = "Sheet1"
    timezone: str = "Europe/Zurich"
    default_leads_file: str = ""
    template_base: str = "email"
    cc_threshold: int = 3
    subject_line: str = "default subject"
    account: str = ""
    template_column: str = "template"
    language_column: str = "language"
    cc_column: str = "cc"


@dataclass
class Config:
    paths: Paths
    defaults: Defaults


def load() -> Config:
    paths = Paths(
        attachments=str(ROOT / "attachments"),
        templates=str(ROOT / "templates"),
        leads=str(ROOT / "leads"),
    )

    defaults = Defaults()

    if TOML.exists():
        data = tomllib.loads(TOML.read_text())
        if "paths" in data:
            for key, value in data["paths"].items():
                setattr(paths, key, value)
        if "defaults" in data:
            for key, value in data["defaults"].items():
                setattr(defaults, key, value)

    for envvar, attr in [
        ("ATTACHMENTS_DIR", "attachments"),
        ("TEMPLATES_DIR", "templates"),
        ("LEADS_DIR", "leads"),
    ]:
        if os.getenv(envvar):
            setattr(paths, attr, os.getenv(envvar))

    for envvar, attr in [
        ("MAILER_DELAY", "delay_seconds"),
        ("MAILER_SHEET", "sheet_name"),
        ("MAILER_TZ", "timezone"),
        ("DEFAULT_LEADS_FILE", "default_leads_file"),
        ("TEMPLATE_BASE", "template_base"),
        ("CC_THRESHOLD", "cc_threshold"),
        ("SUBJECT_LINE", "subject_line"),
        ("DEFAULT_ACCOUNT", "account"),
        ("TEMPLATE_COLUMN", "template_column"),
        ("LANGUAGE_COLUMN", "language_column"),
        ("CC_COLUMN", "cc_column"),
    ]:
        if os.getenv(envvar):
            setattr(defaults, attr, os.getenv(envvar))

    return Config(paths=paths, defaults=defaults)
