from pathlib import Path
import os
import tomllib

ROOT = Path(__file__).resolve().parent.parent
TOML = ROOT / "settings.toml"


def load():
    cfg = {
        "paths": {
            "attachments": str(ROOT / "attachments"),
            "templates": str(ROOT / "templates"),
            "leads": str(ROOT / "leads"),
        },
        "defaults": {
            "delay_seconds": 2.5,
            "sheet_name": "Sheet1",
            "timezone": "Europe/Zurich",
            "default_leads_file": "",
            "template_base": "email",
            "cc_threshold": 3,
            "subject_line": "default subject",
            "account": "",
            "template_column": "template",
            "language_column": "language",
            "cc_column": "cc",
        },
    }
    if TOML.exists():
        data = tomllib.loads(TOML.read_text())
        for section in ("paths", "defaults"):
            if section in data:
                cfg[section].update(data[section])
    # env overrides
    for envvar, pathkey in [
        ("ATTACHMENTS_DIR", "attachments"),
        ("TEMPLATES_DIR", "templates"),
        ("LEADS_DIR", "leads"),
    ]:
        if os.getenv(envvar):
            cfg["paths"][pathkey] = os.getenv(envvar)
    for envvar, key in [
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
            cfg["defaults"][key] = os.getenv(envvar)
    return cfg
