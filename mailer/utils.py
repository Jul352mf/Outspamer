from __future__ import annotations

from pathlib import Path
import re

import pandas as pd

from .settings import Config

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
    "cc": "cc",
}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    new_cols = []
    for col in df.columns:
        key = col.strip().lower()
        key = _COLUMN_MAP.get(key, key)
        new_cols.append(key)
    df.columns = new_cols
    return df


def resolve_leads_path(cfg: Config, user_path: str | Path | None) -> Path:
    default_path = cfg.defaults.default_leads_file
    if user_path is None or user_path == "":
        if not default_path:
            raise FileNotFoundError("No leads file provided and no default set.")
        user_path = default_path
    p = Path(user_path)
    if p.is_absolute():
        if p.exists():
            return p
        raise FileNotFoundError(p)
    if p.exists():
        return p.resolve()
    candidate = Path(cfg.paths.leads) / p
    if candidate.exists():
        return candidate.resolve()
    raise FileNotFoundError(
        f"{user_path} (looked in cwd and {cfg.paths.leads})"
    )


def extract_subject(html: str) -> str | None:
    match = re.search(r"<title>(.*?)</title>", html, re.IGNORECASE | re.DOTALL)
    if match:
        return match.group(1).strip()
    match = re.search(r"Subject:\s*(.*?)(?:<|\n)", html, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None
