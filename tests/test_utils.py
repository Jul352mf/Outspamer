from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd


import types

# stub win32 modules so mailer can be imported on non-Windows systems
sys.modules.setdefault("pythoncom", types.ModuleType("pythoncom"))
win32com = types.ModuleType("win32com")
win32com.client = types.ModuleType("client")
win32com.client.Dispatch = lambda *a, **k: None
sys.modules.setdefault("win32com", win32com)
sys.modules.setdefault("win32com.client", win32com.client)

import mailer
from mailer import _normalize_columns, _resolve_leads_path, send_campaign


def test_normalize_columns():
    df = pd.DataFrame(columns=["Email Adresse", "Vorname", "Firma"])
    out = _normalize_columns(df)
    assert list(out.columns) == ["email", "vorname", "company"]


def test_resolve_leads_path(tmp_path, monkeypatch):
    leads_dir = tmp_path / "leads"
    leads_dir.mkdir()
    file = leads_dir / "foo.xlsx"
    file.touch()

    monkeypatch.setitem(mailer.cfg["paths"], "leads", str(leads_dir))
    monkeypatch.setitem(mailer.cfg["defaults"], "default_leads_file", "foo.xlsx")

    assert _resolve_leads_path(None) == file
    assert _resolve_leads_path("foo.xlsx") == file


def test_send_campaign_calls_outlook(monkeypatch, tmp_path):
    xls = tmp_path / "l.xlsx"
    df = pd.DataFrame({"email": ["test@example.com"], "vorname": ["Foo"]})
    df.to_excel(xls, index=False)

    calls = []

    def fake_send_with_outlook(**kwargs):
        calls.append(kwargs["row"])

    monkeypatch.setattr(mailer, "send_with_outlook", fake_send_with_outlook)
    monkeypatch.setitem(mailer.cfg["defaults"], "default_leads_file", str(xls))

    # create a minimal template expected by send_campaign
    tpl_dir = tmp_path / "tpl"
    tpl_dir.mkdir()
    (tpl_dir / "email.html").write_text("hello {{ vorname }}")
    monkeypatch.setitem(mailer.cfg["paths"], "templates", str(tpl_dir))

    send_campaign(excel_path=str(xls), subject_line="Hi", dry_run=True)
    assert len(calls) == 1
