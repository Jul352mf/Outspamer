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
from mailer import send_campaign
from mailer.utils import normalize_columns, resolve_leads_path


def test_normalize_columns():
    df = pd.DataFrame(columns=["Email Adresse", "Vorname", "Firma"])
    out = normalize_columns(df)
    assert list(out.columns) == ["email", "vorname", "company"]


def test_resolve_leads_path(tmp_path, monkeypatch):
    leads_dir = tmp_path / "leads"
    leads_dir.mkdir()
    file = leads_dir / "foo.xlsx"
    file.touch()

    monkeypatch.setattr(mailer.cfg.paths, "leads", str(leads_dir))
    monkeypatch.setattr(mailer.cfg.defaults, "default_leads_file", "foo.xlsx")

    assert resolve_leads_path(mailer.cfg, None) == file
    assert resolve_leads_path(mailer.cfg, "foo.xlsx") == file


def test_send_campaign_calls_outlook(monkeypatch, tmp_path):
    xls = tmp_path / "l.xlsx"
    df = pd.DataFrame({"email": ["test@example.com"], "vorname": ["Foo"]})
    df.to_excel(xls, index=False)

    calls = []

    def fake_send_with_sendgrid(**kwargs):
        calls.append(kwargs)

    monkeypatch.setattr(mailer, "send_with_sendgrid", fake_send_with_sendgrid)
    monkeypatch.setattr(mailer.cfg.defaults, "default_leads_file", str(xls))

    # create a minimal template expected by send_campaign
    tpl_dir = tmp_path / "tpl"
    tpl_dir.mkdir()
    (tpl_dir / "email.html").write_text("hello {{ salutation }}")
    monkeypatch.setattr(mailer.cfg.paths, "templates", str(tpl_dir))
    monkeypatch.setattr(mailer.cfg.defaults, "template_base", "email")

    monkeypatch.setattr(mailer.cfg.defaults, "subject_line", "Default")

    send_campaign(excel_path=str(xls), subject_line=None, dry_run=True)
    assert len(calls) == 1


def test_template_subject_overrides(monkeypatch, tmp_path):
    xls = tmp_path / "l.xlsx"
    df = pd.DataFrame({"email": ["a@example.com"], "vorname": ["Foo"], "cc": ["c@example.com"]})
    df.to_excel(xls, index=False)

    captured = {}

    def fake_send_with_sendgrid(**kwargs):
        captured.update(kwargs)

    monkeypatch.setattr(mailer, "send_with_sendgrid", fake_send_with_sendgrid)
    monkeypatch.setattr(mailer.cfg.defaults, "default_leads_file", str(xls))
    monkeypatch.setattr(mailer.cfg.defaults, "subject_line", "Default")
    monkeypatch.setattr(mailer.cfg.defaults, "template_base", "email")

    tpl_dir = tmp_path / "tpl2"
    tpl_dir.mkdir()
    (tpl_dir / "email.html").write_text("<p>Subject: Hello {{ vorname }}</p>")
    monkeypatch.setattr(mailer.cfg.paths, "templates", str(tpl_dir))
    monkeypatch.setattr(mailer.cfg.defaults, "template_base", "email")

    send_campaign(excel_path=str(xls), subject_line="CLI", dry_run=True)
    assert captured["subject"] == "Hello Foo"
    assert captured["cc"] == "c@example.com"


def test_cc_threshold(monkeypatch, tmp_path):
    xls = tmp_path / "l2.xlsx"
    df = pd.DataFrame({
        "email": ["a@example.com;b@example.com"],
        "vorname": ["A"],
        "company": ["Foo"],
    })
    df.to_excel(xls, index=False)

    calls = []

    def fake_send_with_sendgrid(**kwargs):
        calls.append(kwargs)

    monkeypatch.setattr(mailer, "send_with_sendgrid", fake_send_with_sendgrid)
    monkeypatch.setattr(mailer.cfg.defaults, "default_leads_file", str(xls))
    monkeypatch.setattr(mailer.cfg.defaults, "cc_threshold", 3)

    tpl_dir = tmp_path / "tpl2"
    tpl_dir.mkdir()
    (tpl_dir / "email.html").write_text("hello {{ salutation }}")
    monkeypatch.setattr(mailer.cfg.paths, "templates", str(tpl_dir))
    monkeypatch.setattr(mailer.cfg.defaults, "template_base", "email")

    send_campaign(excel_path=str(xls), subject_line="Hi", dry_run=True)

    assert len(calls) == 1
    assert calls[0]["cc"] == "b@example.com"


def test_salutation_extraction(tmp_path):
    tpl = tmp_path / "t.html"
    tpl.write_text(
        "<p>Subject: Hello</p><p>No name Salutation: Hi {{company}} team</p>"
        "<p>Name Salutation: Hi {{vorname}}</p><div>Body</div>"
    )
    subj, name_sal, generic_sal, body = mailer.template_utils.extract_subject_and_body(tpl)
    assert subj == "Hello"
    assert name_sal == "Hi {{vorname}}"
    assert generic_sal == "Hi {{company}} team"
    assert "Subject:" not in body


def test_personalize_first_only(monkeypatch, tmp_path):
    xls = tmp_path / "p.xlsx"
    df = pd.DataFrame({
        "email": ["a@example.com;b@example.com"],
        "vorname": ["Alice"],
        "company": ["Foo"],
    })
    df.to_excel(xls, index=False)

    calls = []

    def fake_send_with_sendgrid(**kwargs):
        calls.append(kwargs)

    monkeypatch.setattr(mailer, "send_with_sendgrid", fake_send_with_sendgrid)
    monkeypatch.setattr(mailer.cfg.defaults, "default_leads_file", str(xls))
    monkeypatch.setattr(mailer.cfg.defaults, "cc_threshold", 1)

    tpl_dir = tmp_path / "tpl3"
    tpl_dir.mkdir()
    (tpl_dir / "email.html").write_text(
        "<p>Subject: X</p><p>No name Salutation: Hi {{company}} team</p>"
        "<p>Name Salutation: Hi {{vorname}}</p><div>{{ salutation }}</div>"
    )
    monkeypatch.setattr(mailer.cfg.paths, "templates", str(tpl_dir))
    monkeypatch.setattr(mailer.cfg.defaults, "template_base", "email")

    send_campaign(excel_path=str(xls), subject_line="Hi", dry_run=True)

    assert len(calls) == 2
    assert "Hi Alice" in calls[0]["html_body"]
    assert "Hi Foo team" in calls[1]["html_body"]


def test_language_fallback(monkeypatch, tmp_path):
    xls = tmp_path / "l.xlsx"
    df = pd.DataFrame({"email": ["a@example.com"], "vorname": ["Foo"], "language": ["es"]})
    df.to_excel(xls, index=False)

    calls = []

    def fake_send_with_sendgrid(**kwargs):
        calls.append(kwargs)

    monkeypatch.setattr(mailer, "send_with_sendgrid", fake_send_with_sendgrid)
    monkeypatch.setattr(mailer.cfg.defaults, "default_leads_file", str(xls))

    tpl_dir = tmp_path / "tpl_fb"
    tpl_dir.mkdir()
    (tpl_dir / "email_en.html").write_text("Hello")
    monkeypatch.setattr(mailer.cfg.paths, "templates", str(tpl_dir))
    monkeypatch.setattr(mailer.cfg.defaults, "template_base", "email")

    send_campaign(excel_path=str(xls), subject_line="Hi", dry_run=True)

    assert len(calls) == 1
    assert "Hello" in calls[0]["html_body"]


def test_follow_up_english(monkeypatch, tmp_path):
    xls = tmp_path / "l2.xlsx"
    df = pd.DataFrame({"email": ["a@example.com"], "vorname": ["Foo"], "language": ["fr"]})
    df.to_excel(xls, index=False)

    calls = []

    def fake_send_with_sendgrid(**kwargs):
        calls.append(kwargs)

    monkeypatch.setattr(mailer, "send_with_sendgrid", fake_send_with_sendgrid)
    monkeypatch.setattr(mailer.cfg.defaults, "default_leads_file", str(xls))

    tpl_dir = tmp_path / "tpl_follow"
    tpl_dir.mkdir()
    (tpl_dir / "email_fr.html").write_text("FR")
    (tpl_dir / "email_en.html").write_text("EN")
    monkeypatch.setattr(mailer.cfg.paths, "templates", str(tpl_dir))
    monkeypatch.setattr(mailer.cfg.defaults, "template_base", "email")

    start = "2023-01-01 10:00"
    send_campaign(excel_path=str(xls), subject_line="Hi", send_at=start, dry_run=True)

    assert len(calls) == 2
    assert calls[0]["send_time"].strftime("%Y-%m-%d %H:%M") == start
    assert calls[1]["send_time"].strftime("%Y-%m-%d %H:%M") == "2023-01-01 11:00"


def test_ch_multilang(monkeypatch, tmp_path):
    xls = tmp_path / "l3.xlsx"
    df = pd.DataFrame({"email": ["a@example.com"], "vorname": ["Foo"], "language": ["CH"]})
    df.to_excel(xls, index=False)

    calls = []

    def fake_send_with_sendgrid(**kwargs):
        calls.append(kwargs)

    monkeypatch.setattr(mailer, "send_with_sendgrid", fake_send_with_sendgrid)
    monkeypatch.setattr(mailer.cfg.defaults, "default_leads_file", str(xls))

    tpl_dir = tmp_path / "tpl_ch"
    tpl_dir.mkdir()
    (tpl_dir / "email_de.html").write_text("DE")
    (tpl_dir / "email_it.html").write_text("IT")
    (tpl_dir / "email_fr.html").write_text("FR")
    (tpl_dir / "email_en.html").write_text("EN")
    monkeypatch.setattr(mailer.cfg.paths, "templates", str(tpl_dir))
    monkeypatch.setattr(mailer.cfg.defaults, "template_base", "email")

    start = "2023-01-01 10:00"
    send_campaign(excel_path=str(xls), subject_line="Hi", send_at=start, dry_run=True)

    assert len(calls) == 4
    times = [c["send_time"].strftime("%Y-%m-%d %H:%M") for c in calls]
    assert times == ["2023-01-01 10:00", "2023-01-01 11:00", "2023-01-01 12:00", "2023-01-01 13:00"]
