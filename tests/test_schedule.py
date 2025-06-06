from __future__ import annotations

from pathlib import Path
import sys
import types
from datetime import datetime, timedelta


# ensure project root on path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# stub windows specific modules
sys.modules.setdefault("pythoncom", types.ModuleType("pythoncom"))
sys.modules.setdefault("win32com", types.ModuleType("win32com"))
sys.modules.setdefault("win32com.client", types.ModuleType("client"))


def test_send_with_outlook_schedules(monkeypatch, tmp_path):
    import mailer.outlook as outlook

    send_time = datetime(2023, 1, 1, 12, 0, 0)

    class FakeMail:
        def __init__(self):
            self.DeferredDeliveryTime = None
            self.Attachments = types.SimpleNamespace(Add=lambda path: None)

        def Send(self):
            pass

    fake_mail = FakeMail()

    class FakeOutlook:
        def __init__(self):
            self.Session = types.SimpleNamespace(Accounts=[])

        def CreateItem(self, _):
            return fake_mail

    monkeypatch.setattr(
        sys.modules["pythoncom"], "CoInitialize", lambda: None, raising=False
    )
    monkeypatch.setattr(
        sys.modules["win32com.client"],
        "Dispatch",
        lambda *a, **k: FakeOutlook(),
        raising=False,
    )

    outlook.send_with_outlook(
        row={"email": "a@example.com"},
        html_body="<p>test</p>",
        subject="subj",
        attachments_dir=str(tmp_path),
        send_time=send_time,
        delay_seconds=5,
        index=2,
        dry_run=False,
        send_now_mode=False,
    )

    assert fake_mail.DeferredDeliveryTime == send_time + timedelta(seconds=10)
