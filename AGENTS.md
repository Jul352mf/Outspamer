loaded from an Excel spreadsheet and merged into HTML templates using Jinja2.
The CLI is provided by `send_emails.py` and installed as the `outspamer`
script.

## Layout

- `mailer/` – core package with campaign logic and Outlook integration.
- `templates/` – example HTML templates.
- `settings.toml` – default paths and options. Environment variables with the
  same keys override the values here.
- `tests/` – unit tests using `pytest`. Windows specific modules are stubbed so
  tests run on any platform.
- `Archive/` – old experimental code; not needed for normal development.

## Development

Use **Python 3.11** or newer. After cloning, install dependencies and the
package in editable mode:

```bash
python -m pip install -r requirements.txt
python -m pip install -e .
```

### Style

Code follows `flake8` rules configured in `.flake8` (maximum line length 88 and
ignoring `E501`, `W503` and `E402`). Run `flake8` before committing to ensure
styling is correct.

### Testing

Run unit tests with `pytest`.

```bash
flake8
pytest
```

All changes should keep the tests passing.

## Notes

Real leads and attachments are not included in the repository. Create `leads/`
and `attachments/` folders when using the tool. Logs are written to
`email.log`, which is ignored by git.