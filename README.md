# Outspamer

This project provides a small command line utility for sending personalized email campaigns using Microsoft Outlook. Leads are loaded from an Excel spreadsheet and merged into Jinja2-based HTML templates. Mails can be scheduled or sent immediately and optional attachments are automatically added from the configured folder.

**Important:** This code is intended for controlled outreach and testing. Make sure that any usage complies with local regulations and the terms of your email provider.

## Features

- Reads contact data from an Excel sheet
- Language-specific templates using Jinja2
- Optional send scheduling with configurable delays
- Supports multiple Outlook accounts
- Simple CLI interface powered by [Typer](https://typer.tiangolo.com)
- Optional Tkinter GUI for quick campaigns

## Requirements

- Python 3.11 or newer
- Microsoft Outlook installed on Windows (for the COM interface)
- See `requirements.txt` for Python package dependencies

## Installation

1. Create and activate a virtual environment (optional but recommended):

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Windows: .venv\Scripts\activate
   ```

2. Install dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

## Configuration

The application reads basic paths and defaults from `settings.toml`. You can adjust directories for attachments, templates and leads as well as the default delay between mails. Environment variables with the same names can override the TOML settings.

```
[paths]
attachments = "attachments"
templates   = "templates"
leads       = "leads"

[defaults]
delay_seconds     = 2.5
sheet_name        = "Sheet1"
timezone          = "Europe/Zurich"
default_leads_file = "sample_leads.xlsx"
template_base     = "email"
```

## Usage

The main entry point is `send_emails.py` which exposes a `run` command via Typer.

```bash
python send_emails.py run --subject "Subject line" \
    --leads path/to/leads.xlsx \
    --template-base email \
    --sheet Sheet1
```

Use `--help` for a complete list of options. For testing you can add `--dry-run` to render mails without sending them.

### GUI Application

For a more user-friendly option run the Tkinter based GUI:

```bash
python send_emails_gui.py
```

The GUI lets you browse for the leads Excel file, fill in the campaign details
and start sending in one click while showing live log output.

## Example Templates

HTML templates for different languages can be found in the `templates/` directory. Place any files you want attached to every email in the `attachments/` directory.

## Notes

An `Archive` folder contains older experiments and is not required for normal operation. Some log output is stored in `email.log` after running the script.

