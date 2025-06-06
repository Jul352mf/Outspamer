# Outspamer

This project provides a small command line utility for sending personalized email campaigns using Microsoft Outlook. Leads are loaded from an Excel spreadsheet and merged into Jinja2-based HTML templates. Mails can be scheduled or sent immediately and optional attachments are automatically added from the configured folder.

**Important:** This code is intended for controlled outreach and testing. Make sure that any usage complies with local regulations and the terms of your email provider.

## Features

- Reads contact data from an Excel sheet
- Language-specific templates using Jinja2
- Optional send scheduling with configurable delays
- Supports multiple Outlook accounts
- Simple CLI interface powered by [Typer](https://typer.tiangolo.com)

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

2. Install dependencies and the package itself:

   ```bash
   python -m pip install -e .
   ```

3. Create the working directories for your own leads and any attachments. These
   folders are ignored by git and therefore not present after cloning:

   ```bash
   mkdir -p leads attachments templates
   ```

## Configuration

The application reads basic paths and defaults from `settings.toml`. You can adjust directories for attachments, templates and leads as well as the default delay between mails. Environment variables with the same names can override the TOML settings.
The bundled `sample_leads.xlsx` file contains **fictitious** contact details for testing only.

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
account           = ""
template_column   = "template"
language_column   = "language"
cc_column         = "cc"
```

The optional `account` field selects a specific Outlook account by display name
or SMTP address. `template_column` allows specifying per-row template names. If
absent, the template is built from `template_base` and the value in
`language_column`. `cc_column` can hold additional recipients; if omitted,
addresses separated by semicolons in the `email` column will be used.


## Usage

After installation the `outspamer` command becomes available:

```bash
outspamer --subject "Subject line" \
    --leads path/to/leads.xlsx \
    --template-base email \
    --sheet Sheet1
```

Use `--help` for a complete list of options. For testing you can add `--dry-run` to render mails without sending them.

### Web Interface

You can also control the campaign through a small Flask based web app. Run

```bash
python webapp.py
```

and open `http://localhost:5000` in your browser. The form lets you upload the
Excel leads file, pick templates and other options with a more friendly UI.

## Testing

Run `pytest` to execute the unit tests:

```bash
pytest
```

## Example Templates

HTML templates for different languages can be found in the `templates/` directory. Place any files you want attached to every email in the `attachments/` directory.

## Notes

An `Archive` folder contains older experiments and is not required for normal operation. Real leads and attachments are not included in the repository, so create `leads/` and `attachments/` yourself. When the tool runs it logs activity to `email.log`, which is ignored by git.


## License

This project is licensed under the [MIT License](LICENSE).
