from __future__ import annotations
from pathlib import Path
from bs4 import BeautifulSoup

PROCESSED_MARKER = "<!-- processed -->"


def process_template_file(path: Path) -> None:
    """Rewrite exported Notion HTML into a clean template.

    Runs only once per file (flagged with PROCESSED_MARKER).
    Removes <title> and the Notion <header> block.
    """
    text = path.read_text(encoding="utf-8")
    if PROCESSED_MARKER in text:
        return

    soup = BeautifulSoup(text, "html.parser")

    # 1  Kill the <title> element (browser tab title)
    if soup.title:
        soup.title.decompose()

    # 2  Kill Notion’s header (page title + description)
    header = soup.find("header")
    if header:
        header.decompose()
    # -or-, if you want to be extra safe:
    for tag in soup.select("header, h1.page-title, p.page-description"):
        tag.decompose()

    html_text = str(soup)

    # Insert the marker right after <head>, so Outlook stays happy
    marker = f"\n    {PROCESSED_MARKER}\n"
    if "<head>" in html_text:
        html_text = html_text.replace("<head>", f"<head>{marker}", 1)
    else:
        # Fallback: just append it at the very end
        html_text += marker

    path.write_text(html_text, encoding="utf-8")


def extract_subject_and_body(path: Path) -> tuple[str | None, str | None, str | None, str]:
    """Return subject, name and generic salutations, and the cleaned HTML body."""
    text = path.read_text(encoding="utf-8")
    soup = BeautifulSoup(text, "html.parser")

    subject = None
    name_sal = None
    generic_sal = None

    # Look at actual text-bearing elements only; skip wrappers
    for el in list(soup.find_all(["p", "span", "h1", "h2", "h3", "div"])):
        if el.string is None:        # has child tags → wrapper div, ignore
            continue

        content = el.string.strip()
        if not content:
            continue

        if content.startswith("Subject:") and subject is None:
            subject = content[len("Subject:"):].strip()
            el.decompose()
        elif content.startswith("Name Salutation:") and name_sal is None:
            name_sal = content[len("Name Salutation:"):].strip()
            el.decompose()
        elif content.startswith("No name Salutation:") and generic_sal is None:
            generic_sal = content[len("No name Salutation:"):].strip()
            el.decompose()

    body = str(soup)
    return subject, name_sal, generic_sal, body
