from __future__ import annotations
from pathlib import Path
from bs4 import BeautifulSoup

PROCESSED_MARKER = "<!-- processed -->"


def process_template_file(path: Path) -> None:
    """Rewrite exported Notion HTML into a clean template.

    The processed file is written back only once, identified by the
    ``PROCESSED_MARKER`` comment at the top.
    """
    text = path.read_text(encoding="utf-8")
    if PROCESSED_MARKER in text:
        return
    soup = BeautifulSoup(text, "html.parser")
    if soup.title:
        soup.title.decompose()
    cleaned = PROCESSED_MARKER + "\n" + str(soup)
    path.write_text(cleaned, encoding="utf-8")


def extract_subject_and_body(path: Path) -> tuple[str | None, str | None, str | None, str]:
    """Return subject, name and generic salutations and HTML body.

    The function looks for lines in the HTML that start with ``Subject:``,
    ``Name Salutation:`` and ``No name Salutation:``.  When found, those
    elements are removed from the returned body and their text is returned
    separately.
    """
    text = path.read_text(encoding="utf-8")
    soup = BeautifulSoup(text, "html.parser")

    subject = None
    name_sal = None
    generic_sal = None

    for el in list(soup.find_all(["p", "div", "span", "h1", "h2", "h3"])):
        content = el.get_text(strip=True)
        if not content:
            continue
        if content.startswith("Subject:"):
            subject = content[len("Subject:"):].strip()
            el.decompose()
        elif content.startswith("Name Salutation:"):
            name_sal = content[len("Name Salutation:"):].strip()
            el.decompose()
        elif content.startswith("No name Salutation:"):
            generic_sal = content[len("No name Salutation:"):].strip()
            el.decompose()

    body = str(soup)
    return subject, name_sal, generic_sal, body
