from __future__ import annotations
from pathlib import Path
from typing import Iterable
import re

from bs4 import BeautifulSoup

PROCESSED_MARKER = "<!-- processed -->"

# ---------------------------------------------------------------------------
#  Public helpers
# ---------------------------------------------------------------------------

def process_template_file(path: Path) -> None:
    """Rewrite exported Notion HTML into a lean, email‑ready template.

    The function is *idempotent*: once a template was processed it is skipped on
    subsequent invocations – we tag the file with ``PROCESSED_MARKER``.

    Processing steps in order::

        1. Remove <title> and the Notion <header> block.
        2. Strip <style>/<link rel="stylesheet"> blocks and <script> tags.
        3. Remove all <img> tags.
        4. Remove *all* ``class``/``id`` attributes – CSS classes are useless
           after we removed the style sheets.
        5. Reduce inline ``style="…"`` attributes to an allow‑list of safe
           properties that render reliably in major email clients.
        6. Add a base <body> inline style (font, colour, margins).
        7. Inject ``PROCESSED_MARKER`` right after the opening <head> tag so
           Outlook doesn’t choke on the comment if it’s at the very end.

    The resulting HTML contains *only* semantic tags plus minimal inline style
    attributes – which is what Gmail, Outlook, Apple Mail *and* mobile clients
    are happiest with.
    """

    text = path.read_text(encoding="utf-8")
    if PROCESSED_MARKER in text:
        return

    soup = BeautifulSoup(text, "html.parser")

    # 1 ────────────────────────────────────────────────────────────────────
    _strip_title_and_header(soup)

    # 2–5 ─────────────────────────────────────────────────────────────────
    _strip_css_js_images(soup)

    # 6 ───────────────────────────────────────────────────────────────────
    _ensure_base_body_style(soup)

    # --------------------------------------------------------------------
    # Serialise & mark as processed
    # --------------------------------------------------------------------
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

    # Look at actual text‑bearing elements only; skip wrappers
    for el in list(soup.find_all(["p", "span", "h1", "h2", "h3", "div"])):
        if el.string is None:  # has child tags → wrapper div, ignore
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

# ---------------------------------------------------------------------------
#  Private helpers
# ---------------------------------------------------------------------------

_ALLOWED_INLINE_PROPERTIES = {
    # Lists
    "list-style-type",
    # Text alignment
    "text-align",
    # Spacing – tolerated by most clients
    "margin",
    "margin-left",
    "margin-right",
    "margin-top",
    "margin-bottom",
}

_BASE_BODY_STYLE = (
    "margin:0;"
    "padding:0;"
    "font-family:Arial,Helvetica,sans-serif;"
    "font-size:14px;"
    "line-height:1.4;"
    "color:#333333;"
)


def _strip_title_and_header(soup: BeautifulSoup) -> None:
    """Remove <title>, Notion header and page title/description."""
    if soup.title:
        soup.title.decompose()

    header = soup.find("header")
    if header:
        header.decompose()

    for tag in soup.select("header, h1.page-title, p.page-description"):
        tag.decompose()


def _strip_css_js_images(soup: BeautifulSoup) -> None:
    """Remove CSS/JS, images, classes/ids and prune inline styles."""
    # 2  Strip style / link / script
    for tag in soup.find_all(["style", "script"]):
        tag.decompose()

    for link in soup.find_all("link", rel=lambda r: r and "stylesheet" in r):
        link.decompose()

    # 3  Remove images entirely – keeps weight low & prevents CID hassles
    for img in soup.find_all("img"):
        img.decompose()

    # 4+5  Walk all tags and clean attrs
    for el in soup.find_all(True):  # True == any tag
        # Remove class/id: they have no meaning once CSS is gone
        el.attrs.pop("class", None)
        el.attrs.pop("id", None)

        # Clean up inline styles (if any)
        if "style" in el.attrs:
            clean_style = _filter_inline_style(el["style"], _ALLOWED_INLINE_PROPERTIES)
            if clean_style:
                el["style"] = clean_style
            else:
                del el.attrs["style"]


_STYLE_ITEM_RE = re.compile(r"\s*([\w-]+)\s*:\s*([^;]+)")


def _filter_inline_style(style_str: str, allowed: set[str]) -> str:
    """Return a pruned inline‑style string containing only *allowed* properties."""
    items: list[str] = []
    for prop, val in _STYLE_ITEM_RE.findall(style_str):
        if prop in allowed:
            items.append(f"{prop}:{val.strip()}")
    return ";".join(items)


def _ensure_base_body_style(soup: BeautifulSoup) -> None:
    """Add/merge a sane default inline style on <body>."""
    if not soup.body:
        return

    existing = soup.body.get("style", "")
    if existing:
        # Merge – user‑defined props win
        soup.body["style"] = f"{_BASE_BODY_STYLE}{existing}"
    else:
        soup.body["style"] = _BASE_BODY_STYLE
