#!/usr/bin/env python3
r"""Resume & cover-letter build system.

Pipeline:
  Markdown source  →  structured data  →  HTML  →  PDF (via headless Chrome)
                                            \─→ DOCX (via python-docx)

Outputs (filenames are derived from the H1 of the source, never hardcoded):
  - Master resume:        ./<Name>_Resume.{pdf,docx}
  - Master + photo:        ./<Name>_Resume_Photo.{pdf,docx}
  - Tailored:              applications/<slug>.{pdf,docx}
  - Cover letter:          applications/<slug>-cover-letter.{pdf,docx}

Usage:
  uv run python build.py                                # master, no photo
  uv run python build.py --photo photo.jpg              # master, with photo
  uv run python build.py --all                          # build both variants
  uv run python build.py --all --photo photo.jpg        # both variants, with photo
  uv run python build.py --check                        # ATS verification
  uv run python build.py applications/acme-swe.md
  uv run python build.py applications/acme-swe.md --photo photo.jpg

Why this design:
  - Chrome → PDF gives selectable, vector text and full CSS paged-media control
    (works with Workday, Greenhouse, Lever, iCIMS, Taleo, Ashby).
  - python-docx emits real Word styles (Heading 1/2) so DOCX-aware ATS get the
    section hierarchy for free.
  - No magic. Edit resume.md, run this, send the PDF.
"""

from __future__ import annotations

import argparse
import base64
import copy
import datetime as _dt
import html
import io
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import cast

import pymupdf as fitz
from PIL import Image
from docx import Document

# `docx.Document` is a factory *function*; `docx.document.Document` is the class.
# Annotations need the class — using the function makes pyright reject the
# parameter type outright.
from docx.document import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Inches
from docx.styles.style import ParagraphStyle
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent
TEMPLATES = ROOT / "templates"
RESUME_HTML = TEMPLATES / "resume.html"
COVER_HTML = TEMPLATES / "cover.html"
STYLES_CSS = TEMPLATES / "styles.css"
MASTER_MD = ROOT / "resume.md"

ACCENT = RGBColor(0x0F, 0x2A, 0x4A)  # matches CSS #0F2A4A
BODY_FONT = "Calibri"
FALLBACK_FONT = "Carlito"

# No personal name lives in this source tree. The candidate's name is always
# read from the H1 of the Markdown being built, so the repository can be
# published with sample content and carry no personal data of its own.
DEFAULT_NAME = "Your Name"
# Filenames use underscores: ATS portals and downstream tooling handle
# underscores more reliably than spaces.
DEFAULT_FILENAME_SUFFIX = "_Resume"


def _name_from_md(md_path: Path) -> str:
    """Read the H1 of a Markdown source — the author and filename seed."""
    try:
        for line in HTML_COMMENT_RE.sub(
            "", md_path.read_text(encoding="utf-8")
        ).splitlines():
            if line.lstrip().startswith("# "):
                return _strip_md_inline(line.lstrip()[2:].strip()) or DEFAULT_NAME
    except OSError:
        pass
    return DEFAULT_NAME


def _stem_for(name: str) -> str:
    """'Jordan Ellis' -> 'Jordan_Ellis_Resume'."""
    slug = re.sub(r"[^A-Za-z0-9]+", "_", (name or "").strip()).strip("_")
    return f"{slug or 'Resume'}{DEFAULT_FILENAME_SUFFIX}"


# Recognized resume section headings (case-insensitive; mapped to canonical keys)
CANONICAL_SECTIONS = {
    "summary": "summary",
    "professional summary": "summary",
    "profile": "summary",
    "experience": "experience",
    "work experience": "experience",
    "professional experience": "experience",
    "employment": "experience",
    "education": "education",
    "academic background": "education",
    "skills": "skills",
    "technical skills": "skills",
    "core skills": "skills",
    "projects": "projects",
    "selected projects": "projects",
    "certifications": "certifications",
    "certification": "certifications",
    "licenses": "certifications",
    "publications": "publication",
    "publication": "publication",
    "research": "publication",
    "volunteering": "volunteering",
    "volunteer": "volunteering",
    "leadership": "volunteering",
    "activities": "volunteering",
    "languages": "languages",
    "language": "languages",
    "references": "references",
    # fallback: anything honour-ish lands in certifications
    "honors": "certifications",
    "awards": "certifications",
    "achievements": "certifications",
}

# Render order. Skills sit before Experience so a recruiter scanning a
# broad-spectrum application sees capabilities before job history.
SECTION_ORDER = [
    "summary",
    "skills",
    "experience",
    "education",
    "certifications",
    "publication",
    "projects",
    "volunteering",
    "languages",
    "references",
]

# ---------------------------------------------------------------------------
# Markdown → structured data
# ---------------------------------------------------------------------------

LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
STRONG_RE = re.compile(r"\*\*(.+?)\*\*")
EM_RE = re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)")


@dataclass
class Entry:
    title: str = ""
    org: str = ""
    meta: str = ""
    bullets: list[str] = field(default_factory=list)
    # Sub-paragraphs that are not bullets but render after bullets (rare)


@dataclass
class SkillRow:
    cat: str = ""
    items_html: str = ""


@dataclass
class Section:
    name: str = ""  # canonical key (e.g. "experience")
    label: str = ""  # human label
    entries: list[Entry] = field(default_factory=list)
    skill_rows: list[SkillRow] = field(default_factory=list)
    bullets: list[str] = field(default_factory=list)  # honors list


@dataclass
class Resume:
    name: str = ""
    subtitle: str = ""  # italic line under name (e.g. headline)
    contact_lines: list[list[tuple[str, str]]] = field(default_factory=list)
    open_to: str = ""
    summary: str = ""
    sections: dict[str, Section] = field(default_factory=dict)
    section_order: list[str] = field(default_factory=list)


def md_inline_to_html(text: str) -> str:
    """Convert markdown inline to HTML.

    Order matters: apply bold/em first so that `**[link](url)**` becomes
    `<strong>[link](url)</strong>`, then convert the link patterns into
    `<a>` tags. This way the `<strong>` wrappers survive the link
    replacement.
    """
    if not text:
        return ""
    text = _md_inline_no_links(text)
    out = []
    pos = 0
    for m in LINK_RE.finditer(text):
        before = text[pos : m.start()]
        out.append(before)
        out.append(f'<a href="{html.escape(m.group(2), quote=True)}">{m.group(1)}</a>')
        pos = m.end()
    out.append(text[pos:])
    return "".join(out)


def _md_inline_no_links(text: str) -> str:
    """Apply **bold** and *italic*; assumes links have already been split out."""
    if not text:
        return ""
    parts = STRONG_RE.split(text)
    for i, p in enumerate(parts):
        if i % 2 == 1:
            parts[i] = f"<strong>{_apply_em(p)}</strong>"
        else:
            parts[i] = _apply_em(p)
    return "".join(parts)


def _apply_em(text: str) -> str:
    parts = EM_RE.split(text)
    for i, p in enumerate(parts):
        if i % 2 == 1:
            parts[i] = f"<em>{p}</em>"
    return parts[0] + "".join(parts[1:]) if parts else ""


def _strip_md_inline(text: str) -> str:
    """Plain-text version (no markdown markers, no link markup)."""
    text = LINK_RE.sub(r"\1", text)
    text = STRONG_RE.sub(r"\1", text)
    text = EM_RE.sub(r"\1", text)
    return text


def _classify_section(label: str) -> str:
    """Map a free-form section label to a canonical section key."""
    norm = label.strip().lower()
    if norm in CANONICAL_SECTIONS:
        return CANONICAL_SECTIONS[norm]
    # Containment match
    for k, v in CANONICAL_SECTIONS.items():
        if k in norm:
            return v
    return norm  # unknown → keep as-is


HTML_COMMENT_RE = re.compile(r"<!--.*?-->", re.DOTALL)


def _read_md(md_path: Path) -> str:
    """Read a markdown source and drop HTML comments.

    HTML comments are used in the templates for inline authoring guidance
    ("put something specific about the company here"). They are instructions
    to the human, never content — so they must not reach the PDF or DOCX.
    """
    return HTML_COMMENT_RE.sub("", md_path.read_text(encoding="utf-8"))


def parse_resume(md_path: Path) -> Resume:
    """Parse the master resume markdown into a Resume dataclass.

    Schema:
      # Name                              (first H1 — required)
      Subtitle line (italic, gray)        (optional, first paragraph)
      Contact line(s)                     (one or more lines until first ## or blank+##)
      Open-to line                        (optional, italic gray)
      ## Section Title
      ... bullets or entries ...
      ### Entry title
      Org · Location                      (italic line)
      date - date                         (meta line)
      - bullet
      - bullet
    """
    text = _read_md(md_path)
    lines = text.splitlines()

    r = Resume()
    current_section: Section | None = None
    current_entry: Entry | None = None

    i = 0
    # Find first H1
    while i < len(lines) and not lines[i].lstrip().startswith("# "):
        i += 1
    if i >= len(lines):
        raise ValueError(f"{md_path}: no top-level name (# ...) found")
    r.name = lines[i].lstrip()[2:].strip()
    i += 1

    # Collect preamble until first ##
    while i < len(lines) and not lines[i].lstrip().startswith("## "):
        raw = lines[i].rstrip()
        s = raw.strip()
        if not s:
            i += 1
            continue
        if s.startswith("# "):
            # Sub-name? Ignore (we only support one H1).
            i += 1
            continue
        # Heuristic: subtitle is the FIRST preamble line that isn't contact-y
        # Contact lines are recognised by containing an @, phone-like digits with +,
        # or "|", or "linkedin.com" / "github.com" / a URL
        looks_contact = bool(
            "@" in s
            or re.search(r"\+?\d[\d\s().-]{6,}", s)
            or "linkedin.com" in s.lower()
            or "github.com" in s.lower()
            or re.search(r"https?://", s)
            or "|" in s
        )
        if not r.subtitle and not looks_contact:
            # Could be subtitle or open-to
            if s.lower().startswith("open to"):
                r.open_to = _strip_md_inline(s)
            else:
                r.subtitle = _strip_md_inline(s)
        elif s.lower().startswith("open to"):
            r.open_to = _strip_md_inline(s)
        else:
            # Contact line — split by '|'
            line_pieces = [
                _parse_contact_piece(p.strip())
                for p in re.split(r"\s*\|\s*", s)
                if p.strip()
            ]
            if line_pieces:
                r.contact_lines.append(line_pieces)
        i += 1

    # Sections
    while i < len(lines):
        raw = lines[i].rstrip()
        s = raw.strip()
        if not s:
            i += 1
            continue
        if s.startswith("## "):
            label = s[3:].strip()
            key = _classify_section(label)
            sec = Section(name=key, label=label)
            r.sections[key] = sec
            r.section_order.append(key)
            current_section = sec
            current_entry = None
            i += 1
            continue
        if s.startswith("### "):
            if current_section is None:
                raise ValueError(
                    f"{md_path}: '###' heading before any '##' section — "
                    "every entry must belong to a section"
                )
            # Keep raw markdown so `### [Title](url)` still renders a link.
            current_entry = Entry(title=s[4:].strip())
            current_section.entries.append(current_entry)
            i += 1
            continue
        if s.startswith("- "):
            content = s[2:].strip()
            if current_entry is not None:
                current_entry.bullets.append(content)
            elif current_section is not None:
                # Bullets at section level (no entry) — supported by honors,
                # projects, and skills sections
                current_section.bullets.append(content)
            i += 1
            continue
        # A non-bullet line inside a section. Keep the RAW markdown here —
        # `### Org](url)` must survive so the employer name stays clickable.
        # Renderers (HTML via md_inline_to_html, DOCX via _add_runs_linked)
        # convert it at output time.
        if current_entry is not None and not current_entry.org:
            current_entry.org = s
            i += 1
            continue
        if current_entry is not None and not current_entry.meta:
            current_entry.meta = s
            i += 1
            continue
        # Skill line: **Category:** items  OR  Category: items
        if current_section is not None and current_section.name == "skills":
            m = re.match(r"^\*\*(.+?):\*\*\s*(.*)$", s)
            if m:
                current_section.skill_rows.append(
                    SkillRow(
                        cat=_strip_md_inline(m.group(1).strip()),
                        items_html=md_inline_to_html(m.group(2).strip()),
                    )
                )
                i += 1
                continue
            # If it's a "Category: items" line at section start, also treat as skill row
            m2 = re.match(r"^([A-Z][A-Za-z &/+]+):\s*(.+)$", s)
            if m2 and not current_section.skill_rows and not current_section.entries:
                current_section.skill_rows.append(
                    SkillRow(
                        cat=_strip_md_inline(m2.group(1).strip()),
                        items_html=md_inline_to_html(m2.group(2).strip()),
                    )
                )
                i += 1
                continue
        # Summary-style free paragraph
        if current_section is not None and current_section.name == "summary":
            if current_section.bullets:
                current_section.bullets.append(s)
            else:
                current_section.bullets = [s]
            r.summary = (
                (r.summary + " " + _strip_md_inline(s)).strip()
                if r.summary
                else _strip_md_inline(s)
            )
            i += 1
            continue
        # Otherwise: append as plain section paragraph (rare)
        if (
            current_section is not None
            and current_section.name == "honors"
            and current_entry is None
        ):
            current_section.bullets.append(s)
        i += 1

    # Normalize summary from parsed section
    if "summary" in r.sections:
        # bullets now holds the summary paragraph(s)
        joined = " ".join(r.sections["summary"].bullets).strip()
        r.summary = joined or r.summary

    return r


def _parse_contact_piece(piece: str) -> tuple[str, str]:
    """Split a contact piece into (display, link)."""
    m = LINK_RE.search(piece)
    if m:
        return _strip_md_inline(m.group(1)), m.group(2)
    # Detect email
    if re.fullmatch(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", piece):
        return piece, "mailto:" + piece
    # Detect bare URL
    if re.fullmatch(r"https?://\S+", piece):
        return piece, piece
    # Detect phone
    if re.match(r"^\+?[\d\s().-]{7,}$", piece):
        digits = re.sub(r"\D", "", piece)
        return piece, "tel:" + digits
    return piece, ""


# ---------------------------------------------------------------------------
# HTML rendering
# ---------------------------------------------------------------------------


def render_resume_html(r: Resume, photo_data_uri: str = "") -> tuple[str, str]:
    """Render a Resume into (html, title_for_metadata).

    `photo_data_uri` is a `data:image/jpeg;base64,...` URI (or empty for the
    no-photo variant). When provided, the resume template inserts a
    right-floated headshot in the header.
    """
    tpl = RESUME_HTML.read_text(encoding="utf-8")
    css = STYLES_CSS.read_text(encoding="utf-8")
    title = f"{r.name} — Resume"

    out = tpl
    out = out.replace("__TITLE__", html.escape(title))
    out = out.replace("__CSS__", css)
    out = out.replace("__NAME__", html.escape(r.name))

    if photo_data_uri:
        photo_html = (
            '<img class="photo" alt="" '
            f'src="{photo_data_uri}" '
            'width="113" height="150">'
        )
        out = out.replace("__PHOTO__", photo_html)
        out = out.replace('class="header"', 'class="header has-photo"', 1)
    else:
        out = out.replace("__PHOTO__", "")

    # Subtitle (italic gray line under name)
    if r.subtitle:
        out = out.replace(
            "__SUBTITLE__", f'<p class="subtitle">{md_inline_to_html(r.subtitle)}</p>'
        )
    else:
        out = out.replace("__SUBTITLE__", "")

    # Contact: each piece becomes <span class="ct">label</span>; if multiple pieces in line, join with sep
    contact_html_parts = []
    for line in r.contact_lines:
        for i, (label, link) in enumerate(line):
            if i > 0:
                contact_html_parts.append('<span class="sep">·</span>')
            if link:
                contact_html_parts.append(
                    f'<a href="{html.escape(link, quote=True)}">{html.escape(label)}</a>'
                )
            else:
                contact_html_parts.append(html.escape(label))
    contact_html = "\n".join(contact_html_parts) if contact_html_parts else ""
    out = out.replace("__CONTACT__", contact_html)

    if r.open_to:
        out = out.replace(
            "__OPEN_TO__", f'<p class="open-to">{md_inline_to_html(r.open_to)}</p>'
        )
    else:
        out = out.replace("__OPEN_TO__", "")

    # All body sections rendered in canonical order into a single slot.
    out = out.replace("__SECTIONS__", _render_all_sections(r))

    return out, title


def _render_all_sections(r: Resume) -> str:
    """Render every recognised section into `<section>` blocks, in order."""
    blocks: list[str] = []
    seen = set()
    for key in SECTION_ORDER:
        if key in r.sections:
            seen.add(key)
            sec = r.sections[key]
            inner = _render_section_body(sec, r)
            if not inner:
                continue
            label = (sec.label or key.title()).strip()
            blocks.append(
                f'<section class="section {key}">'
                f'<h2 class="section-title">{html.escape(label)}</h2>'
                f"{inner}</section>"
            )
    # Any section the schema didn't recognise still gets rendered, after the
    # known ones, so nothing in the source is silently dropped.
    for key in r.section_order:
        if key in seen:
            continue
        sec = r.sections[key]
        inner = _render_section_body(sec, r)
        if not inner:
            continue
        label = (sec.label or key.title()).strip()
        blocks.append(
            f'<section class="section {key}">'
            f'<h2 class="section-title">{html.escape(label)}</h2>'
            f"{inner}</section>"
        )
    return "\n".join(blocks)


def _render_section_body(sec: Section, r: Resume) -> str:
    """Render the inner body of one section (no heading, no <section> wrapper)."""
    key = sec.name

    if key == "summary":
        if not r.summary:
            return ""
        return f"<p>{md_inline_to_html(r.summary)}</p>"

    if key == "skills":
        if sec.skill_rows:
            rows = []
            for sr in sec.skill_rows:
                rows.append(
                    '<div class="skill-row">'
                    f'<span class="cat">{html.escape(sr.cat)}</span>'
                    '<span class="items"> · ' + sr.items_html + "</span>"
                    "</div>"
                )
            return "\n".join(rows)
        if sec.bullets:
            return "".join(
                f'<div class="skill-row">{md_inline_to_html(b)}</div>'
                for b in sec.bullets
            )
        if sec.entries:
            return _render_entries(sec)
        return ""

    if key in ("experience", "education"):
        return _render_entries(sec)

    if key == "projects":
        if sec.entries:
            return _render_entries(sec)
        if sec.bullets:
            return (
                '<ul class="honors">'
                + "".join(f"<li>{md_inline_to_html(b)}</li>" for b in sec.bullets)
                + "</ul>"
            )
        return ""

    if key == "languages":
        # Rendered as plain text lines (e.g. "Burmese — Native · English — Proficient")
        if sec.bullets:
            return (
                '<p class="languages">'
                + " · ".join(md_inline_to_html(b) for b in sec.bullets)
                + "</p>"
            )
        if sec.entries:
            return _render_entries(sec)
        return ""

    # certifications / publication / volunteering / references / anything else
    if sec.entries:
        return _render_entries(sec)
    if sec.bullets:
        return (
            '<ul class="honors">'
            + "".join(f"<li>{md_inline_to_html(b)}</li>" for b in sec.bullets)
            + "</ul>"
        )
    return ""


def _render_entries(sec: Section | None) -> str:
    if not sec:
        return ""
    if sec.name == "projects":
        return _render_project_entries(sec)
    out = []
    for entry in sec.entries:
        parts = [f'<div class="entry">']
        if entry.title:
            parts.append(
                f'<h3 class="entry-title">{md_inline_to_html(entry.title)}</h3>'
            )
        if entry.org:
            parts.append(f'<p class="entry-org">{md_inline_to_html(entry.org)}</p>')
        if entry.meta:
            parts.append(f'<p class="entry-meta">{md_inline_to_html(entry.meta)}</p>')
        if entry.bullets:
            bullets_html = "".join(
                f"<li>{md_inline_to_html(b)}</li>" for b in entry.bullets
            )
            parts.append(f'<ul class="entry-bullets">{bullets_html}</ul>')
        parts.append("</div>")
        out.append("\n".join(parts))
    return "\n".join(out)


def _render_project_entries(sec: Section) -> str:
    """Projects render as: title (with optional inline link), description, bullets."""
    out = []
    for entry in sec.entries:
        parts = ['<div class="entry project">']
        if entry.title:
            parts.append(
                f'<h3 class="entry-title">{md_inline_to_html(entry.title)}</h3>'
            )
        # For projects, use `org` as the plain description line (not italic)
        if entry.org:
            parts.append(f'<p class="entry-desc">{md_inline_to_html(entry.org)}</p>')
        elif entry.meta:
            parts.append(f'<p class="entry-desc">{md_inline_to_html(entry.meta)}</p>')
        if entry.bullets:
            bullets_html = "".join(
                f"<li>{md_inline_to_html(b)}</li>" for b in entry.bullets
            )
            parts.append(f'<ul class="entry-bullets">{bullets_html}</ul>')
        parts.append("</div>")
        out.append("\n".join(parts))
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Cover letter
# ---------------------------------------------------------------------------


@dataclass
class CoverLetter:
    name: str = ""
    contact_lines: list[list[tuple[str, str]]] = field(default_factory=list)
    recipient_name: str = "Hiring Manager"
    recipient_title: str = ""
    company: str = ""
    address: str = ""
    role: str = ""  # the role being applied for (from H1, informational)
    greeting: str = "Hiring Manager"
    body_paragraphs: list[str] = field(default_factory=list)
    signoff_name: str = ""


def parse_cover_letter(md_path: Path) -> CoverLetter:
    """Parse a cover letter markdown file.

    Supported shape (the one `coverletter.py` writes):

        # Company — Role                       H1: company (+ optional role)
        City, Country | email | phone ...       contact line
        Dear Hiring Manager,                   greeting

        Paragraph one...                       body (blank-line separated)

        Paragraph two...

        Yours sincerely,                       sign-off
        Your Name                              signatory name

    An optional `## Recipient` block (name / title / company / address lines)
    is also honoured when present.
    """
    text = _read_md(md_path)
    lines = text.splitlines()

    c = CoverLetter()

    # --- H1: company (and optional role) -----------------------------------
    i = 0
    while i < len(lines) and not lines[i].lstrip().startswith("# "):
        i += 1
    if i < len(lines):
        title_line = lines[i].lstrip()[2:].strip()
        if "—" in title_line:
            parts = [p.strip() for p in title_line.split("—", 1)]
        elif " - " in title_line:
            parts = [p.strip() for p in title_line.split(" - ", 1)]
        else:
            parts = [title_line, ""]
        c.company = parts[0]
        # The second half of the H1 is the ROLE BEING APPLIED FOR, not the
        # recipient's job title — so it must not go into the recipient block.
        # Store it separately; recipient_title is only ever set from an
        # explicit "## Recipient" block.
        if len(parts) > 1:
            c.role = parts[1]
        i += 1

    # --- Everything after H1 ------------------------------------------------
    rest = lines[i:]

    # Contact line = first non-empty line that smells like contact info.
    contact_idx = None
    for idx, raw in enumerate(rest):
        s = raw.strip()
        if not s:
            continue
        looks_contact = (
            "@" in s
            or "linkedin.com" in s.lower()
            or "github.com" in s.lower()
            or re.search(r"\+?\d[\d\s().-]{6,}", s) is not None
            or "|" in s
        )
        if looks_contact:
            for piece in re.split(r"\s*\|\s*", s):
                piece = piece.strip()
                if piece:
                    c.contact_lines.append([_parse_contact_piece(piece)])
            contact_idx = idx
            break

    body_start = 0 if contact_idx is None else contact_idx + 1
    tail = rest[body_start:]

    # --- Optional `## Recipient` block --------------------------------------
    recipient_block: list[str] = []
    j = 0
    while j < len(tail):
        s = tail[j].strip()
        if s.startswith("## "):
            label = s[3:].strip().lower()
            j += 1
            block: list[str] = []
            while j < len(tail):
                s2 = tail[j].strip()
                if s2.startswith("## ") or s2.startswith("# "):
                    break
                if s2:
                    block.append(s2)
                j += 1
            if "recipient" in label or "address" in label:
                recipient_block = block
            # `## Body` (or any other heading) simply starts the body.
            continue
        break

    if recipient_block:
        c.recipient_name = recipient_block[0]
        if len(recipient_block) > 1:
            c.recipient_title = recipient_block[1]
        if len(recipient_block) > 2:
            c.company = c.company or recipient_block[2]
        if len(recipient_block) > 3:
            c.address = "\n".join(recipient_block[3:])

    # --- Split remaining text into paragraphs --------------------------------
    paras: list[str] = []
    cur: list[str] = []
    for raw in tail[j:]:
        s = raw.strip()
        if not s:
            if cur:
                paras.append(" ".join(cur).strip())
                cur = []
        elif s.startswith("## ") or s.startswith("# "):
            if cur:
                paras.append(" ".join(cur).strip())
                cur = []
        else:
            cur.append(s)
    if cur:
        paras.append(" ".join(cur).strip())

    # --- Greeting ------------------------------------------------------------
    if paras and paras[0].lower().startswith("dear "):
        m = re.match(r"^Dear\s+(.+?),?\s*$", paras[0], re.IGNORECASE)
        if m:
            c.greeting = m.group(1).rstrip(",").strip()
            paras = paras[1:]

    # --- Sign-off ------------------------------------------------------------
    signoff_name = ""
    for pat in (
        r"^yours\s+sincerely,?\s*(.*)$",
        r"^sincerely,?\s*(.*)$",
        r"^best\s+regards,?\s*(.*)$",
        r"^regards,?\s*(.*)$",
        r"^yours\s+faithfully,?\s*(.*)$",
    ):
        if paras:
            m = re.match(pat, paras[-1], re.IGNORECASE)
            if m:
                signoff_name = m.group(1).strip()
                paras = paras[:-1]
                break
    # Sign-off on its own line, name on the next
    if (
        not signoff_name
        and len(paras) >= 2
        and re.match(
            r"^(yours\s+(sincerely|faithfully)|sincerely|best\s+regards|regards),?$",
            paras[-2],
            re.IGNORECASE,
        )
    ):
        signoff_name = paras[-1]
        paras = paras[:-2]

    if not signoff_name:
        signoff_name = DEFAULT_NAME
    c.signoff_name = signoff_name
    c.name = signoff_name

    if not paras:
        paras = [
            "[Paste 1-3 paragraphs connecting your experience to this role and "
            "company. Mirror 3-5 exact phrases from the job ad.]"
        ]
    c.body_paragraphs = paras
    return c


def render_cover_html(c: CoverLetter) -> tuple[str, str]:
    tpl = COVER_HTML.read_text(encoding="utf-8")
    css = STYLES_CSS.read_text(encoding="utf-8")
    out = tpl
    title = f"{c.name or DEFAULT_NAME} — Cover Letter"
    out = out.replace("__TITLE__", html.escape(title))
    out = out.replace("__CSS__", css)
    out = out.replace("__NAME__", html.escape(c.name or DEFAULT_NAME))
    # contact — one line, "·" between pieces, same as the resume
    contact_parts = []
    flat = [(label, link) for line in c.contact_lines for label, link in line]
    for i, (label, link) in enumerate(flat):
        if i > 0:
            contact_parts.append('<span class="sep">·</span>')
        if link:
            contact_parts.append(
                f'<a href="{html.escape(link, quote=True)}">{html.escape(label)}</a>'
            )
        else:
            contact_parts.append(html.escape(label))
    contact_html = " ".join(contact_parts)
    out = out.replace("__CONTACT__", contact_html)

    out = out.replace("__DATE__", html.escape(_today_str()))
    out = out.replace(
        "__RECIPIENT_NAME__", html.escape(c.recipient_name or "Hiring Manager")
    )
    out = out.replace("__RECIPIENT_TITLE__", html.escape(c.recipient_title or ""))
    out = out.replace("__COMPANY__", html.escape(c.company or ""))
    address_html = "<br>".join(
        html.escape(part) for part in (c.address or "").split("\n") if part
    )
    out = out.replace("__ADDRESS__", address_html)
    out = out.replace(
        "__RECIPIENT_GREETING__", html.escape(c.greeting or "Hiring Manager")
    )

    body_paras = "".join(
        f"<p>{md_inline_to_html(p)}</p>" for p in c.body_paragraphs if p.strip()
    )
    out = out.replace("__BODY__", body_paras)
    return out, title


def _today_str() -> str:
    return _dt.date.today().strftime("%d %B %Y")


# ---------------------------------------------------------------------------
# Photo processing
# ---------------------------------------------------------------------------


def process_photo(
    photo_path: Path,
    build_dir: Path,
    stem: str,
    out_w_px: int = 360,
    out_h_px: int = 480,
    aspect: tuple[int, int] = (3, 4),
) -> tuple[str, Path]:
    """Center-crop the photo to `aspect` (default 3:4 portrait) and resize.

    Returns (base64_data_uri_for_html, file_path_for_docx). The file is also
    written to `.build/<stem>.jpg` so downstream builds reuse the same crop.
    """
    img = Image.open(photo_path)
    if img.mode != "RGB":
        img = img.convert("RGB")

    aw, ah = aspect
    src_w, src_h = img.size
    target_ratio = aw / ah
    src_ratio = src_w / src_h

    if src_ratio > target_ratio:
        # source is wider than target → crop sides
        new_w = int(round(src_h * target_ratio))
        left = (src_w - new_w) // 2
        img = img.crop((left, 0, left + new_w, src_h))
    elif src_ratio < target_ratio:
        # source is taller than target → crop top/bottom
        new_h = int(round(src_w / target_ratio))
        top = (src_h - new_h) // 2
        img = img.crop((0, top, src_w, top + new_h))

    # Resize to the target pixel box (300 DPI on a 30mm × 40mm print area).
    # Image.LANCZOS is the deprecated module-level alias; Pillow has been
    # nudging everyone to Image.Resampling since 9.1.
    img = img.resize((out_w_px, out_h_px), Image.Resampling.LANCZOS)

    build_dir.mkdir(exist_ok=True)
    out_file = build_dir / f"{stem}.jpg"
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=88, optimize=True)
    out_file.write_bytes(buf.getvalue())

    data_uri = "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode(
        "ascii"
    )
    return data_uri, out_file


# ---------------------------------------------------------------------------
# PDF via headless Chrome
# ---------------------------------------------------------------------------


def _chrome_path() -> str | None:
    candidates = []
    if platform.system() == "Windows":
        candidates += [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        ]
        # Also PATH
        for n in ("chrome", "msedge"):
            p = shutil.which(n)
            if p:
                candidates.append(p)
    else:
        candidates += [
            "/usr/bin/google-chrome",
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            shutil.which("google-chrome") or "",
            shutil.which("chromium") or "",
            shutil.which("msedge") or "",
        ]
    for c in candidates:
        if c and Path(c).exists():
            return c
    return None


def html_to_pdf(
    html_path: Path,
    pdf_path: Path,
    title: str,
    author: str,
    subject: str = "Resume",
    keywords: str = "",
) -> None:
    chrome = _chrome_path()
    if not chrome:
        raise RuntimeError(
            "No Chrome/Edge found. Install Google Chrome or Microsoft Edge."
        )
    # Use a temp profile to avoid lock conflicts
    with tempfile.TemporaryDirectory() as profile:
        cmd = [
            chrome,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--no-pdf-header-footer",
            f"--user-data-dir={profile}",
            f"--print-to-pdf={pdf_path}",
            f"file:///{html_path.as_posix()}",
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if not pdf_path.exists():
            raise RuntimeError(
                f"Chrome failed to write PDF.\nstdout:\n{result.stdout}\nstderr:\n{result.stderr}"
            )
    # Patch metadata: title/author/subject/keywords
    _set_pdf_metadata(
        pdf_path, title=title, author=author, subject=subject, keywords=keywords
    )


def _set_pdf_metadata(
    pdf: Path, title: str, author: str, subject: str, keywords: str
) -> None:
    """Set /Info /Title /Author /Subject /Keywords via pymupdf.

    Chrome opens the file briefly; on Windows we have to close it first, edit
    metadata in a copy, then atomically replace. Chrome (and any PDF viewer
    that has the file open) can still hold a handle for a moment, so retry the
    swap a few times before giving up.
    """
    tmp_out = pdf.with_suffix(".pdf.tmp")
    doc = fitz.open(pdf)
    doc.set_metadata(
        {
            "title": title,
            "author": author,
            "subject": subject,
            "keywords": keywords,
            "creator": "cv-resume-consolidation",
            "producer": "cv-resume-consolidation (Chrome headless)",
        }
    )
    doc.save(tmp_out.as_posix(), garbage=4, deflate=True, clean=True)
    doc.close()

    last_err: Exception | None = None
    for attempt in range(8):
        try:
            os.replace(tmp_out, pdf)
            return
        except (PermissionError, OSError) as exc:  # transient Windows handle
            last_err = exc
            time.sleep(0.35 * (attempt + 1))
    # Final fallback: overwrite in place by copying bytes.
    try:
        pdf.write_bytes(tmp_out.read_bytes())
        tmp_out.unlink(missing_ok=True)
        return
    except Exception as exc:  # pragma: no cover - last resort
        last_err = exc
    if last_err:
        print(
            f"warning: could not swap metadata into {pdf.name} ({last_err}).\n"
            "         Is the PDF open in a viewer? Close it and rebuild.",
            file=sys.stderr,
        )


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------


def build_resume_docx(
    r: Resume, out_path: Path, photo_path: Path | None = None
) -> None:
    """Build the DOCX so it matches the PDF page-for-page.

    Every metric below is deliberately mirrored from `templates/styles.css`.
    If the CSS changes, these must change with it — otherwise the DOCX drifts
    onto a second page while the PDF stays on one (which is exactly the bug
    this function was rewritten to fix).

    PDF reference values:
      margins 9mm v / 13mm h · body 10pt / 1.3 · name 26pt · subtitle 10pt
      contact 9pt · open-to 9pt · section 8.8pt · entry title 10.4pt
      org 9.4pt · meta 9pt · bullets 9.6pt · skills 9.6pt
    """
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)  # A4
    sec.top_margin = Inches(0.35)  # 9mm  — was 0.55in, a big part of the overflow
    sec.bottom_margin = Inches(0.35)  # 9mm
    sec.left_margin = Inches(0.51)  # 13mm
    sec.right_margin = Inches(0.51)  # 13mm

    # Typed as BaseStyle, which carries no .font / .paragraph_format; the
    # object behind "Normal" is always a ParagraphStyle.
    normal = cast(ParagraphStyle, doc.styles["Normal"])
    normal.font.name = BODY_FONT
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.line_spacing = 1.3

    # Optional photo: FLOATING, anchored to the right margin, so the name /
    # contact text wraps beside it exactly as it does in the PDF. An inline
    # picture would stack above the text and add ~36mm of height.
    if photo_path and photo_path.exists():
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        _add_floating_picture(
            p, photo_path.as_posix(), width_cm=2.7, height_cm=3.6, align="right"
        )

    # Name
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.08
    run = p.add_run(r.name)
    run.font.size = Pt(26)
    run.font.bold = True
    run.font.color.rgb = ACCENT

    if r.subtitle:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(r.subtitle)
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x4A, 0x5A, 0x72)

    # Contact: combined on one line for docx
    contact_flat = []
    for line in r.contact_lines:
        for label, link in line:
            contact_flat.append((label, link))
    if contact_flat:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        for i, (label, link) in enumerate(contact_flat):
            if i > 0:
                sep = p.add_run("  ·  ")
                sep.font.size = Pt(9)
                sep.font.color.rgb = RGBColor(0x98, 0xA4, 0xB6)
            if link:
                # Contact pieces carry an explicit URL (mailto:/tel:/https:).
                # Emit a real Word hyperlink so it stays clickable in the DOCX.
                _add_hyperlink(p, link, label, size=9)
            else:
                _add_runs_linked(p, label, size=9)
        for run in p.runs:
            if not run.font.underline:
                run.font.color.rgb = RGBColor(0x2A, 0x2A, 0x2A)
    if r.open_to:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(r.open_to)
        run.font.size = Pt(9)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0x4A, 0x5A, 0x72)

    # Sections in canonical order (mirrors SECTION_ORDER used by the HTML
    # renderer, so PDF and DOCX always agree).
    section_order = list(SECTION_ORDER)
    extra = [k for k in r.section_order if k not in section_order]
    for key in section_order + extra:
        if key not in r.sections:
            continue
        sec_data = r.sections[key]
        if key == "summary":
            _add_section_heading(doc, sec_data.label or "Summary")
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            _add_runs_linked(p, r.summary, size=10)
            continue
        if key == "skills":
            _add_section_heading(doc, sec_data.label or "Skills")
            if sec_data.skill_rows:
                for sr in sec_data.skill_rows:
                    p = doc.add_paragraph()
                    p.paragraph_format.space_after = Pt(2)
                    cat = p.add_run(sr.cat)
                    cat.font.bold = True
                    cat.font.size = Pt(9.6)
                    cat.font.color.rgb = ACCENT
                    p.add_run(" · ").font.color.rgb = RGBColor(0x98, 0xA4, 0xB6)
                    _add_runs_linked(
                        p,
                        sr.items_html.replace("<strong>", "").replace("</strong>", ""),
                        size=9.6,
                    )
            elif sec_data.bullets:
                for b in sec_data.bullets:
                    p = doc.add_paragraph(style="List Bullet")
                    p.paragraph_format.space_after = Pt(1)
                    _add_runs_linked(p, b, size=9.6)
            continue
        if key == "languages":
            _add_section_heading(doc, sec_data.label or "Languages")
            if sec_data.bullets:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(0)
                for i, b in enumerate(sec_data.bullets):
                    if i:
                        p.add_run(" · ")
                    _add_runs_linked(p, b, size=9.8)
            continue
        if key == "references":
            _add_section_heading(doc, sec_data.label or "References")
            for b in sec_data.bullets:
                p = doc.add_paragraph()
                _add_runs_linked(p, b, size=9.6)
            continue
        if key in ("certifications", "publication", "volunteering", "projects"):
            _add_section_heading(doc, sec_data.label or key.title())
            if sec_data.entries:
                for entry in sec_data.entries:
                    p = doc.add_paragraph()
                    p.paragraph_format.space_after = Pt(0)
                    _add_runs_linked(p, entry.title, size=10.4, bold=True)
                    for b in entry.bullets:
                        bp = doc.add_paragraph(style="List Bullet")
                        bp.paragraph_format.space_after = Pt(1)
                        _add_runs_linked(bp, b, size=9.6)
            for b in sec_data.bullets:
                p = doc.add_paragraph(style="List Bullet")
                p.paragraph_format.space_after = Pt(1)
                _add_runs_linked(p, b, size=9.6)
            continue
        # experience / education (and anything else with entries)
        _add_section_heading(doc, sec_data.label or key.title())
        for entry in sec_data.entries:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(0)
            _add_runs_linked(p, entry.title, size=10.4, bold=True)
            if entry.org:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(0)
                _add_runs_linked(p, entry.org, size=9.4)
                for run in p.runs:
                    run.font.italic = True
                    run.font.color.rgb = RGBColor(0x4A, 0x5A, 0x72)
            if entry.meta:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(1)
                _add_runs_linked(p, entry.meta, size=9)
                for run in p.runs:
                    run.font.color.rgb = RGBColor(0x4A, 0x5A, 0x72)
            for b in entry.bullets:
                p = doc.add_paragraph(style="List Bullet")
                p.paragraph_format.space_after = Pt(1)
                _add_runs_linked(p, b, size=9.6)

    _apply_exact_line_spacing(doc)
    doc.save(out_path.as_posix())


def _apply_exact_line_spacing(doc: DocxDocument) -> None:
    """Convert Word 'multiple' line spacing into exact points matching the CSS.

    CSS `line-height: 1.3` means 1.3 x font-size. Word's "multiple 1.3" means
    1.3 x the font's *natural* line height (~1.22em for Calibri) — about 1.59em,
    roughly 22% looser. That gap alone pushed the DOCX onto a second page while
    the PDF stayed on one. Writing `w:line` in points with `w:lineRule="exact"`
    reproduces the CSS value precisely, so both outputs paginate identically.
    """
    for p in doc.paragraphs:
        sizes = [r.font.size.pt for r in p.runs if r.font.size is not None]
        if not sizes:
            continue
        size = max(sizes)
        factor = 1.08 if size >= 20 else 1.3  # display type uses 1.08 (see .name)
        p.paragraph_format.line_spacing = Pt(round(size * factor, 1))


def _add_section_heading(doc: DocxDocument, label: str) -> None:
    """Use Heading 2 style; this keeps ATS DOCX parsers happy.

    Metrics mirror `.section-title` in the CSS: 8.8pt bold uppercase, navy,
    ~5pt space before, ~2pt after, and a 1.5pt rule underneath (`w:sz` is in
    eighths of a point, so 12 = 1.5pt — matching the PDF's 2px rule).
    """
    p = doc.add_paragraph(style="Heading 2")
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(label.upper())
    run.font.bold = True
    run.font.size = Pt(8.8)
    run.font.color.rgb = ACCENT
    # Character spacing approximates the CSS letter-spacing: 0.18em
    rPr = run._element.get_or_add_rPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:val"), "32")  # ~0.18em at 8.8pt, in twentieths of a point
    rPr.append(spacing)
    # Bottom rule
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")  # 12/8 = 1.5pt
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "0F2A4A")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_runs_linked(p: Paragraph, text: str, size: float, bold: bool = False) -> None:
    """Add inline runs honouring **bold** and [text](url) to a docx paragraph."""
    pos = 0
    for m in LINK_RE.finditer(text):
        before = text[pos : m.start()]
        _add_runs(p, before, size, bold)
        _add_hyperlink(p, m.group(2), m.group(1), size, bold)
        pos = m.end()
    _add_runs(p, text[pos:], size, bold)


def _add_runs(p: Paragraph, text: str, size: float, bold: bool) -> None:
    parts = STRONG_RE.split(text)
    for i, seg in enumerate(parts):
        if not seg:
            continue
        run = p.add_run(seg)
        run.font.name = BODY_FONT
        run.font.size = Pt(size)
        if bold or i % 2 == 1:
            run.bold = True


def _add_floating_picture(
    paragraph: Paragraph,
    image_path: str,
    width_cm: float,
    height_cm: float,
    align: str = "right",
) -> None:
    """Insert a picture as a FLOATING (anchored) shape with square text wrap.

    python-docx's `add_picture` only produces `<wp:inline>` shapes, which sit in
    the paragraph flow — that is why the photo ended up stacked above the text
    and pushed the resume onto a second page. Converting the inline shape to
    `<wp:anchor>` anchored to the right margin makes the text wrap around it,
    matching the PDF layout.

    Child order inside `<wp:anchor>` is schema-significant:
    simplePos, positionH, positionV, extent, effectExtent, wrap*, docPr,
    cNvGraphicFramePr, graphic.
    """
    run = paragraph.add_run()
    shape = run.add_picture(image_path, width=Cm(width_cm), height=Cm(height_cm))
    inline = shape._inline
    drawing = inline.getparent()

    anchor = OxmlElement("wp:anchor")
    anchor.set("distT", "0")
    anchor.set("distB", "0")
    anchor.set("distL", "114300")  # 0.125in in EMU
    anchor.set("distR", "114300")
    anchor.set("simplePos", "0")
    anchor.set("relativeHeight", "251658240")
    anchor.set("behindDoc", "0")
    anchor.set("locked", "0")
    anchor.set("layoutInCell", "1")
    anchor.set("allowOverlap", "1")

    simple = OxmlElement("wp:simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    anchor.append(simple)

    ph = OxmlElement("wp:positionH")
    ph.set("relativeFrom", "margin")
    ph_align = OxmlElement("wp:align")
    ph_align.text = align
    ph.append(ph_align)
    anchor.append(ph)

    pv = OxmlElement("wp:positionV")
    pv.set("relativeFrom", "paragraph")
    pv_off = OxmlElement("wp:posOffset")
    pv_off.text = "0"
    pv.append(pv_off)
    anchor.append(pv)

    # Copy the geometry + graphic from the inline shape we just built.
    for child in inline:
        tag = child.tag.split("}")[-1]
        if tag in ("extent", "effectExtent"):
            anchor.append(copy.deepcopy(child))
    # Force the extent to the size we asked for (add_picture can round).
    emu_per_cm = 360000
    for ext in anchor.findall(qn("wp:extent")):
        ext.set("cx", str(int(width_cm * emu_per_cm)))
        ext.set("cy", str(int(height_cm * emu_per_cm)))

    wrap = OxmlElement("wp:wrapSquare")
    wrap.set("wrapText", "bothSides")
    anchor.append(wrap)

    for child in inline:
        tag = child.tag.split("}")[-1]
        if tag in ("docPr", "cNvGraphicFramePr"):
            anchor.append(copy.deepcopy(child))

    graphic = inline.find(qn("a:graphic"))
    if graphic is not None:
        anchor.append(copy.deepcopy(graphic))

    drawing.replace(inline, anchor)


def _add_hyperlink(
    p: Paragraph, url: str, text: str, size: float, bold: bool = False
) -> None:
    part = p.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), BODY_FONT)
    rFonts.set(qn("w:hAnsi"), BODY_FONT)
    rPr.append(rFonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)
    if bold:
        b = OxmlElement("w:b")
        rPr.append(b)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1A1A1A")
    rPr.append(color)
    # No underline: matches the PDF, where links are embedded and quiet.
    # Recruiters hover/click; the text itself stays clean.
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "none")
    rPr.append(u)
    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    p._p.append(hyperlink)


def build_cover_docx(c: CoverLetter, out_path: Path) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    sec.top_margin = Inches(0.7)
    sec.bottom_margin = Inches(0.7)
    sec.left_margin = Inches(0.8)
    sec.right_margin = Inches(0.8)

    # Typed as BaseStyle, which carries no .font / .paragraph_format; the
    # object behind "Normal" is always a ParagraphStyle.
    normal = cast(ParagraphStyle, doc.styles["Normal"])
    normal.font.name = BODY_FONT
    normal.font.size = Pt(11)
    normal.paragraph_format.line_spacing = 1.4

    p = doc.add_paragraph()
    run = p.add_run(c.name or DEFAULT_NAME)
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = ACCENT

    if c.contact_lines:
        p = doc.add_paragraph()
        flat = [(label, link) for line in c.contact_lines for label, link in line]
        for i, (label, link) in enumerate(flat):
            if i > 0:
                sep = p.add_run("  ·  ")
                sep.font.size = Pt(10)
                sep.font.color.rgb = RGBColor(0x98, 0xA4, 0xB6)
            if link:
                # Real Word hyperlink so the DOCX stays clickable, same as
                # the resume. Passing only `label` would drop the URL.
                _add_hyperlink(p, link, label, size=10)
            else:
                _add_runs_linked(p, label, size=10)
        p.paragraph_format.space_after = Pt(6)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    p.add_run(_today_str())

    for line in (c.recipient_name, c.recipient_title, c.company, c.address):
        if line:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            p.add_run(line)

    doc.add_paragraph()  # spacer

    p = doc.add_paragraph()
    run = p.add_run(f"Dear {c.greeting or 'Hiring Manager'},")
    doc.add_paragraph()  # spacer

    for para in c.body_paragraphs:
        if not para.strip():
            continue
        p = doc.add_paragraph()
        _add_runs_linked(p, para, size=11)
        p.paragraph_format.space_after = Pt(8)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Yours sincerely,")
    p = doc.add_paragraph()
    run = p.add_run(c.signoff_name or (c.name or DEFAULT_NAME))
    run.bold = True

    _apply_exact_line_spacing(doc)
    doc.save(out_path.as_posix())


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="build")
    ap.add_argument(
        "source",
        nargs="?",
        default=str(MASTER_MD),
        help="Markdown source file (resume.md by default)",
    )
    ap.add_argument(
        "--check",
        action="store_true",
        help="Run ATS sanity checks (single page, keywords, selectable text)",
    )
    ap.add_argument(
        "--preview",
        action="store_true",
        help="Also write an HTML preview alongside the PDF",
    )
    ap.add_argument("--out", help="Override output PDF path (DOCX mirrors it)")
    ap.add_argument("--no-docx", action="store_true", help="Skip DOCX output")
    ap.add_argument("--no-pdf", action="store_true", help="Skip PDF output")
    ap.add_argument(
        "--keywords", default="", help="Comma-separated keywords to verify (ATS check)"
    )
    ap.add_argument(
        "--photo",
        default=None,
        help="Path to a headshot; centers+crops to 3:4 and embeds in the header",
    )
    ap.add_argument(
        "--all",
        action="store_true",
        help="Build both variants (with and without photo). Use with --photo.",
    )
    args = ap.parse_args(argv)

    src = Path(args.source).resolve()
    if not src.exists():
        print(f"error: source not found: {src}", file=sys.stderr)
        return 2

    is_cover = "cover" in src.stem.lower() or "coverletter" in src.name.lower()
    build_dir = ROOT / ".build"
    build_dir.mkdir(exist_ok=True)

    # Photo preprocessing (if requested) — done once and shared by both
    # variants so we only crop the image a single time.
    photo_uri = ""
    photo_file: Path | None = None
    if args.photo:
        photo_path = Path(args.photo).resolve()
        if not photo_path.exists():
            print(f"error: photo not found: {photo_path}", file=sys.stderr)
            return 2
        photo_uri, photo_file = process_photo(
            photo_path,
            build_dir,
            stem=src.stem or "headshot",
        )

    # Variants: (photo?, label_for_stem)
    variants: list[tuple[bool, str | None]]
    if args.all:
        variants = [(False, ""), (True, "_Photo")]
    elif args.photo:
        variants = [(True, "_Photo")]
    else:
        variants = [(False, "")]

    # Cover letters don't take a photo; the loop only runs once with photo=False.
    if is_cover:
        variants = [(False, "")]

    # Author name and output filename are both seeded from the source H1.
    # A cover letter's H1 is the company, so its author comes from the master.
    master_name = _name_from_md(MASTER_MD) if MASTER_MD.exists() else DEFAULT_NAME
    doc_name = master_name if is_cover else _name_from_md(src)

    rc = 0
    for want_photo, suffix in variants:
        # Decide output base.
        if args.out:
            base = Path(args.out).resolve()
            if suffix and not str(base).endswith(suffix):
                base = base.with_name(base.stem + suffix + base.suffix)
        elif is_cover:
            base = (src.parent / src.stem).resolve()
        elif src == MASTER_MD:
            base = (ROOT / (_stem_for(doc_name) + (suffix or ""))).resolve()
        else:
            base = (src.parent / (src.stem + (suffix or ""))).resolve()
        pdf_path = base.with_suffix(".pdf")
        docx_path = base.with_suffix(".docx")

        use_photo = want_photo and photo_uri
        # `want_photo and photo_file` widens to Path | Literal[False] | None,
        # which is not a valid Path | None. Spell it out instead.
        use_photo_file = photo_file if want_photo else None

        if is_cover:
            c = parse_cover_letter(src)
            if not c.name:
                c.name = master_name
            html_str, title = render_cover_html(c)
            html_path = build_dir / (src.stem + (suffix or "") + ".html")
            html_path.write_text(html_str, encoding="utf-8")
            if not args.no_pdf:
                html_to_pdf(
                    html_path,
                    pdf_path,
                    title=title,
                    author=c.name,
                    subject="Cover Letter",
                    keywords="resume, cover letter, " + c.company,
                )
                print(f"PDF  -> {pdf_path}")
            if not args.no_docx:
                build_cover_docx(c, docx_path)
                print(f"DOCX -> {docx_path}")
        else:
            r = parse_resume(src)
            html_str, title = render_resume_html(
                r,
                photo_data_uri=photo_uri if use_photo else "",
            )
            html_path = build_dir / (src.stem + (suffix or "") + ".html")
            html_path.write_text(html_str, encoding="utf-8")
            if not args.no_pdf:
                keywords = args.keywords or f"resume, CV, {r.name}"
                html_to_pdf(
                    html_path,
                    pdf_path,
                    title=title,
                    author=r.name,
                    subject="Resume",
                    keywords=keywords,
                )
                print(f"PDF  -> {pdf_path}")
            if not args.no_docx:
                build_resume_docx(r, docx_path, photo_path=use_photo_file)
                print(f"DOCX -> {docx_path}")

        if args.preview and suffix == "":
            shutil.copy(html_path.as_posix(), (ROOT / "preview.html").as_posix())
            print(f"HTML -> preview.html")

        if args.check and not is_cover and suffix == "":
            rc = ats_check(pdf_path, args.keywords, name=doc_name) or rc

    return rc


def _page_text(page: fitz.Page) -> str:
    """Plain text of one PDF page.

    pymupdf's annotations type `get_text()` as `str | list | dict` because the
    same method serves the "dict"/"json"/"rawdict" output modes. The default
    "text" mode is always a str; narrow it here so callers stay clean.
    """
    text = page.get_text("text")
    return text if isinstance(text, str) else str(text)


def ats_check(pdf_path: Path, keywords_arg: str = "", name: str = "") -> int:
    """Verify the produced PDF parses cleanly and contains key info.

    Notes:
    - Section headers use letter-spacing for visual rhythm; pymupdf's text
      extraction returns them with inter-character spaces (e.g. "S U M M A R Y").
      We normalise whitespace before substring matching.
    - Contact links are verified by **link target**, not display text. The
      visible label may be shortened ("yourname", "your-handle") while the URI is the
      full URL, so matching on text alone would give false failures.
    - Single-page is a soft warning; everything else is hard.
    - Keyword matches are warnings — they inform tailoring, not correctness.
    """
    doc = fitz.open(pdf_path)
    n_pages = len(doc)
    text = "\n".join(_page_text(p) for p in doc)

    # --- link targets (clickability) ---
    uris: list[str] = []
    for p in doc:
        for l in p.get_links():
            u = (l.get("uri") or "").strip()
            if u:
                uris.append(u)
    uri_blob = " ".join(uris).lower()
    n_links = len(uris)

    compact = re.sub(r"\s+", "", text).lower()
    page_warn_only = n_pages > 1

    # Phone: an international-style number anywhere (+NN ...)
    phone_in_text = re.search(r"\+\d{1,3}[\s\d().-]{7,}", text) is not None

    checks: dict[str, bool] = {
        f"single page (n={n_pages})": n_pages == 1,
        "selectable text (>1500 chars)": len(compact) > 1500,
        # Expected name is read from the source H1. If it is unknown the check
        # passes rather than failing on a name we never asserted.
        "name present": (not name or re.sub(r"\s+", "", name).lower() in compact),
        "email text present": re.search(r"[\w.+-]+@[\w-]+\.[\w.]+", text) is not None,
        "phone text present": phone_in_text,
        "link: mailto": any(u.lower().startswith("mailto:") for u in uris),
        "link: tel": any(u.lower().startswith("tel:") for u in uris),
        "link: linkedin": "linkedin.com" in uri_blob,
        "link: github": "github.com" in uri_blob,
        "has embedded links (>=10)": n_links >= 10,
        "section Summary": "summary" in compact,
        "section Experience": "experience" in compact,
        "section Education": "education" in compact,
        "section Skills": "skills" in compact,
        "no stray markdown": not re.search(r"\*\*|\[.*\]\(.*\)", text),
    }
    if keywords_arg:
        for kw in [k.strip() for k in keywords_arg.split(",")]:
            checks[f"kw:{kw}"] = re.sub(r"\s+", "", kw).lower() in compact

    failed = False
    for k, v in checks.items():
        is_page_warn = k.startswith("single page") and page_warn_only
        is_kw_warn = k.startswith("kw:")
        soft = is_page_warn or is_kw_warn
        prefix = "PASS " if v else ("WARN " if soft else "FAIL ")
        print(prefix + k)
        if not v and not soft:
            failed = True
    if failed:
        print("ATS check FAILED (structural)")
        return 1
    page_word = "page" if n_pages == 1 else "pages"
    print(
        f"ATS extraction OK ({len(text)} chars, {n_pages} {page_word}, {n_links} links)"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
