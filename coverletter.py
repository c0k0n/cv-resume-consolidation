#!/usr/bin/env python3
"""Cover letter generator.

Generates a Markdown file for a specific job application, then build.py turns
it into a polished PDF + DOCX.

Usage:
  uv run python coverletter.py "Company" "Role" [body_file_or_text]

Examples:
  uv run python coverletter.py "Northwind Logistics" "ITSM Analyst" \
      "I have long admired Northwind's ... In my last role I handled 210 tickets ..."

If the third argument points to an existing file, its contents are read; else
the third argument is treated as inline body text. To skip and write a
placeholder, omit it.

Output:
  applications/<company>-<role>-cover-letter.md
"""

from __future__ import annotations

import re
import sys
from collections.abc import Callable
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TPL_PATH = ROOT / "coverletter-template.md"


MASTER_MD = ROOT / "resume.md"
HTML_COMMENT_RE = re.compile(r"<!--.*?-->", re.DOTALL)

# Used only when the master resume is absent — a fork of this repo has no
# personal data in it, so the letter must still generate cleanly.
PLACEHOLDER_CONTACT = (
    "+00 000 0000 | [you@example.com](mailto:you@example.com) | "
    "[linkedin.com/in/your-handle](https://www.linkedin.com/in/your-handle) | "
    "[github.com/your-handle](https://github.com/your-handle) | City, Country"
)
PLACEHOLDER_NAME = "Your Name"


def slugify(s: str) -> str:
    s = s.lower().strip()
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


def _from_master(pick: Callable[[str], str | None]) -> str | None:
    """Return the first line of the master resume that `pick` accepts."""
    if not MASTER_MD.exists():
        return None
    text = HTML_COMMENT_RE.sub("", MASTER_MD.read_text(encoding="utf-8"))
    for line in text.splitlines():
        value = pick(line.strip())
        if value:
            return value
    return None


def _contact_line() -> str:
    """Reuse the contact line from the master so letters always match it."""

    def pick(s: str) -> str | None:
        if not s or s.startswith("#"):
            return None
        if "@" in s or "|" in s or re.search(r"\+\d", s):
            return s
        return None

    return _from_master(pick) or PLACEHOLDER_CONTACT


def _signatory() -> str:
    """The candidate's name, read from the master resume's H1."""

    def pick(s: str) -> str | None:
        return s[2:].strip() if s.startswith("# ") else None

    return _from_master(pick) or PLACEHOLDER_NAME


def main() -> None:
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)

    company, role = sys.argv[1], sys.argv[2]
    body = sys.argv[3] if len(sys.argv) > 3 else None

    if body and len(body) < 200:
        # Treat as a path; fall back to inline if the file doesn't exist.
        p = Path(body)
        if p.exists():
            body = p.read_text(encoding="utf-8")

    if not body:
        body = (
            "[PARAGRAPH 1 — why this company. Name something specific: a "
            "product, a market, something they published. Generic praise "
            "reads as spam.]\n\n"
            "[PARAGRAPH 2 — your fit. Pick 2-3 requirements from the job ad "
            "and show the evidence. Mirror their wording; ATS and humans "
            "both reward it.]"
        )

    out_dir = ROOT / "applications"
    out_dir.mkdir(exist_ok=True)
    fname = f"{slugify(company)}-{slugify(role)}-cover-letter.md"
    out = out_dir / fname

    tpl = TPL_PATH.read_text(encoding="utf-8")
    body = body.strip()
    rendered = (
        tpl.replace("{Company}", company)
        .replace("{company}", company)
        .replace("{Role}", role)
        .replace("{role}", role)
        .replace("{CONTACT_LINE}", _contact_line())
        .replace("{NAME}", _signatory())
        .replace("{BODY_PARAGRAPHS}", body)
    )
    out.write_text(rendered, encoding="utf-8")
    print(f"Cover letter -> {out}")
    print()
    print("Next:")
    print(f"  uv run python build.py {out.relative_to(ROOT)}")
    print()
    print("Or edit the markdown first to polish the body paragraphs,")
    print("then run the build command above.")


if __name__ == "__main__":
    main()
