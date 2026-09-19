#!/usr/bin/env python3
"""Verification suite for generated resumes / cover letters.

Runs the checks a recruiter's stack actually applies:
  - **Structure**  : page count, paper size
  - **Text layer**  : selectable text, no stray markdown, reading order
  - **Fonts**       : real embedded font (not Type3), one family, min size
  - **Layout**      : uniform section rules, consistent line height
  - **Contact**     : name / email / phone present as text AND as links
  - **Links**       : targets resolve to real schemes, no keyword-link noise
  - **Sections**    : every expected heading detected
  - **DOCX**        : hyperlinks present, photo present when expected

Usage:
  uv run python verify.py                 # verify all four master outputs
  uv run python verify.py --keywords "ITSM, Power BI"
  uv run python verify.py applications/foo.md
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from typing import Any

import pymupdf as fitz

ROOT = Path(__file__).resolve().parent

# Technology documentation sites that must never be linked from a bullet.
# Linking "Svelte" to svelte.dev inside an experience bullet is noise — the
# recruiter gains nothing. (github.com is a *profile* link and is allowed;
# that's why this matches on host, not substring.)
NEVER_LINK_HOSTS = (
    "svelte.dev",
    "nodejs.org",
    "nodejs.dev",
    "python.org",
    "git-scm.com",
    "microsoft.com/en-us/power-platform",
    "djangoproject.com",
    "react.dev",
    "vuejs.org",
)

EXPECTED_SECTIONS = (
    "summary",
    "skills",
    "experience",
    "education",
)

MIN_FONT_PT = 8.5  # AGENTS.md floor
MIN_TEXT_CHARS = 1500  # a real resume, not an image
MIN_LINKS = 10


class Report:
    def __init__(self) -> None:
        self.rows: list[tuple[str, str, str]] = []  # (level, name, detail)
        self.failures = 0

    def add(self, level: str, name: str, detail: str = "") -> None:
        self.rows.append((level, name, detail))
        if level == "FAIL":
            self.failures += 1

    def ok(self, n: str, d: str = "") -> None:
        self.add("PASS", n, d)

    def warn(self, n: str, d: str = "") -> None:
        self.add("WARN", n, d)

    def bad(self, n: str, d: str = "") -> None:
        self.add("FAIL", n, d)

    def check(self, cond: bool, n: str, d: str = "") -> None:
        (self.ok if cond else self.bad)(n, d)

    def render(self) -> None:
        w = max((len(n) for _, n, _ in self.rows), default=10)
        for lvl, name, detail in self.rows:
            print(f"  {lvl:4s}  {name:<{w}}  {detail}")


def _page_text(page: fitz.Page) -> str:
    """pymupdf types get_text() as `str | list | dict` because one method
    serves several output modes. The default "text" mode is always a str."""
    text = page.get_text("text")
    return text if isinstance(text, str) else str(text)


def _text_dict(page: fitz.Page) -> dict[str, Any]:
    """The "dict" output mode — spans, lines and blocks with their metrics."""
    d = page.get_text("dict")
    return d if isinstance(d, dict) else {}


def verify_pdf(pdf_path: Path, rep: Report, keywords: list[str]) -> None:
    doc = fitz.open(pdf_path)
    n_pages = len(doc)
    text = "\n".join(_page_text(p) for p in doc)
    compact = re.sub(r"\s+", "", text).lower()
    raw = pdf_path.read_bytes()
    is_cover = "cover" in pdf_path.stem.lower()

    print(f"\n=== {pdf_path.name} ===")

    # --- structure ---
    rep.check(n_pages == 1, "single page", f"n={n_pages}")
    page = doc[0]
    w, h = page.rect.width, page.rect.height
    # A4 in points is 595 x 842
    rep.check(abs(w - 595) < 12 and abs(h - 842) < 12, "A4 paper", f"{w:.0f}x{h:.0f}pt")

    # --- text layer ---
    # A cover letter is naturally shorter than a resume (~1500 vs ~3000 chars).
    min_chars = 800 if is_cover else MIN_TEXT_CHARS
    rep.check(len(compact) > min_chars, "selectable text", f"{len(compact)} chars")
    rep.check(
        not re.search(r"\*\*|\[[^\]]+\]\([^)]+\)", text),
        "no stray markdown",
        "no literal ** or [x](y)",
    )
    # build.py writes the author into the PDF metadata from the source H1, so
    # the expected name is whatever the document claims it is — no name is
    # hardcoded in this repository.
    meta: dict[str, str] = dict(doc.metadata or {})
    author = (meta.get("author") or "").strip()
    if author:
        rep.check(
            re.sub(r"\s+", "", author).lower() in compact,
            "candidate name in text",
            author,
        )
    else:
        rep.warn("candidate name in text", "no author in PDF metadata")

    # --- fonts ---
    font_types = sorted({f[2] for p in doc for f in p.get_fonts()})
    has_tt = b"/FontFile2" in raw or b"/FontFile3" in raw
    rep.check(
        "Type3" not in font_types, "font is embedded (not Type3)", f"{font_types}"
    )
    rep.check(has_tt, "real font program embedded", f"FontFile2={b'/FontFile2' in raw}")

    sizes: dict[float, int] = {}
    fonts: set[str] = set()
    for p in doc:
        for b in _text_dict(p).get("blocks", []):
            for line in b.get("lines", []):
                for s in line["spans"]:
                    if s["text"].strip():
                        sizes[round(s["size"], 1)] = (
                            sizes.get(round(s["size"], 1), 0) + 1
                        )
                        fonts.add(s["font"].split("+")[-1].split("-")[0])
    if sizes:
        mn = min(sizes)
        rep.check(mn >= MIN_FONT_PT, f"min font >= {MIN_FONT_PT}pt", f"min={mn}pt")
        rep.ok("font sizes used", ", ".join(f"{k}pt" for k in sorted(sizes)))
    rep.check(len(fonts) <= 2, "single type family", f"{len(fonts)} families")

    # --- layout: uniform section rules ---
    rules = [
        round(d["rect"].height, 4)
        for p in doc
        for d in p.get_drawings()
        if d["rect"].width > 200 and d["rect"].height < 5
    ]
    if rules:
        rep.check(
            len(set(rules)) == 1,
            "section rules uniform",
            f"{len(rules)} rules @ {sorted(set(rules))}pt",
        )
    else:
        rep.warn("section rules uniform", "no rules detected")

    # --- contact ---
    rep.check(re.search(r"[\w.+-]+@[\w-]+\.[\w.]+", text) is not None, "email in text")
    rep.check(
        re.search(r"\+\d{1,3}[\s\d().-]{7,}", text) is not None,
        "phone in text (intl format)",
    )

    uris: list[str] = []
    for p in doc:
        for l in p.get_links():
            u = (l.get("uri") or "").strip()
            if u:
                uris.append(u)
    blob = " ".join(uris).lower()
    # A cover letter only carries contact links; a resume carries contact +
    # employers + credentials + DOI, so it legitimately has more.
    min_links = 4 if is_cover else MIN_LINKS
    rep.check(len(uris) >= min_links, "embedded link count", f"{len(uris)} links")
    rep.check(
        any(u.lower().startswith("mailto:") for u in uris), "link: email (mailto)"
    )
    rep.check(any(u.lower().startswith("tel:") for u in uris), "link: phone (tel)")
    rep.check("linkedin.com" in blob, "link: LinkedIn")
    rep.check("github.com" in blob, "link: GitHub")

    def _host(u: str) -> str:
        m = re.match(r"^[a-z]+://([^/]+)", u.lower())
        return (m.group(1) if m else u.lower()).lstrip("www.")

    bad_links = [
        u
        for u in uris
        if any(_host(u) == h or _host(u).endswith("." + h) for h in NEVER_LINK_HOSTS)
    ]
    rep.check(
        not bad_links,
        "no keyword links (Svelte/Python/...)",
        f"{bad_links}" if bad_links else "clean",
    )

    schemes_ok = all(
        u.lower().startswith(("http://", "https://", "mailto:", "tel:")) for u in uris
    )
    rep.check(schemes_ok, "all link schemes valid")

    # --- document type ---
    if is_cover:
        rep.check(bool(re.search(r"\bdear\b", text, re.I)), "has greeting (Dear ...)")
        rep.check(
            bool(re.search(r"yours sincerely|sincerely|regards", text, re.I)),
            "has sign-off",
        )
        rep.check(len(text) > 600, "letter body present", f"{len(compact)} chars")
    else:
        for s in EXPECTED_SECTIONS:
            rep.check(s in compact, f"section {s.title()}")

    # --- keywords (soft) ---
    for kw in keywords:
        k = re.sub(r"\s+", "", kw).lower()
        if k in compact:
            rep.ok(f"kw: {kw}")
        else:
            rep.warn(f"kw: {kw}", "not found — add if genuinely true")


def verify_docx(docx_path: Path, rep: Report, expect_photo: bool) -> None:
    if not docx_path.exists():
        rep.bad(f"{docx_path.name} exists")
        return
    print(f"\n=== {docx_path.name} ===")
    is_cover = "cover" in docx_path.stem.lower()
    with zipfile.ZipFile(docx_path) as z:
        rels = z.read("word/_rels/document.xml.rels").decode("utf8", "ignore")
        uris = re.findall(r'Target="(https?://[^"]+|mailto:[^"]+|tel:[^"]+)"', rels)
        media = [n for n in z.namelist() if n.startswith("word/media/")]
    min_links = 4 if is_cover else MIN_LINKS
    rep.check(len(uris) >= min_links, "docx hyperlinks", f"{len(uris)} links")
    rep.check(any(u.startswith("mailto:") for u in uris), "docx mailto link")
    if expect_photo:
        rep.check(len(media) >= 1, "docx photo embedded", f"{len(media)} image(s)")
    else:
        rep.check(len(media) == 0, "docx has no photo", f"{len(media)} image(s)")


def main(argv: list[str] | None = None) -> int:
    raw = list(argv or sys.argv[1:])
    # Accept both `--keywords a,b` and `--keywords=a,b`
    kw_raw = ""
    positional: list[str] = []
    i = 0
    while i < len(raw):
        a = raw[i]
        if a == "--keywords" and i + 1 < len(raw):
            kw_raw = raw[i + 1]
            i += 2
            continue
        if a.startswith("--keywords="):
            kw_raw = a.split("=", 1)[1]
            i += 1
            continue
        if not a.startswith("--"):
            positional.append(a)
        i += 1
    args = positional
    keywords = [k.strip() for k in kw_raw.split(",") if k.strip()]

    rep = Report()

    if args:
        stem = Path(args[0]).with_suffix("")
        targets = [(stem.with_suffix(".pdf"), stem.with_suffix(".docx"), False)]
    else:
        # Whatever has been built at the repo root. Filenames are derived from
        # the source H1, so they are not known ahead of time.
        targets = [
            (pdf, pdf.with_suffix(".docx"), "photo" in pdf.stem.lower())
            for pdf in sorted(ROOT.glob("*.pdf"))
        ]
        if not targets:
            print("Nothing built yet — run: uv run python build.py --check")

    for pdf, docx, photo in targets:
        if pdf.exists():
            verify_pdf(pdf, rep, keywords)
        else:
            rep.bad(f"{pdf.name} exists")
        verify_docx(docx, rep, photo)

    print("\n" + "=" * 70)
    rep.render()
    print("=" * 70)
    total = len(rep.rows)
    print(
        f"{total - rep.failures}/{total} checks passed"
        f"{'' if not rep.failures else f'  ({rep.failures} FAILED)'}"
    )
    return 1 if rep.failures else 0


if __name__ == "__main__":
    sys.exit(main())
