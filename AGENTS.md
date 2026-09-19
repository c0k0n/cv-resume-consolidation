# AGENTS.md

Operating rules for this repository. Read before editing anything.

## What this is

A résumé and cover-letter production system. Markdown in; ATS-safe,
single-page PDF and DOCX out. PDF goes through headless Chrome, DOCX through
`python-docx`. The design system lives entirely in `templates/styles.css`.

## The privacy invariant (do not break this)

This repository is **public**. It ships sample content only. The paths that hold
a real person's data are gitignored, anchored to the root:

| Path | Rule |
| --- | --- |
| `/resume.md` | never commit |
| `/MASTER_RECORD.md` | never commit |
| `/applications/` | never commit |
| `/photo.jpg` | never commit |
| `/.workbuddy-ai/` | never commit |
| `examples/`, `templates/`, `docs/` | safe to commit |

Consequences:

- **No personal name, email, phone, employer or URL may be hardcoded in any
  tool.** The candidate's name is always read from the H1 of the Markdown being
  built. If you find yourself writing a literal name into `build.py`,
  `verify.py`, `coverletter.py`, `check-docx.ps1` or a template, you are doing it
  wrong — derive it.
- Generated `.pdf`/`.docx` are gitignored. Do not force-add them.
- Docs and graphics must use the sample persona, never real details.

## File map

| Path | Role |
| --- | --- |
| `resume.md` | Master résumé. **Private, gitignored.** The file you edit for real content. |
| `MASTER_RECORD.md` | Full career record — the fact source behind any résumé. **Private, gitignored.** |
| `examples/resume.md` | Public sample. The starting point for anyone forking. |
| `examples/cover-letter.md` | A finished sample letter. |
| `coverletter-template.md` | Skeleton every generated letter starts from. Uses `{CONTACT_LINE}` and `{NAME}` tokens. |
| `applications/` | Per-job résumés and letters. **Private, gitignored.** |
| `build.py` | Markdown → HTML → PDF (Chrome) + DOCX, plus `--check`. Supports `--photo` and `--all`. |
| `coverletter.py` | Drafts a letter for a company/role; pulls contact and name from `resume.md`. |
| `verify.py` | Output test suite: ATS, fonts, layout, links. |
| `check-docx.ps1` | Word COM pagination check — the only way to catch DOCX overflow. |
| `templates/` | `resume.html`, `cover.html`, `styles.css`. The design system. |
| `README.md` | Public showcase and usage guide. |

## Rules when editing

1. **Truth over polish.** Never strengthen a title, metric or claim beyond what
   the record supports. Every line must be defensible in an interview.
2. **Single page is the goal.** After any content edit, run
   `uv run python build.py --check`. If a variant overflows, **cut content** —
   never shrink type below 8.5pt body or cram with negative spacing.
3. **ATS-safe formatting only**: single column, no tables/text boxes/icons for
   layout, standard section headings, bullets starting with verbs or numbers,
   metrics kept, dates in `Month Year` form.
4. **Tailoring**: copy the master to `applications/<target>.md`, adjust the
   summary, reorder or extend skill rows, and mirror 3–5 keywords that are
   genuinely true. Then
   `uv run python build.py applications/<target>.md --check --keywords "k1, k2"`.
5. Tooling is managed with **uv** (`pyproject.toml` + `uv.lock`). Run scripts
   with `uv run python build.py ...`.
6. PDF rendering needs headless **Chrome or Edge** (`--print-to-pdf`).

## Markdown schema — résumé

```
# Jordan Ellis                             ← H1: name (seeds filename + author)
IT Operations, Data & Support Professional ← optional subtitle
+60 … | [you@example.com](mailto:…) | …    ← contact line, split on "|"
Open to graduate, office, …                ← optional italic line

## Summary                                 ← free-text paragraph
...

## Skills
**Category:** comma, separated, items      ← `**Category:**` renders navy

## Experience
### Role Title                             ← H3: entry
[Org](url) — City, Country · Mon YYYY – Mon YYYY   ← italic meta line
- bullet
- bullet

## Education
### Degree
[University](url), Country · Mon YYYY – Mon YYYY
CGPA 3.68/4.0 | Dean's List

## Certifications / Publication / Volunteering / Languages / References
- **Cert Name** — detail with link.
```

Headings are classified by keyword into canonical buckets (`summary`,
`experience`, `education`, `skills`, `projects`, `certifications`,
`publication`, `volunteering`, `languages`, `references`). Unrecognised
headings still render, after the known ones, so nothing is silently dropped.
Render order is `SECTION_ORDER` in `build.py`, shared by both renderers.

## Markdown schema — cover letter

```
# Company — Role                  ← H1; the second half is the role applied for
Contact line                      ← injected from resume.md by coverletter.py
Dear Hiring Manager,

Paragraph one — why THIS company, specifically.

Paragraph two — your fit against 2-3 requirements.

Yours sincerely,
Your Name
```

An optional `## Recipient` block overrides name / title / company / address.
The H1's second half is the **role**, never the recipient's title.

## Content rules

1. **No duplicated facts.** A number appears in exactly one place. CGPA belongs
   in Education, not the Summary.
2. **Link organisations, not keywords.** Link employer, university, credential,
   DOI, inbox, GitHub profile. Never link a technology name inside a bullet.
3. **No credential score flexing.** "EF SET English Certificate", not
   "C2 Proficient (76/100)".
4. **Links carry no underline**, in both PDF and DOCX.
5. **Link text must be the full URL** (`linkedin.com/in/you`), not a bare
   handle — ATS reads display text, not the link target.
6. **No "sage" voice.** No "quick to pick up new systems", "calm under
   pressure", "comfortable across…". State facts, not character traits.
7. **Filenames use underscores** — they survive every upload form.

## Hard-won rendering facts (do not re-derive)

1. **Never put `"Inter"` in the font stack.** Chrome renders it as **Type3**
   (glyph-outline) in `--print-to-pdf`, and some ATS parsers cannot read Type3.
   Calibri and Arial embed as real TrueType. The stack is
   `"Calibri", "Carlito", "Segoe UI", Arial, Helvetica, sans-serif`.
2. **Size rules and bullet markers in whole CSS pixels.** Chrome lays out at
   96dpi; `height: 1pt` = 1.333px gets rounded to 1px or 2px depending on the
   element's Y position, so rules rendered at 0.75pt and 1.5pt on one page.
   `height: 2px` snaps identically. Rules are drawn as filled `::after` boxes,
   not CSS borders, because Chrome clamps thin border strokes to ~0.75pt.
3. **Word "multiple 1.3" is not CSS `line-height: 1.3`.** Word multiplies the
   font's natural line height (~1.22em), giving ~1.59em — about 22% looser,
   which silently pushed the DOCX to two pages while the PDF stayed on one.
   `_apply_exact_line_spacing()` writes points with `w:lineRule="exact"`.
   Mirror any CSS line-height change there, and in the DOCX metrics block.
4. **A DOCX photo must be a floating `<wp:anchor>`**, not `add_picture()`
   (`<wp:inline>` stacks above the text and adds ~36mm).
5. **Chrome holds the PDF handle briefly** after `--print-to-pdf`. Metadata
   rewriting writes a `.tmp`, retries `os.replace`, then falls back to a byte
   copy.
6. **Apply bold before extracting links.** `**[text](url)**` leaves literal
   `**` in the output if done the other way round.
7. **`entry.org` / `.meta` / `.title` keep raw markdown.** Strip at parse time
   and employer links silently stop working.
8. **Section-heading letter-spacing** makes pymupdf return "S U M M A R Y" —
   ATS checks normalise whitespace before matching.
9. **`break-after: avoid` on section titles** orphans the title. Use
   `break-inside: avoid` on `.section`.
10. **`.header { display: flow-root; }`** — without a block formatting context
    the floated headshot escapes the header and overlaps the first heading.

## Verification split

`verify.py` checks the **PDF** only. `python-docx` has no layout engine, so DOCX
pagination must be checked with `check-docx.ps1` (Word COM). Neither substitutes
for the other.

## Standing brief

Produce the best résumé/cover-letter pair either output can offer: passes ATS,
parses cleanly for every parser and LLM screener, and reads well to a human.
When studying competitor output, take **principles**, never structure or
phrasing. Keep the directory lean — no scratch renders, no superseded variants,
no history of experiments.
