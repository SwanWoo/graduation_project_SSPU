# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SSPU graduation thesis tool. Converts `final_paper.md` (a Markdown thesis draft) into a formatted `.docx` Word document by injecting content into a university-provided Word template at the XML level.

All active work is in `thesis2docx.py`.

## Commands

```bash
uv sync                    # Install dependencies (pymupdf)
uv run python thesis2docx.py   # Run conversion → outputs 毕业论文_生成.docx
```

No build step, no tests, no linting. Requires `xelatex` (TeX Live) on PATH for formula rendering; falls back to raw LaTeX text if absent.

## Architecture

```
final_paper.md → thesis2docx.py → 毕业论文_生成.docx
                                (reads word/document.xml as template)
```

The `.docx` template is a ZIP of OpenXML files. `thesis2docx.py` manipulates the XML directly with `xml.etree.ElementTree`. The template is fully extracted into `template/` (preserving the ZIP internal structure). The output ZIP is reassembled from these files plus the modified `document.xml`.

### Pipeline (main function, line ~1860)

1. `parse_markdown()` — splits `.md` into typed blocks (heading, paragraph, bullet, numbered, table, image, formula, codeblock, blockquote)
2. `inject_cover_data()` — fills P4-P13 with cover info (title, student ID, name, etc.)
3. `replace_abstracts()` — cleans template formatting instructions from P31-P42, inserts Chinese/English abstract text and keywords
4. `build_toc_entries()` — generates PAGEREF-field-based TOC from body headings, replaces P44+ area
5. Body content loop — converts Markdown blocks to Word paragraphs and appends after the last section break
6. `_update_rels_file()` — patches `document.xml.rels` with image/formula relationship entries
7. ZIP assembly — writes modified XML into new `.docx`, patches `settings.xml` with `updateFields=true`

## Template Paragraph Positions (hardcoded into document.xml)

```
P0-P3:    Cover page header/instructions
P4-P13:   Cover page fields (题目, 学号, 姓名, 班级, 专业, 学部(院), 入学时间, 指导教师, 日期)
P7,P10,P13: Have AlternateContent dropdowns — fully rebuilt, not preserve-and-replace
P14:      Section break (cover → declaration)
P15-P28:  Declaration page
P30:      Section break (declaration → abstract)
P31:      Chinese thesis title (黑体, sz=36)
P32:      "摘要" label (黑体, bold, sz=32)
P33-P35:  Chinese abstract body (sz=24, firstLine=480)
P36:      Chinese keywords
P37:      Formatting note (cleaned/emptied)
P38:      Page break (CN → EN abstract)
P39:      English title (TNR, sz=36, bold, centered)
P40:      "ABSTRACT" (TNR, sz=32, bold, centered)
P41:      English abstract body (TNR, sz=24)
P42:      English keywords
P43:      Section break (abstract → TOC)
P44:      TOC title ("目录")
P45-P58:  TOC entries (replaced dynamically)
P68:      Section break (TOC → body)
P69+:     Body content (all replaced from Markdown)
```

These indices are **fragile** — template changes require re-indexing.

## Style Mapping (template-specific IDs from word/styles.xml)

| Element | Style | Notes |
|---|---|---|
| H1 | `'11'` | Centered, bold, Times New Roman |
| H2 | `'ad'` | Left-aligned, no indent |
| H3 | `'a'` | Left-aligned, line=360 |
| Body | `'aa'` | firstLineChars=200, firstLine=480 (2-char indent) |
| Table caption (`表\d`/`续表\d`) | `'af4'` | Auto-detected in `build_body_paragraph()` |
| Figure caption / blockquote | `'af4'` | Centered, 黑体, sz=21 |
| Image caption | `'af4'` | After downloaded image |
| Bullet | `'aa'` + numId=2 | |
| Numbered | `'aa'` + numId=1 | |
| Code block | `'aa'` | Courier New, no indent |
| TOC L1/L2/L3 | `'12'`/`'20'`/`'30'` | |

**Important:** Style `af3` (Title) inherits bold from its definition. Runs that need to be non-bold must explicitly set `<w:b val="0"/>` (use `unbold=True` parameter in `make_text_run()`).

## Key Patterns

- **`_clean_paragraph_keep_ppr(p)`** — strips all children except `pPr` from a template paragraph, also removes illegal `<w:rPr>` elements that the template leaves inside `<w:pPr>`
- **Namespace registration order matters** — `register_namespaces()` registers 17 namespaces; `r:` prefix registered last for relationship references
- **`rels_counter` dict** — `{'count': 0, '_rels_entries': []}` passed mutably to track image/formula relationships; consumed by `_update_rels_file()`
- **Formula rendering** — XeLaTeX compiles LaTeX to PDF, PyMuPDF converts to PNG at 600 DPI. Numbered formulas use center-tab + right-tab layout for `(2-1)` style numbering. Un-numbered formulas use `jc=center`.
- **Media files accumulate** — `template/word/media/` grows with each run (formula PNGs + downloaded images). Gitignored via `.gitignore`.
- **Bookmark counter** — global `_bookmark_counter`, scans template for max existing `_Toc` bookmark ID

## Unit Conventions

- Word half-points: sz=44=22pt, sz=36=18pt, sz=32=16pt, sz=24=12pt, sz=21=10.5pt
- Indent: `firstLine=480` twips = 2 Chinese chars
- EMU for images: `pixels / DPI * 72 * 9525`

## Important Files

- `thesis2docx.py` — the converter (all active code)
- `final_paper.md` — thesis source Markdown
- `template/word/document.xml` — template body XML (read-only reference)
- `template/word/styles.xml` — template style definitions
- `pyproject.toml` — uv project config (dependency: `pymupdf>=1.24`)
