---
name: deepseek-itinerary
description: Generate professionally formatted travel itinerary DOCX files from source .doc itinerary documents. Use when you have a .doc itinerary table and need a polished .docx output with day-by-day breakdowns, sight descriptions, meal/hotel annotations, and a cover page.
---

# Itinerary Generator

Generate beautiful, professional .docx travel itineraries from source .doc files.

## Quick Start

```bash
# 1. Extract text from .doc source
antiword -m UTF-8 "input.doc"

# 2. Run the generation script
python .reasonix/skills/deepseek-itinerary/scripts/generate.py
```

## Workflow

### 1. Read the source .doc

Use `antiword` to extract text from .doc files (python-docx cannot read old .doc format):

```bash
antiword -m UTF-8 "path/to/source.doc"
```

The output is a structured table with columns: Day | Route | Itinerary Content | Meals | Hotel.

### 2. Parse the data

Extract these fields for each day:
- **Day number** and **route** (e.g., D1: Singapore/Shanghai -> Changshu -> Taizhou)
- **Time slots**: morning, afternoon, evening, night
- **Sight name**, including grade tags like "National 5A", "World Heritage"
- **Seasonal notes**: flower seasons, replacement sights
- **Meals**: type, cuisine style, price
- **Hotel**: star rating and name
- **Self-paid items** (optional/at own expense)
- **Guaranteed optional package** items

### 3. Expand descriptions

Convert terse bullet points into **100-150 Chinese characters** of factual description suitable for **foreign travelers** who do not know Chinese culture.

#### Core Rules

1. **Count characters precisely.** 100-150 Chinese characters is the target. Fewer than 100 reads as too thin; more than 150 becomes excessive. Use a counting tool to verify every description.

2. **Describe, do not praise.** Foreign readers have no emotional connection to the place. Tell them concretely what they are looking at and why it matters.

3. **Assume zero cultural knowledge.** Do not assume the reader knows who historical figures are, what architectural terms mean, or why events matter. Every reference needs a one-sentence explanation of its context.

4. **Never comment on physical difficulty or pace.** Phrases like "全程步行轻松" (the walk is easy throughout), "适合拍照" (great for photos), or "值得一游" (worth a visit) are subjective judgments that waste characters and tell the reader nothing factual. Describe the path objectively — its length, surface, and gradient — and let readers decide for themselves.

#### The Four-Part Structure

Every description should answer these four questions, in order:

| # | Question | What to write |
|---|----------|---------------|
| 1 | **What is this?** | Place type and basic facts (size, age, location) |
| 2 | **Why is it famous?** | UNESCO/5A rating, historical significance, unique feature |
| 3 | **What can I see/do?** | Observable features, activities, practical layout |
| 4 | **Any practical info?** | Seasonal timing, self-paid items, note closures |

**Q1 + Q3 are mandatory** for every sight.

#### Verification

After writing all descriptions, run a character count check:

```python
for sight in day.sights:
    cn = sum(1 for c in sight.description if '\u4e00' <= c <= '\u9fff')
    if cn < 100 or cn > 150:
        print(f'WARNING: {sight.name} has {cn} chars (target 100-150)')
```

### 4. Generate the DOCX

Edit `scripts/generate.py` for the specific itinerary, then run it. The script produces:
- **Title page**: centered title, subtitle, highlights
- **Day tables**: each day has a 2-column table (Time | Sight Content) with a header row + time/sight rows, followed by optional [Meals] and [Hotel] summary rows at the bottom
- **Pricing section**: tiered pricing, inclusions/exclusions

### 5. Table Layout Rules

#### Two columns only
Every day itinerary table MUST have exactly two columns:

| Col | Width | Content |
|-----|-------|---------|
| Time | 2.2 cm | morning / afternoon / evening / night |
| Sight | 14.8 cm | sight name + description, transit notes |

There is NO third column. The old 3-column layout (Time / Sight / Meal) is deprecated and MUST NOT be used.

#### Meals: [Meals] row only
- Meal information appears ONLY in a dedicated `[Meals]` row at the bottom of the day table, before the `[Hotel]` row.
- The `[Meals]` row uses two merged cells: left cell contains the label `[Meals]`, right cell contains the meal details.
- Format: `[Meals]  中餐：... | 晚餐：...` (pipe-separated, ordered 早餐/中餐/晚餐).
- If a meal is not arranged, omit it entirely — do not write "无" or "自理".
- If the source has no meal data at all for a day, skip the `[Meals]` row entirely.
- Never write meal info in the sight description cells or in a side column.

#### Hotel: [Hotel] row
- Hotel info appears in a dedicated `[Hotel]` row at the very bottom of the day table.
- Same two-column merged layout: left cell `[Hotel]`, right cell the hotel details.
- If there is no hotel (e.g. departure day), skip the `[Hotel]` row.

#### No blank lines between days
- Every day block must sit directly against the next — no empty paragraphs, no spacer paragraphs, no blank table rows between them.
- The `[Hotel]` row of Day N is immediately followed by the `D-N+1` title paragraph of the next day.
- Day title paragraphs (`D-xx  第x天`) serve as the visual separator between days; do not add extra blank space before or after them.

#### No page breaks between days (mandatory)
- Do NOT insert page breaks between day blocks. All days must flow continuously on the same pages.
- The only permitted page break is between the cover page and the first day (D1).
- Between any two consecutive days (e.g., D2 end → D3 start), there must be NO page break.
- A page break after the last day before the pricing section is also acceptable only if needed to prevent the pricing from being split awkwardly; otherwise omit it.

#### No blank lines within a day
- Time-slot rows (上午/下午/傍晚/晚间) must not be separated by empty paragraphs.
- The `[Meals]` row must sit directly after the last sight row.
- The `[Hotel]` row must sit directly after `[Meals]` (or after the last sight row if no meals).
- Sight name and description paragraphs inside a cell: normal line breaks are fine, but never insert a fully empty paragraph as a spacer.

#### No blank table rows
- Every row in every day itinerary table must contain visible content.
- A row where all cells' text is empty (or whitespace-only) after stripping is forbidden and must be removed before saving.
- Do NOT add empty rows for spacing, separation, or visual padding.
- The last row of every table must be a content row: either a sight row, `[Meals]`, or `[Hotel]`.
- If `[Hotel]` has no data for a day, do not emit the row at all.
- Call `cleanup_empty_rows(table)` on every day table before `doc.save()` as a safety net.

#### Footer: PAGE/NUMPAGES (mandatory)
- Every page must have a centered footer showing current page / total pages, e.g., `3/8`.
- Must use Word dynamic fields `PAGE` and `NUMPAGES` — NEVER hardcode numbers.
- Format: `PAGE/NUMPAGES` with a forward slash, font Arial 8pt, color LIGHT_GRAY.
- Footer must be applied to every section. Use `add_page_number_footer(section)` after section setup.
- Footer must contain NOTHING else — no file names, dates, company names, or itinerary names.
- Do NOT display "报价日期", "quote date", or travel-agency contact info in footers or body text.

#### Regression checklist
Before finalizing output, verify:
- Header row has exactly 2 cells: "时间" and "观光内容" — no "用餐".
- No third column exists anywhere in day tables.
- Every `[Meals]` row uses merged 2-cell layout.
- Every `[Hotel]` row uses merged 2-cell layout.
- No blank paragraph between any two day blocks.
- No blank table row at the end of any day table.
- No blank table row between `[Meals]` and `[Hotel]`.
- No blank table row between `[Hotel]` and end-of-table.

## Styling Standards

### Page Setup
- A4 portrait (21 x 29.7 cm), 2 cm margins

### Colors
```
DEEP_BLUE   = RGBColor(0x1A, 0x3C, 0x6E)   # headers, badges
DARK_GREEN  = RGBColor(0x2E, 0x6B, 0x4F)   # sight names
WARM_GOLD   = RGBColor(0xB8, 0x86, 0x2C)   # decorative, pricing
DARK_GRAY   = RGBColor(0x33, 0x33, 0x33)   # body text
MEDIUM_GRAY = RGBColor(0x66, 0x66, 0x66)   # secondary text
LIGHT_GRAY  = RGBColor(0x99, 0x99, 0x99)   # footers
```

### Cell Backgrounds
```python
BG_LIGHT  = 'EDF2F9'   # time column, headers
BG_GREEN  = 'E8F5ED'   # [Meals] summary row
BG_RED    = 'FFF1F0'   # [Hotel] row
```

### Table Structure (2 columns)
| Col | Width | Content |
|-----|-------|---------|
| Time | 2.2 cm | morning / afternoon / evening / night |
| Sight | 14.8 cm | name + description, transit notes |

**Meals and hotel use merged bottom rows, NOT a third column.**

## Critical Pitfalls (read carefully before coding)

### Emoji characters
Python source files with UTF-8 encoding CANNOT contain surrogate-encoded emojis. Use plain text labels like `[Meals]`, `[Hotel]`.

### Two-column day tables (mandatory)
Day itinerary tables must use exactly 2 columns (Time | Sight Content). There is NO third "用餐" column. Meal info goes ONLY in the `[Meals]` merged row at the bottom of the table. Hotel info goes ONLY in the `[Hotel]` merged row.

### No blank paragraphs between days (mandatory)
Never insert empty paragraphs, spacer paragraphs, or blank lines between day blocks. The `[Hotel]` row of one day must be immediately followed by the next day's title paragraph (`D-xx`).

### No empty table rows (mandatory)
Never generate blank/empty table rows. The last row of every day table must contain content. Call `cleanup_empty_rows()` on every day table before saving to strip any accidental whitespace-only rows.

### No quote-date line in body or footer (mandatory)
Do NOT include "报价日期" (quotation date) anywhere in the document — not in the cover page, footer, header, or pricing section. The travel agency name may appear on the cover page only; it must never appear in footers.

### PAGE/NUMPAGES footer (mandatory)
Every page must have a centered footer with Word dynamic PAGE/NUMPAGES fields (e.g., 3/8). Never hardcode page numbers. Use `add_page_number_footer(section)` on every section. Footer must contain only the page numbers — no text, no dates, no company names.

### No page breaks between days (mandatory)
Do NOT insert page breaks between day blocks. All days flow continuously. The only permitted page break is between the cover page and D1.

### Color `.name` attribute
`RGBColor` objects do NOT have a `.name` attribute. Use hex strings directly:

```python
set_cell_shading(cell, DEEP_BLUE_HEX)  # RIGHT
```

### `cell()` vs `cells[]`
Use `row.cells[i]`, NOT `row.cell(i)`.

### Font setup for Chinese characters
Always set both `rFonts` name AND east-Asian font:

```python
run.font.name = 'Arial'
run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
```

### `WD_ALIGN_VERTICAL.CENTER` serializes to wrong XML (NEW - June 2025)
`WD_ALIGN_VERTICAL.CENTER` is an enum value. When formatted into XML it produces `"CENTER (1)"`, NOT `"center"`. Word rejects the file.

```python
# WRONG
set_cell_va(cell, WD_ALIGN_VERTICAL.CENTER)

# RIGHT - pass the raw string
set_cell_va(cell, 'center')
```

Define `set_cell_va` with a string literal:

```python
def set_cell_va(cell, align='center'):
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall('{%s}vAlign' % NS):
        tcPr.remove(existing)
    tcPr.append(parse_xml('<w:vAlign {} w:val="{}"/>'.format(nsdecls("w"), align)))
```

### `bold=False` generates `<w:b w:val="0"/>` that Word rejects (NEW)
When `run.font.bold = False` is set, python-docx writes `<w:b w:val="0"/>`. Some Word versions refuse to open the file.

```python
# WRONG - writes <w:b w:val="0"/>
run.font.bold = False

# RIGHT - only set bold when True; use bold=None as default
run.font.bold = True         # writes <w:b/>
```

### `set_cell_shading` creates duplicate `<w:shd>` elements (NEW)
Every call appends a new element. Remove existing shd first:

```python
NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
for existing in tcPr.findall('{%s}shd' % NS):
    tcPr.remove(existing)
tcPr.append(shading_elm)
```

### DOCX output rejected by Word: clean up the ZIP package (NEW)
python-docx includes `customXml/`, `stylesWithEffects.xml`, `thumbnail.jpeg` that break Word. After `doc.save()`, run `cleanup_docx(output_path)` which:
1. Removes those parts from the ZIP
2. Removes their references from `[Content_Types].xml`, `_rels/.rels`, `word/_rels/document.xml.rels`

Full code in `scripts/generate.py` template.

### .doc extraction via antiword may fail on WPS-created files (NEW)
Fall back to `olefile + UTF-16LE` binary extraction.

### Run scripts with `PYTHONIOENCODING=utf-8` in Git Bash on Windows (NEW)
The default stdout encoding may be `gbk`. Always prefix:

```bash
PYTHONIOENCODING=utf-8 python .reasonix/skills/deepseek-itinerary/scripts/generate.py
```

## References

- Script template: `.reasonix/skills/deepseek-itinerary/scripts/generate.py`
- Formatting guide: `.reasonix/skills/deepseek-itinerary/references/formatting-guide.md`
