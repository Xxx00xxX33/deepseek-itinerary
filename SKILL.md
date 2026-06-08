---
name: deepseek-itinerary
description: >
  Generate professionally formatted travel itinerary DOCX files from source
  .doc itinerary documents. Use when: (1) you have a .doc itinerary table and
  need a polished .docx output; (2) you need day-by-day travel formatting with
  sight descriptions, meal/hotel annotations, and a cover page; (3) you need to
  expand terse itinerary tables into 100-150 character explanatory descriptions
  for foreign travelers. Works with python-docx on Windows.
---

# Itinerary Generator

Generate beautiful, professional .docx travel itineraries from source .doc files.

## Quick Start

```bash
# 1. Extract text from .doc source
antiword -m UTF-8 "input.doc"

# 2. Run the generation script
python scripts/generate.py
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

#### The Three Rules

1. **Count characters precisely.** 100-150 Chinese characters is the target. Fewer than 100 reads as too thin; more than 150 becomes excessive. Use a counting tool to verify every description. The first version of a real itinerary failed because most descriptions were only 70-98 characters.

2. **Describe, do not praise.** Foreign readers have no emotional connection to the place. Tell them concretely what they are looking at and why it matters.

3. **Assume zero cultural knowledge.** Do not assume the reader knows who historical figures are, what architectural terms mean, or why events matter. Every reference needs a one-sentence explanation of its context.

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
- **Day tables**: each day has a header row + time/sight/meal rows
- **Pricing section**: tiered pricing, inclusions/exclusions

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
BG_GREEN  = 'E8F5ED'   # meal summary row
BG_RED    = 'FFF1F0'   # hotel row
```

### Table Structure (3 columns)
| Col | Width | Content |
|-----|-------|---------|
| Time | 2.2 cm | morning/afternoon |
| Sight | 11.5 cm | name + description |
| Meal | 3.3 cm | meal info |

## Critical Pitfalls (read carefully before coding)

### Emoji characters
Python source files with UTF-8 encoding CANNOT contain surrogate-encoded emojis. Use plain text labels like `[Meal]`, `[Hotel]`.

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
Fall back to `olefile + UTF-16LE` binary extraction. See SKILL.md for code.

### Run scripts with `PYTHONIOENCODING=utf-8` in Git Bash on Windows (NEW)
The default stdout encoding may be `gbk`. Always prefix:

```bash
PYTHONIOENCODING=utf-8 python scripts/generate.py
```