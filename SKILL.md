---
name: itinerary-generator
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

Convert terse bullet points into **100-150 Chinese characters** of factual description suitable for **foreign travelers** who do not know Chinese culture:

| Do | Don't |
|----|-------|
| "What this place is" (type, history, size) | "Beautiful", "stunning", "poetic" |
| "Why it is famous" (UNESCO, 5A, historical) | Vague praise |
| "What visitors can see/do" | Flowery language |
| "Practical info" (seasonal timing, transportation) | Assumed cultural knowledge |

### 4. Generate the DOCX

Run `scripts/generate.py` (view it first — it is a template you must edit for the specific itinerary).

The script produces a document with:
- **Title page**: centered title, subtitle, highlights
- **Day tables**: each day has a header row + time/sight/meal rows
- **Pricing section**: tiered pricing, inclusions/exclusions, optional package

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
BG_WARM   = 'FEF6E9'   # meal column
BG_GREEN  = 'E8F5ED'   # meal summary row
BG_RED    = 'FFF1F0'   # hotel row
```

### Table Structure (3 columns)
| Col | Width | Content |
|-----|-------|---------|
| Time | 2.2 cm | morning/afternoon/evening/night |
| Sight | 11.5 cm | name (bold green) + description (gray) |
| Meal | 3.3 cm | meal info, notes, seasonal tags |

### Special Rows
- **Meal row** (`add_meal`): green background, summarizes all meals for the day
- **Hotel row** (`add_hotel`): red-tinted background
- **Free day row**: merge all 3 columns, light gray background

## Critical Pitfalls (read carefully before coding)

### Emoji characters
Python source files with UTF-8 encoding CANNOT contain surrogate-encoded emojis. Never write:

```python
# WRONG - causes UnicodeEncodeError: surrogates not allowed
'\uD83C\uDF7D'  # 🍽 as surrogate pair
```

Instead use:

```python
# RIGHT
'[Meal]'            # simple text label (preferred - safest)
'\U0001F37D'        # full Unicode codepoint (works in Python 3)
# OR just use plain text labels like [餐食], [住宿]
```

Never use emoji in the source file at all unless you are absolutely certain they are encoded as proper non-surrogate UTF-8. When in doubt, use plain text labels like `[Meal]`, `[Hotel]`, `[Free Time]`.

### Color `.name` attribute
`RGBColor` objects do NOT have a `.name` attribute. Use hex strings directly:

```python
# WRONG
set_cell_shading(cell, DEEP_BLUE.name)  # AttributeError!

# RIGHT
DEEP_BLUE_HEX = '1A3C6E'
set_cell_shading(cell, DEEP_BLUE_HEX)
```

### `cell()` vs `cells[]`
Use `row.cells[i]`, NOT `row.cell(i)`:

```python
# WRONG
cell = table.rows[0].cell(0)  # AttributeError: '_Row' object has no attribute 'cell'

# RIGHT
cell = table.rows[0].cells[0]
```

### Font setup for Chinese characters
Always set both `rFonts` name AND east-Asian font:

```python
run.font.name = 'Arial'
run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
```

### Table grid style
Always set `table.style = 'Table Grid'` to ensure visible borders.

### GBK encoding on Windows
When printing to console on Windows with Chinese locale, avoid non-GBK characters:

```python
# WRONG - may fail with UnicodeEncodeError: 'gbk' codec
print('✔ Document saved')

# RIGHT
print('OK - Document saved')
```
