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

1. **Count characters precisely.** 100-150 Chinese characters is the target. Fewer than 100 reads as too thin; more than 150 becomes excessive. Use a counting tool (`sum(1 for c in text if '\u4e00' <= c <= '\u9fff')`) to verify every description. The first version of a real itinerary failed because most descriptions were only 70-98 characters.

2. **Describe, do not praise.** Foreign readers have no emotional connection to the place — telling them it is "beautiful", "stunning", "poetic", or "breathtaking" tells them nothing. Instead, tell them concretely what they are looking at and why it matters.

3. **Assume zero cultural knowledge.** Do not assume the reader knows who historical figures are (姜太公, 梅兰芳), what architectural terms mean (文人园林, 回廊, 石舫), or why events matter. Every reference needs a one-sentence explanation of its context.

#### The Four-Part Structure

Every description should answer these four questions, in order:

| # | Question | What to write | Example |
|---|----------|---------------|---------|
| 1 | **What is this?** | Place type (temple, lake, museum, garden, street) and basic facts (size, age, location) | "尚湖是常熟最大的湖泊景区，被评为国家5A级旅游景区。" |
| 2 | **Why is it famous?** | UNESCO/5A rating, historical significance, unique feature | "瘦西湖2014年被列入世界文化遗产名录（大运河沿线遗产点）。" |
| 3 | **What can I see/do?** | Observable features, activities, practical layout | "乘坐竹筏沿水道穿行于杉林之间，水面倒映树影，光线从树冠缝隙洒下。" |
| 4 | **Any practical info?** | Seasonal timing, self-paid items, note closures | "周一闭馆" / "4月牡丹节期间数百个品种盛开" |

Not every description needs all four parts, but **Q1 + Q3 are mandatory** for every sight.

#### Writing Examples

**Bad (praise, no substance, 89 chars):**
> 瘦西湖以"瘦"为美，湖面狭长曲折，两岸花柳依水。漫步湖畔，一步一景，处处是诗。

This fails all three rules: it praises, assumes knowledge of Chinese poetry, and tells the reader nothing factual.

**Good (factual, informative, 118 chars):**
> 瘦西湖是扬州最核心的景区，2014年被列入世界文化遗产名录。湖区呈狭长形，沿湖分布五亭桥、白塔、二十四桥等清代建筑，这些建筑各有其历史典故和造型特色。园内植被丰富，4月万花园百花盛开，10月底至11月中旬举办菊花展。

This passes: it names the place type (景区), explains its fame (UNESCO), lists what you see (五亭桥、白塔), and gives practical dates (菊花展).

#### Before-and-After Fixes from Real Iteration

| Sight | Before (bad) | After (fixed) |
|-------|-------------|---------------|
| 梅兰芳纪念馆 | "庭院清雅，梅花点缀，处处流露京剧文化底蕴" | "梅兰芳是京剧'四大名旦'之首，以男性演员演绎女性角色的精湛技艺享誉世界。纪念馆位于泰州故居原址，馆内陈列了大量演出剧照、戏服和生平资料，可以了解这位艺术大师的舞台生涯。" |
| 凤城河 | "桨声灯影里品味泰州水城慢生活" | "凤城河是环绕泰州古城的护城河，夜游乘坐中式画舫沿河道行驶，可以观赏两岸古建筑和桥梁在灯光下的倒影。船程约40分钟，船上播放讲解，介绍沿岸的历史建筑。" |
| 乔园 | "亭榭廊桥、曲径通幽，一步一景" | "乔园是泰州现存最古老的私家园林，始建于明代，属于典型的江南文人园林风格。园内以假山、水池、亭廊为主要元素，空间虽不大但布局紧凑，回廊曲折。" |

#### Verification

After writing all descriptions, run a character count check:

```python
for sight in day.sights:
    cn = sum(1 for c in sight.description if '\u4e00' <= c <= '\u9fff')
    if cn < 100 or cn > 150:
        print(f'WARNING: {sight.name} has {cn} chars (target 100-150)')
```

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
