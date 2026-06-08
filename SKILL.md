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

#### Core Rules

1. **Count characters precisely.** 100-150 Chinese characters is the target. Fewer than 100 reads as too thin; more than 150 becomes excessive. Use a counting tool (`sum(1 for c in text if '\u4e00' <= c <= '\u9fff')`) to verify every description. The first version of a real itinerary failed because most descriptions were only 70-98 characters.

2. **Describe, do not praise.** Foreign readers have no emotional connection to the place — telling them it is "beautiful", "stunning", "poetic", or "breathtaking" tells them nothing. Instead, tell them concretely what they are looking at and why it matters.

3. **Assume zero cultural knowledge.** Do not assume the reader knows who historical figures are (姜太公, 梅兰芳), what architectural terms mean (文人园林, 回廊, 石舫), or why events matter. Every reference needs a one-sentence explanation of its context.

4. **Never comment on physical difficulty or pace.** Phrases like "全程步行轻松" (the walk is easy throughout), "适合拍照" (great for photos), or "值得一游" (worth a visit) are subjective judgments that waste characters and tell the reader nothing factual. Describe the path objectively — its length, surface, and gradient — and let readers decide for themselves. A temple's steps are steps; tell the visitor how many and how steep, not whether they are "easy."

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
- **Day tables**: each day has a header row + time/sight rows
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
BG_GREEN  = 'E8F5ED'   # meal summary row
BG_RED    = 'FFF1F0'   # hotel row
```

### Table Structure (2 columns)

Each day's detail table has **exactly 2 columns**. There is NO third column for meals.

| Col | Width | Content |
|-----|-------|---------|
| Time | 2.5 cm | morning / afternoon / evening / night / 全天 |
| Sight | 14.5 cm | sight name (bold green) + description (gray) |

Meals are **never** placed in the sight rows. They appear only in the dedicated `[Meals]` summary row at the bottom of the table (see below).

### Special Rows
- **Sight rows**: 2 columns — time label (left) + sight name & description (right). No meal info in these rows.
- **Meal row** (`[Meals]`): green background, 2 columns — `[Meals]` label (left) + all meals for the day in `早餐: … | 中餐: … | 晚餐: …` format (right). If no meals at all, omit this row.
- **Hotel row** (`[Hotel]`): red-tinted background, 2 columns — `[Hotel]` label (left) + hotel name (right).
- **Free day row**: merge both columns, light gray background.

**Critical: meals must NOT appear in sight rows or in a third column.** They only appear in the `[Meals]` row. This avoids duplication and keeps the sight content column wide and readable.

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

### `WD_ALIGN_VERTICAL.CENTER` serializes to wrong XML
`WD_ALIGN_VERTICAL.CENTER` is an enum value. When formatted into XML via `.format()` it produces `"CENTER (1)"`, NOT the required `"center"`. Word rejects the file.

```python
# WRONG - generates <w:vAlign w:val="CENTER (1)"/> which Word cannot parse
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

### `bold=False` generates `<w:b w:val="0"/>` that Word rejects
When `run.font.bold = False` is set, python-docx writes `<w:b w:val="0"/>`. Some Word versions (especially Word desktop) treat `w:val="0"` on a boolean element as an error and refuse to open the file.

```python
# WRONG - writes <w:b w:val="0"/> which can break Word
run.font.bold = False

# RIGHT - don't set bold at all when not bold; only set bold=True
run.font.bold = True         # OK - writes <w:b/>
# or simply skip setting it when False:
# (default None = no <w:b> element emitted)
```

Use a default of `bold=None` on helper functions:

```python
def add_run(p, text, bold=None):
    r = p.add_run(text)
    if bold is not None:
        r.font.bold = bold    # only emit <w:b/> when True
    return r
```

### `set_cell_shading` creates duplicate `<w:shd>` elements
Every call to `set_cell_shading` appends a new `<w:shd>` element via `tcPr.append()`. If called twice on the same cell, two `<w:shd>` elements accumulate. Word tolerates this in most cases but some versions may reject it or show incorrect colors.

```python
# WRONG - second append adds a duplicate
tcPr.append(shading_elm)

# RIGHT - remove existing shd first
NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
for existing in tcPr.findall('{%s}shd' % NS):
    tcPr.remove(existing)
tcPr.append(shading_elm)
```

### DOCX output rejected by Word: clean up the ZIP package
python-docx generates files with several extra parts (`customXml/`, `stylesWithEffects.xml`, `docProps/thumbnail.jpeg`) that can cause Word to refuse opening the document. The ZIP itself is valid, but references to missing or incompatible parts in `[Content_Types].xml` and `.rels` files trigger errors.

After `doc.save()`, run a cleanup step that:

1. Removes `customXml/`, `stylesWithEffects.xml`, `thumbnail.jpeg` from the ZIP
2. Removes their references from `[Content_Types].xml`, `_rels/.rels`, and `word/_rels/document.xml.rels`

```python
import zipfile, re, shutil

def cleanup_docx(filepath):
    tmp = filepath + '.tmp'
    with zipfile.ZipFile(filepath, 'r') as zin:
        ct = zin.read('[Content_Types].xml').decode('utf-8')
        ct = re.sub(r'<Override[^>]*?customXml[^>]*?/>', '', ct)
        ct = re.sub(r'<Override[^>]*?stylesWithEffects[^>]*?/>', '', ct)
        ct = re.sub(r'<Default[^>]*?jpeg[^>]*?/>', '', ct)

        doc_rels = zin.read('word/_rels/document.xml.rels').decode('utf-8')
        doc_rels = re.sub(r'<Relationship[^>]*?customXml[^>]*?/>', '', doc_rels)
        doc_rels = re.sub(r'<Relationship[^>]*?stylesWithEffects[^>]*?/>', '', doc_rels)

        main_rels = zin.read('_rels/.rels').decode('utf-8')
        main_rels = re.sub(r'<Relationship[^>]*?thumbnail[^>]*?/>', '', main_rels)

        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                fname = item.filename
                if any(x in fname for x in ['customXml', 'stylesWithEffects',
                                              'thumbnail', 'item1.xml']):
                    continue
                if fname == '[Content_Types].xml':
                    zout.writestr(fname, ct.encode('utf-8'))
                elif fname == 'word/_rels/document.xml.rels':
                    zout.writestr(fname, doc_rels.encode('utf-8'))
                elif fname == '_rels/.rels':
                    zout.writestr(fname, main_rels.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(fname))
    shutil.move(tmp, filepath)

# After doc.save():
cleanup_docx(output_path)
```

### .doc extraction via antiword may fail on WPS-created files
Some `.doc` files created by WPS Office are not recognized by `antiword` (error: "not a Word Document"). In that case fall back to `olefile + UTF-16LE` binary extraction:

```python
import olefile
ole = olefile.OleFileIO(input_path)
data = ole.openstream('WordDocument').read()
text = data.decode('utf-16le', errors='replace')
# Extract Chinese text portions
lines = []
current = []
for c in text:
    if c.isprintable() or c in '\n\r\t':
        current.append(c)
    else:
        if current:
            line = ''.join(current).strip()
            if line and sum(1 for ch in line if '\u4e00' <= ch <= '\u9fff') >= 3:
                lines.append(line)
            current = []
```

### Run scripts with `PYTHONIOENCODING=utf-8` in Git Bash on Windows
When executing Python scripts from Git Bash (MSYS2) on Windows with Chinese text, the default stdout encoding may be `gbk`, causing `UnicodeEncodeError`. Always prefix:

```bash
PYTHONIOENCODING=utf-8 python scripts/generate.py
```

This is not needed when running from PowerShell or cmd with a proper Chinese locale, but does no harm either.
