---
name: deepseek-itinerary
description: >
  Generate professional travel itinerary DOCX from 华宇/供应商报价 .doc files.
  Use when: (1) user has a .doc quote and wants a polished .docx itinerary;
  (2) /deepseek-itinerary; (3) 行程、报价转精美docx; (4) need Voyage-style layout
  (green day bars, paragraph flow, company header). Default style follows
  12D11N东北草原-type docs — NOT the old blue table layout. Source quote is the
  sole fact source: never invent sights, services, meals, hotels, or prices.
---

# Itinerary Generator (Voyage style)

From a supplier **报价 .doc** (or an existing draft DOCX), produce a **readable client-facing .docx** that:

1. **Never invents** sights, services, meal standards, hotels, or prices.
2. Matches **Voyage Ventures paragraph-flow design** (teal day bars, 宿： lines).
3. **Opens reliably in Word** by cloning a known-good template shell.
4. **Replicates fonts/colors exactly** when the user points at a modified sample DOCX.

## Absolute rules (fact fidelity)

| Rule | Detail |
|------|--------|
| Quote is sole source | Every sight, meal, hotel, gift, fee, contact must appear in the quote |
| No invented inclusions | Do not write「含机票」「含于团费」「无必加」unless the quote says so |
| 赠送 vs 行程内 | Only items under 赠送 (or explicitly「赠送…」in day text) are gifts; DIY/表演 in day body are itinerary, not free gifts |
| Prices exact | Copy RMB figures, 单间差, 小费, 儿童政策, 有效期 verbatim — **or omit the whole pricing block if user says「不要包含报价」** |
| No extra nights label | If quote says「6天」do not invent「5晚」unless nights are explicit |
| Output to source dir | Write `.docx` next to the source file (and optional ASCII-name copy) |
| Style clone | If user says「按某某行程排版 / 学习这个修改好的版本 / 完全复刻字体颜色」, open **that DOCX as the template shell** and match its run colors/fonts |

---

## Quick Start

```bash
# 1. Extract quote text (antiword often fails on WPS .doc)
PYTHONIOENCODING=utf-8 python scripts/extract_doc.py "报价.doc"

# 2. Edit scripts/generate.py DAYS/PRICING for this quote, then:
PYTHONIOENCODING=utf-8 python scripts/generate.py
```

**Preferred generation method:** open a **known-good Voyage itinerary .docx** as template → clear body (keep `w:sectPr` + header) → write paragraphs. Do **not** start from empty `Document()` if Word/WPS rejected previous pure-python-docx files.

Default style reference (when user does not specify another):

```
C:\Users\DELL\Downloads\12D11N_东北草原深度游_2026年6月.docx
```

Worked examples:

| Script | Product |
|--------|---------|
| `scripts/generate_chaoshan6_voyage_style.py` | 潮汕 6 天 |
| `scripts/generate_chongqing_sanxia8_voyage_style.py` | 重庆+长江三峡揽月 8D7N 上水 |

---

## Workflow

### 1. Extract source

`python-docx` cannot read legacy `.doc`.

**Try antiword first:**

```bash
antiword -m UTF-8 "path/to/quote.doc"
```

**Fallback (WPS / Windows, recommended):** `olefile` + UTF-16LE scan — see `scripts/extract_doc.py`.

**If source is already `.docx`:** extract all `w:t` text (paragraphs + tables). Some drafts only hold front matter; fill days from the quote or a fuller sibling draft.

Save a clean `_quote_extract.txt` for side-by-side fact check.

### 2. Learn from a user-revised DOCX (when provided)

When the user attaches `…行程 (2).docx` / says「学习这个修改好的版本」:

1. **Extract full text + style map** (per-paragraph: fill color, font, size, bold, text color).
2. **Copy structure exactly:** front matter order, Day header wording (incl. `06Dec Day 1`), meal parentheses, 宿： lines, hotel names.
3. **Copy palette exactly** from that file (not from memory):
   - Title `#1A6B5E` SimSun bold ~20pt
   - Subtitle `#555555` SimSun bold 11pt
   - Highlights / food / gifts often `#008080` (teal)
   - Day bar fill `#2E8B7A`, white text, **Arial Unicode MS**
   - 宿： / 【景点名】 `#1A6B5E` bold
4. **Shell:** `Document(user_revised.docx)` → clear body except `sectPr` → rewrite (keeps header/logo).
5. **Only deviate where user asks** — e.g. replace stub sight lines with foreign-traveler place intros; drop pricing if requested.

### 3. Parse quote fields

Per day:

- Optional date prefix (`06Dec`) when present in the approved draft
- Route / city / drive time (e.g. `厦门/永定（3H）`, `重庆→宜昌（周二登船）`)
- Full itinerary paragraph (keep wording; trim only layout spaces)
- Meals + 餐标 numbers (e.g. `午餐：北京烤鸭60` or draft form `（午：重庆风味｜晚：中式合菜）`)
- Hotel name + star (五星 / 准五星 / 五钻)
- Meal column code: `三餐` / `早午` / `早餐` / dinner only

Global:

- Product title, city list, 独家安排 / 亮点, 各地美食
- 精选酒店总述
- 报价、单间差、16免1、儿童、购物、小费、赠送、备注、联系人
- Skip 报价 block entirely if user forbids prices

### 4. Fact-check before save (mandatory)

Checklist — every item must pass:

```
[ ] 若输出含报价：团费 / 单间差 / 小费 / 儿童政策 / 有效期 数字一致
[ ] 若用户要求无报价：正文无 5850/单房差/FOC 等价格段
[ ] 每日酒店名与星级一致
[ ] 每日餐名与餐标数字一致（含「自理」）
[ ] 景点列表 ⊆ 报价（无新增）
[ ] 赠送栏未把行程内活动误标为赠送
[ ] 未写报价没有的「含机票/含保险/无购物店外费用」等
[ ] 车程 / 车次 / 时段标注与来源一致
[ ] 字体颜色与指定模板 / 修改版一致（Day 底 #2E8B7A 等）
```

### 5. Generate DOCX (Voyage paragraph style)

**Do not use the old 3-column blue table as default.**

Structure:

```
标题（居中，深绿 #1A6B5E，~18–20pt 加粗，SimSun）
城市线（居中，灰 #555555）
特别安排 / 亮点行（常见 #1A6B5E 或 #008080）
各地美食 / 赠送（常见 #008080 加粗）

[可选日期] Day N  路线    （午：…｜晚：…）   ← 白字 + 底色 #2E8B7A + Arial Unicode MS
当日概述段落（两端对齐，紧贴来源）
【景点名】地点介绍…                 ← 名加粗深绿；介绍=地点说明，非「游览XXX」复述
宿：酒店名（星级）                  ← 深绿加粗

[可选] 报价与说明 / 行程说明
联系方式（来源无则写「源报价未列联系人」，勿编造）
```

#### Typography (readability)

- **Delete useless spaces:** `· 集庆楼 / 善庆楼` → `·集庆楼、善庆楼`
- Prefer `→` for routes, `｜` between meal slots, Chinese顿号 `、` in lists
- **No giant space padding** in day headers — use 2–4 spaces as separators only
- Day overview = full source narrative; `【景点】` must **not** merely repeat「游览/参观+景点名」
- Line spacing ~1.15–1.25; consistent space before/after
- Only set `bold=True` when needed — never `bold=False` (Word rejects `w:b w:val="0"`)

#### Colors (Voyage / 东北草原 + common client teal)

```
C_TITLE  = 1A6B5E   # 标题、宿、【景点名】
C_SUB    = 555555   # 城市线
C_ORANGE = C45911   # 无购物强调（仅来源有时）
C_TEAL   = 008080   # 亮点/美食/赠送（见修改版 (2)）
C_DAY_BG = 2E8B7A   # Day 条段落底纹
C_BODY   = 000000
C_WHITE  = FFFFFF   # Day 条文字
Font body: SimSun / 宋体
Font Day:  Arial Unicode MS
```

When cloning a user file, **read colors from XML** — do not hardcode if the sample differs.

#### Template shell (Word-open reliability)

```python
doc = Document(TEMPLATE)  # known-good Voyage itinerary OR user-revised DOCX
# delete all body children except w:sectPr
# re-add paragraphs with shading / fonts
doc.save(output_path)
# also copy to ASCII filename if path has Chinese
```

Keeps company **header/logo**, section margins, theme.

### 6. Output

- Primary: same folder as source, Chinese product name
- Copies: English / short ASCII names for path-sensitive tools
- If primary path is locked (Word open), save `_无报价` / ASCII copy then retry overwrite

---

## Sight introductions (foreign travelers)

When the user wants **景点介绍 / 针对外国游客**:

| Rule | Detail |
|------|--------|
| Length | Default **40–50 汉字（不含标点）** if user specifies that range; otherwise 100–150 only when asked |
| Content | **Place description** — what/where/why it exists (立体交通、水利枢纽、溶洞、诗城…) |
| Not allowed | Pure itinerary restatement:「游览XXX」「参观XXX」「上午经过XXX」as the whole blurb |
| No praise fluff | No「适合拍照/值得一游/轻松/必去」 |
| Facts | Prefer facts already in quote/day text; do not invent new tickets, shows, or free gifts |
| Validate | Count CJK with `[\u4e00-\u9fff]` before save; fail build if out of range |

Example (good):

> 三峡大坝是世界规模最大的水利枢纽。升船机以垂直升降方式运送船舶过坝；停运时改参观博物馆与截流纪念园。

Example (bad — restatement only):

> 上午参观游览三峡大坝，抵达三斗坪换乘游船过升船机。

---

## Critical pitfalls

### Word cannot open DOCX

| Cause | Fix |
|-------|-----|
| python-docx default package extras | Prefer **clone known-good template** |
| `bold=False` → `w:b w:val="0"` | Only set bold when True |
| `WD_ALIGN_VERTICAL.CENTER` → `"CENTER (1)"` | Pass string `'center'` if using cells |
| Duplicate `w:shd` | Remove existing shd before append |
| Pure OOXML wrong Content_Types NS | Must be `.../package/2006/content-types` |
| antiword missing / WPS .doc | Use `scripts/extract_doc.py` |
| Permission denied on save | File open in Word — write alternate name then copy |

### Encoding on Windows

```powershell
$env:PYTHONIOENCODING='utf-8'
# Prefer full path to python.exe if `python` not on PATH:
& 'C:\Users\DELL\AppData\Local\Programs\Python\Python312\python.exe' scripts\generate.py
```

### Spaces and punctuation

- No spaces around `·` `/` when listing co-located sights (or use `、`)
- `宿：酒店 — 五星` → prefer `宿：酒店（五星）` matching source

### Style drift

- Do not mix `#008080` day bars with `#2E8B7A` if the learned sample uses one consistently
- Day bars in current Voyage client samples: fill **`2E8B7A`**, font **Arial Unicode MS**, text white

---

## Scripts

| File | Role |
|------|------|
| `scripts/extract_doc.py` | Extract text from `.doc` via olefile |
| `scripts/generate.py` | **Template** — Voyage-style generator (edit data, run) |
| `scripts/generate_chaoshan6_voyage_style.py` | Worked example (潮汕6天) |
| `scripts/generate_chaoshan6.py` | Legacy blue-table example (deprecated default) |
| `scripts/generate_chongqing_sanxia8_voyage_style.py` | Worked example (重庆+三峡揽月 8D7N；可学习用户修改版样式 + 景点地点介绍) |

## References

- `references/formatting-guide.md` — paragraph helpers, day-bar shading, template clear-body
- `references/fact-check.md` — pre-delivery checklist
