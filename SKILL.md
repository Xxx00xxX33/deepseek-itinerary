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

From a supplier **报价 .doc**, produce a **readable client-facing .docx** that:

1. **Never invents** sights, services, meal standards, hotels, or prices.
2. Matches **Voyage Ventures paragraph-flow design** (teal day bars, 宿： lines).
3. **Opens reliably in Word** by cloning a known-good template shell.

## Absolute rules (fact fidelity)

| Rule | Detail |
|------|--------|
| Quote is sole source | Every sight, meal, hotel, gift, fee, contact must appear in the quote |
| No invented inclusions | Do not write「含机票」「含于团费」「无必加」unless the quote says so |
| 赠送 vs 行程内 | Only items under 赠送 (or explicitly「赠送…」in day text) are gifts; DIY/表演 in day body are itinerary, not free gifts |
| Prices exact | Copy RMB figures, 单间差, 小费, 儿童政策, 有效期 verbatim |
| No extra nights label | If quote says「6天」do not invent「5晚」unless nights are explicit |
| Output to source dir | Write `.docx` next to the source `.doc` (and optional ASCII-name copy) |

If the user says「按某某行程排版」, **clone that DOCX as the style template** (header/logo/colors), still keep quote facts.

---

## Quick Start

```bash
# 1. Extract quote text (antiword often fails on WPS .doc)
PYTHONIOENCODING=utf-8 python scripts/extract_doc.py "报价.doc"

# 2. Edit scripts/generate.py DAYS/PRICING for this quote, then:
PYTHONIOENCODING=utf-8 python scripts/generate.py
```

**Preferred generation method:** open a **known-good Voyage itinerary .docx** as template → clear body → write paragraphs. Do **not** start from empty `Document()` if Word/WPS rejected previous pure-python-docx files.

Default style reference (when user does not specify another):

```
C:\Users\DELL\Downloads\12D11N_东北草原深度游_2026年6月.docx
```

Worked example script: `scripts/generate_chaoshan6_voyage_style.py`.

---

## Workflow

### 1. Extract source .doc

`python-docx` cannot read legacy `.doc`.

**Try antiword first:**

```bash
antiword -m UTF-8 "path/to/quote.doc"
```

**Fallback (WPS / Windows, recommended):** `olefile` + UTF-16LE scan — see `scripts/extract_doc.py`.

Save a clean `_quote_extract.txt` for side-by-side fact check.

### 2. Parse quote fields

Per day:

- Route / city / drive time (e.g. `厦门/永定（3H）`)
- Full itinerary paragraph (keep wording; trim only layout spaces)
- Meals + 餐标 numbers (e.g. `午餐：北京烤鸭60`)
- Hotel name + star (五星 / 准五星)
- Meal column code: `三餐` / `早午` / `早餐` / dinner only

Global:

- Product title, city list, 独家安排 / 亮点, 各地美食
- 精选酒店总述
- 报价、单间差、16免1、儿童、购物、小费、赠送、备注、联系人

### 3. Fact-check before save (mandatory)

Checklist — every item must pass:

```
[ ] 团费 / 单间差 / 小费 / 儿童政策 / 有效期 数字一致
[ ] 每日酒店名与星级一致
[ ] 每日餐名与餐标数字一致（含「自理」）
[ ] 景点列表 ⊆ 报价（无新增）
[ ] 赠送栏未把行程内活动误标为赠送
[ ] 未写报价没有的「含机票/含保险/无购物店外费用」等
[ ] 车程标注与报价括号一致
```

### 4. Generate DOCX (Voyage paragraph style)

**Do not use the old 3-column blue table as default.**

Structure:

```
标题（居中，深绿 #1A6B5E，~18pt 加粗）
城市线（居中，灰 #555555）
***全程不进店 / 报价购物句（橙 #C45911，居中）
独家安排：…（标签加粗青绿）
各地美食：…
赠送：…

Day N  路线（车程）    （早/午/晚）     ← 白字 + 底色 #2E8B7A
当日概述段落（两端对齐，紧贴报价）
【景点名】说明…                    ← 景点名加粗深绿，说明正文
宿：酒店名（星级）                 ← 深绿加粗

报价与说明
联系方式
```

#### Typography (readability)

- **Delete useless spaces:** `· 集庆楼 / 善庆楼` → `·集庆楼、善庆楼` or `·集庆楼、善庆楼、绳庆楼`
- Prefer `→` for routes, `｜` between meal slots, Chinese顿号 `、` in lists
- **No giant space padding** in day headers — use 2–4 spaces as separators only
- Day overview = full quote narrative; `【景点】` = short scan lines (avoid repeating the whole day twice)
- Line spacing ~1.15–1.25; consistent space before/after
- Only set `bold=True` when needed — never `bold=False` (Word rejects `w:b w:val="0"`)

#### Colors (Voyage / 东北草原)

```
C_TITLE  = 1A6B5E   # 标题、宿、标签
C_SUB    = 555555   # 城市线
C_ORANGE = C45911   # 无购物强调
C_HL     = 419D99   # 可选亮点色
C_FOOD   = 2E8B7A   # 美食/赠送/Day底
C_DAY_BG = 2E8B7A   # Day 条段落底纹
C_BODY   = 000000
C_WHITE  = FFFFFF   # Day 条文字
Font: SimSun (body), Arial Unicode MS (Day 条)
```

#### Template shell (Word-open reliability)

```python
doc = Document(TEMPLATE)  # known-good Voyage itinerary
# delete all body children except w:sectPr
# re-add paragraphs with shading / fonts
doc.save(output_path)
# also copy to ASCII filename if path has Chinese
```

Keeps company **header/logo**, section margins, theme — files open in Word/WPS far more reliably than pure OOXML rebuilds or default python-docx packages with `customXml` / `stylesWithEffects` / `thumbnail`.

### 5. Output

- Primary: same folder as source `.doc`, Chinese product name
- Copies: English / short ASCII names for path-sensitive tools
- Optional: HTML twin only if DOCX still fails to open

---

## Description expansion (optional)

When the user wants **foreign-traveler** expansions (100–150 Chinese chars per sight):

1. Only rephrase **facts already in the quote** — no new attractions or services
2. Target 100–150 CJK characters; verify with a counter
3. Describe, do not praise; no「适合拍照/值得一游/轻松」
4. If user demands strict quote fidelity only, **skip expansion** — use quote sentences

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

### Encoding on Windows

```bash
$env:PYTHONIOENCODING='utf-8'   # PowerShell
PYTHONIOENCODING=utf-8 python scripts/generate.py
```

### Spaces and punctuation

User-visible junk spaces are a defect. Normalize:

- No spaces around `·` `/` when listing co-located sights (or use `、` instead of `/`)
- `宿：酒店 — 五星` → prefer `宿：酒店（五星）` matching quote

---

## Scripts

| File | Role |
|------|------|
| `scripts/extract_doc.py` | Extract text from `.doc` via olefile |
| `scripts/generate.py` | **Template** — Voyage-style generator (edit data, run) |
| `scripts/generate_chaoshan6_voyage_style.py` | Full worked example (潮汕6天) |

## References

- `references/formatting-guide.md` — paragraph helpers, day-bar shading, template clear-body
- `references/fact-check.md` — pre-delivery checklist
