---
name: deepseek-itinerary
description: >
  Generate professional travel itinerary DOCX from 华宇/供应商报价 .doc files.
  Use when: (1) user has a .doc quote and wants a polished .docx itinerary;
  (2) /deepseek-itinerary; (3) 行程、报价转精美docx; (4) need Voyage-style layout
  (green day bars, paragraph flow, company header). Default style follows
  12D11N东北草原-type docs — NOT the old blue table layout. Default client
  output is NO pricing block; include 团费/单间差 only when user explicitly asks.
  Source quote is the sole fact source: never invent sights, services, meals,
  hotels, or prices.
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
| **Default: no pricing** | Client DOCX **omits** 报价与说明 / 团费 / 单间差 / FOC / 有效期价格段 by default. Include prices **only** when user explicitly asks（含报价 / 要团费 / INCLUDE_PRICING） |
| Prices exact when included | If pricing is on: copy RMB figures, 单间差, 小费, 儿童政策, 有效期 verbatim |
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
| `scripts/generate_chengdu6_voyage_style.py` / `_en.py` | 漫游成都6日（卧龙）— 生动体验式文案 + 英文版 |

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
4. **Shell:** `Document(user_revised.docx)` → clear body except `sectPr` → rewrite (keeps header/logo).
5. **Default still no pricing.** Only deviate on content the user asks for (e.g. place-style sight intros). Add 报价 block only if they explicitly want prices.

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
- 精选酒店总述（可作行程说明，**默认不写团费数字**）
- 赠送、备注（非价格）、联系人
- 报价、单间差、16免1、儿童、小费、有效期：**仅当用户明确要求含报价时**解析并写入

### 4. Fact-check before save (mandatory)

Checklist — every item must pass:

```
[ ] 默认无报价：正文无 5850/单房差/FOC/团费 RMB… 等价格段
[ ] 仅当用户要求含报价：团费 / 单间差 / 小费 / 儿童政策 / 有效期 数字与报价一致
[ ] 每日酒店名与星级一致
[ ] 每日餐名与餐标数字一致（含「自理」）— 餐标≠团费
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

[默认省略] 报价与说明 / 团费数字 — 仅当用户明确要求「含报价」时写入
联系方式（来源无则省略；勿编造；通常也不放在默认无报价稿）
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

## Sight introductions & day narrative (Singapore travelers)

Client readers are mostly **Singapore travelers**. Dry encyclopedia-style blurbs（「始建于X代，现存建筑为清代重建…」）do not interest them. **Default style = vivid, experiential copy**: what they will see, hear, taste and feel there — still with one or two solid facts as backbone.

### Sight intros（【景点】后的描述）

| Rule | Detail |
|------|--------|
| Length | Default **~70–100 汉字**. Use 40–50 only when the user explicitly asks for short intros |
| Hook first | Open with an image or identity:「藏身闹市的千年古刹」「都江堰夜晚最梦幻的一幕」「想看最地道的成都日常，就来人民公园」 |
| Sensory scene | Concrete things on site: 红灯笼、青石板、竹椅盖碗茶、熊猫啃竹子打滚、灯光倒映江面 |
| One fact backbone | Keep 1–2 verifiable facts (年代、人物、世界遗产、名称由来) — woven in, not listed |
| Memorable detail | Legends / icons travelers will look for: 摸铜羊讨好运、IFS「爬墙大熊猫」、三国主题街 |
| Explain local terms | Dialect or cultural words get a bracket gloss for overseas readers:「摆龙门阵（聊天）」「巴适（舒服惬意）」 |
| Not allowed | Whole blurb = itinerary restatement（「游览XXX」「上午参观XXX」） |
| Facts guard | Describing what exists at a place（小吃、相亲角、酒吧街）is fine; never phrase it as an included service, meal, show or gift unless the quote says so. Quote-listed inclusions（电瓶车+讲解器、含茶）may be woven into the scene |

Example (good — vivid):

> 【卧龙神树坪基地（电瓶车+讲解器）】藏在卧龙国家级自然保护区深山里的大熊猫基地，云雾缭绕、竹林满坡，比城市里的熊猫园更接近它们的野外家园。乘电瓶车进入园区，戴上讲解器，看“滚滚”们啃竹子、爬树、打滚，萌态十足。

> 【鹤茗茶社（含茶）】人民公园里的老字号露天茶馆，临湖摆满竹椅。捧一碗盖碗茶，看身边的成都人“摆龙门阵”（聊天），亲身体会“一杯茶坐一下午”的巴适（舒服惬意）生活。

Example (bad — dry, encyclopedia only):

> 成都保存较完整的佛教寺院，始建于隋代，现存殿宇为清代重建，是川西著名禅宗丛林。

Example (bad — restatement only):

> 上午参观游览三峡大坝，抵达三斗坪换乘游船过升船机。

### Day overview（Day 条下的概述段）

- Open with a **theme line** for the day:「今天走进成都的老城时光。」「今天是期待已久的熊猫之旅！」
- Chain sights with light connectors（随后、再到、最后）and a short purpose per stop（探访道家文化、感受市井烟火）— sight order must match the quote.
- Do **not** invent clock times（上午/傍晚/晚上）unless the quote has them or the sight is inherently nocturnal（夜景、蓝眼泪）.
- Free day: add a gentle suggestion tied to facts already in the quote (e.g. hotel location) — never an optional tour or paid activity.
- Departure day: warm close（带着满满的回忆，返回温暖的家——新加坡）.

### English version

When asked for an English DOCX: same structure (Day bars, 【Sight】 paragraphs, `Stay:` lines), same vivid tone, same fact / no-pricing checks.

- Sight names: English or pinyin name + Chinese in brackets, e.g. `Wenshu Monastery (文殊院)`; keep quote inclusions in the name, e.g. `(incl. shuttle bus & audio guide)`.
- Hotels: the hotel's own English name when known, with Chinese name in brackets; otherwise pinyin + Chinese. Star grade `4-Diamond`.
- Meals in Day bar: `(B: Hotel ｜ L: Included ｜ D: Included)`, `X` for not included.
- Body font: a Latin font (Arial) with SimSun as East Asian fallback for the bracketed Chinese.
- Glossed local terms work well in English too: "*bai long men zhen* (a good long chat)".

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
| `scripts/generate_chengdu6_voyage_style.py` | Worked example (漫游成都6日；新加坡游客生动文案、Day 条 keep-with-next 防分页) |
| `scripts/generate_chengdu6_voyage_style_en.py` | English edition of the Chengdu example (Arial + SimSun fallback, `Stay:` lines) |

## References

- `references/formatting-guide.md` — paragraph helpers, day-bar shading, template clear-body
- `references/fact-check.md` — pre-delivery checklist
