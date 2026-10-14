# deepseek-itinerary

Generate **client-facing travel itinerary DOCX** from 华宇 / supplier **报价 `.doc`** files (or an approved draft DOCX).

Default layout is **Voyage Ventures paragraph style** (teal day bars, 宿： lines, company header) — the same family as `12D11N_东北草原深度游`.  
**Not** the old blue 3-column table template.

## Absolute rule

The source quote is the **sole fact source**. Do not invent sights, services, meal standards, hotels, gifts, or prices.

**Default client output has NO pricing block** (no 团费 / 单房差 / FOC). Include prices only when the user explicitly asks（含报价 / 要团费）.

## Layout

```
deepseek-itinerary/
├── SKILL.md                              # Agent instructions (read this first)
├── agents/openai.yaml                    # Skill metadata
├── scripts/
│   ├── extract_doc.py                    # .doc text via olefile (WPS-safe)
│   ├── generate.py                       # Voyage-style generator template
│   ├── generate_chaoshan6_voyage_style.py  # Worked example (潮汕 6 天)
│   └── generate_chongqing_sanxia8_voyage_style.py  # 重庆+三峡揽月 8D7N
└── references/
    ├── formatting-guide.md               # Paragraph helpers, day-bar shading
    └── fact-check.md                     # Pre-delivery checklist
```

## How it works

1. **Extract** quote text — `antiword` if available, else `scripts/extract_doc.py` (olefile + UTF-16LE). For `.docx`, dump all `w:t`.
2. **Learn style** — if user provides a revised sample DOCX, map its fonts/colors/day-bar fill and clone that shell.
3. **Parse** days, meals, hotels, pricing (optional), 赠送, contacts.
4. **Fact-check** against the quote (`references/fact-check.md`).
5. **Generate** DOCX by **cloning a known-good Voyage / user sample** as template shell (header/logo preserved), then writing paragraphs.

## Sight introductions

When the user asks for 外国游客 / 景点介绍:

- Write **vivid, experiential copy** for Singapore travelers: hook + on-site scene + 1–2 facts — not dry encyclopedia text, not「游览XXX」restatements.
- Default length **~70–100 Chinese characters**; 40–50 only when the user asks for short intros.
- Gloss dialect / cultural terms for overseas readers（摆龙门阵（聊天））.
- English DOCX on request: same structure and tone (see SKILL.md).
- Do not invent services not in the quote.

## Quick start

```bash
# Extract
PYTHONIOENCODING=utf-8 python scripts/extract_doc.py "报价.doc"

# Edit DAYS / optional PRICING in scripts/generate.py, then:
PYTHONIOENCODING=utf-8 python scripts/generate.py
```

Default style template (when user does not specify another):

```
C:\Users\DELL\Downloads\12D11N_东北草原深度游_2026年6月.docx
```

If the user says「按某某行程排版」or「学习这个修改好的版本」, open **that file** as `TEMPLATE` / shell instead.

Output goes **next to the source file**, with an optional ASCII-name copy.

## Prerequisites

- Python 3.8+
- `python-docx` (`pip install python-docx`)
- `olefile` (`pip install olefile`) — for WPS / Windows `.doc` extraction
- Optional: `antiword` for Unix-style `.doc` extraction

## Key techniques

| Topic | Approach |
|-------|----------|
| Word-open reliability | Clone known-good DOCX shell; clear body; rewrite paragraphs |
| Style fidelity | Read fill/font/color from user sample XML when asked to 复刻 |
| Quote fidelity | Checklist before save; 赠送 ≠ itinerary activity |
| **No pricing (default)** | Omit 报价与说明 / 团费数字 unless user explicitly requests prices |
| Readability | Strip junk spaces around `·` `/`; use `、` / `→` / `｜` |

## Deprecated

- Old **blue table** layout and pure OOXML rebuilds — keep only as historical reference if present (`generate_chaoshan6.py`). Prefer `generate.py` / `*_voyage_style.py`.
