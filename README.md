# deepseek-itinerary

Generate **client-facing travel itinerary DOCX** from 华宇 / supplier **报价 `.doc`** files.

Default layout is **Voyage Ventures paragraph style** (teal day bars, 宿： lines, company header) — the same family as `12D11N_东北草原深度游`.  
**Not** the old blue 3-column table template.

## Absolute rule

The source quote is the **sole fact source**. Do not invent sights, services, meal standards, hotels, gifts, or prices.

## Layout

```
deepseek-itinerary/
├── SKILL.md                              # Agent instructions (read this first)
├── agents/openai.yaml                    # Skill metadata
├── scripts/
│   ├── extract_doc.py                    # .doc text via olefile (WPS-safe)
│   ├── generate.py                       # Voyage-style generator template
│   ├── generate_chaoshan6_voyage_style.py  # Worked example (潮汕 6 天)
│   └── generate_shanwei_huaxiang.py      # Advanced: page break + 60-70 char attractive intros
└── references/
    ├── formatting-guide.md               # Paragraph helpers, day-bar shading
    └── fact-check.md                     # Pre-delivery checklist
```

## How it works

1. **Extract** quote text — `antiword` if available, else `scripts/extract_doc.py` (olefile + UTF-16LE).
2. **Parse** days, meals, hotels, pricing, 赠送, contacts.
3. **Fact-check** against the quote (`references/fact-check.md`).
4. **Generate** DOCX by **cloning a known-good Voyage itinerary** as template shell (header/logo preserved), then writing paragraphs.

## Quick start

```bash
# Extract
PYTHONIOENCODING=utf-8 python scripts/extract_doc.py "报价.doc"

# Edit DAYS / PRICING in scripts/generate.py, then:
PYTHONIOENCODING=utf-8 python scripts/generate.py
```

Default style template (when user does not specify another):

```
C:\Users\DELL\Downloads\12D11N_东北草原深度游_2026年6月.docx
```

Output goes **next to the source `.doc`**, with an optional ASCII-name copy.

## Prerequisites

- Python 3.8+
- `python-docx` (`pip install python-docx`)
- `olefile` (`pip install olefile`) — for WPS / Windows `.doc` extraction
- Optional: `antiword` for Unix-style `.doc` extraction

## Key techniques

| Topic | Approach |
|-------|----------|
| Word-open reliability | Clone known-good DOCX shell; clear body; rewrite paragraphs |
| Day headers | Paragraph shading `#2E8B7A`, white text — not giant space padding |
| Bold | Only set `bold=True`; never `bold=False` |
| Chinese fonts | `SimSun` body, `Arial Unicode MS` day bars, set `w:eastAsia` |
| Quote fidelity | Checklist before save; 赠送 ≠ itinerary activity |
| Readability | Strip junk spaces around `·` `/`; use `、` / `→` / `｜` |

## Deprecated

- Old **blue table** layout and pure OOXML rebuilds — keep only as historical reference if present (`generate_chaoshan6.py`). Prefer `generate.py` / `generate_chaoshan6_voyage_style.py`.
