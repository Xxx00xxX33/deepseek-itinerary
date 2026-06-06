# deepseek-itinerary

AI-powered itinerary generator skill for Codex/Reasonix.

Convert raw `.doc` travel itinerary tables into beautifully formatted `.docx` documents with:
- Professional cover page with title, subtitle, and highlights
- Day-by-day breakdown with time slots, sight descriptions, meals, and hotels
- Factual 100-150 character sight descriptions suitable for foreign travelers
- Color-coded table styling (blue headers, green meals, red-tinted hotels)
- Pricing section with tiered rates, inclusions/exclusions

## What's Inside

```
deepseek-itinerary/
├── SKILL.md                    # Main instructions for Codex
├── agents/openai.yaml        # Skill metadata
├── scripts/generate.py        # Python template for generating .docx
└── references/formatting-guide.md  # python-docx reference patterns
```

## How It Works

1. **Extract** the itinerary table from `.doc` using `antiword`
2. **Parse** day-by-day data: sights, meals, hotels, seasonal notes
3. **Expand** terse bullet points into 100-150 character factual descriptions
4. **Generate** a polished `.docx` with the template script

## Key Techniques Documented

- Reading `.doc` files (antiword, not python-docx)
- python-docx table formatting (shading, borders, merged cells)
- Chinese font handling with `w:eastAsia`
- Unicode encoding pitfalls (surrogate pairs, GBK console output)
- Emoji-safe text rendering

## Prerequisites

- Python 3.8+
- `python-docx` (`pip install python-docx`)
- `antiword` (for `.doc` extraction)
