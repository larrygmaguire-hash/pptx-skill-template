# Automated PowerPoint Generation with Claude Code

A skill template and workflow for generating branded PowerPoint presentations programmatically using Claude Code and Python.

---

## What This Is

This is a customisable Claude Code skill that generates PowerPoint presentations from your branded template. It handles:

- Template-based slide generation
- Consistent typography and colours
- Arrow bullet formatting
- Dissolve animations on click
- Paragraph-by-paragraph content builds

**Prerequisites:**
- Claude Code (Anthropic's agentic coding tool)
- Python 3.9+
- `python-pptx` and `lxml` packages
- A prepared PowerPoint template (`.potx` or `.pptx`)

---

## Included Files

| File | Purpose |
|------|---------|
| `README.md` | This guide |
| `SKILL-TEMPLATE.md` | Customisable skill file for your `.claude/skills/` folder |
| `presentation-generator.py` | Reference Python implementation |
| `style-template.json` | Style specification template |
| `PREPARATION-GUIDE.md` | How to prepare your PowerPoint template |

---

## Quick Start

1. **Prepare your PowerPoint template** — Follow `PREPARATION-GUIDE.md`
2. **Customise the style JSON** — Edit `style-template.json` with your brand specs
3. **Customise the skill** — Edit `SKILL-TEMPLATE.md` with your layouts
4. **Install the skill** — Copy to `.claude/skills/your-skill-name/`
5. **Test and refine** — Work with Claude Code to perfect the workflow

---

## Important Notes

### This Requires Iteration

The workflow shown here is a **starting point**. You will need to work with Claude Code to:

- Adjust layout indices to match your template
- Fine-tune typography settings
- Modify animation behaviour
- Add or remove slide types

Expect 2-4 iterations before the workflow matches your exact requirements.

### Template Preparation is Critical

The Python script reads layouts from your template by index number. If your template isn't set up correctly, the script will fail or produce incorrect results.

**Do not skip the preparation guide.**

---

## How It Works

```
Your Template (.potx)          Your Content (outline/brief)
        │                                │
        ▼                                ▼
┌─────────────────────────────────────────────┐
│            Claude Code + Python              │
│  ┌───────────────────────────────────────┐  │
│  │  1. Convert .potx to .pptx            │  │
│  │  2. Add slides using layout indices   │  │
│  │  3. Set text with typography          │  │
│  │  4. Format bullets (Wingdings arrows) │  │
│  │  5. Add dissolve animations           │  │
│  └───────────────────────────────────────┘  │
└─────────────────────────────────────────────┘
                      │
                      ▼
          Generated Presentation (.pptx)
```

---

## Customisation Points

| Element | Where to Change | Notes |
|---------|-----------------|-------|
| Layouts | `SKILL-TEMPLATE.md` + `style-template.json` | Must match your template |
| Typography | `style-template.json` | Font, sizes, line spacing |
| Colours | `style-template.json` | Hex codes for each background type |
| Bullets | `presentation-generator.py` | Wingdings character, margins |
| Animations | `presentation-generator.py` | Effect type, duration, triggers |

---

## Support

This is a self-service resource. For help:

1. Work through the preparation guide carefully
2. Use Claude Code to debug issues
3. Join the [GenAI Skills Academy](https://www.skool.com/genai-skills-academy-1964) community for discussion

---

*Part of the PowerPoint AI Mastery course*
*Created by Larry G. Maguire | Larry G. Maguire Human Performance*
