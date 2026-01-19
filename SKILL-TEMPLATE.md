---
name: your-presentation-skill
description: Create branded PowerPoint presentations using your template. Use when generating slide decks.
---

# [Your Brand] Presentation Creation

Generate professional branded presentations using the standard template and styling.

## When to Use

- Creating presentation decks for [your use cases]
- Any [your brand] branded presentation

## When NOT to Use

- Generic presentations without your branding
- Non-PowerPoint formats

---

## Mode Selection (Optional)

If you have multiple presentation types, ask at the start:

> Is this presentation for [type A], [type B], or [type C]?

| Mode | Structure | Use For |
|------|-----------|---------|
| Type A | [describe] | [use case] |
| Type B | [describe] | [use case] |
| Type C | [describe] | [use case] |

---

## Template

**File:** `[path to your template file]`

Single source of truth for all presentations. Do not create alternative templates.

---

## Typography Specification

| Element | Font | Size | Line Spacing | Case | Style |
|---------|------|------|--------------|------|-------|
| Title | [Your Font] | [size] | 1.2 | [UPPER/Sentence] | Standard |
| Subtitle | [Your Font] | [size] | 1.2 | [case] | [Bold/Standard] |
| Body | [Your Font] | [size] | 1.2 | — | Standard |
| Bullets | [Your Font] | [size] | 1.2 | — | Standard |

### Colour by Background

| Background | Text Colour | Hex |
|------------|-------------|-----|
| Dark slides | White | #FFFFFF |
| Light slides | [Your Accent] | #[HEX] |

---

## Layout Index Mapping (0-based for Python)

**IMPORTANT:** These indices must match YOUR template. Run the extraction script to verify.

| Index | Layout Name | Placeholders | Use For |
|-------|-------------|--------------|---------|
| 0 | [Name] | 0, 1 | [Purpose] |
| 1 | [Name] | 0, 1 | [Purpose] |
| 2 | [Name] | 0, 1, 13 | [Purpose] |
| ... | ... | ... | ... |

**Fixed slides (animations in template):** [list indices, e.g., 5, 17]

**Do not use:** [list any reserved/duplicate indices]

---

## Content Guidelines

### Content Slide Structure — MANDATORY

Content slides have three elements:

1. **Title** (placeholder 0) — [Your case convention], concise topic
2. **Subtitle** (placeholder 1) — Short, succinct (5-10 words max)
3. **Body** (placeholder [your body index]) — Intro paragraph + bullet points

**Body structure:**
- 1-2 sentence intro paragraph explaining context
- Blank line for spacing
- Arrow bullet points (Wingdings Ø character)

**Example:**
```
TITLE: [TOPIC NAME] (placeholder 0)
Subtitle: Short Descriptive Phrase (placeholder 1)

Body (placeholder [index]):
One or two sentences explaining why this topic matters and framing the bullet points that follow.

→ First key point
→ Second key point
→ Third key point
→ Fourth key point (maximum)
```

### White Space — MANDATORY

- **Maximum 4 bullets per slide** — if more needed, split across slides
- **Maximum 10-12 words per bullet** — be concise
- **Paragraph intros: 1-2 sentences only**
- **Trust the template margins** — no cramming content

---

## Python Generation Workflow

### 1. Convert Template

```python
import tempfile
import zipfile
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

TEMPLATE_PATH = Path("[your template path]")
ACCENT_COLOUR = RGBColor([r], [g], [b])  # Your accent colour RGB values

def convert_potx_to_pptx(potx_path: Path, pptx_path: Path):
    """Convert .potx template to .pptx."""
    with zipfile.ZipFile(potx_path, "r") as zin:
        with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = data.replace(
                        b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
                        b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                    )
                zout.writestr(item, data)
```

### 2. Set Simple Text (titles, subtitles)

```python
def set_text_simple(placeholder, text: str, is_light_bg: bool = False, font_size: int = None):
    """Set text with typography and 1.2 line spacing (no bullets)."""
    placeholder.text = text
    for paragraph in placeholder.text_frame.paragraphs:
        paragraph.line_spacing = 1.2
        for run in paragraph.runs:
            run.font.name = "[Your Font]"
            if font_size:
                run.font.size = Pt(font_size)
            if is_light_bg:
                run.font.color.rgb = ACCENT_COLOUR
```

### 3. Set Body with Arrow Bullets

```python
from pptx.oxml.ns import qn
from lxml import etree

def set_body_with_bullets(placeholder, intro_text: str, bullets: list, is_light_bg: bool = False):
    """Set body content with intro paragraph, blank line, then arrow bullets."""
    placeholder.text = ""
    txBody = placeholder._element.find(qn('p:txBody'))
    if txBody is None:
        return

    # Remove existing paragraphs
    for p in txBody.findall(qn('a:p')):
        txBody.remove(p)

    colour_val = "[YOUR_HEX]" if is_light_bg else "FFFFFF"

    # Paragraph 1: Intro text
    intro_p = etree.SubElement(txBody, qn('a:p'))
    intro_pPr = etree.SubElement(intro_p, qn('a:pPr'))
    intro_lnSpc = etree.SubElement(intro_pPr, qn('a:lnSpc'))
    etree.SubElement(intro_lnSpc, qn('a:spcPct'), val="120000")
    intro_r = etree.SubElement(intro_p, qn('a:r'))
    intro_rPr = etree.SubElement(intro_r, qn('a:rPr'), sz="2200", dirty="0")
    intro_fill = etree.SubElement(intro_rPr, qn('a:solidFill'))
    etree.SubElement(intro_fill, qn('a:srgbClr'), val=colour_val)
    etree.SubElement(intro_rPr, qn('a:latin'), typeface="[Your Font]")
    intro_t = etree.SubElement(intro_r, qn('a:t'))
    intro_t.text = intro_text

    # Paragraph 2: Empty line for spacing
    empty_p = etree.SubElement(txBody, qn('a:p'))
    empty_pPr = etree.SubElement(empty_p, qn('a:pPr'))
    empty_lnSpc = etree.SubElement(empty_pPr, qn('a:lnSpc'))
    etree.SubElement(empty_lnSpc, qn('a:spcPct'), val="120000")
    etree.SubElement(empty_p, qn('a:endParaRPr'), dirty="0")

    # Bullet paragraphs with arrow (Wingdings Ø)
    for bullet_text in bullets:
        bullet_p = etree.SubElement(txBody, qn('a:p'))
        bullet_pPr = etree.SubElement(bullet_p, qn('a:pPr'), marL="342900", indent="-342900")
        bullet_lnSpc = etree.SubElement(bullet_pPr, qn('a:lnSpc'))
        etree.SubElement(bullet_lnSpc, qn('a:spcPct'), val="120000")
        etree.SubElement(bullet_pPr, qn('a:buFont'), typeface="Wingdings", pitchFamily="2", charset="2")
        etree.SubElement(bullet_pPr, qn('a:buChar'), char="\u00D8")  # Arrow in Wingdings

        bullet_r = etree.SubElement(bullet_p, qn('a:r'))
        bullet_rPr = etree.SubElement(bullet_r, qn('a:rPr'), sz="2200", dirty="0")
        bullet_fill = etree.SubElement(bullet_rPr, qn('a:solidFill'))
        etree.SubElement(bullet_fill, qn('a:srgbClr'), val=colour_val)
        etree.SubElement(bullet_rPr, qn('a:latin'), typeface="[Your Font]")
        bullet_t = etree.SubElement(bullet_r, qn('a:t'))
        bullet_t.text = bullet_text
```

### 4. Add Slides

```python
LAYOUTS = {
    "title": 0,
    "section": 1,
    "content": 2,
    # Add your layouts here
}

def add_slide(prs: Presentation, layout_key: str):
    """Add slide using layout key."""
    return prs.slides.add_slide(prs.slide_layouts[LAYOUTS[layout_key]])
```

### 5. Add Dissolve Animations

Add 0.5s dissolve on-click animations to each slide after creating it. The function:
- Animates title as a single unit
- Animates subtitle paragraph-by-paragraph
- Animates body paragraph-by-paragraph (intro, then each bullet)

**Key XML attributes:**
- `presetID="9"` — Dissolve effect
- `filter="dissolve"` — Effect filter
- `dur="500"` — 0.5 second duration
- `delay="indefinite"` — On click trigger
- `build="p"` — Paragraph-by-paragraph for multi-paragraph shapes

**Fixed slides:** Do NOT add animations via script — these slides have pre-set animations in the template master.

See `presentation-generator.py` for the full `add_dissolve_animations()` function implementation.

### 6. Data-Driven Outline Structure

The script generates slides dynamically from a structured outline. Modify the OUTLINE dict to expand or reduce sections as needed:

```python
OUTLINE = {
    "title": "PRESENTATION TITLE",
    "subtitle": "Your Subtitle Here",
    "thank_you_subtitle": "Contact information",

    "sections": [
        {
            "name": "SECTION NAME",
            "subtitle": "Section description",
            "section_type": "blue",  # or "pale" or "none"
            "slides": [
                {
                    "type": "content",
                    "title": "SLIDE TITLE",
                    "subtitle": "Short Subtitle (5-10 words)",
                    "intro": "1-2 sentence intro paragraph.",
                    "bullets": ["Point 1", "Point 2", "Point 3"]
                },
                # Add more slides as needed — sections expand automatically
            ]
        },
        # Add more sections as needed
    ]
}
```

**Section options:**
- `"section_type": "blue"` — Default blue section header
- `"section_type": "pale"` — Pale section header (good for caveats/limitations)
- `"section_type": "none"` — Skip section header, just content slides

**Slide types:**
- `"type": "content"` — Title + Subtitle + Body (intro + bullets)
- `"type": "quote"` — Quote slide with attribution

**Layout alternation:** Content slides automatically alternate between white and pale backgrounds unless you specify `"layout": "content_white"` or `"layout": "content_pale"` explicitly.

---

## Output Conventions

### Naming

`DD-MM-YY-Presentation-Title.pptx`

### Location

Ask user or save to project folder.

---

## Checklist Before Generating

- [ ] Presentation type confirmed (if multiple modes)
- [ ] Source content reviewed (outline, brief, or transcript)
- [ ] Section count determined
- [ ] Output location confirmed

## Checklist After Generating

- [ ] Correct layouts used for each slide type
- [ ] Subtitles are short and succinct (5-10 words)
- [ ] Body has intro paragraph before bullets
- [ ] Arrow bullets (Wingdings) used, not text characters
- [ ] Maximum 4 bullets per slide
- [ ] Dissolve animations added to all slides (except fixed slides)
- [ ] Fixed slides use template-inherited animations
- [ ] File saved to correct location with naming convention
