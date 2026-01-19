# PowerPoint Template Preparation Guide

Before Claude Code can generate presentations from your template, you must prepare the template correctly. This guide walks you through every step. Try giving this guide to your AI and ask it to walk you through the process step-by-step.

---

## Why Preparation Matters

Claude Code generates slides by referencing **layout index numbers**. If your template has:
- 19 layouts → indices 0-18
- 5 layouts → indices 0-4

The Python script adds slides like this:
```python
slide = prs.slides.add_slide(prs.slide_layouts[13])  # Uses layout at index 13
```

If index 13 doesn't exist or contains the wrong layout, your deck will fail or look wrong.

---

## Step 1: Open Your Template in Slide Master View

1. Open your branded PowerPoint (or create a new one)
2. Go to **View → Slide Master**
3. You'll see a hierarchy:
   - **Slide Master** (top, larger thumbnail) — controls global styling
   - **Layouts** (below, smaller thumbnails) — individual slide types

---

## Step 2: Audit Your Layouts

For each layout in your template, document:

| Index | Layout Name | Purpose | Placeholders |
|-------|-------------|---------|--------------|
| 0 | Title Slide | Opening slide | Title, Subtitle |
| 1 | Section Header | Section dividers | Title, Subtitle |
| 2 | Content | Body content | Title, Subtitle, Body |
| ... | ... | ... | ... |

### Finding Placeholder Indices

Placeholders also have index numbers. To find them:

1. Click on a layout in Slide Master view
2. Click on each placeholder
3. Note the placeholder type and position

Common placeholder types:
- **Title** — usually index 0
- **Subtitle** — usually index 1
- **Body/Content** — often index 10, 13, or higher

You can also extract this programmatically (see "Extracting Layout Information" below).

---

## Step 3: Design or Refine Your Layouts

### Minimum Recommended Layouts

| Purpose | Description |
|---------|-------------|
| Title slide | Opening with main title and subtitle |
| Section header | Divides major sections |
| Content slide | Title + subtitle + body with bullets |
| Quote slide | For key statements or testimonials |
| CTA slide | Call-to-action (optional, can be fixed content) |
| Thank you/closing | Final slide |

### Design Principles

- **Consistent margins** — same spacing across all layouts
- **Consistent typography** — same fonts, sizes, line heights
- **Clear placeholders** — each placeholder serves one purpose
- **Colour contrast** — text readable on all backgrounds

---

## Step 4: Name Your Layouts Clearly

In Slide Master view:

1. Right-click each layout
2. Select **Rename Layout**
3. Use a clear naming convention with number prefix:

```
01 Main Title Slide
02 Blue Section Slide
03 Pale Section Slide
04 White Content Slide
05 Pale Content Slide
06 Quote Slide
07 CTA Slide
08 Thank You Slide
```

The number prefix helps you track which index corresponds to which layout.

---

## Step 5: Remove Unused Layouts

Delete any layouts you won't use:

1. Right-click the layout in Slide Master view
2. Select **Delete Layout**

Fewer layouts = simpler index mapping = fewer errors.

---

## Step 6: Add Pre-Set Animations (Fixed Slides Only)

For slides with **fixed content that never changes** (e.g., About slide, CTA slide):

1. Exit Slide Master view
2. Create a test slide using that layout
3. Add your preferred animations to all elements
4. Go back to Slide Master view
5. The animations should appear on the layout

**Why only fixed slides?** Dynamic slides (where content changes) get animations added by the Python script. Fixed slides inherit animations from the template to avoid duplication.

---

## Step 7: Save As Template

1. **File → Save As**
2. Choose format: **PowerPoint Template (.potx)**
3. Save to a known location

Alternatively, save as `.pptx` — the Python script can convert either format.

---

## Step 8: Extract Layout Information Programmatically

Run this Python script to extract your template's layout structure:

```python
from pptx import Presentation
from pathlib import Path

template_path = Path("your-template.potx")  # or .pptx
prs = Presentation(str(template_path))

print("LAYOUT INDEX MAPPING")
print("=" * 60)

for idx, layout in enumerate(prs.slide_layouts):
    print(f"\nIndex {idx}: {layout.name}")
    print("-" * 40)

    # Get placeholder info
    for shape in layout.placeholders:
        ph = shape.placeholder_format
        print(f"  Placeholder idx={ph.idx}, type={ph.type}, name={shape.name}")
```

Save the output — you'll need it to configure your skill.

---

## Step 9: Create Your Style Specification

Create a JSON file documenting your template's styling:

```json
{
  "template_file": "your-template.potx",

  "layouts": {
    "title": {
      "index": 0,
      "name": "Main Title Slide",
      "placeholders": [0, 1],
      "use_for": "Opening slide"
    },
    "section": {
      "index": 1,
      "name": "Section Slide",
      "placeholders": [0, 1],
      "use_for": "Section dividers"
    },
    "content": {
      "index": 2,
      "name": "Content Slide",
      "placeholders": [0, 1, 13],
      "use_for": "Body content with bullets"
    }
  },

  "typography": {
    "font_family": "Your Font Name",
    "line_spacing": 1.2,
    "title": { "size": "28pt", "case": "UPPER" },
    "subtitle": { "size": "24pt", "style": "bold" },
    "body": { "size": "22pt" }
  },

  "colours": {
    "accent_hex": "#YOUR_HEX",
    "text_on_dark": "#FFFFFF",
    "text_on_light": "#YOUR_HEX"
  },

  "bullets": {
    "font": "Wingdings",
    "character": "Ø",
    "margin": 342900,
    "indent": -342900
  },

  "animations": {
    "effect": "dissolve",
    "duration_ms": 500,
    "trigger": "on_click",
    "build_paragraphs": true
  }
}
```

---

## Step 10: Test Your Template

Before writing any code:

1. Create a test presentation manually using your template
2. Add slides using each layout
3. Verify:
   - Text appears in correct placeholders
   - Colours are correct on each background
   - Fonts render correctly
   - Fixed slides show their animations

---

## Common Issues

| Problem | Cause | Solution |
|---------|-------|----------|
| Wrong layout appears | Index doesn't match | Re-run extraction script, verify mapping |
| Text in wrong position | Wrong placeholder index | Check placeholder indices for that layout |
| Missing placeholder | Layout doesn't have that placeholder | Use a different layout or add placeholder in Slide Master |
| Animations duplicated | Both template and script add them | Remove from template OR remove from script |
| Font doesn't render | Font not installed | Use web-safe font or install font |

---

## Checklist Before Proceeding

- [ ] Template opened in Slide Master view and audited
- [ ] Unused layouts deleted
- [ ] Remaining layouts named clearly with number prefix
- [ ] Placeholder indices documented for each layout
- [ ] Fixed slides have pre-set animations
- [ ] Template saved as .potx (or .pptx)
- [ ] Layout extraction script run and output saved
- [ ] Style JSON created with all specifications
- [ ] Test presentation created manually to verify template

---

## Next Steps

Once your template is prepared:

1. Customise `SKILL-TEMPLATE.md` with your layout indices
2. Customise `style-template.json` with your specifications
3. Adapt `presentation-generator.py` for your structure
4. Work with Claude Code to test and refine

---

*Part of the PowerPoint AI Mastery course*
*Created by Larry G. Maguire | Larry G. Maguire Human Performance*
