# python-pptx Bullet Style Feature

This directory contains the enhanced python-pptx library with the `bullet_style` property for easy bullet and numbered list creation.

## Quick Start

```python
from pptx import Presentation
from pptx.util import Pt, Inches

# Create presentation
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Add text box
textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
tf = textbox.text_frame

# Title (no bullet)
p = tf.paragraphs[0]
p.text = "My List"
p.font.size = Pt(24)
p.font.bold = True

# Bullet points
p1 = tf.add_paragraph()
p1.text = "First bullet item"
p1.bullet_style = "bullet"
p1.font.size = Pt(18)

# Indented bullet (level 1)
p2 = tf.add_paragraph()
p2.text = "Sub-item"
p2.level = 1
p2.bullet_style = "bullet"
p2.font.size = Pt(16)

# Numbered list
p3 = tf.add_paragraph()
p3.text = "First numbered item"
p3.bullet_style = "number"
p3.font.size = Pt(18)

prs.save("output.pptx")
```

## API Reference

### `paragraph.bullet_style`

Read/write property that controls bullet formatting.

| Value | Result |
|-------|--------|
| `"bullet"` | Character bullet (•) |
| `"number"` | Auto-numbered list (1., 2., 3.) |
| `None` or `""` | No bullet (plain text) |

### `paragraph.level`

Read/write property (0-8) that controls indentation level. Works with `bullet_style` for nested lists.

**Important:** Set `level` before `bullet_style` for proper indentation calculation.

```python
p.level = 0   # Main level (0.125" indent)
p.level = 1   # First sub-level (0.25" indent)
p.level = 2   # Second sub-level (0.375" indent)
```

## Reading Existing Bullets

The `bullet_style` property correctly reads existing bullets from PPTX files:

```python
from pptx import Presentation

prs = Presentation("existing_file.pptx")

for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                print(f"Text: {p.text}")
                print(f"Bullet: {p.bullet_style}")  # Returns "bullet", "number", or None
```

## Test Scripts

### `test_roundtrip_bullets.py` - Main Test Suite

Comprehensive test suite with 4 tests:

```bash
python src/test_roundtrip_bullets.py
```

**Tests:**
1. **Read existing bullets** - Opens PPTX with bullets, verifies `bullet_style` returns correct values
2. **Save and reload** - Round-trip test ensuring bullets survive save/load
3. **Modify existing** - Adds new bullets to existing presentation
4. **Create from scratch** - Creates new presentation with mixed bullet types

### `test_bullet_on_existing_pptx.py` - Real-World Test

Tests bullet functionality on a complex real presentation:

```bash
python src/test_bullet_on_existing_pptx.py
```

### `inspect_bullets.py` - XML Introspection

Shows the underlying XML structure for debugging:

```bash
python src/inspect_bullets.py
```

Output shows:
- Paragraph bullet_style values
- Underlying XML elements (`buChar`, `buAutoNum`)
- Raw XML for paragraphs with bullets

### `demo_bullet_style.py` - Feature Demo

Simple demonstration of all bullet_style capabilities:

```bash
python src/demo_bullet_style.py
```

## Output Files

All test scripts output to `src/demo_test_results/`:

| File | Description |
|------|-------------|
| `PresentationWBullets_roundtrip.pptx` | Exact copy (round-trip verification) |
| `PresentationWBullets_modified.pptx` | Original + new bullets added |
| `PresentationWBullets_created.pptx` | Created from scratch |
| `image_exercise_with_bullets.pptx` | Modified real-world template |
| `bullet_style_demo.pptx` | Feature demonstration |

## Examples

### Mixed Bullet Types

```python
# Plain text header
p = tf.paragraphs[0]
p.text = "Project Status"
p.font.bold = True

# Bullet points
for item in ["Completed design", "In progress: development", "Pending: testing"]:
    p = tf.add_paragraph()
    p.text = item
    p.bullet_style = "bullet"
    p.font.size = Pt(14)

# Numbered action items
for i, action in enumerate(["Review code", "Fix bugs", "Deploy"]):
    p = tf.add_paragraph()
    p.text = action
    p.bullet_style = "number"
    p.font.size = Pt(14)
```

### Nested Bullets

```python
items = [
    ("Main topic 1", 0),
    ("Sub-point A", 1),
    ("Sub-point B", 1),
    ("Detail under B", 2),
    ("Main topic 2", 0),
]

for text, level in items:
    p = tf.add_paragraph()
    p.text = text
    p.level = level  # Set level BEFORE bullet_style
    p.bullet_style = "bullet"
    p.font.size = Pt(18 - level * 2)  # Smaller font for sub-items
```

### Remove Bullets from Existing Paragraphs

```python
for p in tf.paragraphs:
    if p.bullet_style is not None:
        p.bullet_style = None  # Removes bullet formatting
```

## How It Works

The `bullet_style` property maps to OpenXML elements:

| Property Value | XML Element |
|----------------|-------------|
| `"bullet"` | `<a:buChar char="•"/>` |
| `"number"` | `<a:buAutoNum type="arabicPeriod"/>` |
| `None` | `<a:buNone/>` or no element |

When reading existing files, the getter checks which element exists:

```python
@property
def bullet_style(self):
    if pPr.buChar is not None:
        return "bullet"
    if pPr.buAutoNum is not None:
        return "number"
    return None
```

## File Structure

```
src/
├── pptx/
│   ├── text/
│   │   └── text.py          # _Paragraph.bullet_style property
│   └── oxml/
│       ├── __init__.py      # Element class registration
│       └── text.py          # CT_TextCharBullet, CT_TextAutonumberBullet
├── demo_test_results/       # Output directory for test files
├── test_roundtrip_bullets.py
├── test_bullet_on_existing_pptx.py
├── inspect_bullets.py
├── demo_bullet_style.py
├── PresentationWBullets.pptx  # Test input file
└── image_exercise.pptx        # Complex test input
```

## Requirements

```
python >= 3.9
lxml
Pillow
```

Install dependencies:
```bash
pip install lxml Pillow
```

## Known Limitations

1. Bullet character is fixed to `•` (U+2022)
2. Numbering type is fixed to `arabicPeriod` (1., 2., 3.)
3. Bullet font and color are not configurable through this API
4. Start number for numbered lists is not exposed
5. Bullets inherit paragraph alignment from template - use left-aligned text for proper nested indentation

See `IMPLEMENTATION.md` for technical details and future enhancement plans.
