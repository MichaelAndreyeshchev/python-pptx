# bullet_style Implementation Details

This document describes the implementation of the `bullet_style` property in python-pptx, which provides a simple API for adding bullet points and numbered lists to PowerPoint paragraphs.

## Overview

The `bullet_style` property is implemented on the `_Paragraph` class in `src/pptx/text/text.py`. It provides read/write access to paragraph bullet formatting through a simple string-based API.

## API

```python
paragraph.bullet_style = "bullet"   # Character bullet (•)
paragraph.bullet_style = "number"   # Auto-numbered list (1., 2., 3.)
paragraph.bullet_style = None       # No bullet (plain text)

# Reading
style = paragraph.bullet_style      # Returns "bullet", "number", or None
```

## Architecture

### Files Modified

| File | Changes |
|------|---------|
| `src/pptx/text/text.py` | Added `bullet_style` property getter/setter to `_Paragraph` class |
| `src/pptx/oxml/text.py` | Added `CT_TextCharBullet` and `CT_TextAutonumberBullet` element classes |
| `src/pptx/oxml/__init__.py` | Registered `a:buChar` and `a:buAutoNum` element classes |

### XML Structure

The bullet style maps to OpenXML elements within the paragraph properties (`a:pPr`):

```xml
<!-- Character bullet (bullet_style = "bullet") -->
<a:pPr marL="114300" indent="-114300">
  <a:buChar char="•"/>
</a:pPr>

<!-- Auto-numbered (bullet_style = "number") -->
<a:pPr marL="114300" indent="-114300">
  <a:buAutoNum type="arabicPeriod"/>
</a:pPr>

<!-- No bullet (bullet_style = None) -->
<a:pPr>
  <a:buNone/>
</a:pPr>
```

### Element Class Definitions

```python
# src/pptx/oxml/text.py

class CT_TextCharBullet(BaseOxmlElement):
    """`a:buChar` element for character bullet lists."""
    char: str = RequiredAttribute("char", ST_TextTypeface)


class CT_TextAutonumberBullet(BaseOxmlElement):
    """`a:buAutoNum` element for auto-numbered bullet lists."""
    type: str = RequiredAttribute("type", ST_TextTypeface)
    startAt: int | None = OptionalAttribute("startAt", ST_TextFontSize, default=1)
```

### Property Implementation

```python
# src/pptx/text/text.py - _Paragraph class

@property
def bullet_style(self) -> str | None:
    """Bullet style of this paragraph.
    
    Read/write. Valid values are None (or empty string) for no bullet, 
    "bullet" for a character bullet, and "number" for an auto-numbered list.
    """
    pPr = self._p.pPr
    if pPr is None:
        return None
    if pPr.buChar is not None:
        return "bullet"
    if pPr.buAutoNum is not None:
        return "number"
    if pPr.buNone is not None:
        return None
    return None

@bullet_style.setter
def bullet_style(self, value: str | None):
    pPr = self._pPr
    if value is None or value == "":
        pPr._remove_eg_buType()
        pPr.marL = None
        pPr.indent = None
    elif value == "bullet":
        buChar = pPr.get_or_change_to_buChar()
        buChar.char = "\u2022"  # Bullet character •
        pPr.marL = self._calculate_marL(self.level)
        pPr.indent = -114300  # Hanging indent
    elif value == "number":
        buAutoNum = pPr.get_or_change_to_buAutoNum()
        buAutoNum.type = "arabicPeriod"
        pPr.marL = self._calculate_marL(self.level)
        pPr.indent = -114300  # Hanging indent
    else:
        raise ValueError(
            f"bullet_style must be None, '', 'bullet', or 'number', got {value!r}"
        )
```

### Indentation Calculation

Bullet indentation scales with the paragraph's `level` property:

```python
def _calculate_marL(self, level: int) -> int:
    """Calculate left margin based on indentation level.
    
    Each level adds 114300 EMUs (0.125 inch) of indentation.
    Level 0: 114300 EMUs (0.125 inch)
    Level 1: 228600 EMUs (0.25 inch)
    Level 2: 342900 EMUs (0.375 inch)
    etc.
    """
    base_indent = 114300  # 0.125 inch in EMUs
    return base_indent * (level + 1)
```

## How Reading Works

When you open an existing PPTX file with bullets:

1. **Package Extraction**: The PPTX (ZIP file) is opened and XML parts are extracted
2. **XML Parsing**: lxml parses the XML with custom element class lookup
3. **Element Registration**: `a:buChar` → `CT_TextCharBullet`, `a:buAutoNum` → `CT_TextAutonumberBullet`
4. **Property Access**: `bullet_style` getter checks which element exists and returns the appropriate string

### Example XML from Existing File

```xml
<a:p>
  <a:pPr marL="342900" indent="-342900">
    <a:buFont typeface="Arial"/>
    <a:buChar char="•"/>
  </a:pPr>
  <a:r>
    <a:rPr lang="en-US"/>
    <a:t>Bullet 1</a:t>
  </a:r>
</a:p>
```

When `bullet_style` is accessed on this paragraph, it returns `"bullet"` because `pPr.buChar is not None`.

## Round-Trip Preservation

The implementation preserves existing bullet formatting through round-trips:

1. **Read**: Existing XML elements are parsed into element objects
2. **Save**: Element objects are serialized back to XML
3. **Result**: Original bullet formatting is preserved exactly

## Testing

### Test Files

| File | Purpose |
|------|---------|
| `src/test_roundtrip_bullets.py` | Comprehensive test suite (4 tests) |
| `src/test_bullet_on_existing_pptx.py` | Tests on real-world template |
| `src/inspect_bullets.py` | XML introspection utility |
| `src/demo_bullet_style.py` | Feature demonstration |

### Running Tests

```bash
cd C:\Users\micha\python-pptx
python src/test_roundtrip_bullets.py
```

### Output Directory

All generated files are saved to `src/demo_test_results/`:
- `PresentationWBullets_roundtrip.pptx` - Exact copy (round-trip test)
- `PresentationWBullets_modified.pptx` - With added bullets
- `PresentationWBullets_created.pptx` - Created from scratch
- `image_exercise_with_bullets.pptx` - Modified real-world template
- `bullet_style_demo.pptx` - Feature demonstration

## Limitations

1. **Bullet Character**: Currently hardcoded to `•` (U+2022). Custom characters not yet exposed.
2. **Numbering Type**: Currently hardcoded to `arabicPeriod` (1., 2., 3.). Other types not yet exposed.
3. **Bullet Font**: Not configurable through this API.
4. **Bullet Color**: Not configurable through this API.
5. **Start Number**: `startAt` attribute exists but not exposed through the simple API.
6. **Alignment**: Bullets inherit paragraph alignment from the template. For proper indentation with nested bullets, ensure text is left-aligned.

## Future Enhancements

Potential future additions:
- `bullet_char` property for custom bullet characters
- `numbering_type` property for different numbering schemes
- `bullet_color` property for colored bullets
- `start_at` property for numbered list start values
