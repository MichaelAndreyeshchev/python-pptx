"""
Demo script to test the bullet_style feature in python-pptx.

This script demonstrates how to use the new bullet_style property
on paragraphs to create bulleted and numbered lists.

Usage:
    1. Apply the patch to python-pptx:
       cd python-pptx
       git apply python-pptx-bullet-style.patch
       pip install -e .
    
    2. Run this script:
       python demo_bullet_style.py
    
    3. Open the generated 'bullet_style_demo.pptx' in PowerPoint
"""

from pptx import Presentation
from pptx.util import Inches, Pt

# Create a new presentation
prs = Presentation()

# Add a slide with a blank layout
slide_layout = prs.slide_layouts[6]  # Blank layout
slide = prs.slides.add_slide(slide_layout)

# Add a text box
left = Inches(1)
top = Inches(1)
width = Inches(8)
height = Inches(5)

txBox = slide.shapes.add_textbox(left, top, width, height)
tf = txBox.text_frame

# First paragraph - no bullet (default)
p1 = tf.paragraphs[0]
p1.text = "This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default)"
p1.font.size = Pt(14)
print(f"Paragraph 1 - No bullet: bullet_style = {p1.bullet_style}")

# Second paragraph - bullet style (as requested in the example)
p2 = tf.add_paragraph()
p2.text = "$XXk annual software spend, split by XX per segment This is a paragraph without any bullet style (default) This is a paragraph without any bu This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default)llet style (default) This is a paragraph without any bullet style (default)"
p2.level = 0
p2.bullet_style = "bullet"
p2.space_before = Pt(6)  # Add spacing before bullet items
p2.font.size = Pt(14)
print(f"Paragraph 2 - Bullet: bullet_style = {p2.bullet_style}")

# Third paragraph - another bullet
p3 = tf.add_paragraph()
p3.text = "Cost breakdown by department This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default)"
p3.level = 0
p3.bullet_style = "bullet"
p3.space_before = Pt(6)  # Add spacing before bullet items
p3.font.size = Pt(14)
print(f"Paragraph 3 - Bullet: bullet_style = {p3.bullet_style}")

# Fourth paragraph - numbered style
p4 = tf.add_paragraph()
p4.text = "First action item This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default)"
p4.level = 0
p4.bullet_style = "number"
p4.space_before = Pt(6)  # Add spacing before numbered items
p4.font.size = Pt(14)
print(f"Paragraph 4 - Number: bullet_style = {p4.bullet_style}")

# Fifth paragraph - also numbered (continues numbering)
p5 = tf.add_paragraph()
p5.text = "Second action item"
p5.level = 0
p5.bullet_style = "number"
p5.space_before = Pt(6)  # Add spacing before numbered items
p5.font.size = Pt(14)
print(f"Paragraph 5 - Number: bullet_style = {p5.bullet_style}")

# Sixth  paragraph - no bullet (default)
p6 = tf.add_paragraph()
p6.text = "This is a paragraph without any bullet style (default) This is a paragraph without any bullet style (default)"
p6.font.size = Pt(14)
print(f"Paragraph 6 - No bullet: bullet_style = {p6.bullet_style}")


# Save the presentation
output_path = "bullet_style_demo.pptx"
prs.save(output_path)
print(f"\nPresentation saved to: {output_path}")
print("Open the file in PowerPoint to verify the bullet styles are correct.")
