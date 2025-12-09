"""
Test script demonstrating bullet_style on an existing PowerPoint presentation.

This script opens image_exercise.pptx and adds:
- Bullet points
- Numbered lists  
- Regular text entries

to various text boxes in the presentation.
"""

from pptx import Presentation
from pptx.util import Pt

def main():
    # Open the existing presentation
    print("Opening image_exercise.pptx...")
    prs = Presentation('image_exercise.pptx')
    
    slide = prs.slides[0]
    
    # Track which shapes we modified
    modifications = []
    
    # Find text boxes and add different content types
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
            
        tf = shape.text_frame
        
        # === SHAPE 7: "Textplatzhalter 2" - Add bullet list ===
        if shape.name == "Textplatzhalter 2" and shape.top < 3000000:  # First one
            tf.clear()  # Clear existing content
            
            # Title paragraph (no bullet)
            p1 = tf.paragraphs[0]
            p1.text = "Key Features"
            p1.font.size = Pt(14)
            p1.font.bold = True
            
            # Bullet points
            p2 = tf.add_paragraph()
            p2.text = "Cloud-native architecture"
            p2.bullet_style = "bullet"
            p2.font.size = Pt(11)
            p2.space_before = Pt(6)
            
            p3 = tf.add_paragraph()
            p3.text = "Real-time analytics"
            p3.bullet_style = "bullet"
            p3.font.size = Pt(11)
            p3.space_before = Pt(4)
            
            p4 = tf.add_paragraph()
            p4.text = "Enterprise security"
            p4.bullet_style = "bullet"
            p4.font.size = Pt(11)
            p4.space_before = Pt(4)
            
            modifications.append(f"Shape '{shape.name}': Added bullet list")
            
        # === SHAPE 8: Second "Textplatzhalter 2" - Add numbered list ===
        elif shape.name == "Textplatzhalter 2" and 3500000 < shape.top < 4500000:
            tf.clear()
            
            p1 = tf.paragraphs[0]
            p1.text = "Implementation Steps"
            p1.font.size = Pt(14)
            p1.font.bold = True
            
            p2 = tf.add_paragraph()
            p2.text = "Assessment & planning"
            p2.bullet_style = "number"
            p2.font.size = Pt(11)
            p2.space_before = Pt(6)
            
            p3 = tf.add_paragraph()
            p3.text = "System integration"
            p3.bullet_style = "number"
            p3.font.size = Pt(11)
            p3.space_before = Pt(4)
            
            p4 = tf.add_paragraph()
            p4.text = "Testing & deployment"
            p4.bullet_style = "number"
            p4.font.size = Pt(11)
            p4.space_before = Pt(4)
            
            modifications.append(f"Shape '{shape.name}': Added numbered list")
            
        # === SHAPE 9: Third "Textplatzhalter 2" - Mixed content ===
        elif shape.name == "Textplatzhalter 2" and shape.top > 5000000:
            tf.clear()
            
            p1 = tf.paragraphs[0]
            p1.text = "Summary"
            p1.font.size = Pt(14)
            p1.font.bold = True
            
            p2 = tf.add_paragraph()
            p2.text = "This solution provides comprehensive coverage for all enterprise needs."
            p2.font.size = Pt(10)
            p2.space_before = Pt(6)
            
            modifications.append(f"Shape '{shape.name}': Added regular text")
            
        # === "Textfeld 78" with "Example" - Add bullet list ===
        elif shape.name == "Textfeld 78" and tf.paragraphs[0].text == "Example" and shape.top > 5000000:
            tf.clear()
            
            p1 = tf.paragraphs[0]
            p1.text = "Benefits"
            p1.font.size = Pt(12)
            p1.font.bold = True
            
            p2 = tf.add_paragraph()
            p2.text = "Cost reduction"
            p2.bullet_style = "bullet"
            p2.font.size = Pt(10)
            p2.space_before = Pt(4)
            
            p3 = tf.add_paragraph()
            p3.text = "Improved efficiency"
            p3.bullet_style = "bullet"
            p3.font.size = Pt(10)
            p3.space_before = Pt(4)
            
            modifications.append(f"Shape '{shape.name}': Added bullet benefits")
            
        # === "Textfeld 78" with "Exposure Management" - Add numbered list ===
        elif shape.name == "Textfeld 78" and "Exposure" in tf.paragraphs[0].text:
            tf.clear()
            
            p1 = tf.paragraphs[0]
            p1.text = "Risk Categories"
            p1.font.size = Pt(12)
            p1.font.bold = True
            
            p2 = tf.add_paragraph()
            p2.text = "Market risk"
            p2.bullet_style = "number"
            p2.font.size = Pt(10)
            p2.space_before = Pt(4)
            
            p3 = tf.add_paragraph()
            p3.text = "Credit risk"
            p3.bullet_style = "number"
            p3.font.size = Pt(10)
            p3.space_before = Pt(4)
            
            p4 = tf.add_paragraph()
            p4.text = "Operational risk"
            p4.bullet_style = "number"
            p4.font.size = Pt(10)
            p4.space_before = Pt(4)
            
            modifications.append(f"Shape '{shape.name}': Added numbered risk list")
            
        # === Update "Example" text in Textfeld 74 (right side) ===
        elif shape.name == "Textfeld 74" and tf.paragraphs[0].text == "Example":
            tf.clear()
            
            p1 = tf.paragraphs[0]
            p1.text = "Product Overview"
            p1.font.size = Pt(12)
            p1.font.bold = True
            
            p2 = tf.add_paragraph()
            p2.text = "Our solution addresses key market needs with innovative technology."
            p2.font.size = Pt(10)
            p2.space_before = Pt(6)
            
            p3 = tf.add_paragraph()
            p3.text = "Scalable infrastructure"
            p3.bullet_style = "bullet"
            p3.font.size = Pt(10)
            p3.space_before = Pt(4)
            
            p4 = tf.add_paragraph()
            p4.text = "24/7 support"
            p4.bullet_style = "bullet"
            p4.font.size = Pt(10)
            p4.space_before = Pt(4)
            
            modifications.append(f"Shape '{shape.name}': Added mixed content")
    
    # Save the modified presentation
    output_path = "image_exercise_with_bullets.pptx"
    prs.save(output_path)
    
    # Print summary
    print(f"\n=== Modifications Made ===")
    for mod in modifications:
        print(f"  [OK] {mod}")
    
    print(f"\nSaved modified presentation to: {output_path}")
    print("\nOpen the file in PowerPoint to verify:")
    print("  - Bullet points (•) with proper indentation")
    print("  - Numbered lists (1., 2., 3.) with proper indentation")
    print("  - Regular text paragraphs")
    print("  - Proper spacing between items")


def test_bullet_style_reading():
    """Test that bullet_style correctly reads existing bullets from the modified file."""
    print("\n=== Testing bullet_style reading ===")
    
    try:
        prs = Presentation('image_exercise_with_bullets.pptx')
    except FileNotFoundError:
        print("Run main() first to create the modified file")
        return
    
    slide = prs.slides[0]
    
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        
        tf = shape.text_frame
        has_bullets = False
        
        for p in tf.paragraphs:
            if p.bullet_style is not None:
                has_bullets = True
                break
        
        if has_bullets:
            print(f"\nShape: {shape.name}")
            for i, p in enumerate(tf.paragraphs):
                style = p.bullet_style if p.bullet_style else "none"
                text_preview = p.text[:40] + "..." if len(p.text) > 40 else p.text
                print(f"  P{i}: style={style:8s} text=\"{text_preview}\"")


if __name__ == "__main__":
    main()
    test_bullet_style_reading()

