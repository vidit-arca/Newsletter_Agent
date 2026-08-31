from pptx import Presentation
prs = Presentation('/Users/apple/Desktop/Akshayam/Newsletter_Agent/output/output.pptx')
slide = prs.slides[0]
for i, shape in enumerate(slide.shapes):
    if shape.has_text_frame:
        print(f"Shape {i}: {shape.text!r}")
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.font.color and r.font.color.type == 1: # RGB
                    print(f"  Color: {r.font.color.rgb}")
                else:
                    print("  No explicit RGB color set")
