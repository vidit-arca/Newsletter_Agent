from pptx import Presentation
prs = Presentation('/Users/apple/Desktop/Akshayam/Newsletter_Agent/output/output.pptx')
slide = prs.slides[1]
for i, shape in enumerate(slide.shapes):
    if shape.has_text_frame:
        print(f"Shape {i}: {shape.text[:100]!r}")
