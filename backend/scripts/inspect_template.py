import os
import docx
from docx.shared import Pt, Inches, Cm

template_path = os.path.join("backend", "templates", "template.docx")
if not os.path.exists(template_path):
    # Try absolute path
    template_path = r"e:\Projects\caa-court-affidavit\backend\templates\template.docx"

if not os.path.exists(template_path):
    print(f"Error: Template not found at {template_path}")
    exit(1)

print(f"Inspecting local template: {template_path}")
doc = docx.Document(template_path)

print("\n--- SECTIONS INFO ---")
for i, sec in enumerate(doc.sections):
    print(f"Section {i}:")
    print(f"  Orientation: {sec.orientation}")
    print(f"  Page Width: {sec.page_width.cm if sec.page_width else 'None'} cm")
    print(f"  Page Height: {sec.page_height.cm if sec.page_height else 'None'} cm")
    print(f"  Top Margin: {sec.top_margin.cm if sec.top_margin else 'None'} cm")
    print(f"  Bottom Margin: {sec.bottom_margin.cm if sec.bottom_margin else 'None'} cm")
    print(f"  Left Margin: {sec.left_margin.cm if sec.left_margin else 'None'} cm")
    print(f"  Right Margin: {sec.right_margin.cm if sec.right_margin else 'None'} cm")
    print(f"  Gutter: {sec.gutter.cm if sec.gutter else 'None'} cm")

print("\n--- FIRST 20 PARAGRAPHS FORMATTING ---")
for i in range(min(50, len(doc.paragraphs))):
    p = doc.paragraphs[i]
    text = p.text.strip()
    if text:
        pf = p.paragraph_format
        print(f"P {i}: {repr(text[:60])}")
        print(f"  Alignment: {p.alignment}")
        print(f"  Space Before: {pf.space_before.pt if pf.space_before else 'Default'} pt")
        print(f"  Space After: {pf.space_after.pt if pf.space_after else 'Default'} pt")
        print(f"  Line Spacing: {pf.line_spacing} (Rule: {pf.line_spacing_rule})")
        print(f"  Left Indent: {pf.left_indent.cm if pf.left_indent else 'Default'} cm")
        print(f"  Right Indent: {pf.right_indent.cm if pf.right_indent else 'Default'} cm")
        
        # Check fonts in runs
        run_fonts = []
        for r in p.runs:
            run_fonts.append(f"{r.font.name or 'Default'}:{r.font.size.pt if r.font.size else 'Default'}pt")
        print(f"  Runs: {list(set(run_fonts))}")
