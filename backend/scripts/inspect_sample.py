import docx
from docx.shared import Cm, Pt

def inspect_file(file_path):
    print(f"\n================ INSPECTING {file_path} ================")
    doc = docx.Document(file_path)
    
    print(f"Number of sections: {len(doc.sections)}")
    print(f"Number of tables: {len(doc.tables)}")
    print(f"Number of paragraphs: {len(doc.paragraphs)}")
    
    # Check margins of all sections
    for i, sec in enumerate(doc.sections):
        print(f"\nSection {i}:")
        print(f"  Orientation: {sec.orientation}")
        print(f"  Page Width: {sec.page_width.cm if sec.page_width else 'None'} cm")
        print(f"  Page Height: {sec.page_height.cm if sec.page_height else 'None'} cm")
        print(f"  Top Margin: {sec.top_margin.cm if sec.top_margin else 'None'} cm")
        print(f"  Bottom Margin: {sec.bottom_margin.cm if sec.bottom_margin else 'None'} cm")
        print(f"  Left Margin: {sec.left_margin.cm if sec.left_margin else 'None'} cm")
        print(f"  Right Margin: {sec.right_margin.cm if sec.right_margin else 'None'} cm")
        print(f"  Gutter: {sec.gutter.cm if sec.gutter else 'None'} cm")
        
    print("\n--- DETAILED PARAGRAPH FORMATTING ---")
    non_empty_count = 0
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if text:
            non_empty_count += 1
            pf = p.paragraph_format
            print(f"\nP {i} (Non-empty #{non_empty_count}): {repr(text[:80])}")
            print(f"  Alignment: {p.alignment}")
            print(f"  Space Before: {pf.space_before.pt if pf.space_before else 'Default'} pt")
            print(f"  Space After: {pf.space_after.pt if pf.space_after else 'Default'} pt")
            print(f"  Line Spacing: {pf.line_spacing} (Rule: {pf.line_spacing_rule})")
            print(f"  Left Indent: {pf.left_indent.cm if pf.left_indent else 'Default'} cm")
            print(f"  Right Indent: {pf.right_indent.cm if pf.right_indent else 'Default'} cm")
            
            # Fonts inside runs
            run_fonts = []
            for r in p.runs:
                run_fonts.append(f"{r.font.name or 'Default'}:{r.font.size.pt if r.font.size else 'Default'}pt")
            print(f"  Runs (Font:Size): {list(set(run_fonts))}")

if __name__ == '__main__':
    inspect_file('/app/templates/sample.docx')
