import sys
import docx
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

def format_document(file_path):
    print(f"Formatting document: {file_path}")
    doc = Document(file_path)
    
    # Remove empty paragraphs that precede the second and third notary headers
    # to prevent them from spilling over and creating vacant pages
    paragraphs_to_remove = []
    notary_count = 0
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if text.startswith("Before the Notary Public"):
            notary_count += 1
            if notary_count in [2, 3]:
                j = i - 1
                while j >= 0:
                    prev_p = doc.paragraphs[j]
                    if not prev_p.text.strip():
                        has_pb = False
                        for run in prev_p.runs:
                            for br in run._element.findall(qn('w:br')):
                                if br.get(qn('w:type')) == 'page':
                                    has_pb = True
                                    break
                            if has_pb:
                                break
                        if not has_pb:
                            paragraphs_to_remove.append(prev_p)
                        j -= 1
                    else:
                        break

    for p in paragraphs_to_remove:
        try:
            p._element.getparent().remove(p._element)
        except Exception as e:
            print(f"Error removing empty paragraph: {e}")
            
    # 1. Page Setup (apply A4 and 14cm top margin to all sections)
    for section in doc.sections:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width  = Cm(21.0)   # A4 width
        section.page_height = Cm(29.7)   # A4 height
        section.top_margin    = Cm(14.0)  # 14cm top margin for stamp paper
        section.bottom_margin = Cm(1.0)   # 1cm bottom margin
        section.left_margin   = Cm(1.0)   # 1cm left margin
        section.right_margin  = Cm(1.0)   # 1cm right margin
        section.gutter        = Cm(0.0)

    # 2. Font sizing and spacing adjustments
    # 10.5pt is highly legible and guarantees fitting within the limited 14.7cm printable area
    font_size = 10.5
    
    # Count notary headers to insert page breaks programmatically before the 2nd and 3rd affidavits
    notary_headers_count = 0
    
    # Iterate through all paragraphs and apply strict styling
    prev_text = ""
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        text_lower = text.lower()
        
        # Check if the paragraph contains a page break
        has_page_break = False
        for run in paragraph.runs:
            for br in run._element.findall(qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    has_page_break = True
                    break
            if has_page_break:
                break
                
        # Programmatic page break before the 2nd and 3rd notary public headers
        if text.startswith("Before the Notary Public"):
            notary_headers_count += 1
            if notary_headers_count in [2, 3]:
                paragraph.paragraph_format.page_break_before = True
        
        # If it's a completely empty paragraph, reduce its size to 1pt to prevent unnecessary vertical space
        if not text:
            if has_page_break:
                # Keep page break paragraph height/spacing clean but intact
                pf = paragraph.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = Pt(12)
                continue
                
            pf = paragraph.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = Pt(1)
            continue
            
        # Determine alignment based on legal document standards
        if (
            text.startswith("Before the Notary") or 
            text == "AFFIDAVIT" or 
            "VERIFICATION" in text or
            "testifying the character" in text_lower or
            "(vide item" in text_lower
        ):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif (
            "signature of applicant" in text_lower or 
            "identified by" in text_lower or 
            "mou roy" in text_lower or 
            "debabrata sarkar" in text_lower or 
            "advocate" in text_lower or 
            "adv," in text_lower or 
            "identified by me" in prev_text.lower()
        ):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Save original indentation to check for custom template layout (like witness signature blocks)
        orig_left = paragraph.paragraph_format.left_indent
        orig_first = paragraph.paragraph_format.first_line_indent

        # Determine if it is a list paragraph (needs special indent handling)
        pPr = paragraph._element.pPr
        is_list = False
        if pPr is not None and pPr.find(qn('w:numPr')) is not None:
            is_list = True

        # Paragraph format parameters
        pf = paragraph.paragraph_format
        pf.space_before = Pt(0)  # No extra spacing before
        pf.space_after  = Pt(0)  # No extra spacing after
        
        if is_list:
            # 0.8cm indentation for list text, with -0.8cm first line indent to place
            # numbering exactly at the left margin
            pf.left_indent  = Cm(0.8)
            pf.right_indent = Cm(0)
            pf.first_line_indent = Cm(-0.8)
        else:
            if orig_left is not None and orig_left > Cm(0.1):
                # Keep custom witness block formatting
                pf.left_indent = orig_left
                pf.first_line_indent = orig_first
            else:
                pf.left_indent  = Cm(0)
                pf.first_line_indent = None
            pf.right_indent = Cm(0)
        
        # Compact line spacing (exactly 13pt) to prevent text overflow
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(13)
        
        # Keep paragraph content together
        pf.keep_together = True

        # Apply Bookman Old Style font to all runs
        for run in paragraph.runs:
            run.font.name = 'Bookman Old Style'
            run.font.size = Pt(font_size)
            
            # Apply XML elements to make sure Word/LibreOffice render the font correctly
            rPr = run._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = run._element.makeelement(qn('w:rFonts'), {})
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:ascii'), 'Bookman Old Style')
            rFonts.set(qn('w:hAnsi'), 'Bookman Old Style')
            rFonts.set(qn('w:cs'),    'Bookman Old Style')
            
        if text:
            prev_text = text

    # Apply same formatting inside tables (if any tables are present)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text.strip()
                    if not text:
                        continue
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    pf = paragraph.paragraph_format
                    pf.space_before = Pt(0)
                    pf.space_after  = Pt(0)
                    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                    pf.line_spacing = Pt(13)
                    for run in paragraph.runs:
                        run.font.name = 'Bookman Old Style'
                        run.font.size = Pt(font_size)
                        rPr = run._element.get_or_add_rPr()
                        rFonts = rPr.find(qn('w:rFonts'))
                        if rFonts is None:
                            rFonts = run._element.makeelement(qn('w:rFonts'), {})
                            rPr.insert(0, rFonts)
                        rFonts.set(qn('w:ascii'), 'Bookman Old Style')
                        rFonts.set(qn('w:hAnsi'), 'Bookman Old Style')
                        rFonts.set(qn('w:cs'),    'Bookman Old Style')
                        
    doc.save(file_path)
    print("Formatting completed successfully!")

if __name__ == '__main__':
    if len(sys.argv) > 1:
        format_document(sys.argv[1])
    else:
        print("Error: No file path provided.")
