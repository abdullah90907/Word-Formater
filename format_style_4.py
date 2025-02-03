import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Function to apply formatting
def apply_formatting(paragraph, font_name, font_size, bold=False, italic=False, alignment=None):
    """Applies formatting to the paragraph."""
    for run in paragraph.runs:
        font = run.font
        font.name = font_name
        font.size = Pt(font_size)
        font.bold = bold
        font.italic = italic
        font.color.rgb = RGBColor(0, 0, 0)  # Black text

        # If the run itself is bold (from the input), we should preserve it
        if run.bold:
            font.bold = True

    if alignment:
        paragraph.alignment = alignment

# Function to identify sections based on style from input DOCX
def identify_section(paragraph):
    """Identifies the section type based on the style name from the input document."""
    style_name = paragraph.style.name.lower()

    if "title" in style_name:
        return "title"
    elif "author" in style_name:
        return "author"
    elif "heading 1" in style_name:
        return "heading"
    elif "heading 2" in style_name or "heading 3" in style_name or "heading 4" in style_name or "heading 5" in style_name:
        return "subheading"
    elif "affiliation" in style_name:
        return "affiliation"
    elif "history" in style_name:
        return "history"
    return "body"

# Function to format the document
def format_docx(file_path):
    doc = docx.Document(file_path)
    
    # Set up two-column layout (except for the first page)
    for section in doc.sections:
        section.start_type = WD_SECTION_START.NEW_PAGE
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        
        # Set two equal columns (excluding the title page)
        sectPr = section._sectPr
        cols = OxmlElement('w:cols')
        cols.set(qn('w:num'), '2')
        cols.set(qn('w:space'), '720')  # Adjust spacing between columns
        sectPr.append(cols)
    
    # Format content dynamically based on styles
    for para in doc.paragraphs:
        section_type = identify_section(para)
        
        if section_type == "title":
            apply_formatting(para, "Times New Roman", 18, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        elif section_type == "author":
            apply_formatting(para, "Times New Roman", 10, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        elif section_type == "heading":
            apply_formatting(para, "Times New Roman", 12, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "subheading":
            apply_formatting(para, "Times New Roman", 12, italic=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "affiliation":
            # # Check if bold is applied in the input, and preserve it
            # is_bold = any(run.bold for run in para.runs)
            apply_formatting(para, "Times New Roman", 8, italic=True, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        elif section_type == "history":
            # Check if bold is applied in the input, and preserve it
            is_bold = any(run.bold for run in para.runs)
            apply_formatting(para, "Times New Roman", 8, bold=is_bold, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        else:
            apply_formatting(para, "Times New Roman", 10, alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)
    
    # Save formatted document
    output_filename = "formatted_document.docx"
    doc.save(output_filename)
    return output_filename

