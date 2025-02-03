import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement

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

        if bold or run.bold:
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
        return "heading_1"
    elif "heading 2" in style_name:
        return "heading_2"
    elif "heading 3" in style_name:
        return "heading_3"
    elif "heading 4" in style_name:
        return "heading_4"
    return "body"

# Function to format the document
def format_docx(file_path):
    doc = docx.Document(file_path)

    # Format content dynamically based on styles
    for para in doc.paragraphs:
        section_type = identify_section(para)
        
        if section_type == "title":
            apply_formatting(para, "Times New Roman", 14, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        elif section_type == "author":
            apply_formatting(para, "Times New Roman", 12, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        elif section_type == "heading_1":
            apply_formatting(para, "Times New Roman", 12, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "heading_2":
            apply_formatting(para, "Times New Roman", 12, bold=True, italic=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "heading_3" or section_type == "heading_4":
            apply_formatting(para, "Times New Roman", 12, italic=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        else:
            # For normal text, preserve bold if present in the input
            apply_formatting(para, "Times New Roman", 10, alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)

    # Save formatted document
    output_filename = "formatted_document.docx"
    doc.save(output_filename)
    return output_filename

