import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Function to apply formatting
def apply_formatting(paragraph, font_name, font_size, bold=False, italic=False, underline=False, alignment=None):
    """Applies formatting to the paragraph."""
    for run in paragraph.runs:
        font = run.font
        font.name = font_name
        font.size = Pt(font_size)
        font.bold = bold if bold else run.bold  # Keep existing bold if present
        font.italic = italic
        font.underline = underline
        font.color.rgb = RGBColor(0, 0, 0)  # Black text

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
    elif "affiliation" in style_name:
        return "affiliation"
    elif "abstract" in style_name:
        return "abstract"
    elif "keyword" in style_name:
        return "keyword"
    elif "articletype" in style_name:
        return "articletype"
    elif "doinum" in style_name:
        return "doinum"
    elif "BackMatter" in style_name:
        return "BackMatter"
    elif "heading2" in style_name:
        return "heading_2"
    elif "heading3" in style_name:
        return "heading_3"
    elif "heading4" in style_name:
        return "heading_4"
    return "body"

# Function to format the document
def format_docx(file_path):
    doc = docx.Document(file_path)

    # Format content dynamically based on styles
    for para in doc.paragraphs:
        section_type = identify_section(para)
        
        if section_type == "title":
            apply_formatting(para, "Minion Pro", 14, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "author":
            apply_formatting(para, "Minion Pro", 12, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "affiliation":
            apply_formatting(para, "Minion Pro", 9, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "abstract" or section_type == "keyword":
            apply_formatting(para, "Minion Pro", 10, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
            # Keep bold only if originally present
            for run in para.runs:
                run.font.bold = run.bold
        elif section_type == "BackMatter":
            apply_formatting(para, "Minion Pro", 10, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
            # Keep bold only if originally present
            for run in para.runs:
                run.font.bold = run.bold
        elif section_type == "articletype":
            apply_formatting(para, "Minion Pro", 9, bold=True, underline=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "doinum":
            apply_formatting(para, "Minion Pro", 7, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "heading_1":
            apply_formatting(para, "Minion Pro", 11, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "heading_2":
            apply_formatting(para, "Minion Pro", 11, bold=True, italic=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "heading_3" or section_type == "heading_4":
            apply_formatting(para, "Minion Pro", 11, italic=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        else:
            # For normal text, preserve bold if present in the input
            apply_formatting(para, "Minion Pro", 10, alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)
            for run in para.runs:
                run.font.bold = run.bold

    # Save formatted document
    output_filename = "formatted_document.docx"
    doc.save(output_filename)
    return output_filename
