import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml.ns import qn  

# Function to apply formatting
def apply_formatting(paragraph, font_name, font_size, bold=False, italic=False, underline=False, alignment=None):
    """Applies formatting to the paragraph."""
    for run in paragraph.runs:
        font = run.font
        font.name = font_name
        font.size = Pt(font_size)
        font.bold = bold if bold else run.bold  # Preserve existing bold
        font.italic = italic
        font.underline = underline
        font.color.rgb = RGBColor(0, 0, 0)  # Black text

    if alignment:
        paragraph.alignment = alignment

# Function to set a single-column layout
def set_single_column(doc):
    """Ensures the entire document uses a single-column layout."""
    for section in doc.sections:
        section.start_type = WD_SECTION_START.CONTINUOUS  # Prevent new page breaks
        section.page_width = Inches(8.5)  # Standard single-page width
        section.page_height = Inches(11)  # Standard height (A4)
        
        # Ensure columns are set to 1
        sectPr = section._sectPr
        cols = sectPr.xpath('./w:cols')
        if cols:
            cols[0].set(qn('w:num'), '1')  # Ensure single-column format

# Function to detect title, authors, and abstract index
def identify_sections(doc):
    """Identifies title, authors, and detects where 'Abstract' starts."""
    paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    
    title = paragraphs[0] if paragraphs else ""
    authors = paragraphs[1] if len(paragraphs) > 1 else ""

    abstract_index = next((i for i, text in enumerate(paragraphs) if text.lower().startswith("abstract")), None)
    
    return title, authors, abstract_index

# Function to identify section based on style
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

    # Detect title, authors, and abstract start index
    title, authors, abstract_index = identify_sections(doc)

    # Ensure single-column layout
    set_single_column(doc)

    # Format content dynamically based on styles
    before_abstract = True  # Flag to track text before Abstract

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        
        # Stop "before Abstract" formatting when Abstract is found
        if abstract_index is not None and i >= abstract_index:
            before_abstract = False

        section_type = identify_section(para)

        if text == title:
            apply_formatting(para, "Minion Pro", 14, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif text == authors:
            apply_formatting(para, "Minion Pro", 12, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "affiliation":
            apply_formatting(para, "Minion Pro", 9, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "abstract" or section_type == "keyword":
            apply_formatting(para, "Minion Pro", 10, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "BackMatter":
            apply_formatting(para, "Minion Pro", 10, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
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
        elif before_abstract:
            apply_formatting(para, "Minion Pro", 10, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        else:
            apply_formatting(para, "Minion Pro", 10, alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)

    # Save formatted document
    output_filename = "formatted_document.docx"
    doc.save(output_filename)
    return output_filename
