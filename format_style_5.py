from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.section import WD_SECTION_START
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from lxml import etree
from docx.oxml import OxmlElement
from docx.enum.text import WD_TAB_ALIGNMENT
import re

# Function to apply formatting
def apply_formatting(paragraph, font_name, font_size, bold=False, italic=False, underline=False, alignment=None):
    """Applies formatting to the paragraph, resetting pre-existing properties."""
    # Reset all relevant paragraph formatting properties to neutral values
    paragraph_format = paragraph.paragraph_format
    paragraph_format.alignment = None  # Reset alignment to default
    paragraph_format.left_indent = Cm(0)  # No left indent
    paragraph_format.right_indent = Cm(0)  # No right indent
    paragraph_format.first_line_indent = Cm(0)  # No first-line indent
    paragraph_format.space_before = Pt(0)  # No space before
    paragraph_format.space_after = Pt(0)  # No space after
    paragraph_format.line_spacing = None  # Reset line spacing to default
    paragraph_format.widow_control = False  # Disable widow/orphan control
    paragraph_format.keep_together = False  # Disable keep together
    paragraph_format.keep_with_next = False  # Disable keep with next

    # Explicitly set alignment if provided
    if alignment is not None:
        paragraph.alignment = alignment

    # Apply run formatting
    for run in paragraph.runs:
        font = run.font
        font.name = font_name
        font.size = Pt(font_size)
        font.bold = bold if bold else run.bold  # Preserve existing bold if not overridden
        font.italic = italic
        font.underline = underline
        font.color.rgb = RGBColor(0, 0, 0)  # Black text

# Function to set page margins and size
def set_page_layout(doc):
    """Sets page margins and size."""
    for section in doc.sections:
        # Set margins to 2.54 cm (1 inch)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

        # Set page size to 21.59 cm (width) and 27.95 cm (height)
        section.page_width = Cm(21.59)
        section.page_height = Cm(27.95)

# Function to add numbering on the right side
def add_numbering(doc):
    """Adds line numbers to the right side of the document."""
    for section in doc.sections:
        sectPr = section._sectPr
        lnNumType = OxmlElement('w:lnNumType')
        lnNumType.set(qn('w:countBy'), '1')
        lnNumType.set(qn('w:distance'), '360')
        lnNumType.set(qn('w:restart'), 'continuous')
        lnNumType.set(qn('w:numStart'), '1')
        sectPr.append(lnNumType)


def clear_header(header):
    """Completely clear all content from a header object."""
    header.is_linked_to_previous = False
    # Remove all child elements from the header
    for child in list(header._element):
        header._element.remove(child)
    # Ensure no residual paragraphs or tables remain
    header.paragraphs.clear()
    header.tables.clear()

def add_header_footer(doc):
    """Applies headers: images on page 1, even/odd headers with continuous numbering from page 2 onward."""
    # Ensure at least two pages for testing (unchanged from original)
    if len(doc.paragraphs) < 2:
        doc.add_page_break()

    # Process each section, ensuring continuous numbering and even/odd headers
    for i, section in enumerate(doc.sections):
        sectPr = section._sectPr

        # Step 1: Clear all headers
        for header_type in ['header', 'first_page_header', 'even_page_header']:
            if hasattr(section, header_type) and getattr(section, header_type) is not None:
                header = getattr(section, header_type)
                clear_header(header)

        # Step 2: Apply headers
        if i == 0:
            # Enable "Different First Page" for section 0
            titlePg = sectPr.find(qn('w:titlePg'))
            if titlePg is None:
                titlePg = OxmlElement('w:titlePg')
                sectPr.append(titlePg)

            # First page header (images only)
            header = section.first_page_header
            header.is_linked_to_previous = False
            table = header.add_table(rows=1, cols=2, width=Cm(16.50))
            table.autofit = False
            table.columns[0].width = Cm(10.795)
            table.columns[1].width = Cm(10.795)
            left_cell = table.cell(0, 0)
            left_paragraph = left_cell.paragraphs[0]
            left_run = left_paragraph.add_run()
            left_run.add_picture("left_image.jpg", width=Cm(3))
            left_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            right_cell = table.cell(0, 1)
            right_paragraph = right_cell.paragraphs[0]
            right_run = right_paragraph.add_run()
            right_run.add_picture("right_image.jpg", width=Cm(3))
            right_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

            # Footer (unchanged)
            footer = section.first_page_footer
            footer.is_linked_to_previous = False
            table = footer.add_table(rows=1, cols=2, width=Cm(16.50))
            table.autofit = False
            table.columns[0].width = Cm(5.0)
            table.columns[1].width = Cm(14.50)
            left_cell = table.cell(0, 0)
            left_paragraph = left_cell.paragraphs[0]
            left_run = left_paragraph.add_run()
            left_run.add_picture("footer_image.jpg", width=Cm(3))
            left_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            right_cell = table.cell(0, 1)
            right_paragraph = right_cell.paragraphs[0]
            right_paragraph.text = (
                "Copyright © 2025 The Author(s). Published by Tech Science Press. "
                "This work is licensed under a Creative Commons Attribution 4.0 International License."
            )
            right_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            table.allow_autofit = False

        # Enable Different Odd & Even Pages for all sections
        evenOdd = sectPr.find(qn('w:evenAndOddHeaders'))
        if evenOdd is None:
            evenOdd = OxmlElement('w:evenAndOddHeaders')
            sectPr.append(evenOdd)

        # Even page header (Page number left, Journal right)
        even_header = section.even_page_header
        even_header.is_linked_to_previous = False
        table = even_header.add_table(rows=1, cols=2, width=Cm(16.50))
        table.autofit = False
        table.columns[0].width = Cm(10.795)
        table.columns[1].width = Cm(10.795)
        left_cell = table.cell(0, 0)
        left_paragraph = left_cell.paragraphs[0]
        run = left_paragraph.add_run()
        fld_char_begin = OxmlElement('w:fldChar')
        fld_char_begin.set(qn('w:fldCharType'), 'begin')
        run._r.append(fld_char_begin)
        instr_text = OxmlElement('w:instrText')
        instr_text.text = ' PAGE '
        instr_text.set(qn('xml:space'), 'preserve')
        run._r.append(instr_text)
        fld_char_end = OxmlElement('w:fldChar')
        fld_char_end.set(qn('w:fldCharType'), 'end')
        run._r.append(fld_char_end)
        left_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        right_cell = table.cell(0, 1)
        right_paragraph = right_cell.paragraphs[0]
        right_run = right_paragraph.add_run("Comput Mater Contin. 2025;volume(issue)")
        right_run.font.size = Cm(0.35)  # Match your original size
        right_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

        # Odd page header (Journal left, Page number right)
        odd_header = section.header
        odd_header.is_linked_to_previous = False
        table = odd_header.add_table(rows=1, cols=2, width=Cm(16.50))
        table.autofit = False
        table.columns[0].width = Cm(10.795)
        table.columns[1].width = Cm(10.795)
        left_cell = table.cell(0, 0)
        left_paragraph = left_cell.paragraphs[0]
        left_run = left_paragraph.add_run("Comput Mater Contin. 2025;volume(issue)")
        left_run.font.size = Cm(0.35)  # Match your original size
        left_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        right_cell = table.cell(0, 1)
        right_paragraph = right_cell.paragraphs[0]
        run = right_paragraph.add_run()
        fld_char_begin = OxmlElement('w:fldChar')
        fld_char_begin.set(qn('w:fldCharType'), 'begin')
        run._r.append(fld_char_begin)
        instr_text = OxmlElement('w:instrText')
        instr_text.text = ' PAGE '
        instr_text.set(qn('xml:space'), 'preserve')
        run._r.append(instr_text)
        fld_char_end = OxmlElement('w:fldChar')
        fld_char_end.set(qn('w:fldCharType'), 'end')
        run._r.append(fld_char_end)
        right_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

        # Step 3: Ensure continuous page numbering
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is not None:
            sectPr.remove(pgNumType)
        pgNumType = OxmlElement('w:pgNumType')
        if i == 0:
            pgNumType.set(qn('w:start'), '1')  # Start at 1 for the first section only
        # For subsequent sections, omit w:start to continue numbering
        sectPr.append(pgNumType)

# Function to capitalize and bold "Abstract" and "Keyword"
def capitalize_and_bold_abstract_keyword(doc):
    # Define the target words to capitalize and bold
    target_words = ["abstract", "keyword", "keywords"]
    # Iterate through each paragraph
    for paragraph in doc.paragraphs:
        # Check if any target word exists in the paragraph (case-insensitive)
        if any(word in paragraph.text.lower() for word in target_words):
            print(f"Processing paragraph: {paragraph.text}")

            # Clear the paragraph and rebuild it with formatted runs
            new_runs = []
            for run in paragraph.runs:
                text = run.text
                # Capitalize and bold the target words
                for word in target_words:
                    # Use regex to replace all occurrences (case-insensitive)
                    text = re.sub(
                        re.compile(re.escape(word), re.IGNORECASE),
                        word.upper(),  # Capitalize the word
                        text
                    )
                # Add the modified text to new_runs
                new_runs.append((text, run.bold, run.font.size, run.font.name))

            # Clear the paragraph and recreate it with formatted runs
            paragraph.clear()
            for text, orig_bold, orig_size, orig_font in new_runs:
                # Split text into parts (words and separators) using regex
                parts = re.split(r'(\W+)', text)  # Split by non-word characters
                for part in parts:
                    run = paragraph.add_run(part)
                    run.font.name = orig_font or "Minion Pro"
                    run.font.size = orig_size or Pt(10)
                    # Bold only the target words
                    if part.upper() in [word.upper() for word in target_words]:
                        run.bold = True
                    else:
                        run.bold = orig_bold if orig_bold is not None else False


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

def indent_first_line(paragraph, indent_size=Cm(0.5)):
    """Adds first-line indentation to the paragraph."""
    paragraph.paragraph_format.first_line_indent = indent_size

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

def split_and_center_align_images(doc):
    """Splits images into their own paragraphs and center-aligns/resizes them."""
    # Define the namespaces used in the document
    namespaces = {
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    }

    # Iterate through all paragraphs in the document
    for paragraph in list(doc.paragraphs):  # Use list() to avoid skipping paragraphs
        # Check if the paragraph contains both text and images
        has_text = any(run.text.strip() for run in paragraph.runs)
        has_image = any(run.element.xpath('.//w:drawing') for run in paragraph.runs)

        if has_text and has_image:
            # Create a new paragraph for the image
            new_paragraph = doc.add_paragraph()
            new_paragraph.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            # Move the image to the new paragraph
            for run in paragraph.runs:
                if run.element.xpath('.//w:drawing'):
                    new_run = new_paragraph.add_run()
                    new_run._r.append(run.element)

            # Remove the image from the original paragraph
            for run in list(paragraph.runs):  # Use list() to avoid skipping runs
                if run.element.xpath('.//w:drawing'):
                    paragraph._element.remove(run._element)

            # Resize the image in the new paragraph
            for run in new_paragraph.runs:
                if run.element.xpath('.//w:drawing'):
                    # Convert the drawing element to an lxml element
                    drawing_xml = run.element.xml
                    drawing = etree.fromstring(drawing_xml)

                    # Resize the image (reduce size to 50% of original)
                    for extent in drawing.xpath('.//wp:extent', namespaces=namespaces):
                        cx = int(int(extent.get('cx')) * 0.85)  # Reduce width by 85%
                        cy = int(int(extent.get('cy')) * 0.85)  # Reduce height by 85%
                        extent.set('cx', str(int(cx)))  # Set new width
                        extent.set('cy', str(int(cy)))  # Set new height

                    # Convert the modified XML back to a string
                    updated_xml = etree.tostring(drawing, encoding='unicode')

                    # Parse the updated XML back into an element
                    updated_element = etree.fromstring(updated_xml)

                    # Update the run's XML with the modified drawing
                    run._r.clear()
                    run._r.append(updated_element)

            # Insert the new paragraph (with the image) immediately after the original paragraph
            paragraph._p.addnext(new_paragraph._element)

def indent_first_line(paragraph, indent_size=Cm(0.5)):
    """Adds first-line indentation to a paragraph."""
    paragraph.paragraph_format.first_line_indent = indent_size

def adjust_table_widths(doc):
    """Adjusts all table widths to the maximum usable width based on set_page_layout."""
    # Use fixed usable width based on set_page_layout: 21.59 cm - 2.54 cm - 2.54 cm = 16.51 cm
    usable_width_cm = Cm(16.51)
    usable_width_twips = int(usable_width_cm.cm * 567)  # Convert to twips (1 cm = 567 twips)

    # Iterate through all tables in the document
    for table in doc.tables:
        # Set table width to maximum usable width
        table.width = usable_width_cm
        table.autofit = False  # Disable autofit to enforce our width

        # Adjust column widths proportionally to fit new table width
        total_current_width = sum(col.width for col in table.columns if col.width is not None)
        if total_current_width > 0:  # Avoid division by zero
            width_ratio = usable_width_cm / total_current_width
            for col in table.columns:
                if col.width is not None:
                    col.width = int(col.width * width_ratio)

        # Set table alignment to left (consistent with your document style)
        table.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        # Ensure table width is enforced via XML
        tbl_pr = table._tbl.tblPr
        tbl_w = tbl_pr.find(qn('w:tblW'))
        if tbl_w is None:
            tbl_w = OxmlElement('w:tblW')
            tbl_pr.append(tbl_w)
        tbl_w.set(qn('w:w'), str(usable_width_twips))  # Set width in twips
        tbl_w.set(qn('w:type'), 'dxa')  # Use absolute units


def normalize_inline_spacing(doc):
    """Normalizes excessive spaces in all text runs, including headings, preserving images and structure."""
    for para in doc.paragraphs:
        if para._element.findall(qn('w:tbl')) or para._element.xpath('.//w:drawing|.//w:pict'):
            continue
        full_text = ''.join(run.text for run in para.runs if run.text)
        if full_text:
            normalized_text = re.sub(r'\s+', ' ', full_text).strip()
            if normalized_text != full_text:
                para.clear()
                para.add_run(normalized_text)

def format_docx(file_path):
    """Formats the document and returns the path to the saved file."""
    doc = Document(file_path)

    # Ensure document has at least one paragraph to work with
    if not doc.paragraphs:
        doc.add_paragraph("")

    # Step 1: Handle DOI and Paper Type insertions
    paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    print(f"Paragraphs for DOI check: {paragraphs}")  # Debug: Check contents
    has_doi = any("doi" in p.lower() for p in paragraphs if p and len(p) > 3)  # Avoid false positives
    has_paper_type = any("paper type" in p.lower() or "articletype" in p.lower() for p in paragraphs if p)

    # Insert DOI if absent
    if not has_doi:
        print("Inserting DOI because it’s absent")  # Debug: Confirm execution
        doi_para = doc.add_paragraph("DOI: _____________")
        # Safely insert at the start by reordering body elements
        body_elements = list(doc._body._element)
        body_elements.insert(0, doi_para._p)
        doc._body._element.clear()
        for elem in body_elements:
            doc._body._element.append(elem)
        apply_formatting(doi_para, "Minion Pro", 7, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        blank_para = doc.add_paragraph("")
        doi_para._p.addnext(blank_para._p)  # Blank line after DOI
    else:
        print("DOI already present, skipping insertion")  # Debug: Confirm detection
        for i, para in enumerate(doc.paragraphs):
            if "doi" in para.text.lower():
                apply_formatting(para, "Minion Pro", 7, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
                if i + 1 < len(doc.paragraphs):
                    blank_para = doc.add_paragraph("")
                    para._p.addnext(blank_para._p)
                break

    # Insert Paper Type if absent
    doi_index = next((i for i, para in enumerate(doc.paragraphs) if "doi" in para.text.lower()), -1)
    if not has_paper_type:
        print("Inserting Paper Type because it’s absent")  # Debug: Confirm execution
        paper_type_para = doc.add_paragraph("Paper Type (____________)")
        if doi_index >= 0:
            insert_after = doi_index + 1 if doi_index + 1 < len(doc.paragraphs) else len(doc.paragraphs) - 1
            if insert_after < len(doc.paragraphs):
                doc.paragraphs[insert_after]._p.addnext(paper_type_para._p)
            else:
                doc._body._element.append(paper_type_para._p)
        else:
            doc._body._element.insert(1, paper_type_para._p)  # After initial paragraph if no DOI
        apply_formatting(paper_type_para, "Minion Pro", 9, bold=True, underline=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        blank_para = doc.add_paragraph("")
        paper_type_para._p.addnext(blank_para._p)
    else:
        print("Paper Type already present, skipping insertion")  # Debug: Confirm detection
        for i, para in enumerate(doc.paragraphs):
            if "paper type" in para.text.lower() or "articletype" in para.text.lower():
                apply_formatting(para, "Minion Pro", 9, bold=True, underline=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
                if i + 1 < len(doc.paragraphs):
                    blank_para = doc.add_paragraph("")
                    para._p.addnext(blank_para._p)
                break

    # Step 2: Identify Title and Authors
    all_paragraphs = [para for para in doc.paragraphs]  # Include all paragraphs, including blanks
    title_index = 0
    # Count inserted paragraphs to find Title
    inserted_count = 0
    if not has_doi:
        inserted_count += 2  # DOI + blank
    if not has_paper_type:
        inserted_count += 2  # Paper Type + blank
    # Title is the first non-empty paragraph after inserted sections
    for i, para in enumerate(all_paragraphs):
        if i >= inserted_count and para.text.strip() and not para.text.strip().startswith(("DOI:", "Paper Type")):
            title_index = i
            break
    else:
        title_index = inserted_count if inserted_count < len(all_paragraphs) else 0

    # Find the next non-empty paragraph for Authors
    authors_index = title_index
    for i in range(title_index + 1, len(all_paragraphs)):
        if all_paragraphs[i].text.strip() and not all_paragraphs[i].text.strip().startswith(("DOI:", "Paper Type")):
            authors_index = i
            break

    title = all_paragraphs[title_index].text.strip() if title_index < len(all_paragraphs) else ""
    authors = all_paragraphs[authors_index].text.strip() if authors_index < len(all_paragraphs) else ""
    abstract_index = next((i for i, para in enumerate(all_paragraphs) if para.text.strip().lower().startswith("abstract")), None)

    print(f"Title Index: {title_index}, Title: {title}")  # Debug: Verify title selection
    print(f"Authors Index: {authors_index}, Authors: {authors}")  # Debug: Verify authors selection

    # Ensure single-column layout
    set_single_column(doc)

    # Flags
    in_references_section = False
    before_abstract = True
    title_formatted = False
    authors_formatted = False

    # Step 3: Format paragraphs
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()

        if abstract_index is not None and i >= abstract_index:
            before_abstract = False

        # Format DOI
        if "doi" in text.lower() and i <= 1:
            apply_formatting(para, "Minion Pro", 7, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
            continue
        # Format Paper Type
        elif ("paper type" in text.lower() or "articletype" in text.lower()) and i <= 3:
            apply_formatting(para, "Minion Pro", 9, bold=True, underline=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
            continue
        # Format Title
        elif i == title_index and text and not title_formatted:
            apply_formatting(para, "Minion Pro", 14, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
            title_formatted = True
            continue
        # Format Authors
        elif i == authors_index and text and not authors_formatted:
            apply_formatting(para, "Minion Pro", 12, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
            authors_formatted = True
            continue

        # Skip further formatting if Title or Authors were just formatted
        if title_formatted and i <= authors_index:
            continue

        section_type = identify_section(para)
        if section_type == "affiliation":
            apply_formatting(para, "Minion Pro", 9, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "abstract" or section_type == "keyword":
            apply_formatting(para, "Minion Pro", 10, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "BackMatter":
            apply_formatting(para, "Minion Pro", 10, bold=False, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "heading_1":
            apply_formatting(para, "Minion Pro", 11, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "heading_2":
            apply_formatting(para, "Minion Pro", 11, bold=True, italic=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        elif section_type == "heading_3" or section_type == "heading_4":
            apply_formatting(para, "Minion Pro", 11, italic=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
        else:
            apply_formatting(para, "Minion Pro", 10, alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)

        if text.lower().startswith("References"):
            in_references_section = True
            continue
        if not in_references_section and i > 0:
            previous_para = doc.paragraphs[i - 1]
            previous_section_type = identify_section(previous_para)
            current_section_type = identify_section(para)
            if previous_section_type in ["heading_1", "heading_2", "heading_3", "heading_4"] and \
               current_section_type not in ["heading_1", "heading_2", "heading_3", "heading_4"]:
                if re.match(r'^[A-Za-z]', text) and not re.match(r'^\d', text) and not re.match(r'^[A-Z]\s', text):
                    if not all(run.bold for run in para.runs if run.text.strip()):
                        indent_first_line(para, Cm(0.5))

    normalize_inline_spacing(doc)
    split_and_center_align_images(doc)
    set_page_layout(doc)
    add_numbering(doc)
    add_header_footer(doc)
    capitalize_and_bold_abstract_keyword(doc)
    adjust_table_widths(doc)

    output_filename = "formated.docx"
    doc.save(output_filename)
    return output_filename