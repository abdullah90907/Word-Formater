# import docx
# from docx.shared import RGBColor, Pt, Inches
# from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
# from docx.oxml import OxmlElement
# from docx.oxml.ns import qn
# import re
# from transformers import BertTokenizer, BertForSequenceClassification
# import torch

# # Load BERT model for title detection
# tokenizer = BertTokenizer.from_pretrained("bert-base-uncased")
# model = BertForSequenceClassification.from_pretrained("bert-base-uncased", num_labels=2)

# # Function to apply formatting
# def apply_formatting(paragraph, font_size=12, is_heading=False, bold=False, italic=False, no_indent=False):
#     for run in paragraph.runs:
#         font = run.font
#         font.name = 'Palatino Linotype'
#         font.size = Pt(font_size)
#         font.color.rgb = RGBColor(0, 0, 0)
#         if run.bold:
#             font.bold = True
#         else:
#             font.bold = bold
#         font.italic = italic

#     paragraph.paragraph_format.left_indent = None if no_indent else Inches(2.0)
#     paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

#     if is_heading:
#         for run in paragraph.runs:
#             run.font.bold = True

# # Function to detect decimal-based headings
# def is_decimal_heading(text):
#     return re.match(r'^\d+(\.\d+)+$', text.strip()) is not None

# # Function to identify title from style
# def identify_title_from_style(doc):
#     for para in doc.paragraphs:
#         if "title" in para.style.name.lower():
#             return para.text.strip()
#     return None

# # Function to identify title using BERT
# def identify_title_with_bert(doc):
#     text_data = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
#     inputs = tokenizer(text_data, padding=True, truncation=True, return_tensors="pt")
#     with torch.no_grad():
#         outputs = model(**inputs).logits
#     predicted_label = torch.argmax(outputs, dim=1).tolist()
    
#     for i, label in enumerate(predicted_label):
#         if label == 1:
#             return text_data[i]
#     return None

# # Function to format references section
# def format_references_section(doc):
#     is_bibliography = False
#     for para in doc.paragraphs:
#         text = para.text.strip().lower()
#         if "references" in text or "bibliography" in text:
#             is_bibliography = True
#             apply_formatting(para, font_size=10, bold=True)
#         if is_bibliography:
#             apply_formatting(para, font_size=10, bold=True)

# # Function to format reference items
# def format_reference_items(doc):
#     for para in doc.paragraphs:
#         if "referenceitem" in para.style.name.lower():
#             apply_formatting(para, font_size=10, italic=False, bold=False)

# # Function to format DOCX file
# def format_docx(file_path):
#     doc = docx.Document(file_path)

#     # Step 1: Identify title
#     global detected_title
#     detected_title = identify_title_from_style(doc)
#     if not detected_title:
#         detected_title = identify_title_with_bert(doc)

#     for para in doc.paragraphs:
#         text = para.text.strip()

#         if text == detected_title:
#             apply_formatting(para, font_size=18, bold=True, no_indent=True)
#         elif 'author' in para.style.name.lower():
#             apply_formatting(para, font_size=10, bold=True, no_indent=True)
#         elif 'subtitle' in para.style.name.lower() or 'article' in para.style.name.lower():
#             apply_formatting(para, font_size=10, italic=True, no_indent=True)
#         elif para.style.name.startswith('Heading 1'):
#             apply_formatting(para, font_size=12, is_heading=True)
#         elif para.style.name.startswith('Heading 2') or is_decimal_heading(text):
#             apply_formatting(para, font_size=12, is_heading=True)
#         else:
#             apply_formatting(para, font_size=10)

#     format_references_section(doc)
#     format_reference_items(doc)

#     for section in doc.sections:
#         section.different_first_page_header_footer = True
#         section.header.is_linked_to_previous = False

#         # Remove header content (no image upload)
#         first_page_header = section.first_page_header
#         for para in first_page_header.paragraphs:
#             first_page_header._element.remove(para._element)  # Remove any existing text in header

#         # Ensure only the footer is retained
#         standard_header = section.header
#         for para in standard_header.paragraphs:
#             standard_header._element.remove(para._element)

#     output_filename = "formatted_document_with_textbox.docx"
#     doc.save(output_filename)
#     return output_filename
