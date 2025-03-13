# import docx
# from docx.shared import Pt, RGBColor
# from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
# from transformers import BertTokenizer, BertForSequenceClassification
# import torch

# # Load BERT model for title identification
# tokenizer = BertTokenizer.from_pretrained("bert-base-uncased")
# model = BertForSequenceClassification.from_pretrained("bert-base-uncased", num_labels=2)

# # Function to check if paragraph contains title
# def identify_title_from_style(doc):
#     for para in doc.paragraphs:
#         if "title" in para.style.name.lower():
#             return para.text.strip()
#     return None  # Return None if no title found

# # Function to identify title using BERT
# def identify_title_with_bert(doc):
#     text_data = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
#     inputs = tokenizer(text_data, padding=True, truncation=True, return_tensors="pt")
#     with torch.no_grad():
#         outputs = model(**inputs).logits
#     predicted_label = torch.argmax(outputs, dim=1).tolist()
    
#     # Assume the first predicted "title" is correct
#     for i, label in enumerate(predicted_label):
#         if label == 1:  # Title detected by BERT
#             return text_data[i]
#     return None  # No title detected

# # Function to format text
# def apply_formatting(paragraph, font_name, font_size, bold=False, italic=False, alignment=None):
#     for run in paragraph.runs:
#         font = run.font
#         font.name = font_name
#         font.size = Pt(font_size)
#         font.bold = bold
#         font.italic = italic
#         font.color.rgb = RGBColor(0, 0, 0)

#     if alignment:
#         paragraph.alignment = alignment

# # Function to identify section types
# def identify_section(paragraph):
#     text = paragraph.text.strip()
#     style_name = paragraph.style.name.lower()

#     if text == detected_title:
#         return "title"
#     elif "author" in style_name:
#         return "author"
#     elif "address" in style_name:
#         return "address"
#     elif "email" in style_name:
#         return "email"
#     elif "heading 1" in style_name:
#         return "heading1"
#     elif "heading 2" in style_name:
#         return "heading2"
#     elif "heading 3" in style_name:
#         return "heading3"
#     elif "heading 4" in style_name:
#         return "heading4"
#     elif "referenceitem" in style_name:
#         return "referenceitem"
#     elif "reference" in style_name or "Reference" in style_name:
#         return "reference"
#     return "body"

# # Function to format the DOCX file
# def format_docx(file_path):
#     doc = docx.Document(file_path)

#     # Step 1: Try to identify title from style
#     global detected_title
#     detected_title = identify_title_from_style(doc)

#     # Step 2: If no title found, use BERT model
#     if not detected_title:
#         detected_title = identify_title_with_bert(doc)

#     # Step 3: Apply formatting rules
#     for para in doc.paragraphs:
#         section_type = identify_section(para)

#         if section_type == "title":
#             apply_formatting(para, "Times New Roman", 14, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
#         elif section_type == "author":
#             apply_formatting(para, "Times New Roman", 10, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
#         elif section_type == "address":
#             apply_formatting(para, "Times New Roman", 9, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
#         elif section_type == "email":
#             apply_formatting(para, "Courier", 9, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
#         elif section_type == "heading1":
#             apply_formatting(para, "Times New Roman", 12, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
#         elif section_type == "heading2":
#             apply_formatting(para, "Times New Roman", 10, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
#         elif section_type == "heading3":
#             apply_formatting(para, "Times New Roman", 10, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
#         elif section_type == "heading4":
#             apply_formatting(para, "Times New Roman", 10, italic=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
#         elif section_type == "referenceitem":
#             apply_formatting(para, "Times New Roman", 10, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
#         elif section_type == "reference":
#             apply_formatting(para, "Times New Roman", 10, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)
#         else:
#             apply_formatting(para, "Times New Roman", 10, alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY)

#     # Save the formatted document
#     output_filename = "formatted_document.docx"
#     doc.save(output_filename)
#     return output_filename

