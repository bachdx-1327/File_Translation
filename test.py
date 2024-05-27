# import docx

# def getText(filename):
#     doc = docx.Document(filename)
#     fullText = []
#     for para in doc.paragraphs:
#         fullText.append(para.text)
#     return '\n'.join(fullText)

# content_docs = getText("TestData/BBH 18_5_2024.docx")
# print(content_docs)

from docx import Document
from docx.shared import Inches
import os
import docx2txt

# extract text
text = docx2txt.process("TestData/DGP.docx")

# extract text and write images in /tmp/img_dir
text = docx2txt.process("TestData/DGP.docx", "./TestData_txt/")

image_dir = "./TestData_txt"
# Create a new Document
doc = Document()

# Add extracted text to the new Document
for line in text.split('\n'):
    doc.add_paragraph(line)

# Add extracted images to the new Document
if os.path.exists(image_dir):
    for img_file in os.listdir(image_dir):
        img_path = os.path.join(image_dir, img_file)
        doc.add_picture(img_path)  # Adjust the width as needed

# Save the new Document
output_file = "reconstructed_file.docx"
doc.save(output_file)