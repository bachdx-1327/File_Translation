from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

# Define paths
input_file = "TestData/DGP_1.docx"
image_dir = "img"

# Create image directory if not exists
if not os.path.exists(image_dir):
    os.makedirs(image_dir)

# Load the original document
doc = Document(input_file)

# Store the content with their types (text or image)
content = []

# Helper function to get image parts
def get_image_parts(doc):
    images = {}
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image_part = rel.target_part
            image_bytes = image_part.blob
            image_ext = image_part.content_type.split('/')[-1]
            image_filename = f"image{len(images) + 1}.{image_ext}"
            image_path = os.path.join(image_dir, image_filename)
            with open(image_path, 'wb') as img_file:
                img_file.write(image_bytes)
            images[rel.target_ref] = image_path
    return images

# Get all image parts
images = get_image_parts(doc)

def iter_block_items(parent):
    for child in parent.iter():
        yield child

for block in iter_block_items(doc.element.body):
    print(block.tag)
    if block.tag == qn('w:p'):  # Paragraph
        texts = []
        for run in block.iter(qn('w:t')):
            texts.append(run.text)
        text = ' '.join(texts).strip()
        if text:
            content.append(('text', text))
        print(text)
        print(block.tag)
        print(1)
    # elif block.tag == qn('w:tbl'):  # Table
    #     table_data = []
    #     for row in block.iter(qn('w:tr')):
    #         row_data = []
    #         for cell in row.iter(qn('w:tc')):
    #             cell_texts = []
    #             for para in cell.iter(qn('w:p')):
    #                 for run in para.iter(qn('w:t')):
    #                     cell_texts.append(run.text)
    #             cell_text = ' '.join(cell_texts).strip()
    #             row_data.append(cell_text)
    #         table_data.append(row_data)
    #     content.append(('table', table_data))
    elif block.tag == qn('w:drawing') or block.tag == qn('w:pict'):  # Image
        for blip in block.iter(qn('a:blip')):
            r_id = blip.get(qn('r:embed'))
            if r_id in images:
                pass
            
        content.append(('image', 'Image here'))
        print(block.tag)
        print(2)

# print(qn('w:drawing'))
print(content)