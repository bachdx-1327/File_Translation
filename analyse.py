from docx import Document
import os

# Define paths
input_file = "TestData/DGP.docx"
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

# Iterate through document elements in order
def iter_block_items(parent):
    if parent.tag.endswith(('body', 'tc')):
        for child in parent.iterchildren():
            if child.tag.endswith('p'):
                yield child
            elif child.tag.endswith('tbl'):
                yield child
            elif child.tag.endswith(('drawing', 'pict')):
                yield child
            elif child.tag.endswith('sdt'):
                for subchild in iter_block_items(child):
                    yield subchild

for element in doc.element.getiterator():
    # print(element)
    if element.tag.endswith('p'):  # Paragraph
        texts = []
        for run in element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
            texts.append(run.text)
        if texts:
            content.append(('text', ''.join(texts)))
    elif element.tag.endswith('tbl'):  # Table
        table_data = []
        for row in element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr"):
            row_data = []
            for cell in row.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc"):
                cell_texts = []
                for paragraph in cell.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
                    for run in paragraph.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
                        cell_texts.append(run.text)
                row_data.append(''.join(cell_texts))
            table_data.append(row_data)
        content.append(('table', table_data))
    elif element.tag.endswith(('pic')):  # Image
        for blip in element.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
            embed = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if embed in images:
                content.append(('image', images[embed]))

# Save the content list for reconstruction
with open(os.path.join(image_dir, 'content.txt'), 'w') as f:
    for item in content:
        if item[0] == 'table':
            f.write(f"{item[0]}:{item[1]}\n")
        else:
            f.write(f"{item[0]}:{item[1]}\n")
