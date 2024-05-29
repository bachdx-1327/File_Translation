from docx import Document

def get_run_properties(run):
    properties = {}
    if run.bold:
        properties['bold'] = True
    if run.italic:
        properties['italic'] = True
    if run.underline:
        properties['underline'] = True
    if run.font.size:
        properties['font_size'] = run.font.size.pt
    if run.font.name:
        properties['font_name'] = run.font.name
    return properties

def get_paragraph_properties(paragraph):
    properties = {}
    if paragraph.style.name:
        properties['style'] = paragraph.style.name
    if paragraph.alignment:
        properties['alignment'] = str(paragraph.alignment)
    return properties

def docx_to_dict(docx_path):
    doc = Document(docx_path)
    doc_dict = {'paragraphs': []}
    
    for para in doc.paragraphs:
        para_dict = {'text': para.text, 'properties': get_paragraph_properties(para), 'runs': []}
        for run in para.runs:
            run_dict = {'text': run.text, 'properties': get_run_properties(run)}
            para_dict['runs'].append(run_dict)
        doc_dict['paragraphs'].append(para_dict)
    
    return doc_dict

# Đường dẫn đến tệp docx của bạn
docx_path = 'TestData/DGP_1.docx'
document_dict = docx_to_dict(docx_path)

import pprint
pprint.pprint(document_dict)
