from docx import Document
import json
from pprint import pprint

class DocxParser:
    def __init__(self, docx_path):
        self.docx_path = docx_path
        self.doc = Document(docx_path)
        self.doc_dict = {'paragraphs': []}
    
    def get_run_properties(self, run):
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
        if run.font.color and run.font.color.rgb:
            properties['font_color'] = str(run.font.color.rgb)
        return properties

    def get_paragraph_properties(self, paragraph):
        properties = {}
        if paragraph.style and paragraph.style.name:
            properties['style'] = paragraph.style.name
        if paragraph.alignment:
            properties['alignment'] = str(paragraph.alignment)
        if paragraph.paragraph_format:
            if paragraph.paragraph_format.left_indent:
                properties['left_indent'] = paragraph.paragraph_format.left_indent.pt
            if paragraph.paragraph_format.right_indent:
                properties['right_indent'] = paragraph.paragraph_format.right_indent.pt
            if paragraph.paragraph_format.space_before:
                properties['space_before'] = paragraph.paragraph_format.space_before.pt
            if paragraph.paragraph_format.space_after:
                properties['space_after'] = paragraph.paragraph_format.space_after.pt
            if paragraph.paragraph_format.line_spacing:
                properties['line_spacing'] = paragraph.paragraph_format.line_spacing
        return properties

    def parse_document(self):
        for para in self.doc.paragraphs:
            para_dict = {
                'text': para.text,
                'properties': self.get_paragraph_properties(para),
                'runs': []
            }
            for run in para.runs:
                run_dict = {
                    'text': run.text,
                    'properties': self.get_run_properties(run)
                }
                para_dict['runs'].append(run_dict)
            self.doc_dict['paragraphs'].append(para_dict)
    
    def save_as_json(self, json_path):
        with open(json_path, 'w', encoding='utf-8') as json_file:
            json.dump(self.doc_dict, json_file, ensure_ascii=False, indent=4)

    def get_document_dict(self):
        return self.doc_dict

# Đường dẫn đến tệp docx của bạn
docx_path = 'TestData/DGP_1.docx'
json_path = 'TestData/DGP_1.json'

# Create an instance of the parser and parse the document
parser = DocxParser(docx_path)
parser.parse_document()

# Save the parsed document as a JSON file
parser.save_as_json(json_path)
