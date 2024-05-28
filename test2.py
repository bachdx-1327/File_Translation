from docx import Document
from deeptrans import vi2ja

def translate_docx(input_path, out_path):
    # Loading the docx file 
    doc = Document(input_path)
    
    # Iterate through each paragraph in the document
    for para in doc.paragraphs:
        if para.text.strip(): 
            translate_text = vi2ja(para.text)
            para.text = translate_text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells: 
                for para in cell.paragraphs:
                    if para.text.strip():  # Only translate non-empty paragraphs
                        translated_text = vi2ja(para.text)
                        para.text = translated_text
    
    # Save the translated document
    doc.save(out_path)

input_path = "TestData/BBH 18_5_2024.docx"
out_path = "document_2.docx"

# Translate the docx file
translate_docx(input_path, out_path)