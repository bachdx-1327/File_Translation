from docx import Document
from deeptrans import vi2ja

def translate_run_text(run_text):
    return vi2ja(run_text)

def translate_paragraph(paragraph):
    for run in paragraph.runs:
        if run.text.strip():
            # Print the original text of the run
            print(f"Original text: {run.text}")
            
            # Print the XML element of the run
            print(f"Run element: {run.element}")
            
            # Print the bold attribute
            print(f"Bold: {run.bold}")
            
            # Print the italic attribute
            print(f"Italic: {run.italic}")
            
            # Print the underline attribute
            print(f"Underline: {run.underline}")
            
            # Print the font size
            font_size = run.font.size
            if font_size:
                print(f"Font size: {font_size.pt} pt")
            else:
                print("Font size: None")
            
            # Print the font name
            font_name = run.font.name
            print(f"Font name: {font_name}")
            
            # Print the font color
            font_color = run.font.color
            if font_color and font_color.rgb:
                print(f"Font color: {font_color.rgb}")
            else:
                print("Font color: None")
            
            # Translate the text of the run if it has more than 2 characters
            # if len(run.text) > 2:
            #     translated_text = translate_run_text(run.text)
            #     run.text = translated_text
            #     print(f"Translated text: {translated_text}")

# Path to the input Word document
input_path = "TestData/DGP_3.docx"

# Load the Word document
doc = Document(input_path)

# Print the root element of the document
print(f"Document root element: {doc.element}")

# Iterate over each paragraph in the document and apply the translation function
for paragraph in doc.paragraphs:
    translate_paragraph(paragraph)