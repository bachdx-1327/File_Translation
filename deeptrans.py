from deep_translator import GoogleTranslator

def vi2ja(text):
    translated = GoogleTranslator(source='vi', target='ja').translate(text=text)
    return translated
