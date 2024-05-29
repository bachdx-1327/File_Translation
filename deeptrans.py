from deep_translator import GoogleTranslator

def vi2ja(text):
    translated = GoogleTranslator(source='vi', target='ja').translate(text=text)
    return translated

def ja2vi(text):
    translated = GoogleTranslator(source='ja', target='vi').translate(text=text)
    return translated
