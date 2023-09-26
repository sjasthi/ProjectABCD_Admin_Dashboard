# import the translator https://pypi.org/project/deep-translator/
from deep_translator import GoogleTranslator

# Build the translator
dest_language = 'te' #Telugu   use translator.get_supported_languages()
translator = GoogleTranslator(source= 'auto',target=dest_language)

# Perform the translation
source_text = "Hello! How are you doing?"

# translate
dest_text = translator.translate(source_text)

# print the source and dest text
print(source_text)
print(dest_text)
