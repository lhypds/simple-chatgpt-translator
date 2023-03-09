import os
from dotenv import load_dotenv
from pptx import Presentation
from chatgptapi_translator import ChatGPTAPI
from utils import LANGUAGES, TO_LANGUAGE_CODE


def get_paragraph_text(paragraph):
    """Get the text from a paragraph object"""
    text = ""
    for run in paragraph.runs:
        text += run.text
    return text


def translate_text(text):
    """Translate the text"""
    
    # Ignore some text
    if text.strip() == "ARiGATAYA": return text
    if text.strip() == "ARiGATAYA Entab": return text
    if len(text.strip()) == 1: return text  # Skip single characters
    
    print("GPT translating text...")
    result = translate_model.translate(text)
    return result


def replace_text(paragraph):
    """Replace the text of a paragraph object"""
    if paragraph.text.strip() == "": return
    if len(paragraph.runs) == 0: return
    
    paragraph_text = get_paragraph_text(paragraph)
    print("Paragraph text: " + paragraph_text)
    
    # Process the text
    translate =  translate_text(paragraph_text)
    
    if translate.strip() == paragraph_text.strip(): 
        print("Skipping translation")
        print("-")
        return
    print("-")
    
    # Replace the text
    for i, run in enumerate(paragraph.runs):
        if i == 0:
            run.text = translate
        else:
            run.text = ""


def process_pptx_text(fileBasename):
    """Process the text in a pptx file"""
    prs = Presentation(fileBasename + '.pptx')
    for slide in prs.slides:
        for shape in slide.shapes:  # loop through shapes on slide
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    replace_text(paragraph)
            if shape.has_table:
                for cell in shape.table.iter_cells():
                    for paragraph in cell.text_frame.paragraphs:
                        replace_text(paragraph)
    prs.save(fileBasename + '_translated.pptx')


# Create the translator
translate_model = ChatGPTAPI(
    key=os.getenv("CHATGPTAPI_KEY"),
    language=LANGUAGES.get("zh-hant"))


process_pptx_text(os.getenv("FILE_BASENAME"))
