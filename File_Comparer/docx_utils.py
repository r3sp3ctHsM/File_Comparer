from docx import Document
import tkinter as tk

def extract_text_and_formatting_from_docx(docx_path, result_area):
    """Extract text and formatting from a DOCX file."""
    doc = Document(docx_path)
    text_content = []  # This will store the text content of the paragraphs
    font_styles = []   # This will store the font styles for each run of text

    for para in doc.paragraphs:
        #result_area.insert(tk.END, f"Processing paragraph: {para.text}\n")  # Log paragraph
        text_content.append(para.text)  # Append the paragraph text to text_content

        for run in para.runs:
            font_info = {
                'text': run.text,
                'bold': run.bold,
                'italic': run.italic,
                'underline': run.underline,
                'font_name': run.font.name,
                'font_size': run.font.size
            }
            #result_area.insert(tk.END, f"Processing run: {run.text}, Font info: {font_info}\n")  # Log each run
            font_styles.append(font_info)  # Append the font style information to font_styles

    # Return the text content as a single string and the list of font styles
    return '\n'.join(text_content), font_styles
