import os
import win32com.client

def convert_rtf_to_docx(rtf_path, temp_docx_path):
    """Convert an RTF file to a DOCX file using Microsoft Word."""
    # Normalize the path (removes double slashes, etc.)
    rtf_path = os.path.normpath(rtf_path)
    temp_docx_path = os.path.normpath(temp_docx_path)

    # Ensure the RTF file exists
    if not os.path.exists(rtf_path):
        raise FileNotFoundError(f"The RTF file was not found: {rtf_path}")

    # Create a new instance of Word Application
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False  # Make Word invisible

    try:
        print(f"Opening file at: {rtf_path}")  # Print the path for debugging
        doc = word.Documents.Open(rtf_path)  # Open the RTF file
        doc.SaveAs(temp_docx_path, FileFormat=12)  # Save as DOCX
        doc.Close()
    except Exception as e:
        print(f"Error opening the RTF file: {e}")
    finally:
        word.Quit()  # Make sure to quit Word to avoid leaving processes running
