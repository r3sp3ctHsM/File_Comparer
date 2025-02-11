import os
import tempfile
from rtf_utils import convert_rtf_to_docx
from docx_utils import extract_text_and_formatting_from_docx
from image_utils import compare_images, overlay_differences, merge_images, convert_docx_to_image
from text_utils import compare_texts, compare_fonts
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from PIL import Image

def compare_rtf_and_docx(rtf_path, docx_path, output_folder, result_area):
    """Convert RTF to DOCX and compare the documents."""
    # Normalize paths to avoid double slashes and ensure correct formatting
    rtf_path = os.path.normpath(rtf_path)
    docx_path = os.path.normpath(docx_path)
    output_folder = os.path.normpath(output_folder)

    # Create a temporary file for the converted DOCX
    temp_docx_path = os.path.join(tempfile.gettempdir(), 'converted.docx')

    # Convert RTF to DOCX
    convert_rtf_to_docx(rtf_path, temp_docx_path)

    # Extract text and formatting from the original DOCX file
    original_docx_text, original_docx_fonts = extract_text_and_formatting_from_docx(docx_path, result_area)

    # Extract text and formatting from the converted DOCX file
    converted_docx_text, converted_docx_fonts = extract_text_and_formatting_from_docx(temp_docx_path, result_area)

    # Compare text content
    text_differences = compare_texts(original_docx_text, converted_docx_text)

    # Compare fonts
    font_differences = compare_fonts(original_docx_fonts, converted_docx_fonts)

    # Perform image comparisons if needed
    # Convert documents to images and compare them
    rtf_image_path = os.path.join(output_folder, 'rtf_image.png')
    docx_image_path = os.path.join(output_folder, 'docx_image.png')
    convert_docx_to_image(temp_docx_path, rtf_image_path)
    convert_docx_to_image(docx_path, docx_image_path)

    rtf_image = Image.open(rtf_image_path)
    docx_image = Image.open(docx_image_path)

    diff_image = compare_images(rtf_image, docx_image)
    overlay_image = overlay_differences(docx_image, diff_image)

    output_image_path = os.path.join(output_folder, 'diff_image.png')
    overlay_image.save(output_image_path)

    # Output results
    result = ""
    if text_differences:
        result += f"Differences in content between {os.path.basename(rtf_path)} and {os.path.basename(docx_path)}:\n"
        result += '\n'.join(text_differences) + '\n'
    else:
        result += f"Content is identical between {os.path.basename(rtf_path)} and {os.path.basename(docx_path)}.\n"

    if font_differences:
        result += f"Differences in fonts between {os.path.basename(rtf_path)} and {os.path.basename(docx_path)}:\n"
        result += '\n'.join(font_differences) + '\n'
    else:
        result += f"Fonts are identical between {os.path.basename(rtf_path)} and {os.path.basename(docx_path)}.\n"

    if os.path.exists(output_image_path):
        result += f"Visual differences saved as: {output_image_path}\n"
    else:
        result += f"No visual differences found between {os.path.basename(rtf_path)} and {os.path.basename(docx_path)}.\n"

    result_area.insert(tk.END, result)

def start_comparison(old_docs_entry, new_docs_entry, output_entry, result_area):
    """Start the comparison process based on user input."""
    old_docs_dir = old_docs_entry.get()
    new_docs_dir = new_docs_entry.get()
    output_dir = output_entry.get()

    if not(old_docs_dir and new_docs_dir and output_dir):
        messagebox.showerror("Input Error", "Please provide all directories.")
        return
    
    for rtf_file in os.listdir(old_docs_dir):
        if rtf_file.endswith('.rtf'):
            rtf_path = os.path.join(old_docs_dir, rtf_file)
            # Check for corresponding DOCX file
            docx_file_name = rtf_file.replace('.rtf', '.docx')
            docx_path = os.path.join(new_docs_dir, docx_file_name)

            if os.path.exists(docx_path):
                compare_rtf_and_docx(rtf_path, docx_path, output_dir, result_area)
            else:
                result_area.insert(tk.END, f"{docx_file_name} does not exist for comparison.\n")

def create_gui():
    """Create the main GUI for the application."""
    root = tk.Tk()
    root.title("Document Comparator")

    # Input for Old Documents Directory
    tk.Label(root, text="Old Documents Directory (RTF):").grid(row=0, column=0, padx=10, pady=10)
    old_docs_entry = tk.Entry(root, width=50)
    old_docs_entry.grid(row=0, column=1, padx=10, pady=10)
    tk.Button(root, text="Browse...", command=lambda: old_docs_entry.insert(0, filedialog.askdirectory())).grid(row=0, column=2)

    # Input for New Documents Directory
    tk.Label(root, text="New Documents Directory (DOCX):").grid(row=1, column=0, padx=10, pady=10)
    new_docs_entry = tk.Entry(root, width=50)
    new_docs_entry.grid(row=1, column=1, padx=10, pady=10)
    tk.Button(root, text="Browse...", command=lambda: new_docs_entry.insert(0, filedialog.askdirectory())).grid(row=1, column=2)

    # Input for Output Directory
    tk.Label(root, text="Output Directory:").grid(row=2, column=0, padx=10, pady=10)
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=2, column=1, padx=10, pady=10)
    tk.Button(root, text="Browse...", command=lambda: output_entry.insert(0, filedialog.askdirectory())).grid(row=2, column=2)

    # Result area
    result_area = scrolledtext.ScrolledText(root, width=80, height=20)
    result_area.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

    # Button to start comparison
    tk.Button(root, text="Start Comparison", command=lambda: start_comparison(old_docs_entry, new_docs_entry, output_entry, result_area)).grid(row=3, column=1, pady=10)

    # Start the GUI loop
    root.mainloop()

if __name__ == "__main__":
    create_gui()
