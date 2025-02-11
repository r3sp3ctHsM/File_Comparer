import sys
import os
from PIL import Image, ImageChops
import win32com.client
from pdf2image import convert_from_path

def compare_images(img1, img2):
    """Compare two images and return a difference images."""
    diff = ImageChops.difference(img1, img2)
    return diff

def overlay_differences(original_image, difference_image):
    """Overlay the differences on the original image."""
    # Convert the difference image to grayscale
    difference_image = difference_image.convert("L")

    # Create a threshold to identify significant differences
    threshold = 30
    difference_image = difference_image.point(lambda p: p > threshold and 255)

    # Create a mask from the difference image
    mask = difference_image.convert("1") # Convert to binary mask

    # Overlay the mask onto the original image with a color (e.g. red)
    overlay = Image.new("RGB", original_image.size, (255, 0, 0, 128)) # Red overlay with transparency
    original_with_overlay = Image.composite(overlay, original_image, mask)

    return original_with_overlay

def merge_images(images, output_path):
    """Merge a list of images vertically and save as a single image."""
    widths, heights = zip(*(i.size for i in images))
    total_height = sum(heights)
    max_width = max(widths)

    merged_image = Image.new('RGB', (max_width, total_height))
    y_offset = 0
    for img in images:
        merged_image.paste(img, (0, y_offset))
        y_offset += img.height

    merged_image.save(output_path)

def convert_docx_to_image(docx_path, output_path):
    """Convert DOCX file to an image using Microsoft Word."""
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(docx_path)
    
    # Save as PDF first
    pdf_path = output_path.replace('.png', '.pdf')
    doc.SaveAs(pdf_path, FileFormat=17)  # FileFormat=17 is for PDF
    doc.Close()

    # Quit Word application properly by closing the application
    word.Quit()  # Quit the Word application, not the document

    # Determine poppler path dynamically based on the script's location
    if getattr(sys, 'frozen', False):  # Check if running from a packaged EXE
        # When bundled by PyInstaller, Poppler will be in the temporary folder, inside 'poppler/bin'
        poppler_path = os.path.join(sys._MEIPASS, 'poppler', 'bin')
    else:
        # If running from source, assume 'poppler' folder is one level up from the script
        script_dir = os.path.dirname(os.path.abspath(__file__))  # Get the script's directory
        poppler_path = os.path.join(script_dir, '..', 'poppler', 'bin')  # Build relative path

    # Convert PDF to image using the Poppler path
    images = convert_from_path(pdf_path, poppler_path=poppler_path)
    if images:
        images[0].save(output_path, 'PNG')