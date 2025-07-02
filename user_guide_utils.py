# user_guide_utils.py

import os
from docx2pdf import convert as docx2pdf_convert
from pdf2image import convert_from_path


def word_to_image(docx_path, output_dir, dpi=300):
    """
    Converts a Word document to PDF, then to high-res PNG image(s).
    Returns a list of generated image file paths.
    """
    # 1. Convert .docx to .pdf
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    pdf_path = os.path.join(output_dir, "UserGuide.pdf")
    docx2pdf_convert(docx_path, pdf_path)

    # 2. Convert .pdf to .png (one image per page)
    images = convert_from_path(pdf_path, dpi=dpi)
    image_paths = []
    for i, img in enumerate(images):
        img_path = os.path.join(output_dir, f"UserGuide_page_{i+1}.png")
        img.save(img_path, "PNG")
        image_paths.append(img_path)

    return image_paths


# Example usage (to be called from app.py):
# from user_guide_utils import word_to_image
# image_paths = word_to_image("UserGuide.docx", "output_images")
# (In Streamlit: st.image(image_paths) to display)
