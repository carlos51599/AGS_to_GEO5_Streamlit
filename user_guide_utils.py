# user_guide_utils.py


import os
import subprocess
from pdf2image import convert_from_path


def word_to_images(docx_path, output_dir, dpi=200):
    """
    Converts a Word document to PDF using LibreOffice, then to PNG images (one per page).
    Returns a list of generated image file paths.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    pdf_path = os.path.join(output_dir, "UserGuide.pdf")
    # Convert docx to pdf using LibreOffice
    subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            output_dir,
            docx_path,
        ],
        check=True,
    )
    # LibreOffice names the PDF as <basename>.pdf
    base_pdf = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    pdf_path = os.path.join(output_dir, base_pdf)
    # Convert PDF to images
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
