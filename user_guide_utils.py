# user_guide_utils.py


import mammoth


def word_to_html(docx_path):
    """
    Converts a Word document to HTML (preserving formatting and images as base64).
    Returns the HTML string.
    """
    with open(docx_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value  # The generated HTML
    return html


# Example usage (to be called from app.py):
# from user_guide_utils import word_to_image
# image_paths = word_to_image("UserGuide.docx", "output_images")
# (In Streamlit: st.image(image_paths) to display)
