# user_guide_utils.py


from docx import Document


def word_to_text(docx_path):
    """
    Extracts and returns the full text content of a Word document as a string.
    """
    doc = Document(docx_path)
    text = []
    for para in doc.paragraphs:
        text.append(para.text)
    return "\n".join(text)


# Example usage (to be called from app.py):
# from user_guide_utils import word_to_image
# image_paths = word_to_image("UserGuide.docx", "output_images")
# (In Streamlit: st.image(image_paths) to display)
