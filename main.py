import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
from PIL import Image
import re

# Maximum image size (in bytes) e.g., 5 MB
MAX_IMAGE_SIZE = 5 * 1024 * 1024

def count_existing_images(doc):
    """
    Count images already appended by scanning paragraphs for labels of the form "Image{n}".
    """
    max_index = 0
    pattern = re.compile(r"^Image(\d+)$")
    for para in doc.paragraphs:
        match = pattern.match(para.text.strip())
        if match:
            try:
                num = int(match.group(1))
                if num > max_index:
                    max_index = num
            except ValueError:
                continue
    return max_index

def main():
    st.title("Word Document Image Appender")

    st.header("1. Document Selection")
    col1, col2 = st.columns(2)
    
    # Create New Document button
    if col1.button("Create New Document"):
        new_doc = Document()
        st.session_state['doc'] = new_doc
        st.session_state['doc_name'] = "New Document.docx"
        st.success("New document created!")
    
    # Select Existing Document using file uploader (only .docx allowed)
    existing_file = col2.file_uploader("Select Existing Document", type=["docx"], key="existing_doc")
    if existing_file is not None:
        try:
            existing_doc = Document(existing_file)
            st.session_state['doc'] = existing_doc
            st.session_state['doc_name'] = existing_file.name
            st.success(f"Loaded document: {existing_file.name}")
        except Exception as e:
            st.error(f"Error loading document: {e}")

    if 'doc_name' in st.session_state:
        st.info(f"**Current Document:** {st.session_state['doc_name']}")

    st.markdown("---")
    st.header("2. Image Management")

    # Allow selection of multiple image files (JPG, PNG, GIF)
    uploaded_images = st.file_uploader(
        "Select Images", 
        type=["jpg", "jpeg", "png", "gif"], 
        accept_multiple_files=True, 
        key="uploaded_images"
    )

    if uploaded_images:
        st.subheader("Preview Selected Images")
        for image_file in uploaded_images:
            if image_file.size > MAX_IMAGE_SIZE:
                st.error(f"**{image_file.name}** exceeds the maximum file size limit.")
            else:
                try:
                    img = Image.open(image_file)
                    st.image(img, caption=image_file.name, width=200)
                except Exception as e:
                    st.error(f"Error previewing {image_file.name}: {e}")

    # Append Images button
    if st.button("Append Images"):
        if 'doc' not in st.session_state:
            st.error("Please create or select a document first.")
        elif not uploaded_images:
            st.error("Please select at least one image to append.")
        else:
            doc = st.session_state['doc']
            # Determine the starting label index by counting existing "Image{n}" labels.
            current_index = count_existing_images(doc)
            
            for image_file in uploaded_images:
                if image_file.size > MAX_IMAGE_SIZE:
                    st.error(f"**{image_file.name}** exceeds the maximum file size limit. Skipping this file.")
                    continue
                try:
                    # Reset the file pointer in case it was read during preview.
                    image_file.seek(0)
                    # Add a blank paragraph for spacing.
                    doc.add_paragraph("")
                    # Insert the image with a fixed width (maintaining aspect ratio).
                    doc.add_picture(image_file, width=Inches(4))
                    # Increment label count and add label directly below the image.
                    current_index += 1
                    doc.add_paragraph(f"Image{current_index}")
                    # Additional spacing.
                    doc.add_paragraph("")
                except Exception as e:
                    st.error(f"Error appending {image_file.name}: {e}")
            
            # Save the updated document into a BytesIO object.
            doc_io = BytesIO()
            try:
                doc.save(doc_io)
                doc_io.seek(0)
                st.session_state['doc_io'] = doc_io
                st.success("Images appended successfully!")
            except Exception as e:
                st.error(f"Error saving document: {e}")

    # Provide a download button if a document is available.
    if 'doc_io' in st.session_state:
        st.download_button(
            label="Download Updated Document",
            data=st.session_state['doc_io'],
            file_name=st.session_state.get('doc_name', 'document.docx'),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

if __name__ == "__main__":
    main()
