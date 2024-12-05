import streamlit as st
from io import BytesIO
from pptx import Presentation

def extract_text_from_pptx(file):
    """Extract text from a PowerPoint file."""
    presentation = Presentation(file)
    text = []
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text.append(shape.text)
    return "\n".join(text)

st.title("PPT Text Extractor")

# File uploader widget
uploaded_file = st.file_uploader("Choose a PPT file", type=["pptx"])

if uploaded_file is not None:
    # Read the file
    file_bytes = uploaded_file.read()
    
    # Extract text from the uploaded PPT
    text = extract_text_from_pptx(BytesIO(file_bytes))
    
    # Display the extracted text
    st.subheader("Extracted Text")
    st.text_area("Text Content", text, height=300)
    
    # Add a copy button
    st.download_button(
        label="Download Text",
        data=text,
        file_name="extracted_text.txt",
        mime="text/plain",
    )
