import streamlit as st
from io import BytesIO
from pptx import Presentation

# Function to extract text from PPTX
def extract_text_from_pptx(file):
    """Extract text from a PowerPoint file."""
    presentation = Presentation(file)
    text = []
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text.append(shape.text)
    return "\n".join(text)

# Set page layout
st.set_page_config(page_title="PPT Text Extractor", layout="centered")

# Title and introduction
st.markdown("""
    <div style="text-align:center;">
        <h1 style="color:#4CAF50;">PowerPoint Text Extractor</h1>
        <p style="font-size:20px; color:#777;">Developed by Muhammad Haris</p>
    </div>
    <br>
    <p style="text-align:center;">Upload a PowerPoint (.pptx) file to extract its text content.</p>
""", unsafe_allow_html=True)

# File uploader widget
uploaded_file = st.file_uploader("Choose a PPT file", type=["pptx"], label_visibility="collapsed")

# If a file is uploaded, process it
if uploaded_file is not None:
    # Read the file
    file_bytes = uploaded_file.read()
    
    # Extract text from the uploaded PPT
    text = extract_text_from_pptx(BytesIO(file_bytes))
    
    # Display the extracted text
    st.subheader("Extracted Text")
    st.text_area("Text Content", text, height=300, key="extracted_text", disabled=True)

    # Add a download button
    st.download_button(
        label="Download Extracted Text",
        data=text,
        file_name="extracted_text.txt",
        mime="text/plain",
        use_container_width=True
    )

    # Show message for successful extraction
    st.markdown("""
        <div style="text-align:center; margin-top:20px;">
            <p style="color: #4CAF50;">Text successfully extracted!</p>
        </div>
    """, unsafe_allow_html=True)
    
# Footer with developer information
st.markdown("""
    <div style="text-align:center; font-size:12px; color:#888; margin-top:40px;">
        <p>&copy; 2024 Muhammad Haris. All rights reserved.</p>
    </div>
""", unsafe_allow_html=True)
