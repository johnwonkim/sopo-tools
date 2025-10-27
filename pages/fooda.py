import streamlit as st

st.title("Fooda Label Generator")

uploaded_file = st.file_uploader("Upload PDF file containing orders", type=['pdf'])
if uploaded_file:
    # Read the file contents
    file_contents = uploaded_file.read()
    
    # Create download button with the file contents
    st.download_button(
        label="Download Labels",
        data=file_contents,
        file_name=uploaded_file.name,
        mime=uploaded_file.type or "application/octet-stream"
    )