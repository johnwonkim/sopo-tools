import streamlit as st
from openai import OpenAI
import base64
import fitz  # PyMuPDF

st.title("Fooda Label Generator")

# Initialize OpenAI client (using Streamlit secrets for API key)
client = OpenAI(api_key=st.secrets["openai_api_key"])

uploaded_file = st.file_uploader("Upload PDF file containing orders", type=['pdf'])

if uploaded_file and client:
    with st.spinner("Converting PDF to images..."):
        # Read the PDF file
        file_contents = uploaded_file.read()
        
        # Convert PDF pages to images using PyMuPDF
        pdf_document = fitz.open(stream=file_contents, filetype="pdf")
        
        if len(pdf_document) == 0:
            st.error("PDF appears to be empty or corrupted.")
        else:
            st.info(f"PDF has {len(pdf_document)} page(s). Converting to images...")
            
            # Convert each page to an image and encode as base64
            image_urls = []
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                # Render page to an image (pixmap)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x zoom for better quality
                
                # Convert to PNG bytes
                img_bytes = pix.tobytes("png")
                
                # Encode to base64
                base64_image = base64.b64encode(img_bytes).decode('utf-8')
                image_urls.append(f"data:image/png;base64,{base64_image}")
            
            pdf_document.close()
            
            st.success(f"Converted {len(image_urls)} page(s) to images")
            
            # Make LLM call with images attached
            with st.spinner("Extracting order information from PDF using LLM..."):
                try:
                    # Build messages with images
                    messages = [
                        {
                            "role": "system",
                            "content": "You are a helpful assistant that extracts order information from PDF documents. Extract all relevant order details including customer names, items, quantities, prices, order dates, and delivery information. Format the response in a clear, structured way. Note that the order information may appear on multiple pages; de-duplicate the information in the result. Also, note that the order information may appear on multiple pages; de-duplicate the information in the result. Also, note that some pages may not be relevant to the order information."
                        },
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "text",
                                    "text": "Extract order information from the following PDF pages. Analyze all pages and provide a comprehensive extraction of all order details. The result should be a list of orders, in the format of: <name> - <menu item> - <modifications, or mods>"
                                }
                            ] + [
                                {
                                    "type": "image_url",
                                    "image_url": {"url": img_url}
                                }
                                for img_url in image_urls
                            ]
                        }
                    ]
                    
                    # Use gpt-4o for vision capabilities (gpt-4o-mini doesn't support vision)
                    response = client.chat.completions.create(
                        model="gpt-4o",  # gpt-4o supports vision
                        messages=messages,
                        temperature=0.1  # Lower temperature for more consistent extraction
                    )
                    
                    extracted_info = response.choices[0].message.content
                    
                    st.success("Order information extracted successfully!")
                    st.markdown("### Extracted Order Information:")
                    st.markdown(extracted_info)
                    
                    # Option to download the extracted information
                    st.download_button(
                        label="Download Extracted Information",
                        data=extracted_info.encode('utf-8'),
                        file_name="extracted_order_info.txt",
                        mime="text/plain"
                    )
                    
                except Exception as e:
                    st.error(f"Error calling OpenAI API: {str(e)}")
                    st.info("Make sure your API key is valid, you have credits available, and you're using a model that supports vision (gpt-4o).")

elif uploaded_file and not client:
    st.warning("Please enter your OpenAI API key in the sidebar to process the PDF.")