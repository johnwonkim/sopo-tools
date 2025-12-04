import streamlit as st
from openai import OpenAI
import base64
import fitz  # PyMuPDF
import re
import json
import sys
import os

# Add parent directory to path to import from grubhub
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from grubhub import create_labels_doc

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
                    "content": "You are a helpful assistant that extracts order label information from PDF documents. Extract all relevant order label details including customer names, items, quantities, order dates, and delivery information. Return the response as a JSON array of orders. Note that the order information may appear on multiple pages; de-duplicate the information in the result. Some pages may not be relevant to the order label information."
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": """Extract order label information from the following PDF pages. Analyze all pages and provide a comprehensive extraction of all order details.

    Return the result as a JSON array of orders with the following structure:
    [
        {
            "order_group_id": "Order Group ID (will usually be an alphanumeric string of length 3, and be identical for all orders.)",
            "order_id": "Order ID (will usually be a number, and be unique for each customer)",
            "name": "Customer Name (e.g. John Doe)",
            "order_date": "Order Date (in the format of YYYY-MM-DD)",
            "delivery_name": "Delivery Name/Address",
            "items": [
                {
                "quantity": "Quantity of the item",
                "menu_item": "Menu Item Name (e.g. Beef Bulgogi Signature Plate)",
                "modifications": "Modifications or special instructions (if any)"
                }
            ]
        }
    ]

    If information is not available, use empty strings or omit fields."""
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

        except Exception as e:
            st.error(f"Error calling OpenAI API: {str(e)}")
            st.info("Make sure your API key is valid, you have credits available, and you're using a model that supports vision (gpt-4o).")

    with st.spinner("Parsing extracted order information and generating labels..."):
        # Parse JSON from response
        try:
            # Try to extract JSON from the response (may have markdown code blocks)
            json_match = re.search(r'\[.*\]', extracted_info, re.DOTALL)
            if json_match:
                orders_data = json.loads(json_match.group())
            else:
                orders_data = json.loads(extracted_info)
            
            # Convert to label format
            labels = []
            for order in orders_data:
                order_group_id = order.get("order_group_id", "")
                order_id = order.get("order_id", "")
                name = order.get("name", "Unknown")
                order_date = order.get("order_date", "")
                delivery_name = order.get("delivery_name", "")
                items = order.get("items", [])
                
                if not items:
                    continue
                
                total_items = len(items)
                
                for idx, item in enumerate(items, start=1):
                    qty = item.get("quantity", "1")
                    menu_item = item.get("menu_item", "")
                    modifications = item.get("modifications", "")
                    
                    # Build item text
                    item_lines = [f"{qty} {menu_item}"]
                    if modifications:
                        item_lines.append(f"Mods: {modifications}")
                    
                    item_text = "\n".join(item_lines)
                    
                    # Build label in the format expected by create_labels_doc
                    label_lines = [f"{name} - {order_group_id}"]
                    label_lines.append(f'{order_group_id} - {order_id}')
                    label_lines.append(order_date)
                    label_lines.extend([
                        "-------------------------------------",
                        f"Item {idx} out of {total_items}",
                        item_text,
                        "-------------------------------------",
                        "Fooda"
                    ])
                    if delivery_name:
                        label_lines.append(f"Deliver to: {delivery_name}")
                    
                    labels.append("\n".join(label_lines))
            
            if labels:
                # Generate docx file using create_labels_doc
                st.info(f"Generated {len(labels)} labels")
                doc_bytes = create_labels_doc(labels, "Fooda Labels")
                
                # Download button for docx file
                st.download_button(
                    label="Download Labels (DOCX)",
                    data=doc_bytes,
                    file_name="fooda_labels.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("No orders found in the extracted information.")
        
        except json.JSONDecodeError as e:
            st.warning("Could not parse JSON from LLM response. Showing raw text.")
            st.download_button(
                label="Download Extracted Information (TXT)",
                data=extracted_info.encode('utf-8'),
                file_name="extracted_order_info.txt",
                mime="text/plain"
            )
        except Exception as e:
            st.error(f"Error processing orders: {str(e)}")
            st.download_button(
                label="Download Extracted Information (TXT)",
                data=extracted_info.encode('utf-8'),
                file_name="extracted_order_info.txt",
                mime="text/plain"
            )

elif uploaded_file and not client:
    st.warning("Please enter your OpenAI API key in the sidebar to process the PDF.")