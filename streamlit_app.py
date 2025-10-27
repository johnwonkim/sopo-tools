import streamlit as st

st.title("SOPO Internal Tools")
st.write("Select a tool from the left sidebar to get started")

fooda_page = st.Page("pages/fooda.py", title="Fooda Label Generator")
grubhub_page = st.Page("pages/grubhub.py", title="Grubhub Label Generator")
grubhub_pdf_page = st.Page("pages/grubhub_pdf.py", title="Grubhub PDF Generator")

pg = st.navigation([fooda_page, grubhub_page, grubhub_pdf_page])
pg.run()