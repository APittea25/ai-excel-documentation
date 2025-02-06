import streamlit as st
import pandas as pd

# App title
st.title("📊 Excel File Uploader")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

# If a file is uploaded, read and display it
if uploaded_file is not None:
    st.success(f"✅ File '{uploaded_file.name}' uploaded successfully!")
    
    # Read Excel file
    df = pd.read_excel(uploaded_file, sheet_name=None)  # Read all sheets
    sheet_names = df.keys()
    
    # Let user select a sheet to display
    selected_sheet = st.selectbox("Select a sheet", sheet_names)

    # Show the selected sheet
    st.write(f"### Preview of {selected_sheet}")
    st.dataframe(df[selected_sheet].head())  # Show first few rows

else:
    st.warning("⚠️ Please upload an Excel file to proceed.")

