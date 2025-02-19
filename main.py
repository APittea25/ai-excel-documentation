import streamlit as st
import pandas as pd
import openai

openai_api_key = st.secrets["OPENAI_API_KEY"]

if not openai_api_key:
    st.error("⚠️ OpenAI API key is missing. Add it to Streamlit Secrets.")
else:
    openai.api_key = openai_api_key

# App title
st.title("📊 AI-Powered Excel Documentation")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success(f"✅ File '{uploaded_file.name}' uploaded successfully!")
    
    # Read Excel file
    df = pd.read_excel(uploaded_file, sheet_name=None)  # Read all sheets
    sheet_names = df.keys()
    
    # Let user select a sheet
    selected_sheet = st.selectbox("Select a sheet", sheet_names)

    # Show preview
    st.write(f"### Preview of {selected_sheet}")
    st.dataframe(df[selected_sheet].head())

    # Generate AI-powered documentation
    st.write("### 📝 AI-Generated Documentation")

    # Prepare data for AI
    sample_data = df[selected_sheet].head().to_dict()

    # Create prompt for AI
    prompt = f"Analyze this Excel sheet and describe its structure, column meanings, and any insights:\n{sample_data}"

    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        ai_summary = response["choices"][0]["message"]["content"]
        st.write(ai_summary)
    except Exception as e:
        st.error("⚠️ Error fetching AI response. Check your API key and limits.")

else:
    st.warning("⚠️ Please upload an Excel file to proceed.")

