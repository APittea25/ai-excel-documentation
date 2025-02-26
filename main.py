import streamlit as st
import pandas as pd
import openai
import graphviz

# Get OpenAI API Key from Streamlit Secrets
openai_api_key = st.secrets.get("OPENAI_API_KEY")

# Initialize OpenAI client
if openai_api_key:
    client = openai.OpenAI(api_key=openai_api_key)
else:
    st.error("⚠️ OpenAI API key is missing. Add it to Streamlit Secrets.")
    st.stop()

# App title
st.title("📊 AI-Powered Excel Documentation")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    st.success(f"✅ File '{uploaded_file.name}' uploaded successfully!")
    
    # Read Excel file
    df = pd.ExcelFile(uploaded_file)
    sheet_names = df.sheet_names
    
    # Let user select a sheet
    selected_sheet = st.selectbox("Select a sheet", sheet_names)

    # Load selected sheet
    sheet_data = df.parse(selected_sheet)
    
    # Show preview
    st.write(f"### Preview of {selected_sheet}")
    st.dataframe(sheet_data.head())
    
    # Generate Flow Diagram of Sheets
    st.write("### 🔄 Spreadsheet Flow Diagram")
    flow = graphviz.Digraph()
    
    for sheet in sheet_names:
        if sheet == selected_sheet:
            flow.node(sheet, color="lightblue", style="filled")  # Highlight selected tab
        else:
            flow.node(sheet)
    
    for sheet in sheet_names:
        # Example: If sheets reference each other, we could add logic to detect links
        # Here, we assume sheets are connected sequentially for simplicity
        if sheet != sheet_names[-1]:
            flow.edge(sheet, sheet_names[sheet_names.index(sheet) + 1])
    
    st.graphviz_chart(flow)
    
    # Generate AI-powered documentation
    st.write("### 📝 AI-Generated Documentation")
    
    # Prepare data for AI
    sample_data = sheet_data.head().to_dict()
    prompt = f"Analyze this Excel sheet and describe its structure, column meanings, and any insights:\n{sample_data}"
    
    try:
        response = client.chat.completions.create(
            model="gpt-4",  # Using the cheaper model
            messages=[{"role": "user", "content": prompt}]
        )
        ai_summary = response.choices[0].message.content
        st.write(ai_summary)
    except Exception as e:
        st.error(f"⚠️ OpenAI API Error: {e}")
else:
    st.warning("⚠️ Please upload an Excel file to proceed.")

