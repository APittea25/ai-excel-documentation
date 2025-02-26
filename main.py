import streamlit as st
import pandas as pd
import openai
import graphviz
import openpyxl

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
    
    # Toggle button for full sheet preview
    show_full = st.checkbox("Show Full Sheet")
    if show_full:
        st.dataframe(sheet_data)
    
    # Generate Flow Diagram of Sheets
    st.write("### 🔄 Spreadsheet Flow Diagram")
    flow = graphviz.Digraph()
    
    # Detect formula-based relationships
    wb = openpyxl.load_workbook(uploaded_file, data_only=False)
    sheet_links = {}
    
    for sheet in sheet_names:
        ws = wb[sheet]
        sheet_links[sheet] = set()
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    for ref_sheet in sheet_names:
                        if ref_sheet in cell.value:
                            sheet_links[sheet].add(ref_sheet)
    
    for sheet in sheet_names:
        if sheet == selected_sheet:
            flow.node(sheet, color="lightblue", style="filled")  # Highlight selected tab
        else:
            flow.node(sheet)
    
    for sheet, links in sheet_links.items():
        for linked_sheet in links:
            flow.edge(sheet, linked_sheet)
    
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
    
    # Generate AI-Powered Code Representation of the Sheet with Formulas
    st.write("### 🖥️ AI-Generated Python Code Replicating Excel Formulas")
    formula_prompt = f"Generate a Python script using pandas that replicates the formulas in the following Excel sheet:\n{sample_data}\nInclude any necessary calculations that reflect Excel formulas."
    
    try:
        formula_response = client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": formula_prompt}]
        )
        generated_code = formula_response.choices[0].message.content
        st.code(generated_code, language='python')
    except Exception as e:
        st.error(f"⚠️ OpenAI API Error: {e}")
    
    # Add chat input for follow-up questions
    st.write("### 💬 Ask AI Further Questions")
    user_query = st.text_area("Ask a question about this spreadsheet:")
    if st.button("Submit Question") and user_query:
        query_prompt = f"Based on this dataset, answer the following question: {user_query}\n{sample_data}"
        try:
            query_response = client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "user", "content": query_prompt}]
            )
            st.write(query_response.choices[0].message.content)
        except Exception as e:
            st.error(f"⚠️ OpenAI API Error: {e}")
else:
    st.warning("⚠️ Please upload an Excel file to proceed.")
