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
    
    # Generate AI responses for all sheets on upload
    if 'ai_responses' not in st.session_state or st.button("🔄 Refresh AI Responses"):
        st.session_state.ai_responses = {}
        for sheet in sheet_names:
            sheet_data = df.parse(sheet)
            sample_data = sheet_data.head().to_dict()
            prompt = f"Analyze this Excel sheet and describe its structure, column meanings, and any insights:\n{sample_data}"
            formula_prompt = f"Generate a Python script using pandas that replicates the formulas in the following Excel sheet:\n{sample_data}\nInclude any necessary calculations that reflect Excel formulas."
            
            try:
                response = client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": prompt}]
                )
                ai_summary = response.choices[0].message.content
            except Exception as e:
                ai_summary = f"⚠️ OpenAI API Error: {e}"
            
            try:
                formula_response = client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": formula_prompt}]
                )
                generated_code = formula_response.choices[0].message.content
            except Exception as e:
                generated_code = f"⚠️ OpenAI API Error: {e}"
            
            st.session_state.ai_responses[sheet] = {
                "summary": ai_summary,
                "code": generated_code
            }
    
    # Let user select a sheet
    selected_sheet = st.selectbox("Select a sheet", sheet_names)
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
    
    # Detect formula-based and table-based relationships
    wb = openpyxl.load_workbook(uploaded_file, data_only=False)
    sheet_links = {}
    table_references = {}
    table_to_sheet = {}
    
    # Step 1: Identify all tables and their respective sheets
    for sheet in sheet_names:
        ws = wb[sheet]
        sheet_links[sheet] = set()
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and "[" in cell.value and "]" in cell.value:
                    table_name = cell.value.split("[")[0].replace("=", "").strip()
                    table_to_sheet[table_name] = sheet
    
    # Step 2: Detect references to tables in formulas
    for sheet in sheet_names:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    for table_name, table_sheet in table_to_sheet.items():
                        if table_name in cell.value and table_sheet != sheet:
                            if table_sheet not in sheet_links:
                                sheet_links[table_sheet] = set()
                            sheet_links[table_sheet].add(sheet)
    
    # Step 3: Generate the flow diagram
    for sheet in sheet_names:
        if sheet == selected_sheet:
            flow.node(sheet, color="lightblue", style="filled")
        else:
            flow.node(sheet)
    
    for sheet, links in sheet_links.items():
        for linked_sheet in links:
            flow.edge(linked_sheet, sheet)
    
    st.graphviz_chart(flow)
    
    # Show AI-generated documentation
    st.write("### 📝 AI-Generated Documentation")
    if selected_sheet in st.session_state.ai_responses:
        st.write(st.session_state.ai_responses[selected_sheet].get("summary", "⚠️ No AI response available. Try refreshing AI responses."))
    else:
        st.warning(f"⚠️ AI responses not available for '{selected_sheet}'. Try refreshing AI responses.")

    # Show AI-generated Python code
    st.write("### 🖥️ AI-Generated Python Code Replicating Excel Formulas")
    if selected_sheet in st.session_state.ai_responses:
        st.code(st.session_state.ai_responses[selected_sheet].get("code", "⚠️ No AI-generated code available. Try refreshing AI responses."), language='python')
    else:
        st.warning(f"⚠️ AI-generated code not available for '{selected_sheet}'. Try refreshing AI responses.")
else:
    st.warning("⚠️ Please upload an Excel file to proceed.")
