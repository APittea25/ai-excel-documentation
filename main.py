import streamlit as st
import pandas as pd
import openai
import graphviz
import openpyxl
import re

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

            # Documentation Prompt
            prompt = f"Analyze this Excel sheet and describe its structure, column meanings, and any insights:\n{sample_data}"
            # Technical Spec Prompt
            tech_spec_prompt = f"""
You're an Excel systems analyst. Based on this sample of the '{sheet}' sheet, generate a technical specification:
- Describe key inputs and their types
- Outline the calculation logic and dependencies
- Explain what the outputs represent
- Note any assumptions or design features

Sample data:
{sample_data}
"""
            # Python Code Prompt
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
                tech_response = client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": tech_spec_prompt}]
                )
                technical_spec = tech_response.choices[0].message.content
            except Exception as e:
                technical_spec = f"⚠️ OpenAI API Error: {e}"

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
                "technical_spec": technical_spec,
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
if re.search(rf'\b{ref_sheet}!', cell.value, re.IGNORECASE):
                            if ref_sheet not in sheet_links:
                                sheet_links[ref_sheet] = set()
                            sheet_links[ref_sheet].add(sheet)

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
    st.write(st.session_state.ai_responses[selected_sheet].get("summary", "⚠️ No summary available."))

    # Show AI-generated technical specification
    st.write("### 📘 AI-Generated Technical Specification")
    st.write(st.session_state.ai_responses[selected_sheet].get("technical_spec", "⚠️ No technical spec available."))

    # Show AI-generated Python code
    st.write("### 🖥️ AI-Generated Python Code Replicating Excel Formulas")
    st.code(st.session_state.ai_responses[selected_sheet].get("code", "⚠️ No AI-generated code available."), language='python')

else:
    st.warning("⚠️ Please upload an Excel file to proceed.")
