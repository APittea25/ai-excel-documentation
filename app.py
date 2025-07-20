import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
import os
from collections import defaultdict
import graphviz
from docx import Document
import pandas as pd

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")

# ✅ Step 1: Add this here
if "json_summaries" not in st.session_state:
    st.session_state.json_summaries = None

# Print mode controls
if "print_mode" not in st.session_state:
    st.session_state.print_mode = "full"

def set_full():
    st.session_state.print_mode = "full"

def set_summary():
    st.session_state.print_mode = "summary"

col1, col2 = st.columns(2)
with col1:
    st.button("🖨️ Print Full", on_click=set_full)
with col2:
    st.button("🖨️ Print Summary (first 50 lines)", on_click=set_summary)

st.write(f"**Current print mode:** {st.session_state.print_mode}")

st.title("\U0001F4D8 Named Range Coordinates + Formula Remapping")

if "expanded_all" not in st.session_state:
    st.session_state.expanded_all = False

def toggle():
    st.session_state.expanded_all = not st.session_state.expanded_all

st.button("🔁 Expand / Collapse All Named Ranges", on_click=toggle)

# —– Add print-specific CSS to ensure full expansion & wrapping —–
st.markdown("""
<style>
  @media print {
    /* Make sure all expanders are shown (no collapsing in print mode) */
    .streamlit-expanderHeader {
      display: block !important;
    }
    /* Expand all details sections */
    details {
      display: block !important;
    }
    /* Enable wrapping and disable horizontal scrollbars in code blocks */
    .stCodeBlock, pre, code {
      overflow-x: visible !important;
      white-space: pre-wrap !important;
      word-wrap: break-word !important;
    }
  }
</style>
""", unsafe_allow_html=True)

# Allow manual mapping of external references like [1], [2], etc.
st.subheader("Manual Mapping for External References")
external_refs = {}
for i in range(1, 10):
    ref_key = f"[{i}]"
    workbook_name = st.text_input(f"Map external reference {ref_key} to workbook name (e.g., Mortality_Model_Inputs.xlsx)", key=ref_key)
    if workbook_name:
        external_refs[ref_key] = workbook_name

uploaded_files = st.file_uploader("\U0001F4C2 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_display_names = {}
    named_ref_formulas = {}

    for uploaded_file in uploaded_files:
        display_name = uploaded_file.name
        file_display_names[display_name] = uploaded_file
        st.header(f"\U0001F4C4 File: {display_name}")
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

        for name in wb.defined_names:
            dn = wb.defined_names[name]
            if dn.is_external or not dn.attr_text:
                continue
            for sheet_name, ref in dn.destinations:
                try:
                    ws = wb[sheet_name]
                    ref_clean = ref.replace("$", "").split("!")[-1]
                    cells = ws[ref_clean] if ":" in ref_clean else [[ws[ref_clean]]]

                    min_row = min(cell.row for row in cells for cell in row)
                    min_col = min(cell.column for row in cells for cell in row)

                    coord_set = set()
                    for row in cells:
                        for cell in row:
                            r, c = cell.row, cell.column
                            row_offset = r - min_row + 1
                            col_offset = c - min_col + 1
                            all_named_cell_map[(display_name, sheet_name, r, c)] = (name, row_offset, col_offset)
                            coord_set.add((r, c))
                    all_named_ref_info[name] = (display_name, sheet_name, coord_set, min_row, min_col)
                except:
                    continue

    def remap_formula(formula, current_file, current_sheet):
        if not formula:
            return ""

        def cell_address(row, col):
            return f"{get_column_letter(col)}{row}"

        def remap_single_cell(ref, default_file, default_sheet):
        # Use regex to safely extract sheet name and address
            match = re.match(r"(?:'([^']+)'|([^'!]+))!([$A-Z]+[0-9]+)", ref)
            if match:
                sheet_name = match.group(1) or match.group(2)
                addr = match.group(3)
                # Check for external reference like [1]
                external_match = re.match(r"\[(\d+)\]", sheet_name)
                if external_match:
                    external_ref = external_match.group(0)
                    external_file = external_refs.get(external_ref, external_ref)
                    return f"[{external_file}]{ref}"
            else:
                sheet_name = default_sheet
                addr = ref

            addr = addr.replace("$", "").upper()
            match = re.match(r"([A-Z]+)([0-9]+)", addr)
            if not match:
                return ref
            col_str, row_str = match.groups()
            row = int(row_str)
            col = column_index_from_string(col_str)

            key = (default_file, sheet_name, row, col)
            if key in all_named_cell_map:
                name, r_off, c_off = all_named_cell_map[key]
                return f"[{default_file}]{name}[{r_off}][{c_off}]"
            else:
                return f"{sheet_name}!{addr}"  # Removed [default_file] from fallback

        def remap_range(ref, default_file, default_sheet):
            # Handle external references like [1]Sheet!A1
            if ref.startswith("["):
                match = re.match(r"\[(\d+)\]", ref)
                if match:
                    external_ref = match.group(0)
                    external_file = external_refs.get(external_ref, external_ref)
                    return f"[{external_file}]{ref}"

            # Use regex to safely extract sheet name and address or range
            match = re.match(r"(?:'([^']+)'|([^'!]+))!([$A-Z]+[0-9]+(?::[$A-Z]+[0-9]+)?)", ref)
            if match:
                sheet_name = match.group(1) or match.group(2)
                addr = match.group(3)
            else:
                sheet_name = default_sheet
                addr = ref

            addr = addr.replace("$", "").upper()
            if ":" not in addr:
                return remap_single_cell(ref, default_file, default_sheet)

            # It's a range like A1:B2
            start, end = addr.split(":")
            m1 = re.match(r"([A-Z]+)([0-9]+)", start)
            m2 = re.match(r"([A-Z]+)([0-9]+)", end)
            if not m1 or not m2:
                return ref
            start_col = column_index_from_string(m1.group(1))
            start_row = int(m1.group(2))
            end_col = column_index_from_string(m2.group(1))
            end_row = int(m2.group(2))

            label_set = set()
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    key = (default_file, sheet_name, row, col)
                    if key in all_named_cell_map:
                        name, r_off, c_off = all_named_cell_map[key]
                        label_set.add(f"[{default_file}]{name}[{r_off}][{c_off}]")
                    else:
                        label_set.add(f"{sheet_name}!{get_column_letter(col)}{row}")
            return ", ".join(sorted(label_set))

        pattern = r"(?<![A-Za-z0-9_])(?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?|(?<![A-Za-z0-9_])\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?"
        matches = list(re.finditer(pattern, formula))
        replaced_formula = formula
        offset = 0
        for match in matches:
            raw = match.group(0)
            remapped = remap_range(raw, current_file, current_sheet)
            start, end = match.start() + offset, match.end() + offset
            replaced_formula = replaced_formula[:start] + remapped + replaced_formula[end:]
            offset += len(remapped) - len(raw)
        return replaced_formula

    for (name, (file_name, sheet_name, coord_set, min_row, min_col)) in all_named_ref_info.items():
        entries = []
        formulas_for_graph = []

        try:
            file_bytes = file_display_names[file_name]
            wb = load_workbook(BytesIO(file_bytes.getvalue()), data_only=False)
            ws = wb[sheet_name]
            min_col_letter = get_column_letter(min([c for (_, c) in coord_set]))
            max_col_letter = get_column_letter(max([c for (_, c) in coord_set]))
            min_row_num = min([r for (r, _) in coord_set])
            max_row_num = max([r for (r, _) in coord_set])
            ref_range = f"{min_col_letter}{min_row_num}:{max_col_letter}{max_row_num}"
            cell_range = ws[ref_range] if ":" in ref_range else [[ws[ref_range]]]

            for row in cell_range:
                for cell in row:
                    row_offset = cell.row - min_row + 1
                    col_offset = cell.column - min_col + 1
                    label = f"{name}[{row_offset}][{col_offset}]"

                    try:
                        formula = None
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            formula = cell.value.strip()
                        elif hasattr(cell, 'value') and hasattr(cell.value, 'text'):
                            formula = str(cell.value.text).strip()
                        elif hasattr(cell, 'value'):
                            formula = str(cell.value)

                        if formula:
                            remapped = remap_formula(formula, file_name, sheet_name)
                            formulas_for_graph.append(remapped)
                        elif cell.value is not None:
                            formula = f"[value] {str(cell.value)}"
                            remapped = formula
                        else:
                            formula = "(empty)"
                            remapped = formula
                    except Exception as e:
                        formula = f"[error reading cell: {e}]"
                        remapped = formula

                    entries.append(f"{label} = {formula}\n → {remapped}")
        except Exception as e:
            entries.append(f"❌ Error accessing {name} in {sheet_name}: {e}")

        named_ref_formulas[name] = formulas_for_graph
        
        limit = None if st.session_state.print_mode == "full" else 50
        snippet = entries if limit is None else entries[:limit]
        
        with st.expander(
            f"📌 Named Range: {name} → {sheet_name} in {file_name}",
            expanded=st.session_state.expanded_all
        ):
            st.code("\n".join(snippet), language="text")
            if limit is not None and len(entries) > limit:
                st.write(f"...and {len(entries) - limit} more lines hidden")
                
    # —– Missing direct cell references (not in any named range) —–
    with st.expander("⚠️ Missing Direct Cell References", expanded=True):
        st.markdown("#### 🔍 Check for A1-style cell references not covered by any named range")

        raw_ref_re = re.compile(r"\b([A-Z]{1,3}[0-9]{1,7})\b")
        missing_refs = defaultdict(set)

        for nm, formulas in named_ref_formulas.items():
            for f in formulas:
                for ref in raw_ref_re.findall(f):
                    if re.search(rf"\[{nm}\]\[\d+\]\[\d+\]", f):
                        continue
                    missing_refs[nm].add(ref)

        if missing_refs:
            for nm, refs in missing_refs.items():
                st.warning(
                    f"In **{nm}**, these direct cell refs weren’t wrapped by a named range: "
                    f"{', '.join(sorted(refs))}"
                )
        else:
            st.success("✅ No missing direct cell references found.")
    
    # Dependency Graph
    st.subheader("🔗 Dependency Graph")
    dot = graphviz.Digraph()
    dot.attr(compound='true', rankdir='LR')

    grouped = defaultdict(list)
    for name, (file, *_rest) in all_named_ref_info.items():
        grouped[file].append(name)

    dependencies = defaultdict(set)
    for target, formulas in named_ref_formulas.items():
        joined = " ".join(formulas)
        for source in named_ref_formulas:
            if source != target and re.search(rf"\b{re.escape(source)}\b", joined):
                dependencies[target].add(source)

    for i, (file_name, nodes) in enumerate(grouped.items()):
        with dot.subgraph(name=f"cluster_{i}") as c:
            c.attr(label=file_name)
            c.attr(style='filled', color='lightgrey')
            for node in nodes:
                c.node(node)

    for target, sources in dependencies.items():
        for source in sources:
            dot.edge(source, target)

    st.graphviz_chart(dot)
# --- JSON Summary Generation Section ---
    st.subheader("🧠 AI-Powered JSON Summary of Named Range Calculations")

    generate_json = st.button("🧾 Generate Summarised JSON Output")

    if generate_json:
        from openai import OpenAI
        import json

        client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

        summaries = {}

        for name, formulas in named_ref_formulas.items():
            if not formulas:
                continue

            prompt = f"""
    You are an expert actuary and spreadsheet analyst.

    Given the following remapped formulas from an Excel named range, summarize the pattern behind the calculations in a general form.
    Each formula follows a remapped structure using notation like [1][2] to indicate row and column indices.

    Please return a JSON object like:
    {{
      "file_name": "MyWorkbook.xlsx",
      "sheet_name": "Inputs",
      "excel_range": "B2:D5",
      "named_range": "MyNamedRange",
      "summary": "Description of what the formula does",
      "general_formula": "for i in range(...): for j in range(...): Result[i][j] = ...",
      "dependencies": ["OtherNamedRange1", "OtherNamedRange2"],
      "notes": "Any caveats, limitations, or variations found"
    }}

    Formulas:
    {formulas[:10]}  # Sample first 10 for context

    Only return the JSON.
    """

            try:
                response = client.chat.completions.create(
                    model="gpt-4",
                    messages=[
                        {"role": "system", "content": "You summarize spreadsheet formulas into structured JSON."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3
                )
                content = response.choices[0].message.content
                parsed = json.loads(content)
                
                # Get file_name, sheet_name, coord_set for this named range
                file_name, sheet_name, coord_set, *_ = all_named_ref_info[name]

                # Calculate Excel range
                min_col_letter = get_column_letter(min([c for (_, c) in coord_set]))
                max_col_letter = get_column_letter(max([c for (_, c) in coord_set]))
                min_row_num = min([r for (r, _) in coord_set])
                max_row_num = max([r for (r, _) in coord_set])
                excel_range = f"{min_col_letter}{min_row_num}:{max_col_letter}{max_row_num}"
                
                parsed["named_range"] = name  # ✅ Ensure correctness
                parsed["file_name"] = file_name
                parsed["sheet_name"] = sheet_name
                parsed["excel_range"] = excel_range
                parsed["dependencies"] = sorted(dependencies.get(name, []))
                summaries[name] = parsed
            except Exception as e:
                summaries[name] = {"named_range": name,"error": str(e)}

        with st.expander("📦 View JSON Output", expanded=False):
            st.json(summaries)

        #prepare content for documentation

        #input data
        input_summaries = {k: v for k, v in summaries.items() if k.startswith("i_")}
        inputs_data = []
        for idx, (name, summary) in enumerate(input_summaries.items(), start=1):
            source_file = summary.get("file_name", "")
            inputs_data.append({
                "No.": idx,
                "Name": name,
                "Type": "",
                "Source": source_file,
                "Info": ""  # To be filled by GPT
            })

        # GPT to populate Info field
        for row in inputs_data:
            prompt = f"""You are an expert actuary and survival modeller.

        You're documenting a spreadsheet input named `{input_name}`, located in sheet `{sheet}`, cell range `{excel_range}`.

        Its name suggests it's related to: "{input_name}"

        Here is the description of how this input is used in the model:
        "{json_summary}"

        And the general formula pattern that references it:
        "{general_formula}"

        Based on this, describe what `{input_name}` represents and its role in the model, using clear and confident actuarial language. Do not use words like "might", "possibly", or "likely".

        Respond with 1–2 precise sentences."""

            try:
                response = client.chat.completions.create(
                    model="gpt-4",
                    messages=[
                        {"role": "system", "content": "You provide concise descriptions of actuarial inputs."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3
                )
                row["Info"] = response.choices[0].message.content.strip()
            except Exception as e:
                row["Info"] = f"Error: {e}"
        inputs_df = pd.DataFrame(inputs_data)
        
        with st.expander("📄 Spreadsheet Document", expanded=False):
            st.title("📄 Model Documentation")

            st.header("## Version Control")

            # Model version control table
            st.subheader("### Model Version Control")
            model_version_df = pd.DataFrame({
                "Version": [""],
                "Date": ["00/00/0000"],
                "Info": [""],
                "Updated by": [""],
                "Reviewed by": [""],
                "Review Date": [""]
            })
            st.dataframe(model_version_df, use_container_width=True)

            # Documentation version control table
            st.subheader("### Documentation Version Control")
            doc_version_df = pd.DataFrame({
                "Version": [""],
                "Date": ["00/00/0000"],
                "Info": [""],
                "Updated by": [""],
                "Reviewed by": [""],
                "Review Date": [""]
            })
            st.dataframe(doc_version_df, use_container_width=True)

            # Ownership section
            st.header("## Ownership")
            st.text("### Owner")
            st.text("### Risk rating (or other client control standard)")
            st.text("### Internal audit history")

            # Purpose
            st.header("## Purpose")
            st.text_area("Describe the purpose of the model:")

            # Inputs table
            st.header("## Inputs")
            row_height = 35
            max_height = 500
            calculated_height = min(len(inputs_df) * row_height + 35, max_height)
            st.dataframe(inputs_df, use_container_width=True)

            # Outputs
            st.header("## Outputs")
            st.text_area("Describe the outputs of the model:")

            # Logic
            st.header("## Logic")
            st.text_area("Describe the logic used in the model:")

            # Checks and validation
            st.header("## Checks and Validation")
            st.text_area("Describe the checks and validation steps:")

            # Assumptions and limitations
            st.header("## Assumptions and Limitations")
            st.text_area("List assumptions and limitations:")

            # TAS Compliance
            st.header("## TAS Compliance")
            st.text_area("Describe how the model complies with TAS:")

        # JSON download
        json_str = json.dumps(summaries, indent=2)
        st.download_button("📥 Download JSON Summary", data=json_str, file_name="named_range_summaries.json", mime="application/json")


        doc = Document()
        doc.add_heading("Named Range JSON Summary", 0)
        for name, summary in summaries.items():
            doc.add_heading(name, level=1)
            for key, value in summary.items():
                if isinstance(value, (list, dict)):
                    value = json.dumps(value, indent=2)
                doc.add_paragraph(f"{key}: {value}")

        docx_io = BytesIO()
        doc.save(docx_io)
        docx_io.seek(0)

        st.download_button(
            "📄 Download Summary as Word Document",
            data=docx_io,
            file_name="named_range_summary.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    else:
        st.info("Press the button above to generate a GPT-based JSON summary of calculations.")

else:
    st.info("⬆️ Upload one or more .xlsx files to begin.")
