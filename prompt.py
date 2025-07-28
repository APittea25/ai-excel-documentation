# prompt.py
def build_json_summary_prompt(named_range, formulas):
    return f"""
You are an expert actuary and spreadsheet analyst.

Given the following remapped formulas from an Excel named range, summarize the pattern behind the calculations in a general form.
Each formula follows a remapped structure using notation like [1][2] to indicate row and column indices.

Please return a JSON object like:
{{
  "file_name": "MyWorkbook.xlsx",
  "sheet_name": "Inputs",
  "excel_range": "B2:D5",
  "named_range": "{named_range}",
  "summary": "Description of what the formula does",
  "general_formula": "for i in range(...): for j in range(...): Result[i][j] = ...",
  "dependencies": ["OtherNamedRange1", "OtherNamedRange2"],
  "notes": "Any caveats, limitations, or variations found"
}}

Formulas:
{formulas[:10]}  # Sample first 10 for context

Only return the JSON.
"""


def build_purpose_prompt(summaries, example=None):
    joined_descriptions = "\n".join(
        f"{k}: {v.get('summary', '')}" for k, v in summaries.items() if "summary" in v
    )
    joined_formulas = "\n".join(
        f"{k}: {v.get('general_formula', '')}" for k, v in summaries.items() if "general_formula" in v
    )

    base = f"""You are an expert actuary and spreadsheet modeller.

You are reviewing an Excel model based on the **Lee-Carter mortality framework**.

The model uses named ranges and formulas structured to perform actuarial calculations.
"""

    if example:
        base += f"\n--- Example Purpose Statement ---\n{example.strip()}\n"

    base += f"""
Below are descriptions of how various parts of the model behave:

--- Summaries ---
{joined_descriptions}

--- Formula patterns ---
{joined_formulas}

Using this information, write a **concise and confident purpose statement** for documentation. Your paragraph should follow this structure:

1. Start with a clear sentence about what the model is designed to do (e.g. project mortality, simulate survival rates).
2. Describe what kinds of inputs it uses (e.g. mortality trends, drift terms, random simulations).
3. Summarize the types of outputs produced (e.g. annuity rates, survival curves).

Use actuarial language. Do not say “likely”, “possibly”, or “may”. Be direct and factual.
"""
    return base


def build_input_prompt(input_name, summary_json, hint_sentence=None, example=None):
    
    base = f"""You are an expert actuary and survival modeller.

You are reviewing a spreadsheet model based on the Lee-Carter mortality model or a closely related framework.

You're now documenting the spreadsheet input named `{input_name}`, located in sheet `{summary_json.get("sheet_name", "")}`, cell range `{summary_json.get("excel_range", "")}`.

Its name suggests it's related to: "{input_name}".
    
However, if the hint below provides more precise context, use that as the primary guide:
"{hint_sentence}". 

Here is the description of how this input is used in the model:
"{summary_json.get("summary", "")}"

And here is the general formula pattern that references it:
"{summary_json.get("general_formula", "")}"
"""

    if example:
        base += f'\n--- Example Input Description ---\n{example.strip()}\n'

    base += """
Based on all the above, write a concise, confident description of what `{input_name}` represents and how it contributes to the model.

Use actuarial language. Avoid vague expressions like “might”, “somewhat”, “typically”, or filler phrases like “plays a crucial role” or “is important”. Do not describe patterns in the data (e.g., “decreasing linearly”) unless they are explicitly mentioned.

Respond with one precise sentence, or two if the second adds new technical detail or context.
"""
    return base


def build_output_prompt(output_name, summary_json, hint_sentence=None, example=None):
    
    base = f"""You are an expert actuary and spreadsheet modeller.

You are reviewing an Excel spreadsheet built on the **Lee-Carter mortality model** or a closely related survival modelling framework.

You're documenting the model output named `{output_name}`, located in sheet `{summary_json.get("sheet_name", "")}`, cell range `{summary_json.get("excel_range", "")}`.

Its name suggests it primarily relates to: "{output_name}".

Use the following contextual hint as *additional guidance* where applicable:
"{hint_sentence}"

Here is how this output behaves in the model:
"{summary_json.get("summary", "")}"

And here is the formula structure used to calculate it:
"{summary_json.get("general_formula", "")}"
"""

    if example:
        base += f'\n--- Example Output Description ---\n{example.strip()}\n'

    base += """
Based on all the above, write a concise and confident explanation of what `{output_name}` represents and how it contributes to the model's output.

Use actuarial language. Do **not** include vague expressions like “might”, “possibly”, or “likely”, and avoid filler phrases like “plays a crucial role”, “important component”, or “used to calculate”. Focus instead on what it does and how it connects to the broader modelling framework.

Respond with **one precise sentence**, or two if the second adds useful technical context.
"""
    return base

def build_logic_prompt(name, summary_json, step_number, hint_sentence=None, example=None):
    base = f"""You are an expert actuary and spreadsheet modeller.

You are reviewing a **logic step** in an Excel model built on the **Lee-Carter mortality model** or a similar survival projection framework.

{hint_sentence or ""}

The step being documented is `{name}`, located in sheet `{summary_json.get("sheet_name", "")}`, range `{summary_json.get("excel_range", "")}`. It is labelled as **Step {step_number}** in the spreadsheet.

Here is a description of this step’s behavior in the model:
"{summary_json.get("summary", "")}"

And here is the generalised formula pattern:
"{summary_json.get("general_formula", "")}"

This step depends on the following named ranges:
{', '.join(summary_json.get("dependencies", [])) or "None listed"}
"""

    if example:
        base += f"\n--- Example Logic Description ---\n{example.strip()}\n"

    base += """
Now write a clear and technical description of this step, using the following exact structure:

1. The purpose of this calculation step.
2. The type of calculation it performs and what it is projecting or transforming.
3. Its direct dependencies and how they influence or feed into this step.

Respond in **three numbered sentences**, each covering one of the above points. Use actuarial language. Do **not** include filler phrases (e.g., "this is an important step") or speculative terms like "might" or "possibly".
"""
    return base

    
def build_check_prompt(name, summary_json, hint_sentence=None, example=None):
    base = f"""You are an expert actuary and spreadsheet modeller.

You are reviewing a **validation check** in an Excel model based on the **Lee-Carter mortality framework** or a similar survival model.

{hint_sentence}

The named range being reviewed is `{name}`, located in sheet `{summary_json.get("sheet_name", "")}`, cell range `{summary_json.get("excel_range", "")}`.

Here is a summary of the logic used in this check:
"{summary_json.get("summary", "")}"

And here is the general formula pattern:
"{summary_json.get("general_formula", "")}"
"""

    if example:
        base += f"\n--- Example Check Description ---\n{example.strip()}\n"

    base += """
Write a clear and confident description of what this check is verifying, referencing model outputs, intermediate calculations, or assumptions where relevant.

Avoid vague words like “might” or “appears to”, and do not use generic filler like “this is a check to ensure…”.

Respond with one precise sentence explaining what this check validates or confirms.
"""
    return base


def build_assumptions_prompt(summaries, example=None):
    all_summaries = "\n".join(
        f"{k}: {v.get('summary', '')}" for k, v in summaries.items() if "summary" in v
    )
    all_formulas = "\n".join(
        f"{k}: {v.get('general_formula', '')}" for k, v in summaries.items() if "general_formula" in v
    )

    base_prompt = f"""You are an expert actuary and spreadsheet modeller reviewing a workbook based on the **Lee-Carter mortality model** or a similar mortality projection framework.

Below are summaries and general formulas of the spreadsheet's calculations:

--- Summaries ---
{all_summaries}

--- Formula patterns ---
{all_formulas}
"""

    if example:
        base_prompt += f"\n--- Example Assumptions Paragraph ---\n{example.strip()}\n"

    base_prompt += """
Using this information, write a **short, clear paragraph** that outlines:

1. The key assumptions used in this spreadsheet (e.g. mortality trends, parameter stability, projection horizon).
2. Any notable limitations or modelling constraints (e.g. fixed inputs, deterministic assumptions, lack of stress testing).

Avoid vague phrases like “it might be assumed” or “possibly”. Be direct and professional.
"""
    return base_prompt
