# prompt.py

def build_purpose_prompt(summaries, hint_sentence):
    joined_descriptions = "\n".join(
        f"{k}: {v.get('summary', '')}" for k, v in summaries.items() if "summary" in v
    )
    joined_formulas = "\n".join(
        f"{k}: {v.get('general_formula', '')}" for k, v in summaries.items() if "general_formula" in v
    )

    return f"""You are an expert actuary and spreadsheet modeller.

You are reviewing an Excel model based on the **Lee-Carter mortality framework**.

{hint_sentence}

The model uses named ranges and formulas structured to perform actuarial calculations.

Below are descriptions of how various parts of the model behave:

--- Summaries ---
{joined_descriptions}

--- Formula patterns ---
{joined_formulas}

Using this information, write a **concise and confident purpose statement** for documentation. Your paragraph should follow this structure:

1. Start with a clear sentence about what the model is designed to do (e.g. project mortality, simulate survival rates).
2. Describe what kinds of inputs it uses (e.g. mortality trends, drift terms, random simulations).
3. Summarize the types of outputs produced (e.g. annuity rates, survival curves).
4. Close with a sentence explaining what this model is useful for — pricing, forecasting, risk management, etc.

Use actuarial language. Do not say “likely”, “possibly”, or “may”. Be direct and factual.
"""


def build_input_prompt(input_name, summary_json, hint_sentence):
    return f"""You are an expert actuary and survival modeller.

You are reviewing a spreadsheet model based on the Lee-Carter mortality model or a closely related framework.

{hint_sentence}

You're now documenting the spreadsheet input named `{input_name}`, located in sheet `{summary_json.get("sheet_name", "")}`, cell range `{summary_json.get("excel_range", "")}`.

Its name suggests it's related to: "{input_name}"

Here is the description of how this input is used in the model:
"{summary_json.get("summary", "")}"

And here is the general formula pattern that references it:
"{summary_json.get("general_formula", "")}"

Based on all the above, write a concise, confident description of what `{input_name}` represents and how it contributes to the model.

Use actuarial language. Avoid vague expressions like “might”, “somewhat”, “typically”, or filler phrases like “plays a crucial role” or “is important”. Do not describe patterns in the data (e.g., “decreasing linearly”) unless they are explicitly mentioned.

Respond with one precise sentence, or two if the second adds new technical detail or context.
"""


def build_output_prompt(name, summary_json, hint_sentence):
    return f"""You are an expert actuary and spreadsheet modeller.

You are reviewing an Excel spreadsheet built on the **Lee-Carter mortality model** or a closely related survival modelling framework.

{hint_sentence}

You are documenting the model output named `{name}`, located in sheet `{summary_json.get("sheet_name", "")}`, cell range `{summary_json.get("excel_range", "")}`.

Here is how this output behaves in the model:
"{summary_json.get("summary", "")}"

And here is the formula structure used to calculate it:
"{summary_json.get("general_formula", "")}"

Based on this, write a concise and confident explanation of what `{name}` represents and how it contributes to the model's output.

Use actuarial language. Do **not** include vague expressions like “might”, “possibly”, or “likely”, and avoid filler phrases like “plays a crucial role”, “important component”, or “used to calculate”. Focus instead on what it does and how it connects to the broader modelling framework.

Respond with **one precise sentence**, or two if the second adds useful technical context.
"""


def build_logic_prompt(name, summary_json, step_number, hint_sentence):
    return f"""You are an expert actuary and spreadsheet modeller.

You are reviewing a calculation step in an Excel model built on the **Lee-Carter mortality model** or a similar survival framework.

{hint_sentence}

The step being documented is `{name}`, located in sheet `{summary_json.get("sheet_name", "")}`, range `{summary_json.get("excel_range", "")}`. It is labelled as step {step_number} in the spreadsheet.

Below is a general description of the calculation:
"{summary_json.get("summary", "")}"

And here is the abstracted formula pattern:
"{summary_json.get("general_formula", "")}"

This step depends directly on the following named ranges:
{', '.join(summary_json.get("dependencies", []))}

Write a concise explanation that covers:

1. The purpose of this calculation step.
2. The type of calculation it performs and what is being projected or transformed.
3. Its direct dependencies — inputs or other calculations — and how they flow into this step.

Use confident actuarial language. Avoid generic phrases like “important step”, “plays a role”, “typically used for”, etc. Do not speculate.

Respond with 1–2 clear sentences.
"""


def build_check_prompt(name, summary_json, hint_sentence):
    return f"""You are an expert actuary and spreadsheet modeller.

You are reviewing a **validation check** in an Excel model based on the **Lee-Carter mortality framework** or a similar survival model.

{hint_sentence}

The named range being reviewed is `{name}`, located in sheet `{summary_json.get("sheet_name", "")}`, cell range `{summary_json.get("excel_range", "")}`.

Here is a summary of the logic used in this check:
"{summary_json.get("summary", "")}"

And here is the general formula pattern:
"{summary_json.get("general_formula", "")}"

Write a clear and confident description of what this check is verifying, referencing model outputs, intermediate calculations, or assumptions where relevant.

Avoid vague words like “might” or “appears to”, and do not use generic filler like “this is a check to ensure…”.

Respond with one precise sentence explaining what this check validates or confirms.
"""


def build_assumptions_prompt(summaries, hint_sentence):
    all_summaries = "\n".join(
        f"{k}: {v.get('summary', '')}" for k, v in summaries.items() if "summary" in v
    )
    all_formulas = "\n".join(
        f"{k}: {v.get('general_formula', '')}" for k, v in summaries.items() if "general_formula" in v
    )

    return f"""You are an expert actuary and spreadsheet modeller reviewing a workbook based on the **Lee-Carter mortality model** or a similar mortality projection framework.

{hint_sentence}

Below are summaries and general formulas of the spreadsheet's calculations:

--- Summaries ---
{all_summaries}

--- Formula patterns ---
{all_formulas}

Using this information, write a **short, clear paragraph** that outlines:

1. The key assumptions used in this spreadsheet (e.g. mortality trends, parameter stability, projection horizon).
2. Any notable limitations or modelling constraints (e.g. fixed inputs, deterministic assumptions, lack of stress testing).

Avoid vague phrases like “it might be assumed” or “possibly”. Be direct and professional.
"""
