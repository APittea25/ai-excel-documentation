# prompt.py

def build_json_summary_prompt(named_range, formulas):
    return (
        "You are an expert actuary and spreadsheet analyst.\n\n"
        "Given the following remapped formulas from an Excel named range, summarize the pattern behind the calculations in a general form.\n"
        "Each formula follows a remapped structure using notation like [1][2] to indicate row and column indices.\n\n"
        "Please return a JSON object like:\n"
        "{\n"
        '  "file_name": "MyWorkbook.xlsx",\n'
        '  "sheet_name": "Inputs",\n'
        '  "excel_range": "B2:D5",\n'
        f'  "named_range": "{named_range}",\n'
        '  "summary": "Description of what the formula does",\n'
        '  "general_formula": "for i in range(...): for j in range(...): Result[i][j] = ...",\n'
        '  "dependencies": ["OtherNamedRange1", "OtherNamedRange2"],\n'
        '  "notes": "Any caveats, limitations, or variations found"\n'
        "}\n\n"
        f"Formulas:\n{formulas[:10]}  # Sample first 10 for context\n\n"
        "Only return the JSON."
    )


def build_purpose_prompt(summaries, example=None):
    joined_descriptions = "\n".join(
        f"{k}: {v.get('summary', '')}" for k, v in summaries.items() if "summary" in v
    )
    joined_formulas = "\n".join(
        f"{k}: {v.get('general_formula', '')}" for k, v in summaries.items() if "general_formula" in v
    )

    base = (
        "You are an expert actuary and spreadsheet modeller.\n\n"
        "You are reviewing an Excel model based on the **Lee-Carter mortality framework**.\n\n"
        "The model uses named ranges and formulas structured to perform actuarial calculations.\n"
    )

    if example:
        base += f"\n--- Example Purpose Statement ---\n{example.strip()}\n"

    base += (
        "\nBelow are descriptions of how various parts of the model behave:\n\n"
        "--- Summaries ---\n"
        f"{joined_descriptions}\n\n"
        "--- Formula patterns ---\n"
        f"{joined_formulas}\n\n"
        "Using this information, write a **concise and confident purpose statement** for documentation. Your paragraph should follow this structure:\n\n"
        "1. Start with a clear sentence about what the model is designed to do (e.g. project mortality, simulate survival rates).\n"
        "2. Describe what kinds of inputs it uses (e.g. mortality trends, drift terms, random simulations).\n"
        "3. Summarize the types of outputs produced (e.g. annuity rates, survival curves).\n\n"
        "Use actuarial language. Do not say “likely”, “possibly”, or “may”. Be direct and factual."
    )
    return base


def build_input_prompt(input_name, summary_json, hint_sentence=None, example=None):
    base = (
        "You are an expert actuary and survival modeller.\n\n"
        "You are reviewing a spreadsheet model based on the Lee-Carter mortality model or a closely related framework.\n\n"
        f"You're now documenting the spreadsheet input named `{input_name}`, located in sheet `{summary_json.get('sheet_name', '')}`, "
        f"cell range `{summary_json.get('excel_range', '')}`.\n\n"
        f"Its name suggests it's related to: \"{input_name}\".\n\n"
        f"However, if the hint below provides more precise context, use that as the primary guide:\n\"{hint_sentence}\".\n\n"
        f"Here is the description of how this input is used in the model:\n\"{summary_json.get('summary', '')}\"\n\n"
        f"And here is the general formula pattern that references it:\n\"{summary_json.get('general_formula', '')}\"\n"
    )

    if example:
        base += f"\n--- Example Input Description ---\n{example.strip()}\n"

    base += (
        "\nBased on all the above, write a concise, confident description of what `{input_name}` represents and how it contributes to the model.\n\n"
        "Use actuarial language. Avoid vague expressions like “might”, “somewhat”, “typically”, or filler phrases like “plays a crucial role” or “is important”. "
        "Do not describe patterns in the data (e.g., “decreasing linearly”) unless they are explicitly mentioned.\n\n"
        "Respond with one precise sentence, or two if the second adds new technical detail or context."
    )
    return base


def build_output_prompt(output_name, summary_json, hint_sentence=None, example=None):
    base = (
        "You are an expert actuary and spreadsheet modeller.\n\n"
        "You are reviewing an Excel spreadsheet built on the **Lee-Carter mortality model** or a closely related survival modelling framework.\n\n"
        f"You're documenting the model output named `{output_name}`, located in sheet `{summary_json.get('sheet_name', '')}`, "
        f"cell range `{summary_json.get('excel_range', '')}`.\n\n"
        f"Its name suggests it primarily relates to: \"{output_name}\".\n\n"
        f"Use the following contextual hint as *additional guidance* where applicable:\n\"{hint_sentence}\"\n\n"
        f"Here is how this output behaves in the model:\n\"{summary_json.get('summary', '')}\"\n\n"
        f"And here is the formula structure used to calculate it:\n\"{summary_json.get('general_formula', '')}\"\n"
    )

    if example:
        base += f"\n--- Example Output Description ---\n{example.strip()}\n"

    base += (
        "\nBased on all the above, write a concise and confident explanation of what `{output_name}` represents and how it contributes to the model's output.\n\n"
        "Use actuarial language. Do **not** include vague expressions like “might”, “possibly”, or “likely”, and avoid filler phrases like “plays a crucial role”, "
        "“important component”, or “used to calculate”. Focus instead on what it does and how it connects to the broader modelling framework.\n\n"
        "Respond with **one precise sentence**, or two if the second adds useful technical context."
    )
    return base


def build_logic_prompt(name, summary_json, step_number, hint_sentence=None, example=None):
    base = (
        "You are an expert actuary and spreadsheet modeller.\n\n"
        "You are reviewing a **logic step** in an Excel model built on the **Lee-Carter mortality model** or a similar survival projection framework.\n\n"
        f"{hint_sentence or ''}\n\n"
        f"The step being documented is `{name}`, located in sheet `{summary_json.get('sheet_name', '')}`, "
        f"range `{summary_json.get('excel_range', '')}`. It is labelled as **Step {step_number}** in the spreadsheet.\n\n"
        f"Here is a description of this step’s behavior in the model:\n\"{summary_json.get('summary', '')}\"\n\n"
        f"And here is the generalised formula pattern:\n\"{summary_json.get('general_formula', '')}\"\n\n"
        f"This step depends on the following named ranges:\n{', '.join(summary_json.get('dependencies', [])) or 'None listed'}\n"
    )

    if example:
        base += f"\n--- Example Logic Description ---\n{example.strip()}\n"

    base += (
        "\nNow write a clear and technical description of this step, using the following exact structure:\n\n"
        "1. The purpose of this calculation step.\n"
        "2. The type of calculation it performs and what it is projecting or transforming.\n"
        "3. Its direct dependencies and how they influence or feed into this step.\n\n"
        "Respond in **three numbered sentences**, each covering one of the above points. Use actuarial language. "
        "Do **not** include filler phrases (e.g., \"this is an important step\") or speculative terms like \"might\" or \"possibly\"."
    )
    return base


def build_check_prompt(name, summary_json, hint_sentence=None, example=None):
    base = (
        "You are an expert actuary and spreadsheet modeller.\n\n"
        "You are reviewing a **validation check** in an Excel model based on the **Lee-Carter mortality framework** or a similar survival model.\n\n"
        f"{hint_sentence or ''}\n\n"
        f"The named range being reviewed is `{name}`, located in sheet `{summary_json.get('sheet_name', '')}`, "
        f"cell range `{summary_json.get('excel_range', '')}`.\n\n"
        f"Here is a summary of the logic used in this check:\n\"{summary_json.get('summary', '')}\"\n\n"
        f"And here is the general formula pattern:\n\"{summary_json.get('general_formula', '')}\"\n"
    )

    if example:
        base += f"\n--- Example Check Description ---\n{example.strip()}\n"

    base += (
        "\nWrite a clear and confident description of what this check is verifying, referencing model outputs, intermediate calculations, or assumptions where relevant.\n\n"
        "Avoid vague words like “might” or “appears to”, and do not use generic filler like “this is a check to ensure…”.\n\n"
        "Respond with one precise sentence explaining what this check validates or confirms."
    )
    return base


def build_assumptions_prompt(summaries, example=None):
    all_summaries = "\n".join(
        f"{k}: {v.get('summary', '')}" for k, v in summaries.items() if "summary" in v
    )
    all_formulas = "\n".join(
        f"{k}: {v.get('general_formula', '')}" for k, v in summaries.items() if "general_formula" in v
    )

    base = (
        "You are an expert actuary and spreadsheet modeller reviewing a workbook based on the **Lee-Carter mortality model** "
        "or a similar mortality projection framework.\n\n"
        "Below are summaries and general formulas of the spreadsheet's calculations:\n\n"
        "--- Summaries ---\n"
        f"{all_summaries}\n\n"
        "--- Formula patterns ---\n"
        f"{all_formulas}\n"
    )

    if example:
        base += f"\n--- Example Assumptions Paragraph ---\n{example.strip()}\n"

    base += (
        "\nUsing this information, write a **short, clear paragraph** that outlines:\n\n"
        "1. The key assumptions used in this spreadsheet (e.g. mortality trends, parameter stability, projection horizon).\n"
        "2. Any notable limitations or modelling constraints (e.g. fixed inputs, deterministic assumptions, lack of stress testing).\n\n"
        "Avoid vague phrases like “it might be assumed” or “possibly”. Be direct and professional."
    )
    return base
