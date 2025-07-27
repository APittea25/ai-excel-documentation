# llm_engine.py

from openai import OpenAI
import os

# You could also move these to st.secrets or config later
DEFAULT_MODEL = "gpt-4o"
DEFAULT_TEMPERATURE = 0.3

# Initialize OpenAI client
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def call_chat_model(system_msg, user_prompt, model=DEFAULT_MODEL, temperature=DEFAULT_TEMPERATURE):
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_prompt}
            ],
            temperature=temperature
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"
