import openai
import streamlit as st

# Get API key from Streamlit secrets
openai_api_key = st.secrets.get("OPENAI_API_KEY")

if not openai_api_key:
    st.error("⚠️ OpenAI API key is missing. Add it to Streamlit Secrets.")
else:
    openai.api_key = openai_api_key

    # Test API key with a simple request
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "system", "content": "Say hello"}]
        )
        st.success("✅ API Key is working!")
    except Exception as e:
        st.error(f"⚠️ OpenAI API Error: {e}")

