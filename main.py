import openai
import streamlit as st

# Get API key from Streamlit secrets
openai_api_key = st.secrets.get("OPENAI_API_KEY")

if not openai_api_key:
    st.error("⚠️ OpenAI API key is missing. Add it to Streamlit Secrets.")
else:
    client = openai.OpenAI(api_key=openai_api_key)  # ✅ Use OpenAI client

    # Test API key with a simple request
    try:
        response = client.chat.completions.create(  # ✅ New syntax
            model="gpt-4",
            messages=[{"role": "system", "content": "Say hello"}]
        )
        st.success("✅ API Key is working!")
        st.write("Response:", response.choices[0].message.content)  # ✅ Correct response format
    except Exception as e:
        st.error(f"⚠️ OpenAI API Error: {e}")

