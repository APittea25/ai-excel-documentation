import streamlit as st
import openai

# Get OpenAI API Key from Streamlit Secrets
openai_api_key = st.secrets.get("OPENAI_API_KEY")

if not openai_api_key:
    st.error("⚠️ OpenAI API key is missing. Add it to Streamlit Secrets.")
else:
    client = openai.OpenAI(api_key=openai_api_key)

    # Test OpenAI connection
    try:
       response = client.chat.completions.create(
    model="gpt-3.5-turbo",  # 🔄 Use a cheaper model
    messages=[{"role": "system", "content": "Say hello"}]
)
        )
        st.success("✅ OpenAI API Key is working!")
        st.write("Response:", response.choices[0].message.content)
    except Exception as e:
        st.error(f"⚠️ OpenAI API Error: {e}")

