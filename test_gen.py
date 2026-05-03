import sys
import google.generativeai as genai
import streamlit as st
from app import robust_generate_content, get_safe_gemini_model, load_local_academic_db

# Dummy session state
st.session_state = {"gemini_api_key": ""}

import os
import json
try:
    with open(".streamlit/secrets.toml", "r") as f:
        # crude parse
        for line in f:
            if "gemini_api_key" in line:
                key = line.split("=")[1].strip().strip('"').strip("'")
                st.session_state["gemini_api_key"] = key
                break
except Exception as e:
    print(f"No secrets found: {e}")

prompt = "Hello"
print("Generating...")
res = robust_generate_content(prompt, use_grounding=True)
print("Result:")
print(res[:100] if res else "Failed")
