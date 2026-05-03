import google.generativeai as genai
import streamlit as st
import toml
import os
try:
    secrets = toml.load(".streamlit/secrets.toml")
    genai.configure(api_key=secrets.get("gemini_api_key"))
except:
    pass

for m in genai.list_models():
    print(m.name, m.supported_generation_methods)
