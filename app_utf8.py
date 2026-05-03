import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io
import os
import urllib.request
from PIL import Image
import matplotlib.pyplot as plt
import numpy as np
from mpl_toolkits.mplot3d import Axes3D
import pypandoc
import ssl

# Fix for macOS SSL certificate verify failed error
try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    pass
else:
    ssl._create_default_https_context = _create_unverified_https_context


st.set_page_config(page_title="SNU Chem-Ed Studio", layout="wide", page_icon="