#!/bin/bash
echo "=============================================="
echo "Installing Python requirements for Mac..."
echo "=============================================="
cd "$(dirname "$0")"
pip3 install -r requirements.txt

echo ""
echo "=============================================="
echo "Starting the Streamlit App..."
echo "=============================================="
echo "A browser window should open automatically."
python3 -m streamlit run app.py
