@echo off
title Auto Note to Word App (Windows)
echo ==============================================
echo Installing Python requirements...
echo ==============================================
pip install -r requirements.txt

echo.
echo ==============================================
echo Starting the Streamlit App...
echo ==============================================
echo A browser window should open automatically.
streamlit run app.py
pause
