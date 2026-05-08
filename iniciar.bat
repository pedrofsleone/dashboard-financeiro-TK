@echo off
start "" "http://localhost:8501"
python -m streamlit run "%~dp0app.py" --browser.gatherUsageStats false --server.port 8501
pause
