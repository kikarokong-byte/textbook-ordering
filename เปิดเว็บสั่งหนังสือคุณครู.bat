@echo off
chcp 65001 > nul
echo Starting Teacher Ordering Web App...
.venv\Scripts\streamlit.exe run teacher_app.py
pause
