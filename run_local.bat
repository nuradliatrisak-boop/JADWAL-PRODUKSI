@echo off
cd backend
pip install -r requirements.txt
start http://localhost:8000
uvicorn main:app --port 8000
