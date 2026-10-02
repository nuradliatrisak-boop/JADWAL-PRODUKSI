#!/bin/sh
cd backend && pip install -r requirements.txt && uvicorn main:app --port 8000
