@echo off
chcp 65001 >nul
set OLLAMA_NUM_PARALLEL=4
python -m pip install -r requirements.txt
python translate_batch.py
pause