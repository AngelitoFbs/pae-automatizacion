@echo off
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
set STREAMLIT_SERVER_HEADLESS=true
"C:\Users\TERMINAL141\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" -m streamlit run "C:\Users\TERMINAL141\Documents\Default Project\pae_automatizacion\app.py" --server.port 8501 --server.address 0.0.0.0