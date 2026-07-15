@echo off
setlocal
set LOGFILE=streamlit_log_%date:~-4%%date:~4,2%%date:~7,2%_%time:~0,2%%time:~3,2%%time:~6,2%.txt
set LOGFILE=%LOGFILE: =0%
echo Starting Streamlit -- console output is redirected to %LOGFILE%
echo (clicking/selecting text in THIS window can no longer freeze the app,
echo  since nothing is written to the console for a text-selection to block)
echo App will be at http://localhost:8501
streamlit run fdd_app.py > "%LOGFILE%" 2>&1
