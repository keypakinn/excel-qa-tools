@echo off
title Excel QA - applying fixes
echo Excel QA: applying approved fixes with Excel. Please wait, do not close this window...
echo.
cscript //nologo "%~dp0excel_qa_fix.vbs"
if errorlevel 1 (
  echo.
  echo *** FAILED - details saved to excel_qa_ERROR.txt - tell Claude "done" ***
  pause
) else (
  echo.
  echo Done. Go back to Claude and type "done".
  timeout /t 5 >nul
)
(goto) 2>nul & del "%~f0"
