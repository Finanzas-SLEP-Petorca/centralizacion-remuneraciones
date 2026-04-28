@echo off
cd /d "%~dp0"
where pythonw >nul 2>&1
if %errorlevel%==0 (
  pythonw centralizacion_app.py
) else (
  where py >nul 2>&1
  if %errorlevel%==0 (
    py -3 centralizacion_app.py
  ) else (
    python centralizacion_app.py
  )
)
