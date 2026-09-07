@echo off
rem Crisis-plan generator — double-click entry point (SPEC AC-U1).
rem Safe defaults: DUMMY sample data, local mode (offline, no Google),
rem dummy trigger date. Output goes to private\runs\<today>\ (gitignored).
rem
rem For a real run, use the command line instead, with the real spreadsheet
rem which lives OUTSIDE this folder (see PRIVACY.md):
rem   python tools\generate.py -f C:\path\to\your-name.xlsx --date YYYY-MM-DD

cd /d "%~dp0"
python tools\generate.py -f samples\pat.sample.xlsx --date 2030-01-15 --mode local
echo.
pause
