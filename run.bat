set original_dir=%CD%
set venv_root_dir=%~dp0venv
cd %venv_root_dir%

call %venv_root_dir%\Scripts\activate.bat
python %original_dir%\speller\start.py
call %venv_root_dir%\Scripts\deactivate.bat

cd %original_dir%
pause
exit /B 1