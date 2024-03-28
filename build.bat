python3.10 --version >nul 2>&1 && ( SET python3.10=True ) || ( SET python3.10=False )

if %python3.10%==True (python3.10 -m pip install virtualenv) ELSE (python3 -m pip install virtualenv)
if %python3.10%==True (python3.10 -m venv ./env) ELSE (python3 -m venv ./env)
if %python3.10%==True (python3.10 ./src/get-tcl.py) ELSE (python3 ./src/get-tcl.py)


.\env\Scripts\python -m pip install --upgrade pyinstaller
.\env\Scripts\python -m pip install --upgrade cx_Freeze
.\env\Scripts\python -m pip install -r requirements.txt
@REM call build-pyinstaller.bat
call build-cx-freeze.bat