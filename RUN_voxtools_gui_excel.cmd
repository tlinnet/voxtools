REM Calling environment setup
call "P:\Voxmeter_python_tools\WinPython-64bit-3.6.3.0Qt5\scripts\env_for_icons.bat"

REM #####################################################
REM # The REM is telling windows it is a comment line
REM # 2018-01-25 : By Troels Schwarz-Linnet
REM #
REM #
REM # The REM is telling windows it is a comment line
REM #####################################################

REM # add PYTHONPATH path
REM # python -c "import sys;print(sys.path)"
set PYTHONPATH=P:\Voxmeter_python_tools\voxtools

REM # Run excel GUI
python -m voxtools.gui.excel

REM # Let the window pause
REM pause
set /p=Hit ENTER to continue...


REM ------------------------------------------------------
REM - Possible follow up in cmd
REM ------------------------------------------------------
REM cmd.exe /k
REM Powershell.exe -Command "& {Start-Process PowerShell.exe -ArgumentList '-ExecutionPolicy RemoteSigned -noexit -File ""P:\Voxmeter_python_tools\WinPython-64bit-3.6.3.0Qt\scripts\WinPython_PS_Prompt.ps1""'}"
