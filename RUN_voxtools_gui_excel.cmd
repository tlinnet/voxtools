REM Calling environment setup
call "P:\Voxmeter_python_tools\WinPython-64bit-3.6.3.0Qt5\scripts\env_for_icons.bat"

REM ################################################################################
REM # Copyright (C) Voxmeter A/S - All Rights Reserved
REM #
REM # Voxmeter A/S
REM # Borgergade 6, 4.
REM # 1300 Copenhagen K
REM # Denmark
REM #
REM # Written by Troels Schwarz-Linnet <tsl@voxmeter.dk>, 2018
REM # 
REM # Unauthorized copying of this file, via any medium is strictly prohibited.
REM #
REM # Any use of this code is strictly unauthorized without the written consent
REM # by Voxmeter A/S. This code is proprietary of Voxmeter A/S.
REM # 
REM ################################################################################

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
