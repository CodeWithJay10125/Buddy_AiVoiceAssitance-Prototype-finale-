@echo off
title BUDDY AI - Setup
color 0B

echo.
echo  ██████╗ ██╗   ██╗██████╗ ██████╗ ██╗   ██╗
echo  ██╔══██╗██║   ██║██╔══██╗██╔══██╗╚██╗ ██╔╝
echo  ██████╔╝██║   ██║██║  ██║██║  ██║ ╚████╔╝ 
echo  ██╔══██╗██║   ██║██║  ██║██║  ██║  ╚██╔╝  
echo  ██████╔╝╚██████╔╝██████╔╝██████╔╝   ██║   
echo  ╚═════╝  ╚═════╝ ╚═════╝ ╚═════╝    ╚═╝   
echo.
echo  BUDDY AI Assistant - Dependency Setup
echo  =======================================
echo.

echo [1/6] Upgrading pip...
python -m pip install --upgrade pip --quiet

echo [2/6] Installing psutil (system stats)...
python -m pip install psutil --quiet

echo [3/6] Installing pyttsx3 (text-to-speech)...
python -m pip install pyttsx3 --quiet

echo [4/6] Installing python-docx (Word documents)...
python -m pip install python-docx --quiet

echo [5/6] Installing pywin32 (Windows OS control)...
python -m pip install pywin32 --quiet

echo [6/6] Installing pyautogui (screen automation)...
python -m pip install pyautogui --quiet

echo.
echo  ✓ All dependencies installed!
echo.
echo  IMPORTANT: Edit buddy.py APP_PATHS section if your
echo  app install locations differ from defaults.
echo.
echo  Press any key to launch BUDDY...
pause >nul

python buddy.py
