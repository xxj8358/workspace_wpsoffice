@echo off

rem Set encoding to UTF-8
chcp 65001 >nul

rem Run Python script
python test_copy_table.py

rem Wait for user input before exiting
pause