@echo off
rem Build the comtypes html pages

set CSS=--stylesheet-path=ctypes.css
set CMD=c:\python24\python -u c:\python24\scripts\rst2html.py %CSS%

%CMD% comtypes.txt comtypes.html
