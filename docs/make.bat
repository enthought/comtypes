@echo off
rem Build the comtypes html pages

set CSS=--stylesheet-path=ctypes.css
REM set OPTIONS=--time --source-link --initial-header-level=2
set OPTIONS=--time --source-link
set CMD=c:\python24\python -u c:\python24\scripts\rst2html.py %OPTIONS% %CSS%

%CMD% comtypes.client.txt comtypes.client.html
%CMD% com_interfaces.txt com_interfaces.html

REM start com_interfaces.html
REM start comtypes.client.html