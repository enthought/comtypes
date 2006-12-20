@echo off
rem Build the comtypes html pages

REM set CSS=--stylesheet-path=ctypes.css
set CSS=--stylesheet-path=comtypes.css
set OPTIONS=--time --source-link --initial-header-level=2
REM set OPTIONS=--time --source-link
set CMD=c:\python24\python -u c:\python24\scripts\rst2html.py %OPTIONS% %CSS%

%CMD% comtypes.client.txt comtypes.client.html
%CMD% com_interfaces.txt com_interfaces.html

if not "%1" == "" start %1.html
