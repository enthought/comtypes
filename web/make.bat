@echo off
rem Build the comtypes web pages
set PYTHONPATH=..

c:\python25\python -c "from comtypes.client import CreateObject; help(CreateObject('MSScriptControl.ScriptControl'))" > scriptcontrol.txt

set CSS=--stylesheet-path=ctypes.css --link-stylesheet
REM set CSS=--stylesheet-path=comtypes.css

REM set OPTIONS=--time --source-link --initial-header-level=2
set OPTIONS=--time --initial-header-level=2
REM set OPTIONS=--time --source-link

REM set CMD=c:\python24\python -u c:/python24/scripts/rst2html.py %OPTIONS% %CSS%
set CMD=c:\python25\python -u rst2html.py %OPTIONS% %CSS%

%CMD% index.rst index.html
%CMD% scriptcontrol.rst scriptcontrol.html

if not "%1" == "" start %1.html
