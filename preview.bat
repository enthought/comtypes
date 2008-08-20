@echo off
REM
REM Convert setup.py's long description to HTML and show it.
REM
py25 setup.py --long-description | py25 web\rst2html.py --link-stylesheet --stylesheet=http://www.python.org/styles/styles.css > ~pypi.html
start ~pypi.html
del ~pypi.html
