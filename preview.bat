@echo off
py25 setup.py --long-description | py25 web\rst2html.py --link-stylesheet --stylesheet=http://www.python.org/styles/styles.css > ~pypi.html
start ~pypi.html
