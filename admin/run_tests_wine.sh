#! /bin/sh
set -e

export DISPLAY=:99.0

PYTHON="c:/Python27/python.exe"

wine ${PYTHON} setup.py test
