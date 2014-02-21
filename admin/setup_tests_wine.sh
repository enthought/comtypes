#! /bin/sh
set -e

export DISPLAY=:99.0

PYTHON="c:/Python27/python.exe"
EASY_INSTALL="c:/Python27/Scripts/easy_install.exe"
PIP="c:/Python27/Scripts/pip.exe"

wget http://www.python.org/ftp/python/2.7.6/python-2.7.6.msi
wine msiexec /i python-2.7.6.msi /qn

wget https://pypi.python.org/packages/source/s/setuptools/setuptools-2.2.tar.gz
tar xf setuptools-2.2.tar.gz
(cd setuptools-2.2 && wine ${PYTHON} setup.py install)

wine ${EASY_INSTALL} coverage
