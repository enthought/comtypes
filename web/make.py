#!/usr/bin/env python

"""
Build the comtypes web pages
"""

"""
from comtypes.client import CreateObject
with open('scriptcontrol.txt', 'w') as sc_stream:
	ob = CreateObject('MSScriptControl.ScriptControl')
	sc_stream.write(ob.__doc__)
"""

CSS="--stylesheet-path=ctypes.css --link-stylesheet"
OPTIONS="--time --initial-header-level=2"

from docutils import core
settings = dict(
	stylesheet_path='ctypes.css',
	link_stylesheet=True,
	time=True,
	initial_header_level=2,
)
def publish(name):
	core.publish_file(source_path=name+'.rst', destination_path=name+'.html',
		settings_overrides=settings, writer_name='html')
publish('index')
publish('server')
#publish('scriptcontrol')
