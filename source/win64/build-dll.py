import subprocess

from jaraco.develop import python
from jaraco.develop import vstudio

vs = vstudio.VisualStudio.find()
env = vs.get_vcvars_env()
msbuild = python.find_in_path('msbuild.exe', env['Path'])
cmd = [msbuild, 'AvmcIfc.sln', '/p:Configuration=Debug',
	'/p:Platform=x64']
subprocess.check_call(cmd, env=env)
