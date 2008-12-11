import os, sys
from ctypes import windll

def is_cache():
    try:
        import comtypes.gen
    except ImportError:
        return
    return comtypes.gen.__path__[0]

install_text = """\
When installing a new comtypes version, it is recommended to remove
the comtypes\gen directory and the automatically generated modules
it contains.  This directory and the modules will be regenerated
on demand.

Should the installer delete all the files in this directory?"""

deinstall_text = """\
The comtypes\gen directory contains modules that comtypes
automatically generates.

Should this directory be removed?"""

if len(sys.argv) > 1 and sys.argv[1] == "-install":
    title = "Install comtypes"
    text = install_text
else:
    title = "Remove comtypes"
    text = deinstall_text


IDYES = 6
IDNO = 7
MB_YESNO = 4
MB_ICONWARNING = 48
directory = is_cache()
if directory:
    res = windll.user32.MessageBoxA(0, text, title, MB_YESNO|MB_ICONWARNING)
    if res == IDYES:
        for f in os.listdir(directory):
            fullpath = os.path.join(directory, f)
            os.remove(fullpath)
        os.rmdir(directory)
        print("Removed directory %s" % directory)
    else:
        print("Directory %s NOT removed" % directory)
