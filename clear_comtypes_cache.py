import os
import sys
import shutil

def get_next_cache_dir():
    try:
        print("import comtypes.gen")
        # change working directory to avoid import from local folder
        # during installation process
        os.chdir(os.path.dirname(sys.executable))
        import comtypes
        import comtypes.client
        print("DONE")
        return comtypes.client._code_cache._find_gen_dir()
    except ImportError:
        return None


def _remove(directory):
    shutil.rmtree(directory)
    print('Removed directory "%s"' % directory)


def remove_directory(directory, silent):
    if directory:
        if silent:
            _remove(directory)
        else:
            try:
                confirm = raw_input('Remove directory "%s" ? (y/n)' % directory)
            except NameError:
                confirm = input('Remove cache directory "%s" ? (y/n)' % directory)
            if confirm.lower() == 'y':
                _remove(directory)
            else:
                print('Directory "%s" NOT removed' % directory)
                return False
    return True


if len(sys.argv) > 1 and "-y" in sys.argv[1:]:
    silent = True
else:
    silent = False


# First iteration may get folder with restricted rights.
# Second iteration always gets temp cache folder (writable for all).
directory = get_next_cache_dir()
removed = remove_directory(directory, silent)

if removed:
    directory = get_next_cache_dir()

    # do not request confirmation if already confirmed and removed
    remove_directory(directory, silent=removed)
