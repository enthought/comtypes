import sys
from comtypes.client import GetModule
import winreg

# this makes it easy to load registered type libraries. It uses the .NET
# names that are stored in the registry to function. Not all .NET names
# can be used because there isn't typelib information available for all
# of them.
# Here are a few examples
#
# if you have an Nvidia video card
# comtypes.typelib.DisplayServer.Config
#
# Windows Media Player
# comtypes.typelib.MediaPlayer.MediaPlayer


def _get_reg_value(path, key, wow6432=False):
    d = _read_reg_values(path, wow6432)
    if key in d:
        return d[key]

    return ''


def _read_reg_keys(key, wow6432=False):
    if isinstance(key, tuple):
        root = key[0]
        key = key[1]
    else:
        root = winreg.HKEY_CLASSES_ROOT

    try:
        if wow6432:
            handle = winreg.OpenKey(
                root,
                key,
                0,
                winreg.KEY_READ | winreg.KEY_WOW64_32KEY
            )
        else:
            handle = winreg.OpenKeyEx(root, key)
    except winreg.error:
        return []
    res = []

    for i in range(winreg.QueryInfoKey(handle)[0]):
        res += [winreg.EnumKey(handle, i)]

    winreg.CloseKey(handle)
    return res


def _read_reg_values(key, wow6432=False):
    if isinstance(key, tuple):
        root = key[0]
        key = key[1]
    else:
        root = winreg.HKEY_CLASSES_ROOT

    try:
        if wow6432:
            handle = winreg.OpenKey(
                root,
                key,
                0,
                winreg.KEY_READ | winreg.KEY_WOW64_32KEY
            )
        else:
            handle = winreg.OpenKeyEx(root, key)
    except winreg.error:
        return {}
    res = {}
    for i in range(winreg.QueryInfoKey(handle)[1]):
        name, value, _ = winreg.EnumValue(handle, i)
        res[_convert_mbcs(name)] = _convert_mbcs(value)

    winreg.CloseKey(handle)

    return res


def _convert_mbcs(s):
    dec = getattr(s, "decode", None)
    if dec is not None:
        try:
            s = dec("mbcs")
        except UnicodeError:
            pass
    return s


class _ModuleWrapper(object):

    def __init__(self, mod):

        self.__doc__ = mod.__doc__
        self.__file__ = mod.__file__
        self.__loader__ = mod.__loader__
        self.__name__ = mod.__name__
        self.__package__ = mod.__package__
        # self.__path__ = mod.__path__
        self.__spec__ = mod.__spec__
        self.__original_module__ = mod

        sys.modules[mod.__name__] = self

    def __getattr__(self, item):
        if item in self.__dict__:
            return self.__dict__[item]

        if hasattr(self.__original_module__, item):
            return getattr(self.__original_module__, item)

        if hasattr(self, item.lower()):
            return getattr(self, item.lower())

        raise AttributeError(item)


class _TypeLibLoader(object):

    def __init__(self, keys, name=None):
        self.__loaded = {}
        self.__names = {}
        self.__name = name
        self.__keys = keys

    def __getattr__(self, item):
        if item in self.__loaded:
            return self.__loaded[item]

        if self.__name is None:
            name = item
        else:
            name = self.__name + '.' + item

        if name in self.__keys:
            values = _read_reg_values(name)
            doc = values.get('', None)

            values = _read_reg_values(name + '\\' + 'CurVer')

            if '' in values:
                name = values['']

            values = _read_reg_values(name + '\\' + 'CLSID')

            clsid = values.get('', None)

            if clsid is None:
                raise AttributeError(item)

            values = _read_reg_values('CLSID\\{0}\\Version'.format(clsid))

            version = values.get('', None)
            values = _read_reg_values('CLSID\\{0}\\TypeLib'.format(clsid))
            typelib_clsid = values.get('', None)

            if not typelib_clsid:
                raise AttributeError(item)

            regkey = 'TypeLib\\{0}'.format(typelib_clsid)

            if version is None:
                for key in _read_reg_keys(regkey):
                    try:
                        int(key)
                    except:  # NOQA
                        try:
                            float(key)
                        except:  # NOQA
                            continue
                    if version is None:
                        version = key
                    elif key > version:
                        version = key

            if version is None:
                values = _read_reg_values(
                    'CLSID\\{0}\\InprocServer32'.format(clsid))
                path = values.get('', None)
                if path is None:
                    raise AttributeError(item)

            else:
                regkey += '\\' + version
                for key in _read_reg_keys(regkey):
                    try:
                        int(key)
                    except:  # NOQA
                        try:
                            float(key)
                        except:  # NOQA
                            continue

                    keys = _read_reg_keys(regkey + '\\' + key)

                    if sys.maxsize > 2 ** 32 and 'win64' in keys:
                        path = _read_reg_values(
                            '{0}\\{1}\\win64'.format(regkey, key)
                        )['']
                    elif 'win32' in keys:
                        path = _read_reg_values(
                            '{0}\\{1}\\win32'.format(regkey, key)
                        )['']
                    else:
                        continue
                    break
                else:
                    raise AttributeError(item)

            try:
                mod = GetModule(path.replace('\\\\', '\\'))
            except:
                print(name)
                raise

            mod = _ModuleWrapper(mod)
            items = [n for n in self.__keys if n.startswith(name)]
            item_names = {}
            for itm in items:
                if itm == name:
                    continue

                itm_name = itm.replace(name + '.', '').split('.', 1)[0]

                if itm_name not in item_names:
                    item_names[itm_name] = []

                item_names[itm_name].append(itm)

            for attr_name, items in item_names.items():
                setattr(
                    mod,
                    attr_name,
                    _TypeLibLoader(items, name + '.' + attr_name)
                )

            setattr(mod, '__doc__', doc)

            self.__loaded[item] = mod

            return mod

        items = [n for n in self.__keys if n.startswith(name)]

        if items:
            self.__loaded[item] = res = _TypeLibLoader(items, name)
            return res

        raise AttributeError(item)


typelib = _TypeLibLoader(_read_reg_keys(''))
