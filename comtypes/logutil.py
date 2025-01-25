# logutil.py
import logging
from ctypes import WinDLL
from ctypes.wintypes import LPCSTR, LPCWSTR

_kernel32 = WinDLL("kernel32")

_OutputDebugStringA = _kernel32.OutputDebugStringA
_OutputDebugStringA.argtypes = [LPCSTR]
_OutputDebugStringA.restype = None

_OutputDebugStringW = _kernel32.OutputDebugStringW
_OutputDebugStringW.argtypes = [LPCWSTR]
_OutputDebugStringW.restype = None


class NTDebugHandler(logging.Handler):
    def emit(
        self,
        record,
        writeA=_OutputDebugStringA,
        writeW=_OutputDebugStringW,
    ):
        text = self.format(record)
        if isinstance(text, str):
            writeA(text + "\n")
        else:
            writeW(text + "\n")


logging.NTDebugHandler = NTDebugHandler


def setup_logging(*pathnames):
    import configparser

    parser = configparser.ConfigParser()
    parser.optionxform = str  # use case sensitive option names!

    parser.read(pathnames)

    DEFAULTS = {
        "handler": "StreamHandler()",
        "format": "%(levelname)s:%(name)s:%(message)s",
        "level": "WARNING",
    }

    def get(section, option):
        try:
            return parser.get(section, option, True)
        except (configparser.NoOptionError, configparser.NoSectionError):
            return DEFAULTS[option]

    levelname = get("logging", "level")
    format = get("logging", "format")
    handlerclass = get("logging", "handler")

    # convert level name to level value
    level = getattr(logging, levelname)
    # create the handler instance
    handler = eval(handlerclass, vars(logging))
    formatter = logging.Formatter(format)
    handler.setFormatter(formatter)
    logging.root.addHandler(handler)
    logging.root.setLevel(level)

    try:
        for name, value in parser.items("logging.levels", True):
            value = getattr(logging, value)
            logging.getLogger(name).setLevel(value)
    except configparser.NoSectionError:
        pass
