import comtypes


def name_wrapper_module(tlib):
    """Determine the name of a typelib wrapper module"""
    libattr = tlib.GetLibAttr()
    modname = "_%s_%s_%s_%s" % (
        str(libattr.guid)[1:-1].replace("-", "_"),
        libattr.lcid,
        libattr.wMajorVerNum,
        libattr.wMinorVerNum,
    )
    return "comtypes.gen.%s" % modname


def name_friendly_module(tlib):
    """Determine the friendly-name of a typelib module.
    If cannot get friendly-name from typelib, returns `None`.
    """
    try:
        modulename = tlib.GetDocumentation(-1)[0]
    except comtypes.COMError:
        return
    return "comtypes.gen.%s" % modulename
