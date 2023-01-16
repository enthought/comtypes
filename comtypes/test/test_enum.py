import unittest as ut

import comtypes
import comtypes.client


class Test_Enum(ut.TestCase):
    def test_enum(self):
        comtypes.client.GetModule("msvidctl.dll")
        comtypes.client.GetModule("quartz.dll")
        from comtypes.gen import MSVidCtlLib as vidlib
        from comtypes.gen import QuartzTypeLib as quartz

        # FilgraphManager has the same CLSID as FilterGraph
        filtergraph = comtypes.CoCreateInstance(
            quartz.FilgraphManager._reg_clsid_,
            vidlib.IGraphBuilder,
            comtypes.CLSCTX_INPROC_SERVER,
        )
        enum_filters = filtergraph.EnumFilters()
        # make sure enum_filters is iterable
        [_ for _ in enum_filters]


if __name__ == "__main__":
    ut.main()
