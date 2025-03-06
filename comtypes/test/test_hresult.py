import unittest as ut

import comtypes.hresult

ERROR_OUTOFMEMORY = 14  # 0xE
ERROR_INVALID_PARAMETER = 87  # 0x57
RPC_S_SERVER_UNAVAILABLE = 1722  # 0x6BA


class Test_HRESULT_FROM_WIN32(ut.TestCase):
    def test(self):
        for w32, hr in [
            (ERROR_OUTOFMEMORY, comtypes.hresult.E_OUTOFMEMORY),
            (ERROR_INVALID_PARAMETER, comtypes.hresult.E_INVALIDARG),
            (RPC_S_SERVER_UNAVAILABLE, comtypes.hresult.RPC_S_SERVER_UNAVAILABLE),
        ]:
            with self.subTest(w32=w32, hr=hr):
                self.assertEqual(comtypes.hresult.HRESULT_FROM_WIN32(w32), hr)
