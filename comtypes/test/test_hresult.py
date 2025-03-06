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
            (0, -2147024896),
            (1, -2147024895),
            (-1, -1),
            (0xFFFF - 3, -2146959364),
            (0xFFFF - 2, -2146959363),
            (0xFFFF - 1, -2146959362),
            (0xFFFF + 0, -2146959361),
            (0xFFFF + 1, -2147024896),
            (0xFFFF + 2, -2147024895),
            (0xFFFF + 3, -2147024894),
        ]:
            with self.subTest(w32=w32, hr=hr):
                self.assertEqual(comtypes.hresult.HRESULT_FROM_WIN32(w32), hr)
