 # so we don't crash when the generated files don't exist
from __future__ import annotations

from ctypes import (POINTER, _Pointer, pointer, cast,
                    c_uint32, c_uint16, c_int8, c_ulong, c_wchar_p)
from typing import Optional
import unittest as ut

import comtypes
import comtypes.client

comtypes.client.GetModule("portabledeviceapi.dll")
import comtypes.gen.PortableDeviceApiLib as port_api

comtypes.client.GetModule("portabledevicetypes.dll")
import comtypes.gen.PortableDeviceTypesLib as port_types


def newGuid(*args: int) -> comtypes.GUID:
    guid = comtypes.GUID()
    guid.Data1 = c_uint32(args[0])
    guid.Data2 = c_uint16(args[1])
    guid.Data3 = c_uint16(args[2])
    for i in range(8):
        guid.Data4[i] = c_int8(args[3 + i])
    return guid


def PropertyKey(*args: int) -> _Pointer[port_api._tagpropertykey]:
    assert len(args) == 12
    assert all(isinstance(x, int) for x in args)
    propkey = port_api._tagpropertykey()
    propkey.fmtid = newGuid(*args[0:11])
    propkey.pid = c_ulong(args[11])
    return pointer(propkey)


class Test_IPortableDevice(ut.TestCase):
    # To avoid damaging or changing the environment, do not CREATE, DELETE or UPDATE!
    # Do READ only!
    def setUp(self):
        info = comtypes.client.CreateObject(
            port_types.PortableDeviceValues().IPersist_GetClassID(),
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_types.IPortableDeviceValues,
        )
        mng = comtypes.client.CreateObject(
            port_api.PortableDeviceManager().IPersist_GetClassID(),
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_api.IPortableDeviceManager,
        )
        p_id_cnt = pointer(c_ulong())
        mng.GetDevices(POINTER(c_wchar_p)(), p_id_cnt)
        if p_id_cnt.contents.value == 0:
            self.skipTest("There is no portable device in the environment.")
        dev_ids = (c_wchar_p * p_id_cnt.contents.value)()
        mng.GetDevices(cast(dev_ids, POINTER(c_wchar_p)), p_id_cnt)
        self.device = comtypes.client.CreateObject(
            port_api.PortableDevice().IPersist_GetClassID(),
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_api.IPortableDevice,
        )
        self.device.Open(list(dev_ids)[0], info)

    def test_EnumObjects(self):
        WPD_OBJECT_NAME = PropertyKey(
            0xEF6B490D, 0x5CD8, 0x437A, 0xAF, 0xFC, 0xDA, 0x8B, 0x60, 0xEE, 0x4A, 0x3C, 4)
        WPD_OBJECT_CONTENT_TYPE = PropertyKey(
            0xEF6B490D, 0x5CD8, 0x437A, 0xAF, 0xFC, 0xDA, 0x8B, 0x60, 0xEE, 0x4A, 0x3C, 7)
        WPD_OBJECT_SIZE = PropertyKey(
            0xEF6B490D, 0x5CD8, 0x437A, 0xAF, 0xFC, 0xDA, 0x8B, 0x60, 0xEE, 0x4A, 0x3C, 11)
        WPD_OBJECT_ORIGINAL_FILE_NAME = PropertyKey(
            0xEF6B490D, 0x5CD8, 0x437A, 0xAF, 0xFC, 0xDA, 0x8B, 0x60, 0xEE, 0x4A, 0x3C, 12)
        WPD_OBJECT_PARENT_ID = PropertyKey(
            0xEF6B490D, 0x5CD8, 0x437A, 0xAF, 0xFC, 0xDA, 0x8B, 0x60, 0xEE, 0x4A, 0x3C, 3)
        folderType = newGuid(0x27E2E392, 0xA111, 0x48E0,
                             0xAB, 0x0C, 0xE1, 0x77, 0x05, 0xA0, 0x5F, 0x85)
        functionalType = newGuid(
            0x99ED0160, 0x17FF, 0x4C44, 0x9D, 0x98, 0x1D, 0x7A, 0x6F, 0x94, 0x19, 0x21)

        content = self.device.Content()
        properties = content.Properties()
        propertiesToRead = comtypes.client.CreateObject(
            port_types.PortableDeviceKeyCollection,
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_api.IPortableDeviceKeyCollection,
        )
        propertiesToRead.Add(WPD_OBJECT_NAME)
        propertiesToRead.Add(WPD_OBJECT_CONTENT_TYPE)

        def findDir(curObjectID, level: int) -> Optional[str]:
            # print(contentID)
            if level < 0:
                return None
            values = properties.GetValues(curObjectID, propertiesToRead)
            contenttype = values.GetGuidValue(WPD_OBJECT_CONTENT_TYPE)
            is_folder = contenttype in [folderType, functionalType]
            print(
                f"{values.GetStringValue(WPD_OBJECT_NAME)} ({'folder' if is_folder else 'file'})")
            if is_folder:
                if level == 0:
                    return curObjectID
                # traverse into the children
                enumobj = content.EnumObjects(c_ulong(0), curObjectID, None)
                for x in enumobj:
                    objId = findDir(x, level-1)
                    if objId:
                        return objId
                return None  # not in this part of the tree
            else:
                return None

        # find a directory 2 levels deep, because you can't write on the top ones
        parentId = findDir("DEVICE", 2)
        if not parentId:
            self.fail("Could not find the parent path on the device")

        buffer = "Text file content".encode("utf-8")
        pdv = comtypes.client.CreateObject(
            port_types.PortableDeviceValues,
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_api.IPortableDeviceValues,
        )
        pdv.SetStringValue(WPD_OBJECT_PARENT_ID, parentId)
        pdv.SetUnsignedLargeIntegerValue(WPD_OBJECT_SIZE, len(buffer))
        pdv.SetStringValue(WPD_OBJECT_ORIGINAL_FILE_NAME, "testfile.txt")
        pdv.SetStringValue(WPD_OBJECT_NAME, "testfile.txt")

        optimalTransferSizeBytes = pointer(c_ulong(0))
        (
            fileStream,
            optimalTransferSizeBytes,
            _,
        ) = content.CreateObjectWithPropertiesAndData(
            pdv,
            optimalTransferSizeBytes,
            POINTER(c_wchar_p)(),
        )
        # here: optional calls to fileStream.RemoteWrite(), fileStream.Commit()

        # WARNING: It seems like this does not actually create a file unless the file stream is commited,
        # though I can't promise with certainty that this is the case on all portable devices.


if __name__ == "__main__":
    ut.main()
