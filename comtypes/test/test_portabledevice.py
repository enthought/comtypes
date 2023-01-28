from ctypes import POINTER, _Pointer, pointer, cast, c_ulong, c_wchar_p
from typing import Optional
import unittest as ut

import comtypes
import comtypes.client
from comtypes import GUID

comtypes.client.GetModule("portabledeviceapi.dll")
import comtypes.gen.PortableDeviceApiLib as port_api

comtypes.client.GetModule("portabledevicetypes.dll")
import comtypes.gen.PortableDeviceTypesLib as port_types


def PropertyKey(fmtid: GUID, pid: int) -> "_Pointer[port_api._tagpropertykey]":
    propkey = port_api._tagpropertykey()
    propkey.fmtid = fmtid
    propkey.pid = c_ulong(pid)
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
        WPD_OBJECT_PROPERTIES_V1 = GUID("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C}")
        WPD_OBJECT_NAME = PropertyKey(WPD_OBJECT_PROPERTIES_V1, 4)
        WPD_OBJECT_CONTENT_TYPE = PropertyKey(WPD_OBJECT_PROPERTIES_V1, 7)
        WPD_OBJECT_SIZE = PropertyKey(WPD_OBJECT_PROPERTIES_V1, 11)
        WPD_OBJECT_ORIGINAL_FILE_NAME = PropertyKey(WPD_OBJECT_PROPERTIES_V1, 12)
        WPD_OBJECT_PARENT_ID = PropertyKey(WPD_OBJECT_PROPERTIES_V1, 3)
        folderType = GUID("{27E2E392-A111-48E0-AB0C-E17705A05F85}")
        functionalType = GUID("{99ED0160-17FF-4C44-9D98-1D7A6F941921}")

        content = self.device.Content()
        properties = content.Properties()
        key_collection = comtypes.client.CreateObject(
            port_types.PortableDeviceKeyCollection,
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
            interface=port_api.IPortableDeviceKeyCollection,
        )
        key_collection.Add(WPD_OBJECT_NAME)
        key_collection.Add(WPD_OBJECT_CONTENT_TYPE)

        def findDir(curObjectId: str, level: int) -> Optional[str]:
            """Searches the directory given by the first parameter for a sub-directory at a given level,
            and returns its object ID if it exists"""
            if level < 0:
                return None
            values = properties.GetValues(curObjectId, key_collection)
            contenttype = values.GetGuidValue(WPD_OBJECT_CONTENT_TYPE)
            is_folder = contenttype in [folderType, functionalType]
            # print(f"{values.GetStringValue(WPD_OBJECT_NAME)} ({'folder' if is_folder else 'file'})")
            if is_folder:
                if level == 0:
                    return curObjectId
                # traverse into the children
                enumobj = content.EnumObjects(c_ulong(0), curObjectId, None)
                for x in enumobj:
                    objId = findDir(x, level - 1)
                    if objId:
                        return objId
                return None  # not in this part of the tree
            else:
                return None

        # find a directory 2 levels deep, because you can't write on the top ones
        parentId = findDir("DEVICE", 2)
        if not parentId:
            self.fail("Could not find a directory two levels deep on the device")

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
            pdv, optimalTransferSizeBytes, POINTER(c_wchar_p)()
        )

        # uncomment the code below to actually write the file
        # from ctypes import c_ubyte
        # c_buf = (c_ubyte * len(buffer)).from_buffer(bytearray(buffer))
        # written = fileStream.RemoteWrite(c_buf, c_ulong(len(buffer)))
        # self.assertEqual(written, len(buffer))
        # fileStream.Commit(c_ulong(0))

        # WARNING: It seems like this does not actually create a file unless the file stream is commited,
        # though I can't promise with certainty that this is the case on all portable devices.


if __name__ == "__main__":
    ut.main()
