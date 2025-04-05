# from ctypes import POINTER, Structure, c_int32, c_uint, c_uint32
# from enum import IntFlag

# from comtypes import GUID, STDMETHOD, IUnknown


# class AdviseFlag(IntFlag):
#     """ADVF"""

#     ADVF_NODATA = 1
#     ADVF_PRIMEFIRST = 2
#     ADVF_ONLYONCE = 4
#     ADVF_DATAONSTOP = 64
#     ADVFCACHE_NOHANDLER = 8
#     ADVFCACHE_FORCEBUILTIN = 16
#     ADVFCACHE_ONSAVE = 32


# class StatData(Structure):
#     """STATDATA"""

#     _fields_ = (
#         ("formatetc", FormatEtc),
#         ("__advflags", c_uint32),
#         ("advise_sink", POINTER(IUnknown)),  # TODO: IAdviseSink
#         ("connection", c_uint32),
#     )

#     @property
#     def advflags(self) -> AdviseFlag:
#         return AdviseFlag(self.__advflags)


# class IEnumSTATDATA(IUnknown):
#     _iid_ = GUID("{00000105-0000-0000-C000-000000000046}")


# IEnumSTATDATA._methods_ = [
#     STDMETHOD(c_int32, "Next", (c_uint, POINTER(StatData), POINTER(c_uint32))),
#     STDMETHOD(c_int32, "Skip", (c_uint32,)),
#     STDMETHOD(c_int32, "Reset", ()),
#     STDMETHOD(c_int32, "Clone", (POINTER(POINTER(IEnumSTATDATA)),)),
# ]
