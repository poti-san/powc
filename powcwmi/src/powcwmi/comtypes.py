from ctypes import POINTER, Structure, c_byte, c_int32, c_uint32, c_uint64, c_wchar_p

from comtypes import BSTR, GUID, STDMETHOD, IUnknown
from powc.safearray import SafeArrayPtr
from powc.variant import Variant


class IWbemClassObject(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{dc12a681-737f-11cf-884d-00aa004b2e24}")


class IWbemObjectAccess(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{49353c9a-516b-11d1-aea6-00c04fb68820}")


class IWbemQualifierSet(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{dc12a680-737f-11cf-884d-00aa004b2e24}")


class IWbemServices(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{9556dc99-828c-11cf-a37e-00aa003240c7}")


class IWbemLocator(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{dc12a687-737f-11cf-884d-00aa004b2e24}")


class IWbemObjectSink(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{7c857801-7381-11cf-884d-00aa004b2e24}")


class IEnumWbemClassObject(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{027947e1-d731-11ce-a357-000000000001}")


class IWbemCallResult(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{44aca675-e8fc-11d0-a07c-00c04fb68820}")


class IWbemContext(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{44aca674-e8fc-11d0-a07c-00c04fb68820}")


class WbemCompileStatusInfo(Structure):
    """WBEM_COMPILE_STATUS_INFO"""

    __slots__ = ()
    _fields_ = (
        ("phaseerror", c_int32),
        ("hres", c_int32),
        ("objectnum", c_int32),
        ("firstline", c_int32),
        ("lastline", c_int32),
        ("outflags", c_uint32),
    )


#     MIDL_INTERFACE("1cfaba8c-1523-11d1-ad79-00c04fd8fdff")
#     IUnsecuredApartment : public IUnknown
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE CreateObjectStub(
#             /* [in] */ __RPC__in_opt IUnknown *pObject,
#             /* [out] */ __RPC__deref_out_opt IUnknown **ppStub) = 0;

#     };

#     MIDL_INTERFACE("31739d04-3471-4cf4-9a7c-57a44ae71956")
#     IWbemUnsecuredApartment : public IUnsecuredApartment
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE CreateSinkStub(
#             /* [in] */ __RPC__in_opt IWbemObjectSink *pSink,
#             /* [in] */ DWORD dwFlags,
#             /* [unique][in] */ __RPC__in_opt LPCWSTR wszReserved,
#             /* [out] */ __RPC__deref_out_opt IWbemObjectSink **ppStub) = 0;

#     };

#     MIDL_INTERFACE("eb87e1bc-3233-11d2-aec9-00c04fb68820")
#     IWbemStatusCodeText : public IUnknown
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE GetErrorCodeText(
#             /* [in] */ HRESULT hRes,
#             /* [in] */ LCID LocaleId,
#             /* [in] */ long lFlags,
#             /* [out] */ BSTR *MessageText) = 0;

#         virtual HRESULT STDMETHODCALLTYPE GetFacilityCodeText(
#             /* [in] */ HRESULT hRes,
#             /* [in] */ LCID LocaleId,
#             /* [in] */ long lFlags,
#             /* [out] */ BSTR *MessageText) = 0;

#     };

#     MIDL_INTERFACE("C49E32C7-BC8B-11d2-85D4-00105A1F8304")
#     IWbemBackupRestore : public IUnknown
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE Backup(
#             /* [string][in] */ __RPC__in_string LPCWSTR strBackupToFile,
#             /* [in] */ long lFlags) = 0;

#         virtual HRESULT STDMETHODCALLTYPE Restore(
#             /* [string][in] */ __RPC__in_string LPCWSTR strRestoreFromFile,
#             /* [in] */ long lFlags) = 0;

#     };

#     MIDL_INTERFACE("A359DEC5-E813-4834-8A2A-BA7F1D777D76")
#     IWbemBackupRestoreEx : public IWbemBackupRestore
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE Pause( void) = 0;

#         virtual HRESULT STDMETHODCALLTYPE Resume( void) = 0;

#     };

#     MIDL_INTERFACE("49353c99-516b-11d1-aea6-00c04fb68820")
#     IWbemRefresher : public IUnknown
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE Refresh(
#             /* [in] */ long lFlags) = 0;

#     };

#     MIDL_INTERFACE("2705C288-79AE-11d2-B348-00105A1F8177")
#     IWbemHiPerfEnum : public IUnknown
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE AddObjects(
#             /* [in] */ long lFlags,
#             /* [in] */ ULONG uNumObjects,
#             /* [size_is][in] */ long *apIds,
#             /* [size_is][in] */ IWbemObjectAccess **apObj) = 0;

#         virtual HRESULT STDMETHODCALLTYPE RemoveObjects(
#             /* [in] */ long lFlags,
#             /* [in] */ ULONG uNumObjects,
#             /* [size_is][in] */ long *apIds) = 0;

#         virtual HRESULT STDMETHODCALLTYPE GetObjects(
#             /* [in] */ long lFlags,
#             /* [in] */ ULONG uNumObjects,
#             /* [length_is][size_is][out] */ IWbemObjectAccess **apObj,
#             /* [out] */ ULONG *puReturned) = 0;

#         virtual HRESULT STDMETHODCALLTYPE RemoveAll(
#             /* [in] */ long lFlags) = 0;

#     };

#     MIDL_INTERFACE("49353c92-516b-11d1-aea6-00c04fb68820")
#     IWbemConfigureRefresher : public IUnknown
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE AddObjectByPath(
#             /* [in] */ IWbemServices *pNamespace,
#             /* [string][in] */ LPCWSTR wszPath,
#             /* [in] */ long lFlags,
#             /* [in] */ IWbemContext *pContext,
#             /* [out] */ IWbemClassObject **ppRefreshable,
#             /* [unique][in][out] */ long *plId) = 0;

#         virtual HRESULT STDMETHODCALLTYPE AddObjectByTemplate(
#             /* [in] */ IWbemServices *pNamespace,
#             /* [in] */ IWbemClassObject *pTemplate,
#             /* [in] */ long lFlags,
#             /* [in] */ IWbemContext *pContext,
#             /* [out] */ IWbemClassObject **ppRefreshable,
#             /* [unique][in][out] */ long *plId) = 0;

#         virtual HRESULT STDMETHODCALLTYPE AddRefresher(
#             /* [in] */ IWbemRefresher *pRefresher,
#             /* [in] */ long lFlags,
#             /* [unique][in][out] */ long *plId) = 0;

#         virtual HRESULT STDMETHODCALLTYPE Remove(
#             /* [in] */ long lId,
#             /* [in] */ long lFlags) = 0;

#         virtual HRESULT STDMETHODCALLTYPE AddEnum(
#             /* [in] */ IWbemServices *pNamespace,
#             /* [string][in] */ LPCWSTR wszClassName,
#             /* [in] */ long lFlags,
#             /* [in] */ IWbemContext *pContext,
#             /* [out] */ IWbemHiPerfEnum **ppEnum,
#             /* [unique][in][out] */ long *plId) = 0;

#     };


IWbemClassObject._methods_ = [
    STDMETHOD(c_int32, "GetQualifierSet", (POINTER(POINTER(IWbemQualifierSet)),)),
    STDMETHOD(c_int32, "Get", (c_wchar_p, c_int32, POINTER(Variant), POINTER(c_int32), POINTER(c_int32))),
    STDMETHOD(c_int32, "Put", (c_wchar_p, c_int32, POINTER(Variant), c_int32)),
    STDMETHOD(c_int32, "Delete", (c_wchar_p,)),
    STDMETHOD(c_int32, "GetNames", (c_wchar_p, c_int32, POINTER(Variant), POINTER(SafeArrayPtr))),
    STDMETHOD(c_int32, "BeginEnumeration", (c_int32,)),
    STDMETHOD(c_int32, "Next", (c_int32, POINTER(BSTR), POINTER(Variant), POINTER(c_int32), POINTER(c_int32))),
    STDMETHOD(c_int32, "EndEnumeration", ()),
    STDMETHOD(c_int32, "GetPropertyQualifierSet", (c_wchar_p, POINTER(POINTER(IWbemQualifierSet)))),
    STDMETHOD(c_int32, "Clone", (POINTER(POINTER(IWbemClassObject)),)),
    STDMETHOD(c_int32, "GetObjectText", (c_int32, POINTER(BSTR))),
    STDMETHOD(c_int32, "SpawnDerivedClass", (c_int32, POINTER(POINTER(IWbemClassObject)))),
    STDMETHOD(c_int32, "SpawnInstance", (c_int32, POINTER(POINTER(IWbemClassObject)))),
    STDMETHOD(c_int32, "CompareTo", (c_int32, POINTER(IWbemClassObject))),
    STDMETHOD(c_int32, "GetPropertyOrigin", (c_wchar_p, POINTER(BSTR))),
    STDMETHOD(c_int32, "InheritsFrom", (c_wchar_p,)),
    STDMETHOD(
        c_int32,
        "GetMethod",
        (c_wchar_p, c_int32, POINTER(POINTER(IWbemClassObject)), POINTER(POINTER(IWbemClassObject))),
    ),
    STDMETHOD(c_int32, "PutMethod", (c_wchar_p, c_int32, POINTER(IWbemClassObject), POINTER(IWbemClassObject))),
    STDMETHOD(c_int32, "DeleteMethod", (c_wchar_p,)),
    STDMETHOD(c_int32, "BeginMethodEnumeration", (c_int32,)),
    STDMETHOD(
        c_int32,
        "NextMethod",
        (c_int32, POINTER(BSTR), POINTER(POINTER(IWbemClassObject)), POINTER(POINTER(IWbemClassObject))),
    ),
    STDMETHOD(c_int32, "EndMethodEnumeration", ()),
    STDMETHOD(c_int32, "GetMethodQualifierSet", (c_wchar_p, POINTER(POINTER(IWbemQualifierSet)))),
    STDMETHOD(c_int32, "GetMethodOrigin", (c_wchar_p, POINTER(BSTR))),
]

IWbemObjectAccess._methods_ = [
    STDMETHOD(c_int32, "GetPropertyHandle", (c_wchar_p, POINTER(c_int32), POINTER(c_int32))),
    STDMETHOD(c_int32, "WritePropertyValue", (c_int32, c_int32, POINTER(c_byte))),
    STDMETHOD(c_int32, "ReadPropertyValue", (c_int32, c_int32, POINTER(c_int32), POINTER(c_byte))),
    STDMETHOD(c_int32, "ReadDWORD", (c_int32, POINTER(c_uint32))),
    STDMETHOD(c_int32, "WriteDWORD", (c_int32, c_uint32)),
    STDMETHOD(c_int32, "ReadQWORD", (c_int32, POINTER(c_uint64))),
    STDMETHOD(c_int32, "WriteQWORD", (c_int32, c_uint64)),
    STDMETHOD(c_int32, "GetPropertyInfoByHandle", (c_int32, POINTER(BSTR), POINTER(c_int32))),
    STDMETHOD(c_int32, "Lock", (c_int32,)),
    STDMETHOD(c_int32, "Unlock", (c_int32,)),
]

IWbemQualifierSet._methods_ = [
    STDMETHOD(c_int32, "Get", (c_wchar_p, c_int32, POINTER(Variant), POINTER(c_int32))),
    STDMETHOD(c_int32, "Put", (c_wchar_p, POINTER(Variant), c_int32)),
    STDMETHOD(c_int32, "Delete", (c_wchar_p,)),
    STDMETHOD(c_int32, "GetNames", (c_int32, POINTER(SafeArrayPtr))),
    STDMETHOD(c_int32, "BeginEnumeration", (c_int32,)),
    STDMETHOD(c_int32, "Next", (c_int32, POINTER(BSTR), POINTER(Variant), POINTER(c_int32))),
    STDMETHOD(c_int32, "EndEnumeration", ()),
]

IWbemServices._methods_ = [
    STDMETHOD(
        c_int32,
        "OpenNamespace",
        (BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemServices)), POINTER(POINTER(IWbemCallResult))),
    ),
    STDMETHOD(c_int32, "CancelAsyncCall", (POINTER(IWbemObjectSink),)),
    STDMETHOD(c_int32, "QueryObjectSink", (c_int32, POINTER(POINTER(IWbemObjectSink)))),
    STDMETHOD(
        c_int32,
        "GetObject",
        (
            BSTR,
            c_int32,
            POINTER(IWbemContext),
            POINTER(POINTER(IWbemClassObject)),
            POINTER(POINTER(IWbemCallResult)),
        ),
    ),
    STDMETHOD(c_int32, "GetObjectAsync", (BSTR, c_int32, POINTER(IWbemContext), POINTER(IWbemObjectSink))),
    STDMETHOD(
        c_int32,
        "PutClass",
        (POINTER(IWbemClassObject), c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemCallResult))),
    ),
    STDMETHOD(
        c_int32, "PutClassAsync", (POINTER(IWbemClassObject), c_int32, POINTER(IWbemContext), POINTER(IWbemObjectSink))
    ),
    STDMETHOD(c_int32, "DeleteClass", (BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemCallResult)))),
    STDMETHOD(c_int32, "DeleteClassAsync", (BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemObjectSink)))),
    STDMETHOD(
        c_int32, "CreateClassEnum", (BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IEnumWbemClassObject)))
    ),
    STDMETHOD(
        c_int32, "CreateClassEnumAsync", (BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemObjectSink)))
    ),
    STDMETHOD(
        c_int32,
        "PutInstance",
        (POINTER(IWbemClassObject), c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemCallResult))),
    ),
    STDMETHOD(
        c_int32,
        "PutInstanceAsync",
        (POINTER(IWbemClassObject), c_int32, POINTER(IWbemContext), POINTER(IWbemObjectSink)),
    ),
    STDMETHOD(c_int32, "DeleteInstance", (BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemCallResult)))),
    STDMETHOD(
        c_int32, "DeleteInstanceAsync", (BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemObjectSink)))
    ),
    STDMETHOD(
        c_int32, "CreateInstanceEnum", (BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IEnumWbemClassObject)))
    ),
    STDMETHOD(
        c_int32,
        "CreateInstanceEnumAsync",
        (BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemObjectSink))),
    ),
    STDMETHOD(
        c_int32, "ExecQuery", (BSTR, BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IEnumWbemClassObject)))
    ),
    STDMETHOD(
        c_int32, "ExecQueryAsync", (BSTR, BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemObjectSink)))
    ),
    STDMETHOD(
        c_int32,
        "ExecNotificationQuery",
        (BSTR, BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IEnumWbemClassObject))),
    ),
    STDMETHOD(
        c_int32,
        "ExecNotificationQueryAsync",
        (BSTR, BSTR, c_int32, POINTER(IWbemContext), POINTER(POINTER(IWbemObjectSink))),
    ),
    STDMETHOD(
        c_int32,
        "ExecMethod",
        (
            BSTR,
            BSTR,
            c_int32,
            POINTER(IWbemContext),
            POINTER(IWbemClassObject),
            POINTER(POINTER(IWbemClassObject)),
            POINTER(POINTER(IWbemCallResult)),
        ),
    ),
    STDMETHOD(
        c_int32,
        "ExecMethodAsync",
        (
            BSTR,
            BSTR,
            c_int32,
            POINTER(IWbemContext),
            POINTER(IWbemClassObject),
            POINTER(POINTER(IWbemClassObject)),
            POINTER(POINTER(IWbemObjectSink)),
        ),
    ),
]

#     MIDL_INTERFACE("e7d35cfa-348b-485e-b524-252725d697ca")
#     IWbemObjectSinkEx : public IWbemObjectSink
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE WriteMessage(
#             /* [in] */ ULONG uChannel,
#             /* [in] */ __RPC__in const BSTR strMessage) = 0;

#         virtual HRESULT STDMETHODCALLTYPE WriteError(
#             /* [unique][in] */ __RPC__in_opt IWbemClassObject *pObjError,
#             /* [out] */ __RPC__out unsigned char *puReturned) = 0;

#         virtual HRESULT STDMETHODCALLTYPE PromptUser(
#             /* [in] */ __RPC__in const BSTR strMessage,
#             /* [in] */ unsigned char uPromptType,
#             /* [out] */ __RPC__out unsigned char *puReturned) = 0;

#         virtual HRESULT STDMETHODCALLTYPE WriteProgress(
#             /* [in] */ __RPC__in const BSTR strActivity,
#             /* [in] */ __RPC__in const BSTR strCurrentOperation,
#             /* [in] */ __RPC__in const BSTR strStatusDescription,
#             /* [in] */ ULONG uPercentComplete,
#             /* [in] */ ULONG uSecondsRemaining) = 0;

#         virtual HRESULT STDMETHODCALLTYPE WriteStreamParameter(
#             /* [in] */ __RPC__in const BSTR strName,
#             /* [in] */ __RPC__in VARIANT *vtValue,
#             /* [in] */ ULONG ulType,
#             /* [in] */ ULONG ulFlags) = 0;

#     };

#     MIDL_INTERFACE("b7b31df9-d515-11d3-a11c-00105a1f515a")
#     IWbemShutdown : public IUnknown
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE Shutdown(
#             /* [in] */ LONG uReason,
#             /* [in] */ ULONG uMaxMilliseconds,
#             /* [in] */ __RPC__in_opt IWbemContext *pCtx) = 0;

#     };

#     MIDL_INTERFACE("bfbf883a-cad7-11d3-a11b-00105a1f515a")
#     IWbemObjectTextSrc : public IUnknown
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE GetText(
#             /* [in] */ long lFlags,
#             /* [in] */ IWbemClassObject *pObj,
#             /* [in] */ ULONG uObjTextFormat,
#             /* [in] */ IWbemContext *pCtx,
#             /* [out] */ BSTR *strText) = 0;

#         virtual HRESULT STDMETHODCALLTYPE CreateFromText(
#             /* [in] */ long lFlags,
#             /* [in] */ BSTR strText,
#             /* [in] */ ULONG uObjTextFormat,
#             /* [in] */ IWbemContext *pCtx,
#             /* [out] */ IWbemClassObject **pNewObj) = 0;

#     };

#     MIDL_INTERFACE("6daf974e-2e37-11d2-aec9-00c04fb68820")
#     IMofCompiler : public IUnknown
#     {
#     public:
#         virtual HRESULT STDMETHODCALLTYPE CompileFile(
#             /* [annotation][string][in] */
#             _In_  LPWSTR FileName,
#             /* [annotation][string][in] */
#             _In_  LPWSTR ServerAndNamespace,
#             /* [annotation][string][in] */
#             _In_  LPWSTR User,
#             /* [annotation][string][in] */
#             _In_  LPWSTR Authority,
#             /* [annotation][string][in] */
#             _In_  LPWSTR Password,
#             /* [annotation][in] */
#             _In_  LONG lOptionFlags,
#             /* [annotation][in] */
#             _In_  LONG lClassFlags,
#             /* [annotation][in] */
#             _In_  LONG lInstanceFlags,
#             /* [annotation][out][in] */
#             _Inout_  WBEM_COMPILE_STATUS_INFO *pInfo) = 0;

#         virtual HRESULT STDMETHODCALLTYPE CompileBuffer(
#             /* [annotation][in] */
#             _In_  long BuffSize,
#             /* [annotation][size_is][in] */
#             _In_reads_bytes_(BuffSize)  BYTE *pBuffer,
#             /* [annotation][string][in] */
#             _In_  LPWSTR ServerAndNamespace,
#             /* [annotation][string][in] */
#             _In_  LPWSTR User,
#             /* [annotation][string][in] */
#             _In_  LPWSTR Authority,
#             /* [annotation][string][in] */
#             _In_  LPWSTR Password,
#             /* [annotation][in] */
#             _In_  LONG lOptionFlags,
#             /* [annotation][in] */
#             _In_  LONG lClassFlags,
#             /* [annotation][in] */
#             _In_  LONG lInstanceFlags,
#             /* [annotation][out][in] */
#             _Inout_  WBEM_COMPILE_STATUS_INFO *pInfo) = 0;

#         virtual HRESULT STDMETHODCALLTYPE CreateBMOF(
#             /* [annotation][string][in] */
#             _In_  LPWSTR TextFileName,
#             /* [annotation][string][in] */
#             _In_  LPWSTR BMOFFileName,
#             /* [annotation][string][in] */
#             _In_  LPWSTR ServerAndNamespace,
#             /* [annotation][in] */
#             _In_  LONG lOptionFlags,
#             /* [annotation][in] */
#             _In_  LONG lClassFlags,
#             /* [annotation][in] */
#             _In_  LONG lInstanceFlags,
#             /* [annotation][out][in] */
#             _Inout_  WBEM_COMPILE_STATUS_INFO *pInfo) = 0;

#     };
IWbemLocator._methods_ = [
    STDMETHOD(
        c_int32,
        "ConnectServer",
        (BSTR, BSTR, BSTR, BSTR, c_int32, BSTR, POINTER(IWbemContext), POINTER(POINTER(IWbemServices))),
    ),
]

IWbemObjectSink._methods_ = [
    STDMETHOD(c_int32, "Indicate", (c_int32, POINTER(POINTER(IWbemClassObject)))),
    STDMETHOD(c_int32, "SetStatus", (c_int32, c_int32, BSTR, POINTER(IWbemClassObject))),
]

IEnumWbemClassObject._methods_ = [
    STDMETHOD(c_int32, "Reset", ()),
    STDMETHOD(c_int32, "Next", (c_int32, c_uint32, POINTER(POINTER(IWbemClassObject)), POINTER(c_uint32))),
    STDMETHOD(c_int32, "NextAsync", (c_uint32, POINTER(IWbemObjectSink))),
    STDMETHOD(c_int32, "Clone", (POINTER(POINTER(IEnumWbemClassObject)),)),
    STDMETHOD(c_int32, "Skip", (c_int32, c_uint32)),
]

IWbemCallResult._methods = [
    STDMETHOD(c_int32, "GetResultObject", (c_int32, POINTER(IWbemClassObject))),
    STDMETHOD(c_int32, "GetResultString", (c_int32, POINTER(BSTR))),
    STDMETHOD(c_int32, "GetResultServices", (c_int32, POINTER(POINTER(IWbemServices)))),
    STDMETHOD(c_int32, "GetCallStatus", (c_int32, POINTER(c_int32))),
]

IWbemContext._methods_ = [
    STDMETHOD(c_int32, "Clone", (POINTER(POINTER(IWbemContext)),)),
    STDMETHOD(c_int32, "GetNames", (c_int32, POINTER(SafeArrayPtr))),
    STDMETHOD(c_int32, "BeginEnumeration", (c_int32,)),
    STDMETHOD(c_int32, "Next", (c_int32, POINTER(BSTR), POINTER(Variant))),
    STDMETHOD(c_int32, "EndEnumeration", ()),
    STDMETHOD(c_int32, "SetValue", (c_wchar_p, c_int32, POINTER(Variant))),
    STDMETHOD(c_int32, "GetValue", (c_wchar_p, c_int32, POINTER(Variant))),
    STDMETHOD(c_int32, "DeleteValue", (c_wchar_p, c_int32)),
]
