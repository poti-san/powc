# TODO

# typedef struct _userFLAG_STGMEDIUM
#     {
#     LONG ContextFlags;
#     LONG fPassOwnership;
#     userSTGMEDIUM Stgmed;
#     } 	userFLAG_STGMEDIUM;

# typedef /* [unique] */  __RPC_unique_pointer userFLAG_STGMEDIUM *wireFLAG_STGMEDIUM;

# typedef /* [wire_marshal] */ struct _FLAG_STGMEDIUM
#     {
#     LONG ContextFlags;
#     LONG fPassOwnership;
#     STGMEDIUM Stgmed;
#     } 	FLAG_STGMEDIUM;


# EXTERN_C const IID IID_IAdviseSink;

# #if defined(__cplusplus) && !defined(CINTERFACE)

#     MIDL_INTERFACE("0000010f-0000-0000-C000-000000000046")
#     IAdviseSink : public IUnknown
#     {
#     public:
#         virtual /* [local] */ void STDMETHODCALLTYPE OnDataChange(
#             /* [annotation][unique][in] */
#             _In_  FORMATETC *pFormatetc,
#             /* [annotation][unique][in] */
#             _In_  STGMEDIUM *pStgmed) = 0;

#         virtual /* [local] */ void STDMETHODCALLTYPE OnViewChange(
#             /* [in] */ DWORD dwAspect,
#             /* [in] */ LONG lindex) = 0;

#         virtual /* [local] */ void STDMETHODCALLTYPE OnRename(
#             /* [annotation][in] */
#             _In_  IMoniker *pmk) = 0;

#         virtual /* [local] */ void STDMETHODCALLTYPE OnSave( void) = 0;

#         virtual /* [local] */ void STDMETHODCALLTYPE OnClose( void) = 0;

#     }
