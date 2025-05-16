"""COMセキュリティ機能を提供します。"""

from ctypes import c_int32, c_uint32, c_void_p, c_wchar_p
from enum import IntEnum
from typing import Any

from .core import _ole32, check_hresult


class RpcImpLevel(IntEnum):
    """RPC偽装レベル。*。"""

    DEFAULT = 0
    ANONYMOUS = 1
    IDENTIFY = 2
    IMPERSONATE = 3
    DELEGATE = 4


class RpcAuthnService(IntEnum):
    """RPC認証サービス。*。"""

    NONE = 0
    DCE_PRIVATE = 1
    DCE_PUBLIC = 2
    DEC_PUBLIC = 4
    GSS_NEGOTIATE = 9
    WINNT = 10
    GSS_SCHANNEL = 14
    GSS_KERBEROS = 16
    DPA = 17
    MSN = 18
    DIGEST = 21
    KERNEL = 20
    NEGO_EXTENDER = 30
    PKU2U = 31
    LIVE_SSP = 32
    LIVEXP_SSP = 35
    CLOUD_AP = 36
    MSONLINE = 82
    MQ = 100
    DEFAULT = 0xFFFFFFFF


class RpcAuthnLevel(IntEnum):
    """RPC認証レベル。*。"""

    DEFAULT = 0
    NONE = 1
    CONNECT = 2
    CALL = 3
    PKT = 4
    PKT_INTEGRITY = 5
    PKT_PRIVACY = 6


class OleAuthnCaps(IntEnum):
    """OLE認証キャパシティ。"""

    NONE = 0
    MUTUAL_AUTH = 0x1
    STATIC_CLOAKING = 0x20
    DYNAMIC_CLOAKING = 0x40
    ANY_AUTHORITY = 0x80
    MAKE_FULLSIC = 0x100
    DEFAULT = 0x800
    SECURE_REFS = 0x2
    ACCESS_CONTROL = 0x4
    APPID = 0x8
    DYNAMIC = 0x10
    REQUIRE_FULLSIC = 0x200
    AUTO_IMPERSONATE = 0x400
    DISABLE_AAA = 0x1000
    NO_CUSTOM_MARSHAL = 0x2000
    RESERVED1 = 0x4000


class RpcAuthzService(IntEnum):
    """RPC許可サービス。RPC_C_AUTHZ_*。"""

    NONE = 0
    NAME = 1
    DCE = 2
    DEFAULT = 0xFFFFFFFF


_CoInitializeSecurity = _ole32.CoInitializeSecurity
_CoInitializeSecurity.argtypes = (
    c_void_p,
    c_int32,
    c_void_p,
    c_void_p,
    c_uint32,
    c_uint32,
    c_void_p,
    c_uint32,
    c_void_p,
)

_CoSetProxyBlanket = _ole32.CoSetProxyBlanket
_CoSetProxyBlanket.argtypes = (c_void_p, c_uint32, c_uint32, c_wchar_p, c_uint32, c_uint32, c_void_p, c_uint32)


def com_init_security(
    authnlevel: RpcAuthnLevel = RpcAuthnLevel.DEFAULT,
    implevel: RpcImpLevel = RpcImpLevel.IMPERSONATE,
    oleauthn: OleAuthnCaps = OleAuthnCaps.NONE,
    allow_toolate: bool = True,
):
    hr = _CoInitializeSecurity(None, -1, None, None, int(authnlevel), int(implevel), None, int(oleauthn), None)

    RPC_E_TOO_LATE = 0x80010119
    if allow_toolate and hr == RPC_E_TOO_LATE:
        print("COM security was initialized.")
        return
    check_hresult(hr)


def com_set_securityblanket(
    obj: Any,
    authnsvc: RpcAuthnService = RpcAuthnService.DEFAULT,
    authzsvc: RpcAuthzService = RpcAuthzService.DEFAULT,
    authnlevel: RpcAuthnLevel = RpcAuthnLevel.DEFAULT,
    implevel: RpcImpLevel = RpcImpLevel.IMPERSONATE,
    oleauthn: OleAuthnCaps = OleAuthnCaps.NONE,
):
    check_hresult(
        _CoSetProxyBlanket(
            obj, int(authnsvc), int(authzsvc), c_wchar_p(-1), int(authnlevel), int(implevel), -1, int(oleauthn)
        ),
    )
