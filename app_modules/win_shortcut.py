"""Windows COM(IShellLinkW)을 직접 호출해 바탕화면 바로가기를 생성한다.

powershell.exe 같은 외부 프로세스를 전혀 실행하지 않으므로,
PATH 환경변수가 깨져 있거나 PowerShell 실행이 차단된 PC에서도 동작한다.
(기존 subprocess 방식은 PATH 탐색 실패 시 WinError 2로 앱 전체가 종료됐다.)
"""

import ctypes
from ctypes import POINTER, WINFUNCTYPE, Structure, byref, c_ulong, c_void_p, c_wchar_p
from ctypes.wintypes import BOOL, BYTE, DWORD, MAX_PATH, WORD

_ole32 = ctypes.oledll.ole32      # oledll: HRESULT 실패 시 OSError 발생
_shell32 = ctypes.windll.shell32

_CLSID_SHELL_LINK = "{00021401-0000-0000-C000-000000000046}"
_IID_ISHELL_LINK_W = "{000214F9-0000-0000-C000-000000000046}"
_IID_IPERSIST_FILE = "{0000010B-0000-0000-C000-000000000046}"

_CLSCTX_INPROC_SERVER = 0x1
_COINIT_APARTMENTTHREADED = 0x2
_RPC_E_CHANGED_MODE = -2147417850   # 0x80010106

# vtable 인덱스 (0~2번은 IUnknown: QueryInterface / AddRef / Release)
_QUERY_INTERFACE = 0
_RELEASE = 2
_ISL_SET_DESCRIPTION = 7
_ISL_SET_WORKING_DIRECTORY = 9
_ISL_SET_PATH = 20
_IPF_SAVE = 6                       # IPersistFile::Save

_CSIDL_DESKTOPDIRECTORY = 0x0010


class _GUID(Structure):
    _fields_ = [("Data1", DWORD), ("Data2", WORD), ("Data3", WORD), ("Data4", BYTE * 8)]


def _guid(text):
    """문자열 형태의 GUID를 구조체로 변환"""
    result = _GUID()
    _ole32.CLSIDFromString(text, byref(result))
    return result


def _method(interface_ptr, index, *argtypes):
    """COM 인터페이스 포인터의 vtable에서 index번째 메서드를 호출 가능한 함수로 변환"""
    vtable = ctypes.cast(interface_ptr, POINTER(POINTER(c_void_p))).contents
    return WINFUNCTYPE(ctypes.HRESULT, c_void_p, *argtypes)(vtable[index])


def _release(interface_ptr):
    """IUnknown::Release (반환값이 참조 카운트라 HRESULT로 검사하면 안 됨)"""
    vtable = ctypes.cast(interface_ptr, POINTER(POINTER(c_void_p))).contents
    WINFUNCTYPE(c_ulong, c_void_p)(vtable[_RELEASE])(interface_ptr)


def get_desktop_dir():
    """실제 바탕화면 경로 조회 (OneDrive 리다이렉트 환경 포함 대응)"""
    buf = ctypes.create_unicode_buffer(MAX_PATH)
    _shell32.SHGetFolderPathW(None, _CSIDL_DESKTOPDIRECTORY, None, 0, buf)
    return buf.value


def create_shortcut(shortcut_path, target_path, working_dir=None, description=None):
    """IShellLinkW로 .lnk 파일을 생성한다. 실패 시 OSError를 발생시킨다."""
    need_uninitialize = True
    try:
        _ole32.CoInitializeEx(None, _COINIT_APARTMENTTHREADED)
    except OSError as exc:
        # 다른 스레딩 모델로 이미 초기화된 경우 그대로 사용
        if getattr(exc, "winerror", None) != _RPC_E_CHANGED_MODE:
            raise
        need_uninitialize = False

    try:
        link = c_void_p()
        _ole32.CoCreateInstance(
            byref(_guid(_CLSID_SHELL_LINK)), None, _CLSCTX_INPROC_SERVER,
            byref(_guid(_IID_ISHELL_LINK_W)), byref(link),
        )
        try:
            _method(link, _ISL_SET_PATH, c_wchar_p)(link, target_path)
            if working_dir:
                _method(link, _ISL_SET_WORKING_DIRECTORY, c_wchar_p)(link, working_dir)
            if description:
                _method(link, _ISL_SET_DESCRIPTION, c_wchar_p)(link, description)

            persist = c_void_p()
            _method(link, _QUERY_INTERFACE, POINTER(_GUID), POINTER(c_void_p))(
                link, byref(_guid(_IID_IPERSIST_FILE)), byref(persist)
            )
            try:
                _method(persist, _IPF_SAVE, c_wchar_p, BOOL)(persist, shortcut_path, True)
            finally:
                _release(persist)
        finally:
            _release(link)
    finally:
        if need_uninitialize:
            ctypes.windll.ole32.CoUninitialize()
