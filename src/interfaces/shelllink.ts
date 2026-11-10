import * as ctypes from "node-ctypes";
import { GUID } from "../guid.ts";
import { COMInterface, HRESULT, STDMETHOD } from "../interface.ts";
import type { CoClass } from "../typelib/coclass.ts";
import { IPersistFile } from "./persist.ts";
import { IUnknown } from "./unknown.ts";

const { POINTER, c_char_p, c_int32, c_uint16, c_uint32, c_void_p, c_wchar_p } = ctypes;

/**
 * Shell link (Unicode). Output strings are written to caller buffers:
 * pass `create_unicode_buffer(n)` and its length in characters.
 * https://learn.microsoft.com/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishelllinkw
 */
export const IShellLinkW = COMInterface({
  name: "IShellLinkW",
  iid: "{000214F9-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [
    STDMETHOD(HRESULT, "GetPath", [c_void_p, c_int32, c_void_p, c_uint32]),
    STDMETHOD(HRESULT, "GetIDList", [POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "SetIDList", [c_void_p]),
    STDMETHOD(HRESULT, "GetDescription", [c_void_p, c_int32]),
    STDMETHOD(HRESULT, "SetDescription", [c_wchar_p]),
    STDMETHOD(HRESULT, "GetWorkingDirectory", [c_void_p, c_int32]),
    STDMETHOD(HRESULT, "SetWorkingDirectory", [c_wchar_p]),
    STDMETHOD(HRESULT, "GetArguments", [c_void_p, c_int32]),
    STDMETHOD(HRESULT, "SetArguments", [c_wchar_p]),
    STDMETHOD(HRESULT, "GetHotkey", [POINTER(c_uint16)]),
    STDMETHOD(HRESULT, "SetHotkey", [c_uint16]),
    STDMETHOD(HRESULT, "GetShowCmd", [POINTER(c_int32)]),
    STDMETHOD(HRESULT, "SetShowCmd", [c_int32]),
    STDMETHOD(HRESULT, "GetIconLocation", [c_void_p, c_int32, POINTER(c_int32)]),
    STDMETHOD(HRESULT, "SetIconLocation", [c_wchar_p, c_int32]),
    STDMETHOD(HRESULT, "SetRelativePath", [c_wchar_p, c_uint32]),
    STDMETHOD(HRESULT, "Resolve", [c_void_p, c_uint32]),
    STDMETHOD(HRESULT, "SetPath", [c_wchar_p]),
  ],
});

/** Shell link (ANSI): as {@link IShellLinkW}, with `create_string_buffer` buffers. */
export const IShellLinkA = COMInterface({
  name: "IShellLinkA",
  iid: "{000214EE-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [
    STDMETHOD(HRESULT, "GetPath", [c_void_p, c_int32, c_void_p, c_uint32]),
    STDMETHOD(HRESULT, "GetIDList", [POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "SetIDList", [c_void_p]),
    STDMETHOD(HRESULT, "GetDescription", [c_void_p, c_int32]),
    STDMETHOD(HRESULT, "SetDescription", [c_char_p]),
    STDMETHOD(HRESULT, "GetWorkingDirectory", [c_void_p, c_int32]),
    STDMETHOD(HRESULT, "SetWorkingDirectory", [c_char_p]),
    STDMETHOD(HRESULT, "GetArguments", [c_void_p, c_int32]),
    STDMETHOD(HRESULT, "SetArguments", [c_char_p]),
    STDMETHOD(HRESULT, "GetHotkey", [POINTER(c_uint16)]),
    STDMETHOD(HRESULT, "SetHotkey", [c_uint16]),
    STDMETHOD(HRESULT, "GetShowCmd", [POINTER(c_int32)]),
    STDMETHOD(HRESULT, "SetShowCmd", [c_int32]),
    STDMETHOD(HRESULT, "GetIconLocation", [c_void_p, c_int32, POINTER(c_int32)]),
    STDMETHOD(HRESULT, "SetIconLocation", [c_char_p, c_int32]),
    STDMETHOD(HRESULT, "SetRelativePath", [c_char_p, c_uint32]),
    STDMETHOD(HRESULT, "Resolve", [c_void_p, c_uint32]),
    STDMETHOD(HRESULT, "SetPath", [c_char_p]),
  ],
});

export const CLSID_ShellLink: GUID = new GUID("{00021401-0000-0000-C000-000000000046}");

/**
 * The shell link coclass: `CreateObject(ShellLink)` returns `IShellLinkW`.
 * Python equivalent: `comtypes.shelllink.ShellLink`.
 */
export const ShellLink: CoClass<typeof IShellLinkW> = Object.freeze({
  kind: "coclass",
  name: "ShellLink",
  clsid: CLSID_ShellLink,
  defaultInterface: IShellLinkW,
  sourceInterface: null,
  interfaces: Object.freeze([IShellLinkW, IShellLinkA, IPersistFile]),
  sourceInterfaces: Object.freeze([]),
  canCreate: true,
});

/** `IShellLinkW::GetPath` flags. */
export const SLGP_SHORTPATH = 0x1;
export const SLGP_UNCPRIORITY = 0x2;
export const SLGP_RAWPATH = 0x4;
export const SLGP_RELATIVEPRIORITY = 0x8;

/** `IShellLinkW::Resolve` flags. */
export const SLR_NO_UI = 0x1;
export const SLR_UPDATE = 0x4;
export const SLR_NOUPDATE = 0x8;
export const SLR_NOSEARCH = 0x10;
export const SLR_NOTRACK = 0x20;
