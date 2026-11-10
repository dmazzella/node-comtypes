import * as ctypes from "node-ctypes";
import { COMMETHOD } from "../commethod.ts";
import { COMInterface, HRESULT, STDMETHOD } from "../interface.ts";
import { IUnknown } from "./unknown.ts";

const { POINTER, c_int32, c_size_t, c_uint32, c_void_p } = ctypes;

/** Size in HIMETRIC units (`SIZEL`). */
export class SIZEL extends ctypes.Structure {
  static _fields_ = [
    ["cx", c_int32],
    ["cy", c_int32],
  ] as const;
  declare cx: number;
  declare cy: number;
}

/** `DVASPECT` values. */
export const DVASPECT_CONTENT = 1;
export const DVASPECT_THUMBNAIL = 2;
export const DVASPECT_ICON = 4;
export const DVASPECT_DOCPRINT = 8;

/**
 * Drawing of an object's presentation. Methods with device contexts, target
 * devices, rectangles and advise sinks take raw arguments. Python equivalent:
 * `comtypes.viewobject.IViewObject`.
 */
export const IViewObject = COMInterface({
  name: "IViewObject",
  iid: "{0000010D-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [
    // dwDrawAspect, lindex, pvAspect, ptd, hdcTargetDev, hdcDraw, lprcBounds, lprcWBounds, pfnContinue, dwContinue
    STDMETHOD(HRESULT, "Draw", [c_uint32, c_int32, c_void_p, c_void_p, c_void_p, c_void_p, c_void_p, c_void_p, c_void_p, c_size_t]),
    // dwDrawAspect, lindex, pvAspect, ptd, hicTargetDev, ppColorSet (LOGPALETTE**)
    STDMETHOD(HRESULT, "GetColorSet", [c_uint32, c_int32, c_void_p, c_void_p, c_void_p, POINTER(c_void_p)]),
    COMMETHOD([], HRESULT, "Freeze", [["in"], c_uint32, "dwDrawAspect"], [["in"], c_int32, "lindex"], [["in"], c_void_p, "pvAspect"], [["out"], c_uint32, "pdwFreeze"]),
    COMMETHOD([], HRESULT, "Unfreeze", [["in"], c_uint32, "dwFreeze"]),
    // aspects, advf, pAdvSink (IAdviseSink*)
    STDMETHOD(HRESULT, "SetAdvise", [c_uint32, c_uint32, c_void_p]),
    STDMETHOD(HRESULT, "GetAdvise", [POINTER(c_uint32), POINTER(c_uint32), POINTER(c_void_p)]),
  ],
});

export const IViewObject2 = COMInterface({
  name: "IViewObject2",
  iid: "{00000127-0000-0000-C000-000000000046}",
  extends: IViewObject,
  methods: [COMMETHOD([], HRESULT, "GetExtent", [["in"], c_uint32, "dwDrawAspect"], [["in"], c_int32, "lindex"], [["in"], c_void_p, "ptd"], [["out"], SIZEL, "lpsizel"])],
});
