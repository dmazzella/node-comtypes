import * as ctypes from "node-ctypes";
import { GUID } from "../guid.ts";
import { COMInterface, HRESULT, STDMETHOD } from "../interface.ts";
import { IUnknown } from "./unknown.ts";

const { POINTER, c_int32, c_void_p } = ctypes;

export const IClassFactory = COMInterface({
  name: "IClassFactory",
  iid: "{00000001-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [STDMETHOD(HRESULT, "CreateInstance", [c_void_p, POINTER(GUID), POINTER(c_void_p)]), STDMETHOD(HRESULT, "LockServer", [c_int32])],
});
