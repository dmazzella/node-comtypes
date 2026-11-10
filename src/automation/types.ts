import * as ctypes from "node-ctypes";

const { c_int32, c_uint16, c_uint32, c_void_p } = ctypes;

/** Arguments of `IDispatch::Invoke`. */
export class DISPPARAMS extends ctypes.Structure {
  static _fields_ = [
    ["rgvarg", c_void_p],
    ["rgdispidNamedArgs", c_void_p],
    ["cArgs", c_uint32],
    ["cNamedArgs", c_uint32],
  ] as const;
  declare rgvarg: bigint | null;
  declare rgdispidNamedArgs: bigint | null;
  declare cArgs: number;
  declare cNamedArgs: number;
}

/** Exception raised by `IDispatch::Invoke` (`DISP_E_EXCEPTION`). BSTR fields are owned by the caller. */
export class EXCEPINFO extends ctypes.Structure {
  static _fields_ = [
    ["wCode", c_uint16],
    ["wReserved", c_uint16],
    ["bstrSource", c_void_p],
    ["bstrDescription", c_void_p],
    ["bstrHelpFile", c_void_p],
    ["dwHelpContext", c_uint32],
    ["pvReserved", c_void_p],
    ["pfnDeferredFillIn", c_void_p],
    ["scode", c_int32],
  ] as const;
  declare wCode: number;
  declare wReserved: number;
  declare bstrSource: bigint | null;
  declare bstrDescription: bigint | null;
  declare bstrHelpFile: bigint | null;
  declare dwHelpContext: number;
  declare pvReserved: bigint | null;
  declare pfnDeferredFillIn: bigint | null;
  declare scode: number;
}

export class SAFEARRAYBOUND extends ctypes.Structure {
  static _fields_ = [
    ["cElements", c_uint32],
    ["lLbound", c_int32],
  ] as const;
  declare cElements: number;
  declare lLbound: number;
}
