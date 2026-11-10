import * as ctypes from "node-ctypes";

const { c_int32, c_uint32, c_void_p } = ctypes;

/** Remote machine for `CoCreateInstanceEx`. */
export class COSERVERINFO extends ctypes.Structure {
  static _fields_ = [
    ["dwReserved1", c_uint32],
    ["pwszName", c_void_p],
    ["pAuthInfo", c_void_p],
    ["dwReserved2", c_uint32],
  ] as const;
  declare dwReserved1: number;
  declare pwszName: bigint | null;
  declare pAuthInfo: bigint | null;
  declare dwReserved2: number;
}

/** Interface request/result for `CoCreateInstanceEx`. */
export class MULTI_QI extends ctypes.Structure {
  static _fields_ = [
    ["pIID", c_void_p],
    ["pItf", c_void_p],
    ["hr", c_int32],
  ] as const;
  declare pIID: bigint | null;
  declare pItf: bigint | null;
  declare hr: number;
}
