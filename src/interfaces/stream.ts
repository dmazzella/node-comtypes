import * as ctypes from "node-ctypes";
import { GUID } from "../guid.ts";
import { COMMETHOD, SELF } from "../commethod.ts";
import { COMInterface, HRESULT } from "../interface.ts";
import { IUnknown } from "./unknown.ts";

const { c_int64, c_uint32, c_uint64, c_void_p } = ctypes;

export class FILETIME extends ctypes.Structure {
  static _fields_ = [
    ["dwLowDateTime", c_uint32],
    ["dwHighDateTime", c_uint32],
  ] as const;
  declare dwLowDateTime: number;
  declare dwHighDateTime: number;
}

/** `IStream::Stat` result; free `pwcsName` with CoTaskMemFree unless `STATFLAG_NONAME`. */
export class STATSTG extends ctypes.Structure {
  static _fields_ = [
    ["pwcsName", c_void_p],
    ["type", c_uint32],
    ["cbSize", c_uint64],
    ["mtime", FILETIME],
    ["ctime", FILETIME],
    ["atime", FILETIME],
    ["grfMode", c_uint32],
    ["grfLocksSupported", c_uint32],
    ["clsid", GUID],
    ["grfStateBits", c_uint32],
    ["reserved", c_uint32],
  ] as const;
  declare pwcsName: bigint | null;
  declare type: number;
  declare cbSize: bigint;
  declare mtime: FILETIME;
  declare ctime: FILETIME;
  declare atime: FILETIME;
  declare grfMode: number;
  declare grfLocksSupported: number;
  declare clsid: GUID;
  declare grfStateBits: number;
  declare reserved: number;
}

/** `pv` buffers are raw memory (`Buffer`); the byte counts are returned. */
export const ISequentialStream = COMInterface({
  name: "ISequentialStream",
  iid: "{0C733A30-2A1C-11CE-ADE5-00AA0044773D}",
  extends: IUnknown,
  methods: [COMMETHOD([], HRESULT, "Read", [["in"], c_void_p, "pv"], [["in"], c_uint32, "cb"], [["out"], c_uint32, "pcbRead"]), COMMETHOD([], HRESULT, "Write", [["in"], c_void_p, "pv"], [["in"], c_uint32, "cb"], [["out"], c_uint32, "pcbWritten"])],
});

export const IStream = COMInterface({
  name: "IStream",
  iid: "{0000000C-0000-0000-C000-000000000046}",
  extends: ISequentialStream,
  methods: [
    // LARGE_INTEGER / ULARGE_INTEGER are passed by value as 64-bit integers.
    COMMETHOD([], HRESULT, "Seek", [["in"], c_int64, "dlibMove"], [["in", "optional"], c_uint32, "dwOrigin", 0], [["out"], c_uint64, "plibNewPosition"]),
    COMMETHOD([], HRESULT, "SetSize", [["in"], c_uint64, "libNewSize"]),
    COMMETHOD([], HRESULT, "CopyTo", [["in"], SELF, "pstm"], [["in"], c_uint64, "cb"], [["out"], c_uint64, "pcbRead"], [["out"], c_uint64, "pcbWritten"]),
    COMMETHOD([], HRESULT, "Commit", [["in", "optional"], c_uint32, "grfCommitFlags", 0]),
    COMMETHOD([], HRESULT, "Revert"),
    COMMETHOD([], HRESULT, "LockRegion", [["in"], c_uint64, "libOffset"], [["in"], c_uint64, "cb"], [["in"], c_uint32, "dwLockType"]),
    COMMETHOD([], HRESULT, "UnlockRegion", [["in"], c_uint64, "libOffset"], [["in"], c_uint64, "cb"], [["in"], c_uint32, "dwLockType"]),
    // Default STATFLAG_NONAME: with STATFLAG_DEFAULT, free pwcsName with CoTaskMemFree.
    COMMETHOD([], HRESULT, "Stat", [["out"], STATSTG, "pstatstg"], [["in", "optional"], c_uint32, "grfStatFlag", 1]),
    COMMETHOD([], HRESULT, "Clone", [["out"], SELF, "ppstm"]),
  ],
});

export const STREAM_SEEK_SET = 0;
export const STREAM_SEEK_CUR = 1;
export const STREAM_SEEK_END = 2;
export const STATFLAG_DEFAULT = 0;
export const STATFLAG_NONAME = 1;
export const STGTY_STORAGE = 1;
export const STGTY_STREAM = 2;
export const STGTY_LOCKBYTES = 3;
export const STGTY_PROPERTY = 4;
