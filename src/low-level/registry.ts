import * as ctypes from "node-ctypes";
import { POINTER_SIZE } from "./native.ts";

// Minimal registry access (advapi32), loaded on first use.

export const HKEY_CLASSES_ROOT = POINTER_SIZE === 8 ? 0xffffffff80000000n : 0x80000000n;
export const HKEY_CURRENT_USER = POINTER_SIZE === 8 ? 0xffffffff80000001n : 0x80000001n;

const KEY_READ = 0x20019;
const RRF_RT_REG_SZ = 0x2;
const REG_SZ = 1;
const ERROR_SUCCESS = 0;
const ERROR_FILE_NOT_FOUND = 2;
const ERROR_MORE_DATA = 234;
const ERROR_NO_MORE_ITEMS = 259;

function bind(library: ctypes.WinDLL) {
  const { c_int32, c_uint32, c_void_p, c_wchar_p } = ctypes;
  return {
    RegGetValueW: library.func("RegGetValueW", c_int32, [c_void_p, c_wchar_p, c_wchar_p, c_uint32, c_void_p, c_void_p, c_void_p]),
    RegOpenKeyExW: library.func("RegOpenKeyExW", c_int32, [c_void_p, c_wchar_p, c_uint32, c_uint32, c_void_p]),
    RegEnumKeyExW: library.func("RegEnumKeyExW", c_int32, [c_void_p, c_uint32, c_void_p, c_void_p, c_void_p, c_void_p, c_void_p, c_void_p]),
    RegCloseKey: library.func("RegCloseKey", c_int32, [c_void_p]),
    RegSetKeyValueW: library.func("RegSetKeyValueW", c_int32, [c_void_p, c_wchar_p, c_wchar_p, c_uint32, c_void_p, c_uint32]),
    RegDeleteTreeW: library.func("RegDeleteTreeW", c_int32, [c_void_p, c_wchar_p]),
  };
}

let advapi32: ReturnType<typeof bind> | null = null;

function api() {
  return (advapi32 ??= bind(new ctypes.WinDLL("advapi32.dll")));
}

/** A `REG_SZ` value (`name` `null` for the default value), or `null` when missing. */
export function readString(root: bigint, key: string, name: string | null = null): string | null {
  const size = new ctypes.c_uint32(0);
  let status = api().RegGetValueW(root, key, name, RRF_RT_REG_SZ, null, null, ctypes.byref(size));
  if (status === ERROR_FILE_NOT_FOUND) return null;
  if (status !== ERROR_SUCCESS && status !== ERROR_MORE_DATA) throw new Error(`RegGetValueW(${key}) failed with error ${status}`);
  const data = Buffer.alloc(size.value + 2);
  size.value = data.length;
  status = api().RegGetValueW(root, key, name, RRF_RT_REG_SZ, null, data, ctypes.byref(size));
  if (status === ERROR_FILE_NOT_FOUND) return null;
  if (status !== ERROR_SUCCESS) throw new Error(`RegGetValueW(${key}) failed with error ${status}`);
  return data.subarray(0, Math.max(0, size.value - 2)).toString("utf16le");
}

/** Names of the subkeys of `key` (empty when the key is missing). */
export function subkeys(root: bigint, key: string): string[] {
  const handle = new ctypes.c_void_p();
  const status = api().RegOpenKeyExW(root, key, 0, KEY_READ, ctypes.byref(handle));
  if (status === ERROR_FILE_NOT_FOUND) return [];
  if (status !== ERROR_SUCCESS) throw new Error(`RegOpenKeyExW(${key}) failed with error ${status}`);
  try {
    const names: string[] = [];
    const name = Buffer.alloc(512);
    for (let index = 0; ; index++) {
      const length = new ctypes.c_uint32(name.length / 2);
      const result = api().RegEnumKeyExW(handle.value, index, name, ctypes.byref(length), null, null, null, null);
      if (result === ERROR_NO_MORE_ITEMS) return names;
      if (result !== ERROR_SUCCESS) throw new Error(`RegEnumKeyExW(${key}) failed with error ${result}`);
      names.push(name.subarray(0, length.value * 2).toString("utf16le"));
    }
  } finally {
    api().RegCloseKey(handle.value);
  }
}

export function writeString(root: bigint, key: string, name: string | null, value: string): void {
  const data = Buffer.from(`${value}\0`, "utf16le");
  const status = api().RegSetKeyValueW(root, key, name, REG_SZ, data, data.length);
  if (status !== ERROR_SUCCESS) throw new Error(`RegSetKeyValueW(${key}) failed with error ${status}`);
}

export function deleteTree(root: bigint, key: string): void {
  const status = api().RegDeleteTreeW(root, key);
  if (status !== ERROR_SUCCESS && status !== ERROR_FILE_NOT_FOUND) throw new Error(`RegDeleteTreeW(${key}) failed with error ${status}`);
}
