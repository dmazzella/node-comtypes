import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, describe, it } from "node:test";
import { inspect } from "node:util";
import * as ctypes from "node-ctypes";

import {
  COMError,
  COMInterface,
  COMPointer,
  CreateObject,
  DISP_E_EXCEPTION,
  DISP_E_TYPEMISMATCH,
  DISP_E_UNKNOWNNAME,
  Dispatch,
  GetActiveObject,
  GetObject,
  HRESULT,
  IDispatch,
  ITypeInfo,
  IUnknown,
  MK_E_SYNTAX,
  MK_E_UNAVAILABLE,
  STDMETHOD,
  TKIND_DISPATCH,
  TKIND_INTERFACE,
  apartment,
  dispatchPointer,
  dispatchTypeInfo,
  enumerate,
  getLivePointerCount,
  getProperty,
  invokeMethod,
  isDispatch,
  memberInfo,
  named,
  putProperty,
  release,
  typeName,
  withFuncDesc,
  withTypeAttr,
  documentation,
} from "@node-ctypes/comtypes";
import type { DispatchObject, InterfacePointer } from "@node-ctypes/comtypes";
import { allocBSTR, oleaut32 } from "@node-ctypes/comtypes/low-level";

const skip = process.platform !== "win32" && "COM requires Windows";
const directory = mkdtempSync(join(tmpdir(), "node-comtypes-dispatch-"));
after(() => rmSync(directory, { recursive: true, force: true }));

function hresultOf(code: number) {
  return (error: unknown) => error instanceof COMError && error.hresult === code;
}

describe("Scripting.Dictionary", { skip }, () => {
  it("reads and writes properties, methods and the default member", () => {
    using _ = apartment();
    using dict = CreateObject("Scripting.Dictionary", { dynamic: true });
    // CompareMode can only change while the dictionary is empty.
    dict.CompareMode = 1;
    assert.equal(dict.CompareMode, 1);
    assert.ok(isDispatch(dict));
    assert.equal(typeName(dict), "IDictionary");
    assert.equal(inspect(dict), "[Dispatch IDictionary]");
    assert.equal(`${dict}`, "[Dispatch IDictionary]");
    dict.Add("number", 1);
    dict.Add("text", "two");
    dict.Add("date", new Date(2024, 0, 15, 10, 30));
    dict.Add("big", 12345678901234n);
    dict.Add("flag", false);
    dict.Add("nothing", null);
    assert.equal(dict.Count, 6);
    assert.equal(dict.Exists("text"), true);
    assert.equal(dict.Exists("missing"), false);
    assert.equal(dict.Item("text"), "two");
    assert.equal(dict("number"), 1);
    assert.equal((dict.Item("date") as Date).getTime(), new Date(2024, 0, 15, 10, 30).getTime());
    assert.equal(dict.Item("big"), 12345678901234n);
    assert.equal(dict.Item("flag"), false);
    assert.equal(dict.Item("nothing"), null);
    putProperty(dict, "Item", "number", 42);
    assert.equal(getProperty(dict, "Item", "number"), 42);
    dict.Remove("flag");
    assert.equal(dict.Count, 5);
    assert.ok("Add" in dict);
    assert.ok(!("NoSuchMember" in dict));
  });

  it("describes members from the type information", () => {
    using _ = apartment();
    using dict = CreateObject("Scripting.Dictionary");
    assert.deepEqual({ ...memberInfo(dict, "Count") }, { memid: 2, name: "Count", method: false, get: true, put: false, putref: false, getArguments: 0 });
    const item = memberInfo(dict, "Item");
    assert.ok(item && item.get && item.put && item.getArguments === 1);
    assert.ok(memberInfo(dict, "Add")?.method);
  });

  it("enumerates through _NewEnum", () => {
    using _ = apartment();
    using dict = CreateObject("Scripting.Dictionary");
    for (const key of ["a", "b", "c"]) dict.Add(key, key.toUpperCase());
    assert.deepEqual([...dict], ["a", "b", "c"]);
    assert.deepEqual([...enumerate(dict)], ["a", "b", "c"]);
    const before = getLivePointerCount();
    for (const key of dict) if (key === "a") break;
    assert.equal(getLivePointerCount(), before, "early exit releases the enumerator");
  });

  it("stores and returns COM objects", () => {
    using _ = apartment();
    using outer = CreateObject("Scripting.Dictionary", { dynamic: true });
    using inner = CreateObject("Scripting.Dictionary", { dynamic: true });
    inner.Add("x", 1);
    outer.Add("inner", inner);
    using back = outer.Item("inner");
    assert.ok(isDispatch(back));
    assert.equal(back.Item("x"), 1);
    assert.ok(dispatchPointer(back).isSameObject(dispatchPointer(inner)));
  });

  it("raises COMError for unknown names, bad arguments and exceptions", () => {
    using _ = apartment();
    using dict = CreateObject("Scripting.Dictionary", { dynamic: true });
    assert.throws(() => dict.NoSuchMember, hresultOf(DISP_E_UNKNOWNNAME));
    dict.Add("a", 1);
    assert.throws(
      () => dict.Add("a", 2),
      (error: unknown) => error instanceof COMError && error.hresult === DISP_E_EXCEPTION && error.details?.scode === (0x800a01c9 | 0) && error.method === "IDispatch::Invoke(Add)",
    );
    assert.throws(() => (dict.CompareMode = "not a number"), hresultOf(DISP_E_TYPEMISMATCH));
    assert.throws(() => dict.Add({}, 1), TypeError);
  });

  it("releases deterministically", () => {
    using _ = apartment();
    const before = getLivePointerCount();
    const dict = CreateObject("Scripting.Dictionary", { dynamic: true });
    assert.equal(getLivePointerCount(), before + 1);
    release(dict);
    assert.equal(getLivePointerCount(), before);
    assert.equal(inspect(dict), "[Dispatch released]");
    assert.throws(() => dict.Count, /released/);
    release(dict);
  });
});

describe("Scripting.FileSystemObject", { skip }, () => {
  it("writes and reads files with named arguments", () => {
    using _ = apartment();
    using fso = CreateObject("Scripting.FileSystemObject");
    const path = join(directory, "text.txt");
    using stream = fso.CreateTextFile(path, named({ Overwrite: true, Unicode: true }));
    stream.WriteLine("first line");
    stream.Write("second");
    stream.Close();
    assert.equal(readFileSync(path, "utf16le").replace(/^﻿/, ""), "first line\r\nsecond");
    assert.ok(fso.FileExists(path));
    using reader = fso.OpenTextFile(path, named({ IOMode: 1, Format: -1 }));
    assert.equal(reader.ReadLine(), "first line");
    assert.equal(reader.ReadAll(), "second");
    reader.Close();
    using file = fso.GetFile(path);
    assert.equal(file.Name, "text.txt");
    assert.ok(file.Size > 0);
    assert.ok(file.DateLastModified instanceof Date);
  });

  it("enumerates collections of objects", () => {
    using _ = apartment();
    writeFileSync(join(directory, "a.txt"), "a");
    writeFileSync(join(directory, "b.txt"), "b");
    using fso = CreateObject("Scripting.FileSystemObject");
    using folder = fso.GetFolder(directory);
    const names = [...folder.Files].map((file) => {
      using owned = file;
      return owned.Name as string;
    });
    assert.ok(names.includes("a.txt") && names.includes("b.txt"));
    const drives = [...fso.Drives].map((drive) => drive.DriveLetter);
    assert.ok(drives.includes(process.env.SystemDrive?.[0] ?? "C"));
  });

  it("reports exceptions with the scode", () => {
    using _ = apartment();
    using fso = CreateObject("Scripting.FileSystemObject", { dynamic: true });
    assert.throws(
      () => fso.GetFile(join(directory, "missing.txt")),
      (error: unknown) => error instanceof COMError && error.hresult === DISP_E_EXCEPTION && error.details?.scode === (0x800a0035 | 0) && /scode 0x800A0035/.test(error.message),
    );
  });

  it("finds dual-interface vtable slots from the type information", () => {
    using _ = apartment();
    using fso = CreateObject("Scripting.FileSystemObject", { dynamic: true });
    // The default interface is IFileSystem3; GetFile is declared by its base IFileSystem.
    assert.deepEqual(vtableSlot(fso, "GetFile"), { iid: "{0AB5A3D0-E5B6-11D0-ABF5-00A0C90FFFC0}", slot: 20 });
  });
});

/** Interface IID and vtable slot of a dual interface method, from the type information. */
function vtableSlot(object: DispatchObject, name: string): { iid: string; slot: number } {
  using dispatchInfo = dispatchTypeInfo(dispatchPointer(object));
  assert.ok(dispatchInfo);
  assert.equal(
    withTypeAttr(dispatchInfo, (attr) => attr.typekind),
    TKIND_DISPATCH,
  );
  // For a dual interface, implemented type -1 is the vtable interface.
  const reference = new ctypes.c_uint32();
  dispatchInfo.GetRefTypeOfImplType(-1 >>> 0, ctypes.byref(reference));
  const out = new ctypes.c_void_p();
  dispatchInfo.GetRefTypeInfo(reference.value, ctypes.byref(out));
  let current: InterfacePointer<typeof ITypeInfo> | null = COMPointer.fromAddress(out.value!, ITypeInfo);
  while (current) {
    using info = current;
    current = null;
    const { functions, implemented, iid } = withTypeAttr(info, (attr) => {
      assert.equal(attr.typekind, TKIND_INTERFACE);
      return { functions: attr.cFuncs, implemented: attr.cImplTypes, iid: attr.guid.toString() };
    });
    for (let index = 0; index < functions; index++) {
      const slot = withFuncDesc(info, index, (desc) => (documentation(info, desc.memid).name === name ? desc.oVft / ctypes.sizeof(ctypes.c_void_p) : -1));
      if (slot >= 0) return { iid, slot };
    }
    if (implemented) {
      const base = new ctypes.c_uint32();
      const next = new ctypes.c_void_p();
      info.GetRefTypeOfImplType(0, ctypes.byref(base));
      info.GetRefTypeInfo(base.value, ctypes.byref(next));
      current = COMPointer.fromAddress(next.value!, ITypeInfo);
    }
  }
  throw new Error(`${name} not found`);
}

describe("IErrorInfo", { skip }, () => {
  it("vtable calls read details through ISupportErrorInfo", () => {
    using _ = apartment();
    using document = CreateObject("MSXML2.DOMDocument.6.0", { dynamic: true });
    const { iid, slot } = vtableSlot(document, "createElement");
    // Placeholders keep the vtable layout up to createElement (IDispatch has 7 slots).
    const IXMLDOMDocumentPart = COMInterface({
      name: "IXMLDOMDocument",
      iid,
      extends: IDispatch,
      methods: [...Array.from({ length: slot - 7 }, (_, index) => STDMETHOD(HRESULT, `slot${index + 7}`)), STDMETHOD(HRESULT, "createElement", [ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)])],
    });
    using vtable = dispatchPointer(document).QueryInterface(IXMLDOMDocumentPart);
    const name = allocBSTR("1invalid");
    try {
      assert.throws(
        () => vtable.createElement(name, ctypes.byref(new ctypes.c_void_p())),
        (error: unknown) => error instanceof COMError && error.method === "IXMLDOMDocument::createElement" && typeof error.details?.description === "string" && error.details.description.length > 0 && error.message.includes(error.details.description),
      );
    } finally {
      oleaut32().SysFreeString(name);
    }
  });

  it("late-bound calls read EXCEPINFO descriptions", () => {
    using _ = apartment();
    using document = CreateObject("MSXML2.DOMDocument.6.0", { dynamic: true });
    assert.throws(
      () => document.createElement("1invalid"),
      (error: unknown) => error instanceof COMError && error.hresult === DISP_E_EXCEPTION && typeof error.details?.description === "string" && error.details.description.length > 0,
    );
  });
});

describe("other automation servers", { skip }, () => {
  it("WScript.Shell expands strings and exposes parameterized properties", () => {
    using _ = apartment();
    using shell = CreateObject("WScript.Shell");
    assert.equal(shell.ExpandEnvironmentStrings("%SystemRoot%").toLowerCase(), process.env.SystemRoot?.toLowerCase());
    using environment = getProperty(shell, "Environment", "PROCESS");
    assert.equal(environment.Item("SystemRoot").toLowerCase(), process.env.SystemRoot?.toLowerCase());
    assert.equal(environment("SystemRoot"), environment.Item("SystemRoot"));
  });

  it("Shell.Application returns folders and items", () => {
    using _ = apartment();
    using app = CreateObject("Shell.Application");
    using folder = app.NameSpace(process.env.SystemRoot);
    assert.ok(typeof folder.Title === "string" && folder.Title.length > 0);
    let count = 0;
    for (const item of folder.Items()) {
      using owned = item;
      assert.equal(typeof owned.Name, "string");
      if (++count === 5) break;
    }
    assert.equal(count, 5);
  });

  it("GetObject binds WMI and enumerates query results", () => {
    using _ = apartment();
    using wmi = GetObject("winmgmts:", { dynamic: true });
    const systems = [...wmi.ExecQuery("SELECT Caption, Version FROM Win32_OperatingSystem")];
    assert.equal(systems.length, 1);
    assert.match(systems[0].Caption, /Windows/);
    assert.match(systems[0].Version, /^\d+\.\d+\.\d+$/);
    for (const system of systems) release(system);
    assert.throws(() => GetObject("::not a moniker::"), hresultOf(MK_E_SYNTAX));
  });

  it("GetActiveObject reports objects that are not running", () => {
    using _ = apartment();
    assert.throws(() => GetActiveObject("Scripting.Dictionary"), hresultOf(MK_E_UNAVAILABLE));
  });

  it("CreateObject can return a vtable interface and Dispatch wraps pointers", () => {
    using _ = apartment();
    using unknown = CreateObject("Scripting.Dictionary", { interface: IUnknown });
    assert.equal(unknown.interface, IUnknown);
    using dict = Dispatch(unknown);
    dict.Add("k", "v");
    using again = Dispatch(dict);
    assert.equal(again.Item("k"), "v");
    assert.equal(invokeMethod(again, "Exists", "k"), true);
  });
});
