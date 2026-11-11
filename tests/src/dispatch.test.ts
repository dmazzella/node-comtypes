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
  sync,
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
  it("reads and writes properties, methods and the default member", async () => {
    using _ = apartment();
    using dict = await CreateObject("Scripting.Dictionary", { dynamic: true });
    // CompareMode can only change while the dictionary is empty.
    dict.CompareMode = 1;
    assert.equal(dict.CompareMode, 1);
    assert.ok(isDispatch(dict));
    assert.equal(typeName(dict), "IDictionary");
    assert.equal(inspect(dict), "[Dispatch IDictionary]");
    assert.equal(`${dict}`, "[Dispatch IDictionary]");
    await dict.Add("number", 1);
    await dict.Add("text", "two");
    await dict.Add("date", new Date(2024, 0, 15, 10, 30));
    await dict.Add("big", 12345678901234n);
    await dict.Add("flag", false);
    await dict.Add("nothing", null);
    assert.equal(dict.Count, 6);
    assert.equal(await dict.Exists("text"), true);
    assert.equal(await dict.Exists("missing"), false);
    assert.equal(dict.Item("text"), "two");
    assert.equal(dict("number"), 1);
    assert.equal((dict.Item("date") as Date).getTime(), new Date(2024, 0, 15, 10, 30).getTime());
    assert.equal(dict.Item("big"), 12345678901234n);
    assert.equal(dict.Item("flag"), false);
    assert.equal(dict.Item("nothing"), null);
    putProperty(dict, "Item", "number", 42);
    assert.equal(getProperty(dict, "Item", "number"), 42);
    await dict.Remove("flag");
    assert.equal(dict.Count, 5);
    assert.ok("Add" in dict);
    assert.ok(!("NoSuchMember" in dict));
  });

  it("describes members from the type information", async () => {
    using _ = apartment();
    using dict = await CreateObject("Scripting.Dictionary");
    assert.deepEqual({ ...memberInfo(dict, "Count") }, { memid: 2, name: "Count", method: false, get: true, put: false, putref: false, getArguments: 0 });
    const item = memberInfo(dict, "Item");
    assert.ok(item && item.get && item.put && item.getArguments === 1);
    assert.ok(memberInfo(dict, "Add")?.method);
  });

  it("enumerates through _NewEnum", async () => {
    using _ = apartment();
    using dict = await CreateObject("Scripting.Dictionary");
    for (const key of ["a", "b", "c"]) await dict.Add(key, key.toUpperCase());
    assert.deepEqual([...dict], ["a", "b", "c"]);
    assert.deepEqual([...enumerate(dict)], ["a", "b", "c"]);
    const before = getLivePointerCount();
    for (const key of dict) if (key === "a") break;
    assert.equal(getLivePointerCount(), before, "early exit releases the enumerator");
  });

  it("stores and returns COM objects", async () => {
    using _ = apartment();
    using outer = await CreateObject("Scripting.Dictionary", { dynamic: true });
    using inner = await CreateObject("Scripting.Dictionary", { dynamic: true });
    await inner.Add("x", 1);
    await outer.Add("inner", inner);
    using back = outer.Item("inner");
    assert.ok(isDispatch(back));
    assert.equal(back.Item("x"), 1);
    assert.ok(dispatchPointer(back).isSameObject(dispatchPointer(inner)));
  });

  it("raises COMError for unknown names, bad arguments and exceptions", async () => {
    using _ = apartment();
    using dict = await CreateObject("Scripting.Dictionary", { dynamic: true });
    assert.throws(() => dict.NoSuchMember, hresultOf(DISP_E_UNKNOWNNAME));
    await dict.Add("a", 1);
    await assert.rejects(dict.Add("a", 2), (error: unknown) => error instanceof COMError && error.hresult === DISP_E_EXCEPTION && error.details?.scode === (0x800a01c9 | 0) && error.method === "IDispatch::Invoke(Add)");
    assert.throws(() => (dict.CompareMode = "not a number"), hresultOf(DISP_E_TYPEMISMATCH));
    await assert.rejects(dict.Add({}, 1), TypeError);
  });

  it("releases deterministically", async () => {
    using _ = apartment();
    const before = getLivePointerCount();
    const dict = await CreateObject("Scripting.Dictionary", { dynamic: true });
    assert.equal(getLivePointerCount(), before + 1);
    release(dict);
    assert.equal(getLivePointerCount(), before);
    assert.equal(inspect(dict), "[Dispatch released]");
    assert.throws(() => dict.Count, /released/);
    release(dict);
  });
});

describe("Scripting.FileSystemObject", { skip }, () => {
  it("writes and reads files with named arguments", async () => {
    using _ = apartment();
    using fso = await CreateObject("Scripting.FileSystemObject");
    const path = join(directory, "text.txt");
    using stream = await fso.CreateTextFile(path, named({ Overwrite: true, Unicode: true }));
    await stream.WriteLine("first line");
    await stream.Write("second");
    await stream.Close();
    assert.equal(readFileSync(path, "utf16le").replace(/^﻿/, ""), "first line\r\nsecond");
    assert.ok(await fso.FileExists(path));
    using reader = await fso.OpenTextFile(path, named({ IOMode: 1, Format: -1 }));
    assert.equal(await reader.ReadLine(), "first line");
    assert.equal(await reader.ReadAll(), "second");
    await reader.Close();
    using file = await fso.GetFile(path);
    assert.equal(file.Name, "text.txt");
    assert.ok(file.Size > 0);
    assert.ok(file.DateLastModified instanceof Date);
  });

  it("enumerates collections of objects", async () => {
    using _ = apartment();
    writeFileSync(join(directory, "a.txt"), "a");
    writeFileSync(join(directory, "b.txt"), "b");
    using fso = await CreateObject("Scripting.FileSystemObject");
    using folder = await fso.GetFolder(directory);
    const names = [...folder.Files].map((file) => {
      using owned = file;
      return owned.Name as string;
    });
    assert.ok(names.includes("a.txt") && names.includes("b.txt"));
    const drives = [...fso.Drives].map((drive) => drive.DriveLetter);
    assert.ok(drives.includes(process.env.SystemDrive?.[0] ?? "C"));
  });

  it("reports exceptions with the scode", async () => {
    using _ = apartment();
    using fso = await CreateObject("Scripting.FileSystemObject", { dynamic: true });
    await assert.rejects(fso.GetFile(join(directory, "missing.txt")), (error: unknown) => error instanceof COMError && error.hresult === DISP_E_EXCEPTION && error.details?.scode === (0x800a0035 | 0) && /scode 0x800A0035/.test(error.message));
  });

  it("finds dual-interface vtable slots from the type information", async () => {
    using _ = apartment();
    using fso = await CreateObject("Scripting.FileSystemObject", { dynamic: true });
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
  sync(dispatchInfo).GetRefTypeOfImplType(-1 >>> 0, ctypes.byref(reference));
  const out = new ctypes.c_void_p();
  sync(dispatchInfo).GetRefTypeInfo(reference.value, ctypes.byref(out));
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
      sync(info).GetRefTypeOfImplType(0, ctypes.byref(base));
      sync(info).GetRefTypeInfo(base.value, ctypes.byref(next));
      current = COMPointer.fromAddress(next.value!, ITypeInfo);
    }
  }
  throw new Error(`${name} not found`);
}

describe("IErrorInfo", { skip }, () => {
  it("vtable calls read details through ISupportErrorInfo", async () => {
    using _ = apartment();
    using document = await CreateObject("MSXML2.DOMDocument.6.0", { dynamic: true });
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
      await assert.rejects(
        vtable.createElement(name, ctypes.byref(new ctypes.c_void_p())),
        (error: unknown) => error instanceof COMError && error.method === "IXMLDOMDocument::createElement" && typeof error.details?.description === "string" && error.details.description.length > 0 && error.message.includes(error.details.description),
      );
    } finally {
      oleaut32().SysFreeString(name);
    }
  });

  it("late-bound calls read EXCEPINFO descriptions", async () => {
    using _ = apartment();
    using document = await CreateObject("MSXML2.DOMDocument.6.0", { dynamic: true });
    await assert.rejects(document.createElement("1invalid"), (error: unknown) => error instanceof COMError && error.hresult === DISP_E_EXCEPTION && typeof error.details?.description === "string" && error.details.description.length > 0);
  });
});

describe("other automation servers", { skip }, () => {
  it("WScript.Shell expands strings and exposes parameterized properties", async () => {
    using _ = apartment();
    using shell = await CreateObject("WScript.Shell");
    assert.equal((await shell.ExpandEnvironmentStrings("%SystemRoot%")).toLowerCase(), process.env.SystemRoot?.toLowerCase());
    using environment = getProperty(shell, "Environment", "PROCESS");
    assert.equal(environment.Item("SystemRoot").toLowerCase(), process.env.SystemRoot?.toLowerCase());
    assert.equal(environment("SystemRoot"), environment.Item("SystemRoot"));
  });

  it("Shell.Application returns folders and items", async () => {
    using _ = apartment();
    using app = await CreateObject("Shell.Application");
    using folder = await app.NameSpace(process.env.SystemRoot);
    assert.ok(typeof folder.Title === "string" && folder.Title.length > 0);
    let count = 0;
    for (const item of await folder.Items()) {
      using owned = item;
      assert.equal(typeof owned.Name, "string");
      if (++count === 5) break;
    }
    assert.equal(count, 5);
  });

  it("GetObject binds WMI and enumerates query results", async () => {
    using _ = apartment();
    using wmi = await GetObject("winmgmts:", { dynamic: true });
    const systems = [...(await wmi.ExecQuery("SELECT Caption, Version FROM Win32_OperatingSystem"))];
    assert.equal(systems.length, 1);
    assert.match(systems[0].Caption, /Windows/);
    assert.match(systems[0].Version, /^\d+\.\d+\.\d+$/);
    for (const system of systems) release(system);
    await assert.rejects(GetObject("::not a moniker::"), hresultOf(MK_E_SYNTAX));
  });

  it("GetActiveObject reports objects that are not running", async () => {
    using _ = apartment();
    await assert.rejects(GetActiveObject("Scripting.Dictionary"), hresultOf(MK_E_UNAVAILABLE));
  });

  it("CreateObject can return a vtable interface and Dispatch wraps pointers", async () => {
    using _ = apartment();
    using unknown = await CreateObject("Scripting.Dictionary", { interface: IUnknown });
    assert.equal(unknown.interface, IUnknown);
    using dict = Dispatch(unknown);
    await dict.Add("k", "v");
    using again = Dispatch(dict);
    assert.equal(again.Item("k"), "v");
    assert.equal(await invokeMethod(again, "Exists", "k"), true);
  });
});
