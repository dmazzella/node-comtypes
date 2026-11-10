import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { setImmediate as tick } from "node:timers/promises";
import { setFlagsFromString } from "node:v8";
import { runInNewContext } from "node:vm";
import * as ctypes from "node-ctypes";

import {
  CLSID_ShellLink,
  COMError,
  COMPointer,
  CoCreateInstance,
  E_NOINTERFACE,
  GUID,
  ICreateErrorInfo,
  IErrorInfo,
  IPersist,
  IPersistFile,
  ISequentialStream,
  IShellLinkW,
  IStream,
  IUnknown,
  STATFLAG_NONAME,
  STATSTG,
  STGTY_STREAM,
  STREAM_SEEK_SET,
  S_OK,
  apartment,
  getLivePointerCount,
  takeErrorInfo,
} from "@node-ctypes/comtypes";
import type { InterfacePointer } from "@node-ctypes/comtypes";

const skip = process.platform !== "win32" && "COM requires Windows";

function createStream(): InterfacePointer<typeof IStream> {
  const ole32 = new ctypes.WinDLL("ole32.dll");
  const CreateStreamOnHGlobal = ole32.func("CreateStreamOnHGlobal", ctypes.c_int32, [ctypes.c_void_p, ctypes.c_int32, ctypes.c_void_p]);
  const out = new ctypes.c_void_p();
  assert.equal(CreateStreamOnHGlobal(null, 1, ctypes.byref(out)), S_OK);
  return COMPointer.fromAddress(out.value!, IStream);
}

/** Current reference count, measured with a raw AddRef/Release pair. */
function refCount(pointer: COMPointer): number {
  pointer.AddRef();
  return pointer.callMethod("Release") as number;
}

describe("COMPointer", { skip }, () => {
  it("QueryInterface returns typed, owned pointers", () => {
    using _ = apartment();
    using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    using file = link.QueryInterface(IPersistFile);
    assert.ok(file instanceof IPersistFile.Pointer);
    assert.ok(file instanceof IPersist.Pointer);
    assert.ok(file instanceof COMPointer);
    assert.equal(file.interface, IPersistFile);
    const clsid: GUID = file.GetClassID();
    assert.ok(clsid.equals(CLSID_ShellLink));
    assert.equal(file.IsDirty(), 1, "S_FALSE is returned, not thrown");
  });

  it("QueryInterface throws COMError and tryQueryInterface returns null", () => {
    using _ = apartment();
    using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    assert.throws(
      () => link.QueryInterface(IStream),
      (error: unknown) => error instanceof COMError && error.hresult === E_NOINTERFACE && error.method === "IUnknown::QueryInterface",
    );
    assert.equal(link.tryQueryInterface(IStream), null);
    using unknown = link.tryQueryInterface(IUnknown);
    assert.ok(unknown);
  });

  it("compares COM identity", () => {
    using _ = apartment();
    using a = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    using b = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    using persist = a.QueryInterface(IPersistFile);
    assert.ok(a.isSameObject(persist));
    assert.ok(persist.isSameObject(a));
    assert.ok(!a.isSameObject(b));
  });

  it("Release drops the owned reference and invalidates the pointer", () => {
    using _ = apartment();
    using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    const base = refCount(link);
    const second = link.QueryInterface(IUnknown);
    assert.equal(refCount(link), base + 1);
    assert.equal(second.Release(), base);
    assert.ok(second.released);
    assert.throws(() => second.Release(), /released/);
    assert.throws(() => second.QueryInterface(IUnknown), /released/);
    second[Symbol.dispose]();
    assert.equal(refCount(link), base);
  });

  it("detach and fromAddress transfer ownership", () => {
    using _ = apartment();
    using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    const base = refCount(link);
    const extra = link.QueryInterface(IPersistFile);
    const address = extra.detach();
    assert.ok(extra.released);
    assert.equal(refCount(link), base + 1, "detached reference is still held");
    using adopted = COMPointer.fromAddress(address, IPersistFile);
    assert.equal(refCount(link), base + 1);
    using borrowed = COMPointer.fromAddress(link.address, IShellLinkW, { addRef: true });
    assert.equal(refCount(link), base + 2);
    assert.ok(adopted.isSameObject(borrowed));
  });

  it("passes as a pointer argument through _as_parameter_", () => {
    using _ = apartment();
    using source = createStream();
    using target = createStream();
    const data = Buffer.from("@node-ctypes/comtypes");
    assert.equal(source.Write(data, data.length), data.length);
    source.Seek(0n, STREAM_SEEK_SET);
    // SELF parameter: the target stream is passed as an IStream*.
    assert.deepEqual(source.CopyTo(target, BigInt(data.length)), [BigInt(data.length), BigInt(data.length)]);
    target.Seek(0n);
    const read = Buffer.alloc(data.length);
    target.Read(read, read.length);
    assert.deepEqual(read, data);
  });

  it("formats itself", () => {
    using _ = apartment();
    const link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    assert.match(String(link), /^<POINTER\(IShellLinkW\) 0x[0-9A-F]{16}>$/);
    link.Release();
    assert.equal(String(link), "<POINTER(IShellLinkW) released>");
  });
});

describe("IStream", { skip }, () => {
  it("writes, seeks, reads, stats and clones", () => {
    using _ = apartment();
    using stream = createStream();
    const data = Buffer.from([1, 2, 3, 4, 5, 6, 7, 8]);
    assert.equal(stream.Write(data, data.length), 8);
    assert.equal(stream.Seek(2n, STREAM_SEEK_SET), 2n);
    const out = Buffer.alloc(4);
    assert.equal(stream.Read(out, 4), 4);
    assert.deepEqual([...out], [3, 4, 5, 6]);
    const stat: STATSTG = stream.Stat(STATFLAG_NONAME);
    assert.ok(stat instanceof STATSTG);
    assert.equal(stat.type, STGTY_STREAM);
    assert.equal(stat.cbSize, 8n);
    using clone = stream.Clone();
    assert.ok(clone instanceof IStream.Pointer, "SELF out parameters are typed as the interface");
    using sequential = clone!.QueryInterface(ISequentialStream);
    const rest = Buffer.alloc(2);
    assert.equal(sequential.Read(rest, 2), 2);
    assert.deepEqual([...rest], [7, 8]);
    assert.equal(stream.Commit(), S_OK);
  });
});

describe("error info", { skip }, () => {
  it("takeErrorInfo reads and clears the thread error object", () => {
    using _ = apartment();
    const oleaut32 = new ctypes.WinDLL("oleaut32.dll");
    const CreateErrorInfo = oleaut32.func("CreateErrorInfo", ctypes.c_int32, [ctypes.c_void_p]);
    const SetErrorInfo = oleaut32.func("SetErrorInfo", ctypes.c_int32, [ctypes.c_uint32, ctypes.c_void_p]);
    const out = new ctypes.c_void_p();
    assert.equal(CreateErrorInfo(ctypes.byref(out)), S_OK);
    using create = COMPointer.fromAddress(out.value!, ICreateErrorInfo);
    create.SetDescription("Something failed");
    create.SetSource("NodeComtypes.Test");
    create.SetHelpContext(42);
    create.SetGUID(IStream.iid);
    using info = create.QueryInterface(IErrorInfo);
    assert.equal(SetErrorInfo(0, info), S_OK);
    const details = takeErrorInfo();
    assert.ok(details);
    assert.equal(details.description, "Something failed");
    assert.equal(details.source, "NodeComtypes.Test");
    assert.equal(details.helpContext, 42);
    assert.equal(details.helpFile, null);
    assert.ok(details.guid?.equals(IStream.iid));
    assert.equal(takeErrorInfo(), null);
  });
});

describe("lifetime", { skip }, () => {
  it("does not leak across many create/QueryInterface/release cycles", () => {
    using _ = apartment();
    const baseline = getLivePointerCount();
    using keeper = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    const base = refCount(keeper);
    for (let i = 0; i < 2000; i++) {
      using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
      using file = link.QueryInterface(IPersistFile);
      using again = keeper.QueryInterface(IPersistFile);
      assert.ok(file.isSameObject(link));
      void again;
    }
    assert.equal(refCount(keeper), base);
    assert.equal(getLivePointerCount(), baseline + 1);
  });

  it("releases forgotten pointers on garbage collection", async () => {
    setFlagsFromString("--expose-gc");
    const gc = runInNewContext("gc") as () => void;
    using _ = apartment();
    using keeper = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    const base = refCount(keeper);
    const live = getLivePointerCount();
    (() => {
      for (let i = 0; i < 10; i++) keeper.QueryInterface(IUnknown);
    })();
    assert.equal(refCount(keeper), base + 10);
    for (let attempt = 0; attempt < 50 && refCount(keeper) !== base; attempt++) {
      gc();
      await tick();
    }
    assert.equal(refCount(keeper), base);
    assert.equal(getLivePointerCount(), live);
  });
});
