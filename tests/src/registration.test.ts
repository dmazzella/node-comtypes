import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { randomUUID } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import { after, describe, it } from "node:test";
import { fileURLToPath } from "node:url";

import { COMError, CreateObject, E_NOINTERFACE, dispatchPointer, GUID, GetModule, apartment, isCoClass, registerTypeLib, registeredTypeLibPath, registeredVersions, unregisterTypeLib } from "@node-ctypes/comtypes";
import type { InterfaceDef } from "@node-ctypes/comtypes";
import { HKEY_CURRENT_USER, readString } from "@node-ctypes/comtypes/low-level";
import { buildTestTypeLib } from "../fixtures/typelib.ts";
import type { TestTypeLibIds } from "../fixtures/typelib.ts";

const skip = process.platform !== "win32" && "COM requires Windows";
const directory = mkdtempSync(join(tmpdir(), "node-comtypes-registration-"));
after(() => rmSync(directory, { recursive: true, force: true }));

const newId = () => `{${randomUUID().toUpperCase()}}`;
const newIds = (): TestTypeLibIds => ({ libid: newId(), counter: newId(), greeter: newId(), clsid: newId() });

function build(ids: TestTypeLibIds): string {
  const file = join(directory, `${ids.libid.slice(1, 9)}.tlb`);
  using _ = apartment();
  buildTestTypeLib(file, ids);
  return file;
}

describe("type library creation and registration", { skip }, () => {
  it("builds a type library with ICreateTypeLib2 and reads it back", () => {
    const ids = newIds();
    const file = build(ids);
    using _ = apartment();
    const Test = GetModule(file);
    const counter = Test.ICounter as InterfaceDef;
    const greeter = Test.IGreeter as InterfaceDef;
    assert.equal(counter.iid.toString(), ids.counter);
    assert.deepEqual(
      counter.methods.map((method) => method.name),
      ["Add", "_get_Name"],
    );
    assert.equal(greeter.base?.name, "IDispatch");
    assert.ok(isCoClass(Test.Greeter));
    assert.equal(Test.Greeter.defaultInterface, greeter);
  });

  it("registers a type library per user and unregisters it", () => {
    const ids = newIds();
    const file = build(ids);
    using _ = apartment();
    const identity = registerTypeLib(file);
    try {
      assert.deepEqual(identity, { libid: ids.libid, major: 1, minor: 0, lcid: 0 });
      assert.equal(registeredTypeLibPath({ libid: ids.libid }).toLowerCase(), file.toLowerCase());
      // [oleautomation] interfaces get the type library marshaler (PSOAInterface) as proxy/stub.
      assert.equal(readString(HKEY_CURRENT_USER, `Software\\Classes\\Interface\\${ids.counter}\\ProxyStubClsid32`), "{00020424-0000-0000-C000-000000000046}");
    } finally {
      unregisterTypeLib(identity);
    }
    assert.deepEqual(registeredVersions(ids.libid), []);
  });
});

describe("LocalServer32 registered with UseCommandLine", { skip }, () => {
  const server = fileURLToPath(new URL("../fixtures/typedserver.ts", import.meta.url));
  const run = (file: string, ids: TestTypeLibIds, ...args: string[]) => execFileSync(process.execPath, [server, file, ids.clsid, ...args], { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });

  it("prints the usage without a switch", () => {
    const script = "const { UseCommandLine } = await import('@node-ctypes/comtypes/server'); process.exitCode = (await UseCommandLine([], { argv: [] })) ? 0 : 3;";
    assert.throws(
      () => execFileSync(process.execPath, ["--input-type=module", "-e", script], { stdio: "pipe" }),
      (error: { status?: number; stderr?: Buffer }) => error.status === 3 && /regserver/.test(String(error.stderr)),
    );
  });

  it("serves typed objects whose custom interfaces cross processes through the registered type library", async () => {
    const ids = newIds();
    const file = build(ids);
    const key = `Software\\Classes\\CLSID\\${ids.clsid}`;
    run(file, ids, "/regserver");
    let pid = 0;
    try {
      assert.match(readString(HKEY_CURRENT_USER, `${key}\\LocalServer32`) ?? "", /typedserver\.ts/);
      assert.equal(readString(HKEY_CURRENT_USER, `${key}\\ProgID`), "NodeComtypesTest.Greeter.1");
      assert.equal(readString(HKEY_CURRENT_USER, `${key}\\VersionIndependentProgID`), "NodeComtypesTest.Greeter");
      assert.equal(readString(HKEY_CURRENT_USER, "Software\\Classes\\NodeComtypesTest.Greeter\\CurVer"), "NodeComtypesTest.Greeter.1");
      assert.equal(readString(HKEY_CURRENT_USER, `${key}\\TypeLib`), ids.libid);
      assert.equal(readString(HKEY_CURRENT_USER, `${key}\\Version`), "1.0");
      assert.ok(GUID.fromProgID("NodeComtypesTest.Greeter").equals(ids.clsid));

      using _ = apartment();
      const Test = GetModule(file);
      // Default: the best interface, from the coclass of the server (IProvideClassInfo).
      using greeter = CreateObject("NodeComtypesTest.Greeter");
      assert.equal((greeter as { interface?: InterfaceDef }).interface, Test.IGreeter);
      const greeting: string = greeter.Greet("node");
      pid = Number(/from (\d+)/.exec(greeting)?.[1]);
      assert.ok(pid > 0 && pid !== process.pid, "served by another process");
      assert.equal(greeter.Count, 1);
      // ICounter is a custom (non-dispatch) interface: its proxy comes from the registered type library.
      using counter = greeter.QueryInterface(Test.ICounter as InterfaceDef) as unknown as Disposable & { Add(a: number, b: number): number; Name: string };
      assert.equal(counter.Add(2, 40), 42);
      assert.equal(counter.Name, `counter ${pid}`);
    } finally {
      run(file, ids, "/unregserver");
    }
    assert.equal(readString(HKEY_CURRENT_USER, `${key}\\LocalServer32`), null);
    assert.equal(readString(HKEY_CURRENT_USER, "Software\\Classes\\NodeComtypesTest.Greeter\\CLSID"), null);
    assert.deepEqual(registeredVersions(ids.libid), [], "the type library is unregistered too");
    const alive = () => {
      try {
        process.kill(pid, 0);
        return true;
      } catch {
        return false;
      }
    };
    for (let attempt = 0; attempt < 100 && alive(); attempt++) await delay(100);
    assert.equal(alive(), false, "the server exits once its objects are released");
  });

  it("cannot marshal custom interfaces without the type library registration", () => {
    const ids = newIds();
    const file = build(ids);
    run(file, ids, "/regserver");
    try {
      using _ = apartment();
      const Test = GetModule(file);
      unregisterTypeLib({ libid: ids.libid, major: 1, minor: 0, lcid: 0 });
      using greeter = CreateObject(ids.clsid, { dynamic: true });
      assert.match(greeter.Greet("node"), /^Hello, node/, "IDispatch is marshaled by the system");
      assert.throws(
        () => dispatchPointer(greeter).QueryInterface(Test.ICounter as InterfaceDef),
        (error: unknown) => error instanceof COMError && error.hresult === E_NOINTERFACE,
      );
    } finally {
      run(file, ids, "/unregserver");
    }
  });
});
