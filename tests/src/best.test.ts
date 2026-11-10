import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { CLSID_ShellLink, CoCreateInstance, CreateObject, GetBestInterface, GetObject, GetModule, IDispatch, IUnknown, apartment, dispatchPointer, getLivePointerCount, getProperty, isDispatch, putProperty, wrap } from "@node-ctypes/comtypes";
import type { InterfaceDef } from "@node-ctypes/comtypes";
import { COMObject } from "@node-ctypes/comtypes/server";

const skip = process.platform !== "win32" && "COM requires Windows";
const nameOf = (object: unknown) => (object as { interface?: InterfaceDef }).interface?.name;

describe("GetBestInterface", { skip }, () => {
  it("types objects from their type library by default", () => {
    using _ = apartment();
    const Scripting = GetModule("scrrun.dll");
    using dict = CreateObject("Scripting.Dictionary");
    assert.equal(nameOf(dict), "IDictionary");
    assert.equal((dict as { interface?: InterfaceDef }).interface, Scripting.IDictionary);
    assert.equal(isDispatch(dict), false);
    dict.Add("a", 1);
    dict.Add("b", 2);
    assert.equal(dict.Count, 2);
    assert.equal(dict.Item("b"), 2);
    assert.deepEqual([...dict], ["a", "b"], "typed dispatch pointers iterate _NewEnum");
    // Late-bound helpers also take typed dispatch pointers.
    putProperty(dict, "Item", "a", 10);
    assert.equal(getProperty(dict, "Item", "a"), 10);
    assert.equal(dispatchPointer(dict), dict);
  });

  it("keeps the dynamic client on request", () => {
    using _ = apartment();
    using dict = CreateObject("Scripting.Dictionary", { dynamic: true });
    assert.equal(isDispatch(dict), true);
    using typed = CreateObject("Scripting.Dictionary", { interface: IDispatch });
    assert.equal(nameOf(typed), "IDispatch");
  });

  it("returns new references and leaves objects without type information as they are", () => {
    using _ = apartment();
    const before = getLivePointerCount();
    {
      using link = CoCreateInstance(CLSID_ShellLink, IUnknown);
      using best = GetBestInterface(link);
      assert.equal(best.interface, IUnknown);
      assert.notEqual(best, link);
      assert.ok(best.isSameObject(link));
      assert.equal(wrap, GetBestInterface);
    }
    assert.equal(getLivePointerCount(), before);
  });

  it("converts IDispatch results of typed calls (wrap_outparam)", () => {
    using _ = apartment();
    using outer = CreateObject("Scripting.Dictionary");
    using inner = CreateObject("Scripting.Dictionary");
    inner.Add("x", 1);
    outer.Add("typed", inner);
    // A VARIANT holding an object with type information: its best interface.
    using back = outer.Item("typed");
    assert.equal(nameOf(back), "IDictionary");
    assert.ok(back.isSameObject(inner));
    assert.equal(back.Item("x"), 1);
    // Without type information (a JavaScript object): a dynamic dispatch object.
    class Plain extends COMObject {
      static _com_interfaces_ = [IDispatch];
      Twice(value: number) {
        return value * 2;
      }
    }
    outer.Add("plain", new Plain());
    using plain = outer.Item("plain");
    assert.equal(isDispatch(plain), true);
    assert.equal(plain.Twice(21), 42);
    // A method declared to return IDispatch (IShellDispatch::Windows).
    using shell = CreateObject("Shell.Application");
    using windows = shell.Windows();
    assert.equal(nameOf(windows), "IShellWindows");
    assert.equal(typeof windows.Count, "number");
  });

  it("types objects returned by GetObject", () => {
    using _ = apartment();
    using wmi = GetObject("winmgmts:");
    assert.equal(nameOf(wmi), "ISWbemServicesEx");
    // WMI class properties are not in the type library: ISWbemObject.Properties_ reads them.
    for (const system of wmi.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")) {
      using owned = system;
      assert.equal(nameOf(owned), "ISWbemObjectEx");
      using property = owned.Properties_.Item("Caption");
      assert.match(property.Value, /Windows/);
    }
  });
});
