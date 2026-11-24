---
title: Python comtypes correspondence
---

# Python comtypes correspondence

`@node-ctypes/comtypes` is modelled closely on Python [comtypes](https://github.com/enthought/comtypes), adapted for ESM, TypeScript, explicit `using` disposal and a non-blocking COM executor.

## API mapping

| Python `comtypes`                                            | `@node-ctypes/comtypes`                                           |
| ------------------------------------------------------------ | ----------------------------------------------------------------- |
| **Interfaces and pointers**                                  |                                                                   |
| `GUID("{...}")`, `GUID.create_new()`                         | `new GUID("{...}")`, `GUID.createNew()`                           |
| `GUID.from_progid()`, `.as_progid()`                         | `GUID.fromProgID()`, `.asProgID()`                                |
| `class IFoo(IUnknown): _iid_, _methods_`                     | `COMInterface({ name, iid, extends, methods })`                   |
| `STDMETHOD(HRESULT, "Name", [args])`                         | `STDMETHOD(HRESULT, "Name", [args])`                              |
| `COMMETHOD([flags], HRESULT, "Name", (["in"], T, "x"), ...)` | `COMMETHOD([flags], HRESULT, "Name", [["in"], T, "x"], ...)`      |
| `(["out"], POINTER(BSTR), "p")`                              | `[["out"], BSTR, "p"]` (`POINTER(T)` also accepted)               |
| `POINTER(IFoo)` parameter / `POINTER(POINTER(IFoo))` out     | `IFoo` (`SELF` for the interface itself)                          |
| `_midlSAFEARRAY(BSTR)`                                       | `SAFEARRAY(VT_BSTR)`                                              |
| `dispid(n)`, `helpstring(text)`                              | `dispid(n)`, `helpstring(text)`                                   |
| `DISPMETHOD`, `DISPPROPERTY`, `_disp_methods_`               | `DISPMETHOD`, `DISPPROPERTY`, `dispMethods`                       |
| `p.Name`, `p._get_Name()`, `p.Name = v`                      | Same                                                              |
| Multiple `out` parameters as a tuple                         | Multiple `out` parameters as a JS array                           |
| `CoInitializeEx()`, `CoCreateInstance(clsid, interface)`     | `CoInitializeEx()` / `apartment()`, `CoCreateInstance(...)`       |
| `p.QueryInterface(IFoo)`                                     | `p.QueryInterface(IFoo)`                                          |
| `p == q`                                                     | `p.isSameObject(q)`                                               |
| `del p` / refcount                                           | `p.Release()` / `using`                                           |
| `COMError(hresult, text, details)`                           | `COMError` with `hresult`, `text`, `details`                      |
| **Client**                                                   |                                                                   |
| `client.CreateObject(progid)`                                | `await CreateObject(progid)` (best interface)                     |
| `obj.Method(...)` (blocking)                                 | `await obj.Method(...)`; blocking: `sync(obj).Method(...)`        |
| `client.CreateObject(progid, dynamic=True)`                  | `await CreateObject(progid, { dynamic: true })`                   |
| `client.GetActiveObject`, `client.GetObject`                 | `await GetActiveObject(...)`, `await GetObject(...)`              |
| `client.GetBestInterface`, `client.wrap`                     | `GetBestInterface`, `wrap`                                        |
| `client.dynamic.Dispatch(p)`                                 | `Dispatch(p)`                                                     |
| `client.Constants(obj)`                                      | `Constants(obj)`                                                  |
| `obj.Item[key] = value`                                      | `putProperty(obj, "Item", key, value)`                            |
| Keyword arguments                                            | Trailing `named({ Name: value })`                                 |
| ByRef arguments (`win32com`: returned values)                | `byRef(value)`, read `.value` after the call                      |
| `automation.VARIANT`, `.value`                               | `VARIANT`, `.value`                                               |
| `comtypes.automation.Missing`                                | `missing`                                                         |
| `client.GetEvents`, `client.ShowEvents`                      | `GetEvents`, `ShowEvents`                                         |
| `client.PumpEvents`                                          | `PumpEvents`, `startMessagePump`, `waitAndPumpMessages`           |
| `CoMarshalInterThreadInterfaceInStream`, GIT                 | `marshalInterface` / `unmarshalInterface`, `GlobalInterfaceTable` |
| **Type libraries**                                           |                                                                   |
| `client.GetModule(path)`, `comtypes.gen`                     | `await GetModule(path, { cacheDir })`, `node-comtypes generate`   |
| Generated `CoClass` subclasses                               | `CoClass` objects; `await CreateObject(coclass)`                  |
| `typeinfo.CreateTypeLib`, `RegisterTypeLib`                  | `createTypeLib`, `registerTypeLib`, `unregisterTypeLib`           |
| **Servers**                                                  |                                                                   |
| `class X(COMObject): _com_interfaces_ = [IFoo]`              | `class X extends COMObject { static _com_interfaces_ = [IFoo] }`  |
| `_outgoing_interfaces_`, `Fire_Event`                        | `_outgoing_interfaces_`, `fireEvent("Name", ...)`                 |
| `_final_release_`, `_reg_clsid_`, `_reg_progid_`             | Same                                                              |
| `_reg_desc_`, `_reg_novers_progid_`                          | Same                                                              |
| `_reg_typelib_`, `_typelib_path_`                            | Same                                                              |
| `ReturnHRESULT(hresult, text)`                               | `throw new ReturnHRESULT(hresult, text)`                          |
| `server.register.UseCommandLine`, `register`                 | `UseCommandLine`, `registerServer`                                |
| `server.localserver.run`                                     | `runLocalServer`                                                  |
| **Other modules**                                            |                                                                   |
| `comtypes.persist`, `comtypes.viewobject`                    | `IPropertyBag`, `IPersistPropertyBag`(2), `IViewObject`(2), ...   |
| `IServiceProvider.QueryService(sid, interface)`              | `queryService(object, sid, interface)`                            |
| `comtypes.shelllink.ShellLink`                               | `ShellLink` (`CreateObject(ShellLink)`), `IShellLinkA`            |

## Key differences from Python

1. **Async by default**: COM methods return promises and execute on the dedicated COM executor thread so the Node.js event loop stays responsive. Use `sync(obj)` whenever you want Python-like synchronous calls on the calling thread.
2. **Deterministic disposal with `using`**: JavaScript does not destroy objects deterministically on scope exit without `using` — prefer `using ptr = ...` or `.Release()` over relying on the `FinalizationRegistry` safety net.
3. **Dynamic fallback for untyped `IDispatch`**: when an `IDispatch` output has no type library, `wrap_outparam` returns a dynamic object rather than a bare `IDispatch` pointer.
