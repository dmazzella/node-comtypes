---
title: Automation and VARIANTs
---

# Automation and VARIANTs

## Creating and binding objects

`CreateObject`, `GetObject` and `GetActiveObject` return the object's **best interface** (`GetBestInterface`), just like Python `comtypes`: the default interface of its coclass (`IProvideClassInfo`), or the interface described by `IDispatch::GetTypeInfo`, built from its type library with `GetModule` (dual interfaces are called directly through the vtable). Objects without type information become [dynamic objects](#dynamic-objects), or remain `IUnknown` pointers when they do not implement `IDispatch`.

```typescript
import { apartment, CreateObject, named, putProperty } from "@node-ctypes/comtypes";

using _ = apartment();

using dict = await CreateObject("Scripting.Dictionary"); // IDictionary, typed from scrrun.dll
dict.CompareMode = 1; // property put (only while the dictionary is empty)
await dict.Add("a", 1); // method
console.log(dict.Count, dict.Item("a")); // property get, parameterized property
putProperty(dict, "Item", "a", 2); // parameterized property put (late bound, by name)
for (const key of dict) console.log(key); // _NewEnum / IEnumVARIANT; for await: on the executor

using fso = await CreateObject("Scripting.FileSystemObject"); // IFileSystem3
using file = await fso.CreateTextFile("C:\\Temp\\a.txt", named({ Overwrite: true })); // ITextStream
```

- **Static types**: `CreateObject("Scripting.Dictionary")` is typed at runtime, with TypeScript type `AutomationObject` (`any` members). For full static TypeScript types, pass a coclass from a [generated module](typelib.md#generated-modules).
- **Wrapped results**: `IDispatch` outputs returned by typed calls (such as a `VARIANT` holding an object or `Application.ActiveSheet`) are wrapped to their best interface too (Python: `wrap_outparam`). Call `GetBestInterface(pointer)` (alias `wrap`) explicitly to obtain a new wrapped reference.
- **Iteration and late-bound helpers**: typed pointers deriving from `IDispatch` iterate their `_NewEnum` collection (`for...of` / `for await...of`), and late-bound helpers (`getProperty`, `putProperty`, `putRefProperty`, `invokeMethod`) accept them directly.
- **Options**:
  - `{ interface: IFoo }` queries and returns that specific interface;
  - `{ dynamic: true }` returns a [dynamic object](#dynamic-objects);
  - `{ home: true }` creates or binds on the home thread (see [Execution model](execution-model.md#creating-at-home));
  - `CreateObject` also accepts `{ clsctx, machine }` (`CoCreateInstanceEx` for remote servers).
- `GetActiveObject(progid)` attaches to an active instance in the Running Object Table; `GetObject(displayName)` binds a moniker (e.g. `"winmgmts:"`, `"LDAP://..."`, or a document file path).

## Constants

`Constants(obj)` (Python: `comtypes.client.Constants`) exposes the constants from an object's type library — enum members and module constants by name, plus enums by name. Accessing an unknown name throws so typos do not silently evaluate to `undefined`:

```typescript
import { Constants, CreateObject } from "@node-ctypes/comtypes";

using fso = await CreateObject("Scripting.FileSystemObject");
const c = Constants(fso);
using temp = await fso.GetSpecialFolder(c.TemporaryFolder); // or c.SpecialFolderConst.TemporaryFolder
```

## Dynamic objects

Pass `{ dynamic: true }` to use late binding through `IDispatch::GetIDsOfNames` + `Invoke` (Python: `dynamic=True`, `comtypes.client.dynamic.Dispatch`). Use dynamic objects for members not declared in a type library, such as WMI class properties:

```typescript
import { CreateObject, GetObject } from "@node-ctypes/comtypes";

using wmi = await GetObject("winmgmts:", { dynamic: true });
for await (const os of await wmi.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")) {
  using owned = os;
  console.log(owned.Caption);
}

using dict = await CreateObject("Scripting.Dictionary", { dynamic: true });
await dict.Add("a", 1);
console.log(dict("a")); // default member (DISPID_VALUE), a property: synchronous
```

| Expression                                 | `IDispatch::Invoke` flags                                                         |
| ------------------------------------------ | --------------------------------------------------------------------------------- |
| `obj.Name`                                 | `DISPATCH_PROPERTYGET`, or returns a method caller                                |
| `obj.Name(...args)`                        | `DISPATCH_METHOD \| DISPATCH_PROPERTYGET` (`Promise` for methods)                 |
| `obj.Name = value`                         | `DISPATCH_PROPERTYPUT` (`DISPATCH_PROPERTYPUTREF` for COM objects when supported) |
| `obj(...args)`                             | Default member (`DISPID_VALUE`)                                                   |
| `for (const x of obj)` / `for await`       | `_NewEnum` (`DISPID_NEWENUM`) and `IEnumVARIANT`                                  |
| `getProperty(obj, name, ...args)`          | `DISPATCH_PROPERTYGET` with arguments                                             |
| `putProperty(obj, name, ...args, value)`   | `DISPATCH_PROPERTYPUT` with arguments                                             |
| `putRefProperty(...)`, `invokeMethod(...)` | `DISPATCH_PROPERTYPUTREF`, `DISPATCH_METHOD` only                                 |

Whether `obj.Name` reads a value or returns a callable is determined from `IDispatch::GetTypeInfo` (cached per type):

- a property whose getter has no required arguments is read directly (`dict.Count`, `range.Value`);
- methods and properties with required arguments (`dict.Item`) return a function (methods return a `Promise` on the executor; parameterized properties run synchronously);
- without type info (e.g. WMI objects), a property read is attempted first, falling back to a callable only when the object rejects the get (`DISP_E_MEMBERNOTFOUND`, `DISP_E_BADPARAMCOUNT`, ...), matching Python `comtypes`.

> **Note:** A property whose getter has only _optional_ arguments is read as a value, so `shell.Environment("PROCESS")` would invoke the default member of the result. Use `getProperty(shell, "Environment", "PROCESS")` instead.

Returned objects are wrapped as dynamic objects again (including inside arrays). Each dynamic object owns its `IDispatch` reference: dispose it with `using` or `release(obj)`. Use `Dispatch(pointer)` to wrap an existing pointer or `COMObject`, and `dispatchPointer(obj)` to obtain the underlying `IDispatch` pointer.

## Special argument wrappers

- `undefined` — omitted optional argument (`missing`: `VT_ERROR` with `DISP_E_PARAMNOTFOUND`).
- `named({ Name: value })` — passes named arguments as the final argument of a call.
- `byRef(value)` — passes a VB `ByRef` argument (`VT_BYREF | VT_VARIANT`, or `byRef(value, VT_BSTR)` for a typed `ByRef`); read `.value` after the call (`const n = byRef(1); await obj.Increment(n); n.value`).
- `typed(vt, value)` — forces an explicit `VARTYPE` (e.g. `typed(VT_UI1, 7)`, `typed(VT_CY, "1.25")`).
- `safeArray(values, { vt, dimensions })` — builds a typed 1-D or N-D `SAFEARRAY`.

## VARIANT and SAFEARRAY conversions

| JavaScript                                                                                   | VARIANT                                                        | Back to JavaScript                                                      |
| -------------------------------------------------------------------------------------------- | -------------------------------------------------------------- | ----------------------------------------------------------------------- |
| `undefined` / `null`                                                                         | `VT_EMPTY` / `VT_NULL`                                         | `undefined` / `null`                                                    |
| `boolean`                                                                                    | `VT_BOOL`                                                      | `boolean`                                                               |
| `number`                                                                                     | `VT_I4` for 32-bit integers, else `VT_R8`                      | `number` for integers up to 32 bits, reals, `VT_CY`, `VT_DECIMAL`       |
| `bigint`                                                                                     | `VT_I8` (`VT_UI8` above 2^63 - 1)                              | `bigint` for `VT_I8` / `VT_UI8`                                         |
| `string`                                                                                     | `VT_BSTR`                                                      | `string`                                                                |
| `Date`                                                                                       | `VT_DATE` (local wall-clock time)                              | `Date`                                                                  |
| `COMPointer`, dynamic object, `COMObject`                                                    | `VT_DISPATCH` (`VT_UNKNOWN` if not an `IDispatch`)             | `COMPointer` (best interface in typed calls, dynamic object if dynamic) |
| `Array`                                                                                      | `VT_ARRAY \| VT_VARIANT`                                       | `Array` (N-D arrays as nested arrays)                                   |
| `Buffer`, typed arrays (`Uint8Array`, `Int32Array`, `Float64Array`, ...)                     | `VT_ARRAY` of matching element `VARTYPE` (1-D zero-copy write) | `Array`                                                                 |
| `ErrorValue`, `missing`                                                                      | `VT_ERROR`                                                     | `ErrorValue`                                                            |
| Record (`record(type, values)`, value from `VT_RECORD`, or record class from a type library) | `VT_RECORD` via `IRecordInfo` (requires type GUID)             | Plain object of fields keeping its `IRecordInfo` under `[recordInfo]`   |
| `typed(vt, value)`                                                                           | The requested `VARTYPE`                                        |                                                                         |
| `safeArray(values, { vt, dimensions })`                                                      | `VT_ARRAY \| vt` (`dimensions: 2` for `v[i][j]`)               |                                                                         |

`VT_BYREF` values are dereferenced automatically when read. A standalone `VARIANT` (`new VARIANT(value)`, `VARIANT.from(buffer)`, `.value`, `.vt`, `.clear()`, `.copy()`, `.changeType(vt, flags, lcid)`, `.exact` for exact decimal text of `VT_DECIMAL` / `VT_CY`) owns its `BSTR`, interface references and `SAFEARRAY`s — always dispose it with `using` or `.clear()`.

`SAFEARRAY`s follow the Windows column-major layout. `createSafeArray`, `readSafeArray` and `destroySafeArray` are exported from `@node-ctypes/comtypes/automation` for vtable interfaces.
