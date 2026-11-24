---
title: Defining interfaces
---

# Defining interfaces

Declare COM interfaces with `COMInterface({ name, iid, extends, methods, dispMethods })`. The `methods` array lists vtable slots in order after the base interface (equivalent to `_methods_` in Python `comtypes`). TypeScript signatures are inferred directly from the method specs, including inherited methods.

## COMMETHOD (IDL flags)

`COMMETHOD(idlflags, restype, name, ...params)` converts inputs and allocates/returns outputs automatically. Each parameter is `[flags, type, name?, default?]`:

```typescript
import { c_int32, c_uint32 } from "node-ctypes";
import { BSTR, COMInterface, COMMETHOD, HRESULT, IDispatch, SELF, VARIANT } from "@node-ctypes/comtypes";

export const IThing = COMInterface({
  name: "IThing",
  iid: "{...}",
  extends: IDispatch,
  methods: [
    COMMETHOD(["propget"], HRESULT, "Name", [["out", "retval"], BSTR, "pVal"]), // thing.Name
    COMMETHOD(["propput"], HRESULT, "Name", [["in"], BSTR, "value"]), // thing.Name = "x"
    COMMETHOD(["propget"], HRESULT, "Item", [["in"], VARIANT, "index"], [["out", "retval"], SELF, "pItem"]), // thing.Item(1): IThing
    COMMETHOD([], HRESULT, "Find", [["in"], BSTR, "what"], [["in", "optional"], c_int32, "start", 0], [["out"], c_int32, "index"], [["out", "retval"], VARIANT, "match"]), // resolves to [index, match]
    COMMETHOD([], HRESULT, "Reset", [["lcid"], c_uint32, "lcid"]), // thing.Reset(): lcid auto-filled
  ],
});
```

### Parameter flags

| Flag combination        | Behavior                                                                                                                                    |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| `in` (or omitted)       | JavaScript input argument                                                                                                                   |
| `out`, `out` + `retval` | Allocated and returned: one output becomes the return value, multiple become a tuple in declaration order; no outputs returns the `HRESULT` |
| `in` + `out`            | Passed in and returned after the call                                                                                                       |
| `optional` / 4th item   | May be omitted (or `undefined`); uses the default value (`missing` for `VARIANT`)                                                           |
| `lcid`                  | Automatically filled with the default LCID, not exposed as a JS argument                                                                    |

A trailing `named({ name: value })` passes arguments by name at runtime. Missing required arguments, extra arguments and unknown names throw a `TypeError` before the native call.

### Parameter types

The type declared in a `COMMETHOD` parameter is the **value** type (`out` adds the indirection automatically; Python-style `["out"], POINTER(T)` is also accepted):

| Parameter type                                                | `in`                                                                                                                                | `out`                                         |
| ------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------- |
| `BSTR`                                                        | `string \| null` (allocated and freed)                                                                                              | `string` (freed automatically)                |
| `LPOLESTR`                                                    | `string \| null`                                                                                                                    | `string \| null` (`CoTaskMemFree`)            |
| `VARIANT`                                                     | Any JS value **by value**; an existing `VARIANT` instance is passed as-is                                                           | Converted JS value, cleared                   |
| Interface (`IFoo`), `SELF`                                    | `COMPointer`, dynamic object or `COMObject` (queried for `IFoo` when needed)                                                        | Owned `InterfacePointer<typeof IFoo> \| null` |
| `SAFEARRAY(vt)`                                               | Array, typed array, or `typed(...)`                                                                                                 | `unknown[] \| null`, destroyed                |
| `REFGUID` / `REFIID` / `REFCLSID`                             | `GUID` or GUID string                                                                                                               | `GUID`                                        |
| `GUID` and other `Structure`s                                 | Structure instance or plain object, by value                                                                                        | Structure instance                            |
| Simple types (`c_int32`, `c_double`, `wintypes.VARIANT_BOOL`) | Same as `node-ctypes`                                                                                                               | `.value`                                      |
| `POINTER(T)`                                                  | Raw pointers (`bigint`, `Buffer`, arrays, structs) pass through; for `VARIANT`, simple and struct types, values use a temporary `T` | Same as `T`                                   |
| `DATE`                                                        | `Date` or OLE Automation date number                                                                                                | `Date`                                        |
| `CURRENCY`                                                    | `number`, exact decimal `string`, or `bigint` (whole units)                                                                         | `number`                                      |
| `lazy(() => IFoo)`                                            | Forward/mutual interface reference                                                                                                  | Owned `InterfacePointer<typeof IFoo> \| null` |

For `POINTER(VARIANT)`, a COM pointer is marshaled as a `VT_DISPATCH` / `VT_UNKNOWN` value (not a raw address), and `undefined` passes a `VT_EMPTY` `VARIANT`.

### Properties and `SELF`

Property accessors are exposed both as `comtypes`-style methods (`_get_Name`, `_set_Name`, `_setref_Name`) and as JavaScript properties:

- a `propget` with no required arguments becomes a JS property, writable when a `propput` / `propputref` exists (`propputref` is selected for COM objects);
- a `propget` with required arguments becomes a method (`thing.Item(1)`), with `_set_Item(1, value)` available for puts.

`SELF` refers to the interface currently being declared (for example `IEnumVARIANT::Clone`, `IStream::Clone`) and preserves the exact `InterfacePointer` type.

## Dispinterfaces (`DISPMETHOD` and `DISPPROPERTY`)

Pure dispinterfaces have no vtable slots beyond `IDispatch`; their members are invoked via `IDispatch::Invoke` using their DISPID. Declare them in `dispMethods` (Python: `_disp_methods_`):

```typescript
import { c_int32 } from "node-ctypes";
import { BSTR, COMInterface, CURRENCY, DISPMETHOD, DISPPROPERTY, IDispatch, VARIANT, dispid } from "@node-ctypes/comtypes";

const Font = COMInterface({
  name: "Font",
  iid: "{BEF6E003-A874-101A-8BBA-00AA00300CAB}",
  extends: IDispatch,
  dispMethods: [
    DISPPROPERTY([dispid(0)], BSTR, "Name"), // font.Name / font.Name = "Arial"
    DISPPROPERTY([dispid(2)], CURRENCY, "Size"),
    DISPPROPERTY([dispid(9), "readonly"], c_int32, "Charset"),
    DISPMETHOD([dispid(10)], VARIANT, "Describe", [["in"], VARIANT, "what"], [["in", "optional"], c_int32, "depth", 1]),
  ],
});
```

`DISPMETHOD(idlflags, valueType, name, ...params)` supports optional arguments, defaults and trailing `named({...})`. Arguments are coerced to their declared types, return `VARIANT`s are decoded, and interface results are queried for their declared interface.

## STDMETHOD (raw vtable slots)

`STDMETHOD(restype, name, argtypes)` passes `node-ctypes` arguments directly without codec wrapping — ideal when the caller manages output buffers directly (such as `IShellLinkW::GetPath` with `create_unicode_buffer`):

```typescript
import { c_int32, c_uint32, c_void_p } from "node-ctypes";
import { COMInterface, HRESULT, IUnknown, STDMETHOD } from "@node-ctypes/comtypes";

export const IShellLinkWPart = COMInterface({
  name: "IShellLinkW",
  iid: "{000214F9-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [STDMETHOD(HRESULT, "GetPath", [c_void_p, c_int32, c_void_p, c_uint32])],
  register: false, // keep the built-in definition registered for this IID
});
// link.GetPath(buffer, 260, null, 0): Promise<number>; sync(link).GetPath(...): number
```

Prefer `c_int32` / `c_uint32` for `LONG` / `DWORD` / `ULONG` (the `wintypes` aliases map 32-bit `long` types to `bigint` in `node-ctypes`). In-memory structures such as `GUID` can be passed to `POINTER(T)` parameters. `pointer.callMethod(name, ...args)` invokes any slot by name. Interface definitions register globally by IID (most recent wins) unless `register: false` is set.

## Built-in interfaces

`@node-ctypes/comtypes/interfaces` (also re-exported from `@node-ctypes/comtypes`) ships ready-to-use declarations:

- **COM core**: `IUnknown`, `IClassFactory`, `IErrorInfo`, `ICreateErrorInfo`, `ISupportErrorInfo`, `IServiceProvider` (with `queryService(object, sid, IFoo)`)
- **Automation**: `IDispatch`, `IEnumVARIANT`, `ITypeInfo`, `ITypeLib`, `ITypeComp`, `IProvideClassInfo`, `IProvideClassInfo2`, `ICreateTypeLib`, `ICreateTypeLib2`, `ICreateTypeInfo`
- **Connection points**: `IConnectionPoint`, `IConnectionPointContainer`, `IEnumConnections`, `IEnumConnectionPoints`
- **Persistence and storage**: `IPersist`, `IPersistFile`, `IPropertyBag`, `IPersistPropertyBag`, `IPropertyBag2`, `IPersistPropertyBag2`, `IErrorLog`, `ISequentialStream`, `IStream`
- **Views**: `IViewObject`, `IViewObject2` (`SIZEL`, `DVASPECT_*`)
- **Shell**: `IShellLinkW`, `IShellLinkA`, `ShellLink` coclass (`CLSID_ShellLink`)

Most methods use `COMMETHOD`; raw `STDMETHOD` declarations are retained where callers pass raw buffers, arrays or descriptors (`IUnknown`, `IDispatch`, `IClassFactory`, `ITypeInfo`, `ITypeLib`, `ITypeComp`, `IServiceProvider`, `IShellLinkW`/`IShellLinkA`, `IEnumVARIANT::Next`, connection point enumerators, `IPropertyBag2`, `IViewObject`, `ICreateTypeInfo`).
