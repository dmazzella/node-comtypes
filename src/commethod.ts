import type { AutomationObject } from "./client/dynamic.ts";
import type { IDispatch } from "./interfaces/dispatch.ts";
import * as ctypes from "node-ctypes";
import { VT_TYPEMASK } from "./automation/constants.ts";
import { defaultLCID } from "./automation/dispatch.ts";
import { NamedArguments } from "./automation/named.ts";
import { currencyToString } from "./automation/decimal.ts";
import { dateToOADate, oaDateToDate } from "./automation/date.ts";
import { Typed, VARIANT, borrowReference, comPointer, comReference, createSafeArray, toCurrency, destroySafeArray, missing, readSafeArray, readVariant } from "./automation/variant.ts";
import { checkHRESULT } from "./errors.ts";
import { GUID, toGUID } from "./guid.ts";
import type { GUIDLike } from "./guid.ts";
import { HRESULT, derivesFrom } from "./interface.ts";
import type { InterfaceDef, InterfacePointer, MethodKind, MethodMarshal, MethodMarshalAsync, MethodSpec } from "./interface.ts";
import { allocBSTR, ole32, oleaut32, readBSTR, takeCoTaskMemString } from "./low-level/native.ts";
import { rawAddRef, rawQueryInterface, rawRelease } from "./low-level/vtable.ts";
import { COMPointer } from "./pointer.ts";
import type { BorrowedReference } from "./automation/variant.ts";

// ── parameter types ───────────────────────────────────────────────────

/** `BSTR`: a JavaScript string; allocated for `in`, read and freed for `out`. */
export interface BSTRType {
  readonly __comtype: "BSTR";
}
/** `LPOLESTR` / `LPWSTR`: a wide string; `out` strings are freed with `CoTaskMemFree`. */
export interface LPOLESTRType {
  readonly __comtype: "LPOLESTR";
}
/** `REFGUID` / `REFIID` / `REFCLSID`: a GUID or GUID string passed by reference. */
export interface RefGuidType {
  readonly __comtype: "REFGUID";
}
/** `SAFEARRAY(vt)*`: JavaScript arrays; `out` arrays are read and destroyed. */
export interface SafeArrayType<V extends number = number> {
  readonly __comtype: "SAFEARRAY";
  readonly vt: V;
}
/** The interface being defined, e.g. the result of `IEnumVARIANT::Clone`. */
export interface SelfMarker {
  readonly __comtype: "SELF";
}
/** An interface resolved when called: for interfaces defined later (mutual references). */
export interface LazyInterface<T extends InterfaceDef<any> = InterfaceDef<any>> {
  readonly __comtype: "LAZY";
  readonly resolve: () => T;
}
/** `DATE` (an OLE Automation date, a double): a JavaScript `Date`. */
export interface DateType {
  readonly __comtype: "DATE";
}
/** `CURRENCY` / `CY` (a 64-bit integer scaled by 10 000): a number (or an exact string / bigint in input). */
export interface CurrencyType {
  readonly __comtype: "CURRENCY";
}

export const BSTR: BSTRType = Object.freeze({ __comtype: "BSTR" });
export const LPOLESTR: LPOLESTRType = Object.freeze({ __comtype: "LPOLESTR" });
export const REFGUID: RefGuidType = Object.freeze({ __comtype: "REFGUID" });
export const REFIID = REFGUID;
export const REFCLSID = REFGUID;
export const SELF: SelfMarker = Object.freeze({ __comtype: "SELF" });
export const DATE: DateType = Object.freeze({ __comtype: "DATE" });
export const CURRENCY: CurrencyType = Object.freeze({ __comtype: "CURRENCY" });

/** Reference to an interface defined later: `[["out", "retval"], lazy(() => IFolder), "folder"]`. */
export function lazy<T extends InterfaceDef<any>>(resolve: () => T): LazyInterface<T> {
  if (typeof resolve !== "function") throw new TypeError("lazy expects a function returning an interface definition");
  return Object.freeze({ __comtype: "LAZY", resolve });
}

/** Python equivalent: `comtypes.automation._midlSAFEARRAY(type)`. */
export function SAFEARRAY<const V extends number>(vt: V): SafeArrayType<V> {
  if (!Number.isInteger(vt) || vt <= 0 || vt & ~VT_TYPEMASK) throw new RangeError(`Invalid SAFEARRAY element type ${vt}`);
  return Object.freeze({ __comtype: "SAFEARRAY", vt });
}

/**
 * Type of a COMMETHOD parameter. It describes the **value**: `out`/`retval`
 * parameters add the pointer (`[["out", "retval"], BSTR, "pVal"]` is a
 * `BSTR*`). Python-style `POINTER(T)` out types are accepted too.
 *
 * - `BSTR`, `LPOLESTR`, `REFGUID`, `SAFEARRAY(vt)`;
 * - an interface definition (`IFoo` is `IFoo*`) or `SELF`;
 * - `VARIANT` (by value for `in`);
 * - any node-ctypes type: simple types, structures (`GUID`), pointers.
 */
export type ComParamType = BSTRType | LPOLESTRType | RefGuidType | SafeArrayType | SelfMarker | LazyInterface | DateType | CurrencyType | InterfaceDef<any> | ctypes.FFIType;

export type ParamFlag = "in" | "out" | "retval" | "lcid" | "optional";

/** `[flags, type, name?, default?]`. Python: `(["in"], type, "name", default)`. */
export type ParamSpec = readonly [flags: readonly ParamFlag[], type: ComParamType, name?: string, defaultValue?: unknown];

export interface DispIdFlag {
  readonly dispid: number;
}
export interface HelpStringFlag {
  readonly helpstring: string;
}
/** `propget`, `propput`, `propputref`, `hidden`, `restricted`, ..., `dispid(n)`, `helpstring(text)`. */
export type IdlFlag = string | DispIdFlag | HelpStringFlag;

export function dispid(id: number): DispIdFlag {
  return Object.freeze({ dispid: id | 0 });
}

export function helpstring(text: string): HelpStringFlag {
  return Object.freeze({ helpstring: text });
}

// ── type-level conversions ────────────────────────────────────────────

/** A COM object accepted for interface parameters: a pointer, a dynamic object or a JavaScript `COMObject`. */
export type ComObject = COMPointer | { readonly [comPointer]: COMPointer } | { readonly [comReference]: () => BorrowedReference | null };

type RawArgument<T> = T extends ctypes.FFIType ? Parameters<ctypes.TypedFFIFunction<readonly [T], null>>[0] : never;
type RawResult<T> = T extends ctypes.FFIType | null ? ReturnType<ctypes.TypedFFIFunction<readonly [], T>> : never;
type SelfDef<Self> = Self extends InterfaceDef<any> ? Self : InterfaceDef;

/** JavaScript value accepted for an input parameter of type `T`. */
export type InValue<T> = T extends BSTRType | LPOLESTRType
  ? string | null
  : T extends DateType
    ? Date | number
    : T extends CurrencyType
      ? number | string | bigint
      : T extends LazyInterface
        ? ComObject | null
        : T extends RefGuidType
          ? GUIDLike
          : T extends SafeArrayType
            ? ArrayLike<unknown> | Typed | null
            : T extends SelfMarker
              ? ComObject | null
              : T extends InterfaceDef<any>
                ? ComObject | null
                : T extends typeof VARIANT
                  ? unknown
                  : T extends typeof GUID
                    ? GUIDLike
                    : T extends ctypes.PointerTypeDef<infer B>
                      ? RawArgument<T> | (B extends typeof VARIANT ? unknown : B extends ctypes.SimpleCDataConstructor ? RawArgument<B> : never)
                      : RawArgument<T>;

type IsExactlyIDispatch<T> = T extends typeof IDispatch ? (typeof IDispatch extends T ? true : false) : false;

/**
 * JavaScript value returned for an output parameter of type `T`. Plain
 * `IDispatch` results are converted at run time (`GetBestInterface`, or a
 * dynamic object without type information): `AutomationObject`.
 */
export type OutValue<T, Self = unknown> = T extends BSTRType
  ? string
  : T extends DateType
    ? Date
    : T extends CurrencyType
      ? number
      : T extends LazyInterface<infer L>
        ? InterfacePointer<L> | null
        : T extends LPOLESTRType
          ? string | null
          : T extends RefGuidType
            ? GUID
            : T extends SafeArrayType
              ? unknown[] | null
              : T extends SelfMarker
                ? InterfacePointer<SelfDef<Self>> | null
                : T extends InterfaceDef<any>
                  ? (IsExactlyIDispatch<T> extends true ? AutomationObject : InterfacePointer<T>) | null
                  : T extends typeof VARIANT
                    ? unknown
                    : T extends ctypes.PointerTypeDef<infer B>
                      ? OutValue<B, Self>
                      : T extends ctypes.StructureClass
                        ? InstanceType<T>
                        : RawResult<T>;

type Flags<S extends ParamSpec> = S[0][number];
type IsOut<S extends ParamSpec> = "out" extends Flags<S> ? true : false;
type IsIn<S extends ParamSpec> = "lcid" extends Flags<S> ? false : "in" extends Flags<S> ? true : IsOut<S> extends true ? false : true;
type IsOptional<S extends ParamSpec> = "optional" extends Flags<S> ? true : S extends readonly [unknown, unknown, unknown, unknown] ? true : false;

/** JavaScript arguments of a COMMETHOD: `in` and `in/out` parameters, `lcid` omitted. */
export type ArgumentsOf<P extends readonly ParamSpec[], Self = unknown> = P extends readonly [infer H extends ParamSpec, ...infer R extends readonly ParamSpec[]]
  ? IsIn<H> extends true
    ? IsOptional<H> extends true
      ? [(InValue<H[1]> | undefined)?, ...ArgumentsOf<R, Self>]
      : [InValue<H[1]>, ...ArgumentsOf<R, Self>]
    : ArgumentsOf<R, Self>
  : [];

/** Values of the `out` parameters, in order. */
export type OutputsOf<P extends readonly ParamSpec[], Self = unknown> = P extends readonly [infer H extends ParamSpec, ...infer R extends readonly ParamSpec[]]
  ? IsOut<H> extends true
    ? [OutValue<H[1], Self>, ...OutputsOf<R, Self>]
    : OutputsOf<R, Self>
  : [];

/** No outputs: the HRESULT (or native result); one: its value; several: a tuple. */
export type CommethodResult<R, P extends readonly ParamSpec[], Self = unknown> =
  OutputsOf<P, Self> extends infer O extends readonly unknown[] ? (O extends readonly [] ? (R extends HRESULT ? number : RawResult<R>) : O extends readonly [infer Only] ? Only : O) : never;

/**
 * JavaScript signature of a COMMETHOD. At runtime a final `named({...})` also
 * passes arguments by name (untyped); typed code skips optional arguments with `undefined`.
 */
export type CommethodFunction<R, P extends readonly ParamSpec[], Self = unknown> = (...args: ArgumentsOf<P, Self>) => CommethodResult<R, P, Self>;

type KindOf<I extends readonly IdlFlag[]> = "propget" extends I[number] ? "propget" : "propput" extends I[number] ? "propput" : "propputref" extends I[number] ? "propputref" : "method";
type NameOf<K extends MethodKind, N extends string> = K extends "propget" ? `_get_${N}` : K extends "propput" ? `_set_${N}` : K extends "propputref" ? `_setref_${N}` : N;

/** Spec returned by {@link COMMETHOD}. */
export type CommethodSpec<I extends readonly IdlFlag[], R, N extends string, P extends readonly ParamSpec[]> = MethodSpec<NameOf<KindOf<I>, N>, (...args: any[]) => any, KindOf<I>, N> & {
  readonly params: readonly ParamInfo[];
  /** Type-level only. */
  readonly __commethod?: { readonly restype: R; readonly params: P };
};

// ── runtime conversions ───────────────────────────────────────────────

interface Holder {
  readonly arg: unknown;
  /** Interface results: handed over to a single-threaded home apartment after an asynchronous call. */
  readonly result?: { readonly kind: "pointer" | "variant" | "safearray"; readonly iid?: GUID };
  read(): unknown;
  /** Release what was not consumed by `read`. */
  free(): void;
}

interface Codec {
  readonly ffi: ctypes.FFIType;
  input(value: unknown, cleanup: Array<() => void>, self: InterfaceDef<any>): unknown;
  /** `null` when the type cannot be an output. */
  readonly output: ((initial: unknown, hasInitial: boolean, self: InterfaceDef<any>) => Holder) | null;
}

/** Description of a COMMETHOD parameter. */
export interface ParamInfo {
  readonly name: string;
  readonly type: ComParamType;
  readonly flags: readonly ParamFlag[];
  readonly input: boolean;
  readonly output: boolean;
  readonly lcid: boolean;
  readonly optional: boolean;
  readonly hasDefault: boolean;
  readonly defaultValue: unknown;
}

function isInterfaceDef(type: unknown): type is InterfaceDef<any> {
  return typeof type === "object" && type !== null && "Pointer" in type && "iid" in type && "slots" in type;
}

function isStructureClass(type: unknown): type is new (...args: any[]) => ctypes.Structure {
  return typeof type === "function" && type.prototype instanceof ctypes.Structure;
}

function isSimpleType(type: unknown): type is ctypes.SimpleCDataConstructor {
  return typeof type === "function" && (type as { _isSimpleCData?: boolean })._isSimpleCData === true;
}

function isPointerType(type: unknown): type is ctypes.PointerTypeDef {
  return typeof type === "object" && type !== null && "_type_" in type && "_pointerTo" in type;
}

function pointerHolder(arg: (out: InstanceType<typeof ctypes.c_void_p>) => unknown, read: (address: bigint | null) => unknown, release: (address: bigint) => void, initial: bigint | null = null): Holder {
  const out = new ctypes.c_void_p(initial ?? 0n);
  let consumed = false;
  return {
    arg: arg(out),
    read() {
      consumed = true;
      return read(out.value);
    },
    free() {
      if (!consumed && out.value) release(out.value);
    },
  };
}

const byReference = (out: InstanceType<typeof ctypes.c_void_p>) => ctypes.byref(out);
const freeBSTR = (address: bigint) => oleaut32().SysFreeString(address);

const bstrCodec: Codec = {
  ffi: ctypes.c_void_p,
  input(value, cleanup) {
    if (value === null || value === undefined) return null;
    const bstr = allocBSTR(String(value));
    cleanup.push(() => freeBSTR(bstr));
    return bstr;
  },
  output(initial, hasInitial) {
    const start = hasInitial && initial !== null && initial !== undefined ? allocBSTR(String(initial)) : null;
    return pointerHolder(
      byReference,
      (address) => {
        try {
          return readBSTR(address) ?? "";
        } finally {
          if (address) freeBSTR(address);
        }
      },
      freeBSTR,
      start,
    );
  },
};

const lpolestrCodec: Codec = {
  ffi: ctypes.c_void_p,
  input(value) {
    return value === null || value === undefined ? null : Buffer.from(`${String(value)}\0`, "utf16le");
  },
  output(_initial, hasInitial) {
    if (hasInitial) throw new TypeError("[in, out] LPOLESTR parameters are not supported");
    return pointerHolder(byReference, takeCoTaskMemString, (address) => ole32().CoTaskMemFree(address));
  },
};

const refGuidCodec: Codec = {
  ffi: ctypes.c_void_p,
  input(value) {
    return (value === undefined || value === null ? new GUID() : toGUID(value as GUIDLike))._buffer;
  },
  output: null,
};

function safeArrayCodec(vt: number): Codec {
  const create = (value: unknown) => (value instanceof Typed ? createSafeArray(value.vt & VT_TYPEMASK, value.value, value.dimensions) : createSafeArray(vt, value, 1));
  return {
    ffi: ctypes.c_void_p,
    input(value, cleanup) {
      if (value === null || value === undefined) return null;
      const psa = create(value);
      cleanup.push(() => destroySafeArray(psa));
      return psa;
    },
    output(initial, hasInitial) {
      const start = hasInitial && initial !== null && initial !== undefined ? create(initial) : null;
      const holder = pointerHolder(
        byReference,
        (psa) => {
          try {
            return readSafeArray(psa ?? 0n, vt);
          } finally {
            destroySafeArray(psa);
          }
        },
        destroySafeArray,
        start,
      );
      // Arrays of objects are handed over to a single-threaded home element by element.
      return { ...holder, result: { kind: "safearray" } };
    },
  };
}

/** Interface reference for `iface`: borrowed when the object already is one, else an owned QueryInterface result. */
function interfaceReference(value: unknown, iface: InterfaceDef<any>): { address: bigint; owned: boolean } | null {
  if (value === null || value === undefined) return null;
  const borrowed = borrowReference(value);
  if (!borrowed) throw new TypeError(`Expected a COM object for ${iface.name}`);
  if (derivesFrom(borrowed.interface, iface.iid)) return { address: borrowed.address, owned: false };
  const [hr, address] = rawQueryInterface(borrowed.address, iface.iid);
  checkHRESULT(hr, `QueryInterface(${iface.name})`);
  return { address: address!, owned: true };
}

const dateCodec: Codec = {
  ffi: ctypes.c_double,
  input: (value) => (value instanceof Date ? dateToOADate(value) : value === undefined || value === null ? 0 : Number(value)),
  output(initial, hasInitial) {
    const holder = new ctypes.c_double(hasInitial && initial !== undefined ? (initial instanceof Date ? dateToOADate(initial) : Number(initial)) : 0);
    return { arg: ctypes.byref(holder), read: () => oaDateToDate(holder.value), free: () => undefined };
  },
};

const currencyCodec: Codec = {
  ffi: ctypes.c_int64,
  input: (value) => (value === undefined || value === null ? 0n : toCurrency(value)),
  output(initial, hasInitial) {
    const holder = new ctypes.c_int64(hasInitial && initial !== undefined ? toCurrency(initial) : 0n);
    return { arg: ctypes.byref(holder), read: () => Number(currencyToString(BigInt(holder.value))), free: () => undefined };
  },
};

/** The interface an interface parameter refers to: a definition, `lazy(...)`, or `SELF` (the declaring interface). */
export function resolveParamInterface(kind: Extract<ParamKind, { kind: "interface" }>, self: InterfaceDef<any>): InterfaceDef<any> {
  return kind.resolve?.() ?? kind.iface ?? self;
}

let outputWrapper: ((value: unknown) => unknown) | null = null;

/**
 * Convert values returned by typed calls (COMMETHOD outputs, DISPMETHOD
 * results). @node-ctypes/comtypes/client installs `GetBestInterface` for plain
 * `IDispatch` pointers, as Python comtypes.client does with `wrap_outparam`.
 */
export function setOutputWrapper(wrapper: ((value: unknown) => unknown) | null): void {
  outputWrapper = wrapper;
}

export function wrapOutput(value: unknown): unknown {
  return outputWrapper ? outputWrapper(value) : value;
}

function interfaceCodec(type: InterfaceDef<any> | SelfMarker | LazyInterface): Codec {
  const resolve = (self: InterfaceDef<any>) => (isInterfaceDef(type) ? type : "resolve" in type ? type.resolve() : self);
  return {
    ffi: ctypes.c_void_p,
    input(value, cleanup, self) {
      const reference = interfaceReference(value, resolve(self));
      if (!reference) return null;
      if (reference.owned) cleanup.push(() => rawRelease(reference.address));
      return reference.address;
    },
    output(initial, hasInitial, self) {
      const iface = resolve(self);
      // [in, out]: the callee may release the input, so it receives its own reference.
      const reference = hasInitial ? interfaceReference(initial, iface) : null;
      if (reference && !reference.owned) rawAddRef(reference.address);
      return { ...pointerHolder(byReference, (address) => (address ? COMPointer.fromAddress(address, iface) : null), rawRelease, reference?.address ?? null), result: { kind: "pointer", iid: iface.iid } };
    },
  };
}

const variantCodec: Codec = {
  ffi: VARIANT,
  input(value, cleanup) {
    if (value instanceof VARIANT) return value;
    const variant = new VARIANT(value === undefined ? missing : value);
    cleanup.push(() => variant.clear());
    return variant;
  },
  output(initial, hasInitial) {
    const variant = hasInitial ? (initial instanceof VARIANT ? initial.copy() : new VARIANT(initial)) : new VARIANT();
    let consumed = false;
    return {
      arg: variant._buffer,
      result: { kind: "variant" },
      read() {
        consumed = true;
        try {
          return readVariant(variant._buffer);
        } finally {
          variant.clear();
        }
      },
      free() {
        if (!consumed) variant.clear();
      },
    };
  },
};

function structureCodec(type: new (...args: any[]) => ctypes.Structure): Codec {
  const guid = type === GUID || type.prototype instanceof GUID;
  const make = (value: unknown): ctypes.Structure => {
    if (value instanceof type) return value;
    if (value === undefined || value === null) return new type();
    if (guid && typeof value === "string") return new type(value);
    return (type as unknown as typeof ctypes.Structure).create(value as Record<string, unknown>);
  };
  return {
    ffi: type as ctypes.StructureClass,
    input: make,
    output(initial, hasInitial) {
      const holder = new type();
      if (hasInitial && initial !== undefined) holder._buffer.set(make(initial)._buffer);
      return { arg: holder._buffer, read: () => holder, free: () => undefined };
    },
  };
}

function simpleCodec(type: ctypes.SimpleCDataConstructor): Codec {
  return {
    ffi: type,
    input: (value) => (value === undefined ? new type().value : value),
    output(initial, hasInitial) {
      const holder = hasInitial && initial !== undefined ? new type(initial as never) : new type();
      return { arg: ctypes.byref(holder), read: () => holder.value, free: () => undefined };
    },
  };
}

// node-ctypes arrays are Proxies over a Buffer: instanceof Uint8Array, but not ArrayBuffer.isView.
function isRawPointerValue(value: unknown): boolean {
  return (
    value === null ||
    typeof value === "bigint" ||
    ArrayBuffer.isView(value) ||
    value instanceof Uint8Array ||
    value instanceof ArrayBuffer ||
    (typeof value === "object" && value !== null && ("_buffer" in value || "contents" in value || "_as_parameter_" in value))
  );
}

/** `POINTER(T)` input: raw pointers pass through, plain values are copied to a temporary `T`. */
function pointerCodec(type: ctypes.PointerTypeDef): Codec {
  const base = type._type_ as unknown;
  const target = base === VARIANT || isSimpleType(base) || isStructureClass(base) ? codecFor(base as ComParamType) : null;
  return {
    ffi: type,
    input(value, cleanup, self) {
      // A COM pointer for VARIANT* is a value (VT_DISPATCH / VT_UNKNOWN), not the address to pass.
      const raw = isRawPointerValue(value) && !(base === VARIANT && value instanceof COMPointer);
      if (!target || raw || (value === undefined && base !== VARIANT)) return value ?? null;
      const holder = target.output!(value, true, self);
      cleanup.push(() => holder.free());
      return holder.arg;
    },
    output: target?.output ?? null,
  };
}

function rawCodec(type: ctypes.FFIType): Codec {
  return { ffi: type, input: (value) => value, output: null };
}

const lcidCodec: Codec = { ffi: ctypes.c_uint32, input: () => defaultLCID, output: null };

/** How a parameter type is converted; shared by COMMETHOD calls and `COMObject` implementations. */
export type ParamKind =
  | { readonly kind: "bstr" }
  | { readonly kind: "lpolestr" }
  | { readonly kind: "refguid" }
  | { readonly kind: "safearray"; readonly vt: number }
  | { readonly kind: "interface"; readonly iface: InterfaceDef<any> | null; readonly resolve?: () => InterfaceDef<any> }
  | { readonly kind: "date" }
  | { readonly kind: "currency" }
  | { readonly kind: "variant" }
  | { readonly kind: "struct"; readonly type: new (...args: any[]) => ctypes.Structure }
  | { readonly kind: "simple"; readonly type: ctypes.SimpleCDataConstructor }
  | { readonly kind: "pointer"; readonly type: ctypes.PointerTypeDef; readonly base: ParamKind }
  | { readonly kind: "raw"; readonly type: ctypes.FFIType };

/** Classify a COMMETHOD parameter type (`iface: null` is `SELF`). */
export function classifyParamType(type: ComParamType): ParamKind {
  if (typeof type === "object" && type !== null && "__comtype" in type) {
    switch (type.__comtype) {
      case "BSTR":
        return { kind: "bstr" };
      case "LPOLESTR":
        return { kind: "lpolestr" };
      case "REFGUID":
        return { kind: "refguid" };
      case "SAFEARRAY":
        return { kind: "safearray", vt: type.vt };
      case "SELF":
        return { kind: "interface", iface: null };
      case "LAZY":
        return { kind: "interface", iface: null, resolve: type.resolve };
      case "DATE":
        return { kind: "date" };
      case "CURRENCY":
        return { kind: "currency" };
    }
  }
  if (type === VARIANT) return { kind: "variant" };
  if (isInterfaceDef(type)) return { kind: "interface", iface: type };
  if (isStructureClass(type)) return { kind: "struct", type };
  if (isSimpleType(type)) return { kind: "simple", type };
  if (isPointerType(type)) return { kind: "pointer", type, base: classifyParamType(type._type_ as ComParamType) };
  return { kind: "raw", type: type as ctypes.FFIType };
}

function codecFor(type: ComParamType): Codec {
  const kind = classifyParamType(type);
  switch (kind.kind) {
    case "bstr":
      return bstrCodec;
    case "lpolestr":
      return lpolestrCodec;
    case "refguid":
      return refGuidCodec;
    case "safearray":
      return safeArrayCodec(kind.vt);
    case "interface":
      return interfaceCodec(kind.resolve ? { __comtype: "LAZY", resolve: kind.resolve } : (kind.iface ?? SELF));
    case "date":
      return dateCodec;
    case "currency":
      return currencyCodec;
    case "variant":
      return variantCodec;
    case "struct":
      return structureCodec(kind.type);
    case "simple":
      return simpleCodec(kind.type);
    case "pointer":
      return pointerCodec(kind.type);
    case "raw":
      return rawCodec(kind.type);
  }
}

/** `REFGUID` out parameters return a GUID. */
function outputCodec(type: ComParamType): Codec {
  if (typeof type === "object" && type !== null && "__comtype" in type && type.__comtype === "REFGUID") return structureCodec(GUID);
  return codecFor(type);
}

// ── COMMETHOD ─────────────────────────────────────────────────────────

const PARAM_FLAGS = new Set<string>(["in", "out", "retval", "lcid", "optional"]);

interface Prepared {
  readonly native: unknown[];
  readonly holders: Holder[];
  readonly cleanup: Array<() => void>;
}

function release(prepared: Prepared): void {
  for (const holder of prepared.holders) holder.free();
  for (const cleanup of prepared.cleanup.reverse()) cleanup();
}

/**
 * Describe a vtable method with IDL flags. Python equivalent: `comtypes.COMMETHOD`.
 *
 * ```ts
 * COMMETHOD([], HRESULT, "GetClassID", [["out"], GUID, "pClassID"]),     // p.GetClassID(): GUID
 * COMMETHOD(["propget"], HRESULT, "Name", [["out", "retval"], BSTR, "p"]), // p.Name / p._get_Name()
 * COMMETHOD(["propput"], HRESULT, "Name", [["in"], BSTR, "value"]),       // p.Name = "x"
 * COMMETHOD([], HRESULT, "Find", [["in"], BSTR, "what"], [["in", "optional"], c_int32, "start", 0], [["out", "retval"], SELF, "found"]),
 * ```
 *
 * - `in` parameters are the JavaScript arguments (positional, or by name in a
 *   final `named({...})`); `optional` ones may be omitted (`VARIANT`: `missing`),
 *   a 4th element is the default value;
 * - `out` parameters are allocated and returned: one value, or an array for
 *   several (`retval` included); without outputs the HRESULT is returned;
 * - `in`+`out` parameters are passed and returned;
 * - `lcid` parameters are filled with the default LCID.
 */
export function COMMETHOD<const I extends readonly IdlFlag[], const R extends ctypes.FFIType | null, const N extends string, const P extends readonly ParamSpec[]>(idlflags: I, restype: R, name: N, ...params: P): CommethodSpec<I, R, N, P> {
  if (typeof name !== "string" || !name) throw new TypeError("COMMETHOD requires a method name");
  if (!Array.isArray(idlflags)) throw new TypeError(`${name}: idlflags must be an array`);
  const kinds = idlflags.filter((flag): flag is MethodKind => flag === "propget" || flag === "propput" || flag === "propputref");
  if (kinds.length > 1) throw new TypeError(`${name}: conflicting accessor flags ${kinds.join(", ")}`);
  const kind: MethodKind = kinds[0] ?? "method";
  const dispidFlag = idlflags.find((flag): flag is DispIdFlag => typeof flag === "object" && flag !== null && "dispid" in flag);

  const infos: ParamInfo[] = params.map((spec, index) => {
    if (!Array.isArray(spec) || spec.length < 2 || !Array.isArray(spec[0])) throw new TypeError(`${name}: parameter ${index} must be [flags, type, name?, default?]`);
    const flags = spec[0] as readonly ParamFlag[];
    for (const flag of flags) if (!PARAM_FLAGS.has(flag)) throw new TypeError(`${name}: unknown parameter flag ${JSON.stringify(flag)}`);
    const output = flags.includes("out");
    const lcid = flags.includes("lcid");
    const hasDefault = spec.length >= 4;
    return Object.freeze({
      name: spec[2] ?? `arg${index}`,
      type: spec[1],
      flags: Object.freeze([...flags]),
      input: !lcid && (flags.includes("in") || !output),
      output,
      lcid,
      optional: flags.includes("optional") || hasDefault,
      hasDefault,
      defaultValue: spec[3],
    });
  });
  const codecs = infos.map((info) => {
    if (info.lcid) return lcidCodec;
    const codec = info.output ? outputCodec(info.type) : codecFor(info.type);
    if (info.output && !codec.output) throw new TypeError(`${name}: unsupported out parameter type for ${info.name}`);
    return codec;
  });
  const argtypes = infos.map((info, index) => (info.output ? ctypes.c_void_p : codecs[index].ffi));
  const inputs = infos.filter((info) => info.input);
  const required = inputs.filter((info) => !info.optional).length;
  const hasOutputs = infos.some((info) => info.output);

  // Converts the arguments; on failure, releases what was already converted.
  const prepare = (args: readonly unknown[], self: InterfaceDef<any>): Prepared => {
    let positional = args;
    let byName: Map<string, unknown> | null = null;
    const last = args[args.length - 1];
    if (last instanceof NamedArguments) {
      positional = args.slice(0, -1);
      byName = new Map(Object.entries(last.values).map(([key, value]) => [key.toLowerCase(), value]));
      for (const key of byName.keys()) if (!inputs.some((info) => info.name.toLowerCase() === key)) throw new TypeError(`${name}: unknown argument ${JSON.stringify(key)}`);
    }
    if (positional.length > inputs.length) throw new TypeError(`${name}: expected at most ${inputs.length} arguments, got ${positional.length}`);
    const prepared: Prepared = { native: [], holders: [], cleanup: [] };
    try {
      let position = 0;
      infos.forEach((info, index) => {
        const codec = codecs[index];
        let value: unknown;
        let supplied = false;
        if (info.input) {
          if (position < positional.length) value = positional[position];
          position++;
          if (value === undefined && byName?.has(info.name.toLowerCase())) value = byName.get(info.name.toLowerCase());
          supplied = value !== undefined;
          if (!supplied && info.hasDefault) {
            value = info.defaultValue;
            supplied = true;
          }
          if (!supplied && !info.optional) throw new TypeError(`${name}: missing argument ${info.name}`);
        }
        if (info.output) {
          const holder = codec.output!(value, info.input && supplied, self);
          prepared.holders.push(holder);
          prepared.native.push(holder.arg);
        } else {
          prepared.native.push(codec.input(value, prepared.cleanup, self));
        }
      });
      return prepared;
    } catch (error) {
      release(prepared);
      throw error;
    }
  };
  const finish = (prepared: Prepared, result: unknown): unknown => {
    if (!hasOutputs) return result;
    const values = prepared.holders.map((holder) => wrapOutput(holder.read()));
    return values.length === 1 ? values[0] : values;
  };

  const marshal: MethodMarshal = (fn, address, args, self) => {
    const prepared = prepare(args, self);
    try {
      return finish(prepared, fn(address, ...prepared.native));
    } finally {
      release(prepared);
    }
  };
  // The converted arguments stay alive (and owned here) until the call completes.
  // Promise steps instead of an async function: this runs for every method call.
  const marshalAsync: MethodMarshalAsync = (run, args, self) => {
    let prepared: Prepared;
    let pending: Promise<unknown>;
    try {
      prepared = prepare(args, self);
    } catch (error) {
      return Promise.reject(error);
    }
    try {
      const results = prepared.holders.flatMap((holder) => (holder.result ? [{ location: ctypes.addressof(holder.arg as Buffer), ...holder.result }] : []));
      pending = run(prepared.native, results);
    } catch (error) {
      release(prepared);
      return Promise.reject(error);
    }
    return pending.then(
      (value) => {
        try {
          return finish(prepared, value);
        } finally {
          release(prepared);
        }
      },
      (error: unknown) => {
        release(prepared);
        throw error;
      },
    );
  };

  const methodName = kind === "propget" ? `_get_${name}` : kind === "propput" ? `_set_${name}` : kind === "propputref" ? `_setref_${name}` : name;
  return Object.freeze({
    name: methodName,
    restype,
    argtypes: Object.freeze(argtypes),
    idlflags: Object.freeze([...idlflags]),
    kind,
    property: name,
    dispid: dispidFlag?.dispid ?? null,
    marshal,
    marshalAsync,
    requiredArguments: required,
    inputArguments: inputs.length,
    params: Object.freeze(infos),
  }) as unknown as CommethodSpec<I, R, N, P>;
}
