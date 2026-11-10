import * as ctypes from "node-ctypes";
import { GUID, toGUID } from "./guid.ts";
import type { GUIDLike } from "./guid.ts";
import { COMPointer, createSlot, invokeSlot } from "./pointer.ts";
import type { Slot, SlotOwner } from "./pointer.ts";
import type { ArgumentsOf, CommethodFunction, OutValue, ParamSpec } from "./commethod.ts";

/** HRESULT restype marker: such methods raise {@link COMError} on failure. */
export const HRESULT = ctypes.HRESULT;
export type HRESULT = typeof ctypes.HRESULT;

/** `method`, or the accessor kind of a `propget` / `propput` / `propputref` method. */
export type MethodKind = "method" | "propget" | "propput" | "propputref";

/**
 * Converts JavaScript arguments, calls the bound vtable function and converts
 * the results (COMMETHOD). `self` is the interface declaring the method.
 */
export type MethodMarshal = (fn: (...args: any[]) => any, address: bigint, args: readonly unknown[], self: InterfaceDef<any>) => unknown;

/** Calls a dispatch member (DISPMETHOD / DISPPROPERTY) through `IDispatch::Invoke`; `put` for property writes. */
export type DispatchCall = (pointer: COMPointer, args: readonly unknown[], self: InterfaceDef<any>, put?: boolean) => unknown;

/**
 * A vtable method. `F` is the JavaScript signature for STDMETHOD; COMMETHOD
 * specs carry their parameter specs (`__params`) so that `SELF` can be
 * resolved by the declaring interface.
 */
export interface MethodSpec<N extends string = string, F extends (...args: any[]) => any = (...args: any[]) => any, K extends MethodKind = MethodKind, P extends string = string> {
  /** Vtable method name: `_get_X` / `_set_X` / `_setref_X` for property accessors, as comtypes. */
  readonly name: N;
  readonly restype: ctypes.FFIType | null;
  /** Native argument types, without `this`. */
  readonly argtypes: readonly ctypes.FFIType[];
  readonly idlflags: readonly unknown[];
  readonly kind: K;
  /** Property name for accessors, otherwise the method name. */
  readonly property: P;
  /** `dispid(n)` IDL flag, if any. */
  readonly dispid: number | null;
  readonly marshal: MethodMarshal | null;
  /** Number of required JavaScript arguments (COMMETHOD), `null` for STDMETHOD. */
  readonly requiredArguments: number | null;
  /** Number of JavaScript input arguments (COMMETHOD), `null` for STDMETHOD. */
  readonly inputArguments: number | null;
  /** Dispatch members (DISPMETHOD / DISPPROPERTY): no vtable slot, called through `IDispatch::Invoke`. */
  readonly dispatch?: DispatchCall | null;
  /** DISPPROPERTY without `readonly`: also writable. */
  readonly writable?: boolean;
  /** Type-level only. */
  readonly __fn?: F;
}

/** JavaScript result of a raw method: the HRESULT (`S_OK`/`S_FALSE`) or the converted return value. */
export type MethodResult<R, A extends readonly ctypes.FFIType[]> = R extends HRESULT ? number : R extends ctypes.FFIType | null ? ReturnType<ctypes.TypedFFIFunction<A, R>> : never;

/** Raw signature of a STDMETHOD: native arguments, HRESULT or native result. */
export type RawFunction<R, A extends readonly ctypes.FFIType[]> = (...args: Parameters<ctypes.TypedFFIFunction<A, R extends ctypes.FFIType | null ? R : never>>) => MethodResult<R, A>;

/**
 * Describe a vtable method with native arguments. Python equivalent:
 * `comtypes.STDMETHOD`. See {@link COMMETHOD} for converted parameters.
 *
 * ```ts
 * STDMETHOD(HRESULT, "SetPath", [c_wchar_p])
 * ```
 */
export function STDMETHOD<const R extends ctypes.FFIType | null, const N extends string, const A extends readonly ctypes.FFIType[] = readonly []>(restype: R, name: N, argtypes?: A): MethodSpec<N, RawFunction<R, A>, "method", N> {
  if (typeof name !== "string" || !name) throw new TypeError("STDMETHOD requires a method name");
  return Object.freeze({
    name,
    restype,
    argtypes: Object.freeze([...(argtypes ?? [])]),
    idlflags: Object.freeze([]),
    kind: "method",
    property: name,
    dispid: null,
    marshal: null,
    requiredArguments: null,
    inputArguments: null,
  });
}

/**
 * Signature of a method spec within interface `Self`: COMMETHOD specs are
 * computed from their parameter specs, with `SELF` resolved to `Self`.
 */
export type MethodFunction<S, Self = unknown> = S extends { readonly __commethod?: { readonly restype: infer R; readonly params: infer P extends readonly ParamSpec[] } }
  ? CommethodFunction<R, P, Self>
  : S extends { readonly __disp?: { readonly restype: infer R; readonly params: infer P extends readonly ParamSpec[] } }
    ? (...args: ArgumentsOf<P, Self>) => [R] extends [null] ? undefined : OutValue<R, Self>
    : S extends { readonly __fn?: infer F }
      ? NonNullable<F>
      : never;

type Getter<S> = S extends { readonly kind: "propget" } ? S : never;
type Setter<S> = S extends { readonly kind: "propput" | "propputref" } ? S : never;
type AccessorNames<M extends readonly MethodSpec[], Self> = { [S in M[number] as S extends Getter<S> ? (MethodFunction<S, Self> extends () => any ? S["property"] : never) : never]: true };
type SetterNames<M extends readonly MethodSpec[]> = Setter<M[number]>["property"] | Extract<M[number], { readonly writable: true }>["property"];

/** Properties of `propget` methods without required arguments, writable with a matching `propput`/`propputref`. */
export type PropertiesOf<M extends readonly MethodSpec[], Self = unknown> = {
  readonly [S in M[number] as S extends Getter<S> ? (MethodFunction<S, Self> extends () => any ? (S["property"] extends SetterNames<M> ? never : S["property"]) : never) : never]: ReturnType<MethodFunction<S, Self>>;
} & {
  -readonly [S in M[number] as S extends Getter<S> ? (MethodFunction<S, Self> extends () => any ? (S["property"] extends SetterNames<M> ? S["property"] : never) : never) : never]: ReturnType<MethodFunction<S, Self>>;
} & {
  /** `propget` with required arguments: a method named after the property. */
  [S in M[number] as S extends Getter<S> ? (MethodFunction<S, Self> extends () => any ? never : S["property"]) : never]: MethodFunction<S, Self>;
} & {
  /** Write-only properties. */
  -readonly [S in M[number] as S extends Setter<S> ? (S["property"] extends keyof AccessorNames<M, Self> ? never : MethodFunction<S, Self> extends (value: any) => any ? S["property"] : never) : never]: Parameters<MethodFunction<S, Self>>[0];
};

/** Method and property signatures of a `methods` tuple, with `SELF` resolved to `Self`. */
export type MethodsOf<M extends readonly MethodSpec[], Self = unknown> = { [S in M[number] as S["name"]]: MethodFunction<S, Self> } & PropertiesOf<M, Self>;

/** Method signatures of an interface, including the inherited ones. */
export type InterfaceMethods<I> = I extends { readonly __methods?: infer T } ? NonNullable<T> : {};

/** A typed pointer to interface `I`: {@link COMPointer} plus the interface methods and properties. */
export type InterfacePointer<I extends InterfaceDef<any> = InterfaceDef> = COMPointer<I> & Omit<InterfaceMethods<I>, keyof COMPointer>;

/** Pointer class generated for an interface. */
export type PointerClass = new (address: bigint, iface: InterfaceDef<any>) => COMPointer;

/** Interface definition returned by {@link COMInterface}. */
export interface InterfaceDef<TMethods extends object = {}> {
  readonly name: string;
  readonly iid: GUID;
  readonly base: InterfaceDef<any> | null;
  /** Methods declared by this interface. */
  readonly methods: readonly MethodSpec[];
  /** All vtable methods, inherited first. */
  readonly allMethods: readonly MethodSpec[];
  /** Method name to vtable slot. */
  readonly methodIndex: ReadonlyMap<string, number>;
  readonly slots: readonly Slot[];
  /** Dispatch members (DISPMETHOD / DISPPROPERTY), inherited first: called through `IDispatch::Invoke`. */
  readonly dispMethods: readonly MethodSpec[];
  /** Pointer class: a {@link COMPointer} subclass with the interface methods. */
  readonly Pointer: PointerClass;
  /** Type-level only: the method signatures. Never set at runtime. */
  readonly __methods?: TMethods;
}

/** Definition type of `COMInterface({ extends: B, methods: M, dispMethods: D })`; recursive so that `SELF` refers to it. */
export interface ComInterfaceDef<B, M extends readonly MethodSpec[], D extends readonly MethodSpec[] = readonly []> extends InterfaceDef<InterfaceMethods<B> & MethodsOf<M, ComInterfaceDef<B, M, D>> & MethodsOf<D, ComInterfaceDef<B, M, D>>> {}

export interface InterfaceSpec<M extends readonly MethodSpec[], B extends InterfaceDef<any> | undefined, D extends readonly MethodSpec[] = readonly []> {
  readonly name: string;
  readonly iid: GUIDLike;
  /** Base interface; every COM interface ultimately derives from `IUnknown`. */
  readonly extends?: B;
  readonly methods?: M;
  /**
   * Dispatch members (`DISPMETHOD`, `DISPPROPERTY`) of a dispinterface: they
   * have no vtable slot and are called through `IDispatch::Invoke`.
   * Python equivalent: `_disp_methods_`.
   */
  readonly dispMethods?: D;
  /** Add to the IID registry used for results (default `true`); synthetic definitions opt out. */
  readonly register?: boolean;
}

function isComObject(value: unknown): boolean {
  if (value instanceof COMPointer) return true;
  return value !== null && (typeof value === "object" || typeof value === "function") && (value as Record<symbol, unknown>)[Symbol.for("node-comtypes.pointer")] instanceof COMPointer;
}

function installProperties(prototype: object, methods: readonly MethodSpec[]): void {
  const groups = new Map<string, { get?: MethodSpec; put?: MethodSpec; putref?: MethodSpec }>();
  for (const method of methods) {
    if (method.kind === "method") continue;
    const group = groups.get(method.property) ?? {};
    group[method.kind === "propget" ? "get" : method.kind === "propput" ? "put" : "putref"] = method;
    groups.set(method.property, group);
  }
  for (const [property, { get, put, putref }] of groups) {
    if (property in COMPointer.prototype || Object.hasOwn(prototype, property)) continue;
    if (get && (get.requiredArguments ?? get.argtypes.length) > 0) {
      // Parameterized property: a method; puts stay available as _set_X(...args, value).
      Object.defineProperty(prototype, property, {
        value: function (this: Record<string, (...args: unknown[]) => unknown>, ...args: unknown[]) {
          return this[get.name](...args);
        },
        writable: true,
        configurable: true,
      });
      continue;
    }
    const settable = [put, putref].filter((setter): setter is MethodSpec => setter !== undefined && (setter.inputArguments ?? setter.argtypes.length) >= 1 && (setter.requiredArguments ?? setter.argtypes.length) <= 1);
    Object.defineProperty(prototype, property, {
      get: get
        ? function (this: Record<string, () => unknown>) {
            return this[get.name]();
          }
        : undefined,
      set: settable.length
        ? function (this: Record<string, (value: unknown) => unknown>, value: unknown) {
            const setter = put && putref ? (isComObject(value) ? putref : put) : settable[0];
            this[setter.name](value);
          }
        : undefined,
      configurable: true,
    });
  }
}

/** Install dispatch members: methods calling `IDispatch::Invoke`, plus the property accessors. */
function installDispatchMembers(prototype: object, members: readonly MethodSpec[], owner: SlotOwner): void {
  const accessors: MethodSpec[] = [];
  for (const member of members) {
    const call = member.dispatch!;
    if (!(member.name in COMPointer.prototype)) {
      Object.defineProperty(prototype, member.name, {
        value: {
          [member.name](this: COMPointer, ...args: unknown[]) {
            return call(this, args, owner.definition!);
          },
        }[member.name],
        writable: true,
        configurable: true,
      });
    }
    accessors.push(member);
    if (member.kind === "propget" && member.writable) {
      // DISPPROPERTY: a _set_X writer and a propput accessor spec.
      const name = `_set_${member.property}`;
      Object.defineProperty(prototype, name, {
        value: function (this: COMPointer, value: unknown) {
          return call(this, [value], owner.definition!, true);
        },
        writable: true,
        configurable: true,
      });
      accessors.push({ ...member, name, kind: "propput", requiredArguments: 1, inputArguments: 1 });
    }
  }
  installProperties(prototype, accessors);
}

/**
 * Define a COM interface. The `methods` order is the vtable order, after the
 * base interface methods. Python equivalent: a `comtypes.IUnknown` subclass
 * with `_iid_` and `_methods_`.
 *
 * ```ts
 * const IPersist = COMInterface({
 *   name: "IPersist",
 *   iid: "{0000010C-0000-0000-C000-000000000046}",
 *   extends: IUnknown,
 *   methods: [COMMETHOD([], HRESULT, "GetClassID", [["out"], GUID, "pClassID"])],
 * });
 * ```
 *
 * `propget` methods without required arguments also become properties
 * (`p.Name`, assignable with a `propput`/`propputref`); with arguments, a
 * method named after the property.
 */
export function COMInterface<const M extends readonly MethodSpec[] = readonly [], B extends InterfaceDef<any> | undefined = undefined, const D extends readonly MethodSpec[] = readonly []>(spec: InterfaceSpec<M, B, D>): ComInterfaceDef<B, M, D> {
  if (typeof spec.name !== "string" || !spec.name) throw new TypeError("COMInterface requires a name");
  const base: InterfaceDef<any> | null = spec.extends ?? null;
  const iid = new GUID(toGUID(spec.iid));
  const methods = Object.freeze([...(spec.methods ?? [])]);
  const own = new Set<string>();
  for (const method of methods) {
    if (own.has(method.name)) throw new TypeError(`${spec.name}: duplicate method ${method.name}`);
    own.add(method.name);
  }
  const dispOwn = Object.freeze([...(spec.dispMethods ?? [])]);
  for (const method of dispOwn) {
    if (!method.dispatch) throw new TypeError(`${spec.name}: dispMethods takes DISPMETHOD / DISPPROPERTY specs (${method.name})`);
    if (own.has(method.name)) throw new TypeError(`${spec.name}: duplicate member ${method.name}`);
    own.add(method.name);
  }
  for (const method of methods) if (method.dispatch) throw new TypeError(`${spec.name}: ${method.name} is a dispatch member: put it in dispMethods`);
  const owner: SlotOwner = { name: spec.name, iid, definition: null };
  const slots = Object.freeze([...(base?.slots ?? []), ...methods.map((method) => createSlot(method, owner))]);
  const allMethods = Object.freeze(slots.map((slot) => slot.spec));
  const methodIndex = new Map<string, number>();
  allMethods.forEach((method, slot) => methodIndex.set(method.name, slot));

  const Parent: PointerClass = base?.Pointer ?? COMPointer;
  const Pointer = class extends Parent {};
  Object.defineProperty(Pointer, "name", { value: `POINTER(${spec.name})` });
  const first = slots.length - methods.length;
  methods.forEach((method, index) => {
    // COMPointer members (QueryInterface, AddRef, Release, ...) take precedence,
    // as in Python comtypes; the raw slot stays reachable through callMethod().
    if (method.name in COMPointer.prototype) return;
    const slot = first + index;
    const fn = {
      [method.name](this: COMPointer, ...args: unknown[]) {
        return invokeSlot(this, slot, args);
      },
    }[method.name];
    Object.defineProperty(Pointer.prototype, method.name, { value: fn, writable: true, configurable: true });
  });
  installProperties(Pointer.prototype, methods);
  installDispatchMembers(Pointer.prototype, dispOwn, owner);
  const dispMethods = Object.freeze([...(base?.dispMethods ?? []), ...dispOwn]);

  const definition = Object.freeze({ name: spec.name, iid, base, methods, allMethods, methodIndex, slots, dispMethods, Pointer });
  owner.definition = definition;
  if (spec.register !== false) registry.set(iid.toString(), definition);
  return definition as ComInterfaceDef<B, M, D>;
}

// IID -> most recent definition. Python equivalent: `comtypes.com_interface_registry`.
const registry = new Map<string, InterfaceDef<any>>();

/** The interface defined with `iid`, if any (the most recent definition wins). */
export function findInterface(iid: GUIDLike): InterfaceDef<any> | undefined {
  return registry.get(toGUID(iid).toString());
}

/** Whether `iface` is `base` or derives from it. */
export function derivesFrom(iface: InterfaceDef<any>, base: GUIDLike): boolean {
  const iid = toGUID(base);
  for (let current: InterfaceDef<any> | null = iface; current; current = current.base) if (current.iid.equals(iid)) return true;
  return false;
}
