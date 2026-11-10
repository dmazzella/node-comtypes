import { DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, VT_CY, VT_TYPEMASK } from "./automation/constants.ts";
import { getIDsOfNames, invoke } from "./automation/dispatch.ts";
import type { DispatchPointer } from "./automation/dispatch.ts";
import { NamedArguments } from "./automation/named.ts";
import { ByRef, Typed, borrowReference, safeArray, typed } from "./automation/variant.ts";
import { classifyParamType, resolveParamInterface, wrapOutput } from "./commethod.ts";
import type { ComParamType, DispIdFlag, IdlFlag, ParamFlag, ParamInfo, ParamKind, ParamSpec } from "./commethod.ts";
import { COMError } from "./errors.ts";
import { DISP_E_MEMBERNOTFOUND } from "./hresult.ts";
import { derivesFrom } from "./interface.ts";
import type { DispatchCall, InterfaceDef, MethodKind, MethodSpec } from "./interface.ts";
import { COMPointer } from "./pointer.ts";

type KindOf<I extends readonly IdlFlag[]> = "propget" extends I[number] ? "propget" : "propput" extends I[number] ? "propput" : "propputref" extends I[number] ? "propputref" : "method";
type NameOf<K extends MethodKind, N extends string> = K extends "propget" ? `_get_${N}` : K extends "propput" ? `_set_${N}` : K extends "propputref" ? `_setref_${N}` : N;

/** Spec of a dispatch member (no vtable slot). */
export type DispSpec<N extends string, K extends MethodKind, P extends string, R, Params extends readonly ParamSpec[], W extends boolean> = MethodSpec<N, (...args: any[]) => any, K, P> & {
  readonly writable: W;
  readonly params: readonly ParamInfo[];
  /** Declared value type of the result (or of the property). */
  readonly valueType: ComParamType | null;
  /** Type-level only. */
  readonly __disp?: { readonly restype: R; readonly params: Params };
};

const PARAM_FLAGS = new Set<string>(["in", "out", "retval", "lcid", "optional"]);

function parameters(member: string, params: readonly ParamSpec[]): ParamInfo[] {
  return params.map((spec, index) => {
    if (!Array.isArray(spec) || spec.length < 2 || !Array.isArray(spec[0])) throw new TypeError(`${member}: parameter ${index} must be [flags, type, name?, default?]`);
    const flags = spec[0] as readonly ParamFlag[];
    for (const flag of flags) if (!PARAM_FLAGS.has(flag)) throw new TypeError(`${member}: unknown parameter flag ${JSON.stringify(flag)}`);
    const hasDefault = spec.length >= 4;
    return Object.freeze({
      name: spec[2] ?? `arg${index}`,
      type: spec[1],
      flags: Object.freeze([...flags]),
      input: !flags.includes("lcid"),
      output: false,
      lcid: flags.includes("lcid"),
      optional: flags.includes("optional") || hasDefault,
      hasDefault,
      defaultValue: spec[3],
    });
  });
}

function dispidOf(member: string, idlflags: readonly IdlFlag[]): number {
  const flag = idlflags.find((candidate): candidate is DispIdFlag => typeof candidate === "object" && candidate !== null && "dispid" in candidate);
  if (!flag) throw new TypeError(`${member}: dispatch members require dispid(n)`);
  return flag.dispid;
}

/** Argument value for a declared parameter type (VARIANT conversion hints). */
function argument(kind: ParamKind | null, value: unknown): unknown {
  if (value === undefined || value === null || value instanceof Typed || value instanceof ByRef || !kind) return value;
  switch (kind.kind) {
    case "currency":
      return typed(VT_CY, value);
    case "bstr":
    case "lpolestr":
      return typeof value === "string" ? value : String(value);
    case "safearray":
      return Array.isArray(value) || ArrayBuffer.isView(value) ? safeArray(value as ArrayLike<unknown>, { vt: kind.vt & VT_TYPEMASK }) : value;
    default:
      return value;
  }
}

/** Typed result: interface results are queried for the declared interface. */
function result(kind: ParamKind | null, value: unknown, self: InterfaceDef<any>): unknown {
  if (!kind || kind.kind !== "interface" || !(value instanceof COMPointer)) return value;
  const target = resolveParamInterface(kind, self);
  if (derivesFrom(value.interface, target.iid)) return value;
  const typedPointer = value.tryQueryInterface(target);
  if (!typedPointer) return value;
  value.Release();
  return typedPointer;
}

function isObject(value: unknown): boolean {
  return borrowReference(value) !== null;
}

function makeCall(member: string, dispid: number, kind: MethodKind, infos: readonly ParamInfo[], valueType: ComParamType | null): DispatchCall {
  const kinds = infos.map((info) => classifyParamType(info.type));
  const resultKind = valueType === null ? null : classifyParamType(valueType);
  const inputs = infos.filter((info) => info.input);
  return (pointer, args, self, put) => {
    let positional = [...args];
    let named: { dispids: number[]; values: unknown[] } | undefined;
    const last = positional[positional.length - 1];
    if (last instanceof NamedArguments) {
      positional = positional.slice(0, -1);
      const names = Object.keys(last.values);
      if (names.length) {
        const ids = getIDsOfNames(pointer as DispatchPointer, [member, ...names]);
        named = { dispids: ids.slice(1), values: names.map((name) => last.values[name]) };
      }
    }
    const writing = put || kind === "propput" || kind === "propputref";
    if (!writing && positional.length > inputs.length) throw new TypeError(`${member}: expected at most ${inputs.length} arguments, got ${positional.length}`);
    const converted = positional.map((value, index) => {
      const info = inputs[index];
      if (value === undefined && info?.hasDefault) return info.defaultValue;
      return argument(info ? kinds[infos.indexOf(info)] : resultKind, value);
    });
    for (let index = converted.length; index < inputs.length; index++) {
      const info = inputs[index];
      if (info.hasDefault) converted[index] = info.defaultValue;
      else if (!info.optional && !named) throw new TypeError(`${member}: missing argument ${info.name}`);
    }
    const dispatch = pointer as DispatchPointer;
    if (writing) {
      const value = converted[converted.length - 1];
      const flags = kind === "propputref" || (kind !== "propput" && isObject(value)) ? DISPATCH_PROPERTYPUTREF : DISPATCH_PROPERTYPUT;
      try {
        invoke(dispatch, dispid, flags, converted, { named, name: member });
      } catch (error) {
        if (!(flags === DISPATCH_PROPERTYPUTREF && kind !== "propputref" && error instanceof COMError && error.hresult === DISP_E_MEMBERNOTFOUND)) throw error;
        invoke(dispatch, dispid, DISPATCH_PROPERTYPUT, converted, { named, name: member });
      }
      return undefined;
    }
    const flags = kind === "propget" ? DISPATCH_PROPERTYGET : DISPATCH_METHOD;
    const value = invoke(dispatch, dispid, flags, converted, { named, name: member });
    return valueType === null ? undefined : wrapOutput(result(resultKind, value, self));
  };
}

/**
 * A method of a dispinterface, called through `IDispatch::Invoke`.
 * Python equivalent: `comtypes.DISPMETHOD`.
 *
 * ```ts
 * DISPMETHOD([dispid(1)], VARIANT, "Item", [["in"], VARIANT, "index"])     // obj.Item(1)
 * DISPMETHOD([dispid(2), "propget"], BSTR, "Caption")                    // obj.Caption
 * DISPMETHOD([dispid(3)], null, "Refresh")                               // returns undefined
 * ```
 *
 * `restype` is the value type of the result (`null` for none); interface
 * results are queried for the declared interface. Parameters are `in`
 * values (`optional`, defaults); a final `named({...})` passes them by name.
 */
export function DISPMETHOD<const I extends readonly IdlFlag[], const R extends ComParamType | null, const N extends string, const P extends readonly ParamSpec[]>(
  idlflags: I,
  restype: R,
  name: N,
  ...params: P
): DispSpec<NameOf<KindOf<I>, N>, KindOf<I>, N, R, P, false> {
  if (typeof name !== "string" || !name) throw new TypeError("DISPMETHOD requires a name");
  const kinds = idlflags.filter((flag): flag is MethodKind => flag === "propget" || flag === "propput" || flag === "propputref");
  if (kinds.length > 1) throw new TypeError(`${name}: conflicting accessor flags ${kinds.join(", ")}`);
  const kind: MethodKind = kinds[0] ?? "method";
  const dispid = dispidOf(name, idlflags);
  const infos = parameters(name, params);
  const inputs = infos.filter((info) => info.input);
  return Object.freeze({
    name: kind === "propget" ? `_get_${name}` : kind === "propput" ? `_set_${name}` : kind === "propputref" ? `_setref_${name}` : name,
    restype: null,
    argtypes: Object.freeze([]),
    idlflags: Object.freeze([...idlflags]),
    kind,
    property: name,
    dispid,
    marshal: null,
    requiredArguments: inputs.filter((info) => !info.optional).length,
    inputArguments: inputs.length,
    dispatch: makeCall(name, dispid, kind, infos, restype),
    writable: false,
    params: Object.freeze(infos),
    valueType: restype,
  }) as unknown as DispSpec<NameOf<KindOf<I>, N>, KindOf<I>, N, R, P, false>;
}

/**
 * A property of a dispinterface (`PROPERTYGET`, and `PROPERTYPUT` unless
 * `readonly`). Python equivalent: `comtypes.DISPPROPERTY`.
 *
 * ```ts
 * DISPPROPERTY([dispid(1)], BSTR, "Name")              // obj.Name, obj.Name = "x"
 * DISPPROPERTY([dispid(2), "readonly"], c_int32, "Count")
 * ```
 */
export function DISPPROPERTY<const I extends readonly IdlFlag[], const T extends ComParamType, const N extends string>(idlflags: I, type: T, name: N): DispSpec<`_get_${N}`, "propget", N, T, readonly [], "readonly" extends I[number] ? false : true> {
  if (typeof name !== "string" || !name) throw new TypeError("DISPPROPERTY requires a name");
  const dispid = dispidOf(name, idlflags);
  const writable = !idlflags.includes("readonly");
  const get = makeCall(name, dispid, "propget", [], type);
  const set = makeCall(name, dispid, "method", [], type);
  return Object.freeze({
    name: `_get_${name}`,
    restype: null,
    argtypes: Object.freeze([]),
    idlflags: Object.freeze([...idlflags]),
    kind: "propget",
    property: name,
    dispid,
    marshal: null,
    requiredArguments: 0,
    inputArguments: 0,
    dispatch: ((pointer, args, self, put) => (put ? set(pointer, args, self, true) : get(pointer, args, self))) as DispatchCall,
    writable,
    params: Object.freeze([]),
    valueType: type,
  }) as unknown as DispSpec<`_get_${N}`, "propget", N, T, readonly [], "readonly" extends I[number] ? false : true>;
}
