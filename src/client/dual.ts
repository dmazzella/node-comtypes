// Dual interfaces reached through IDispatch. Across apartments a dual
// interface needs a registered proxy/stub for its vtable; when QueryInterface
// fails there, the object is still typed: the same members, from the same
// type library, called through IDispatch::Invoke with their DISPIDs.
import { dispid } from "../commethod.ts";
import type { ComParamType, ParamInfo, ParamSpec } from "../commethod.ts";
import { DISPMETHOD } from "../dispmethod.ts";
import { COMInterface, derivesFrom } from "../interface.ts";
import type { InterfaceDef, MethodSpec } from "../interface.ts";
import { IDispatch } from "../interfaces/dispatch.ts";
import { COMPointer } from "../pointer.ts";

const variants = new WeakMap<InterfaceDef<any>, InterfaceDef<any> | null>();

function parameters(spec: MethodSpec): readonly ParamInfo[] | null {
  return (spec as MethodSpec & { readonly params?: readonly ParamInfo[] }).params ?? null;
}

/** The dispatch member equivalent to a dual interface method, or `null` when it has no DISPID. */
function dispatchMember(spec: MethodSpec): MethodSpec | null {
  const params = parameters(spec);
  if (spec.dispid === null || !params) return null;
  const inputs: ParamSpec[] = params.filter((param) => param.input || param.lcid).map((param) => (param.hasDefault ? [param.flags, param.type, param.name, param.defaultValue] : [param.flags, param.type, param.name]) as unknown as ParamSpec);
  const retval = params.find((param) => param.output && param.flags.includes("retval"));
  const flags = spec.kind === "method" ? [dispid(spec.dispid)] : [dispid(spec.dispid), spec.kind];
  return DISPMETHOD(flags as never, (retval?.type ?? null) as ComParamType | null, spec.property, ...(inputs as never[])) as unknown as MethodSpec;
}

/**
 * `iface` (a dual interface from a type library) as a dispinterface: its
 * members called through `IDispatch::Invoke`. `null` when a member has no
 * DISPID (not a dual interface).
 */
export function dispatchVariant(iface: InterfaceDef<any>): InterfaceDef<any> | null {
  if (variants.has(iface)) return variants.get(iface)!;
  let variant: InterfaceDef<any> | null = null;
  if (iface !== IDispatch && derivesFrom(iface, IDispatch.iid)) {
    const own = iface.allMethods.slice(IDispatch.allMethods.length);
    const members = own.map(dispatchMember);
    if (own.length && members.every((member) => member !== null)) {
      variant = COMInterface({ name: iface.name, iid: iface.iid, extends: IDispatch, register: false, dispMethods: members as MethodSpec[] });
    }
  }
  variants.set(iface, variant);
  return variant;
}

/**
 * `iface` of `object`: QueryInterface, else (dual interfaces whose vtable
 * cannot be marshaled) its dispatch variant on the object's `IDispatch`.
 * `null` when neither is available.
 */
export function queryOrDispatch(object: COMPointer, iface: InterfaceDef<any>): COMPointer | null {
  const typed = object.tryQueryInterface(iface);
  if (typed) return typed;
  const variant = iface.allMethods.length === IDispatch.allMethods.length && derivesFrom(iface, IDispatch.iid) ? iface : dispatchVariant(iface);
  if (!variant) return null;
  const dispatch = object.tryQueryInterface(IDispatch);
  if (!dispatch) return null;
  using owned = dispatch;
  return COMPointer.fromAddress(owned.address, variant, { addRef: true });
}
