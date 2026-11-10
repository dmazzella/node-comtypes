// Type library built at run time (CreateTypeLib2) for the registration tests:
//
//   [oleautomation] interface ICounter : IUnknown { Add(long a, long b, [out, retval] long*); [propget] Name([out, retval] BSTR*); }
//   [dual, oleautomation] interface IGreeter : IDispatch { Greet(BSTR name, [out, retval] BSTR*); [propget] Count([out, retval] long*); }
//   coclass Greeter { [default] interface IGreeter; interface ICounter; }
import * as ctypes from "node-ctypes";
import {
  CC_STDCALL,
  COMPointer,
  ELEMDESC,
  FUNCDESC,
  FUNC_PUREVIRTUAL,
  GUID,
  IDispatch,
  IMPLTYPEFLAG_FDEFAULT,
  INVOKE_FUNC,
  INVOKE_PROPERTYGET,
  ITypeInfo,
  IUnknown,
  PARAMFLAG_FIN,
  PARAMFLAG_FOUT,
  PARAMFLAG_FRETVAL,
  TKIND_COCLASS,
  TKIND_INTERFACE,
  TYPEDESC,
  TYPEFLAG_FCANCREATE,
  TYPEFLAG_FDUAL,
  TYPEFLAG_FOLEAUTOMATION,
  VT_BSTR,
  VT_HRESULT,
  VT_I4,
  VT_PTR,
  createTypeLib,
  loadTypeLib,
} from "@node-ctypes/comtypes";
import type { InterfacePointer } from "@node-ctypes/comtypes";

const POINTER_SIZE = ctypes.sizeof(ctypes.c_void_p);

export interface TestTypeLibIds {
  readonly libid: string;
  readonly counter: string;
  readonly greeter: string;
  readonly clsid: string;
}

interface Param {
  readonly name: string;
  readonly vt: number;
  readonly flags: number;
  /** A pointer to `vt` (`[out]` parameters). */
  readonly pointer?: boolean;
}

interface Func {
  readonly name: string;
  readonly invkind: number;
  readonly params: readonly Param[];
}

const outLong = (name: string): Param => ({ name, vt: VT_I4, flags: PARAMFLAG_FOUT | PARAMFLAG_FRETVAL, pointer: true });
const outBSTR = (name: string): Param => ({ name, vt: VT_BSTR, flags: PARAMFLAG_FOUT | PARAMFLAG_FRETVAL, pointer: true });

/** Fill `desc` with `vt`, or a pointer to it. Returns buffers to keep alive. */
function describe(desc: TYPEDESC, vt: number, pointer = false): Buffer[] {
  if (!pointer) {
    desc.vt = vt;
    return [];
  }
  const target = new TYPEDESC();
  target.vt = vt;
  desc.vt = VT_PTR;
  desc.lptdesc = ctypes.addressof(target._buffer);
  return [target._buffer];
}

function addFunctions(info: InterfacePointer<typeof import("@node-ctypes/comtypes").ICreateTypeInfo>, functions: readonly Func[], firstSlot: number): void {
  functions.forEach((func, index) => {
    const keep: Buffer[] = [];
    const params = Buffer.alloc(ctypes.sizeof(ELEMDESC) * Math.max(1, func.params.length));
    func.params.forEach((param, at) => {
      const elem = new ELEMDESC(params.subarray(at * ctypes.sizeof(ELEMDESC)));
      keep.push(...describe(elem.tdesc, param.vt, param.pointer));
      elem.paramdesc.wParamFlags = param.flags;
    });
    const desc = new FUNCDESC();
    desc.memid = index + 1;
    desc.lprgelemdescParam = ctypes.addressof(params);
    desc.funckind = FUNC_PUREVIRTUAL;
    desc.invkind = func.invkind;
    desc.callconv = CC_STDCALL;
    desc.cParams = func.params.length;
    desc.oVft = (firstSlot + index) * POINTER_SIZE;
    desc.elemdescFunc.tdesc.vt = VT_HRESULT;
    info.AddFuncDesc(index, desc);
    const names = [func.name, ...func.params.map((param) => param.name)].map((name) => Buffer.from(`${name}\0`, "utf16le"));
    const array = Buffer.alloc(POINTER_SIZE * names.length);
    names.forEach((name, at) => array.writeBigUInt64LE(BigInt(ctypes.addressof(name)), at * POINTER_SIZE));
    info.SetFuncAndParamNames(index, array, names.length);
    void keep;
  });
}

/** Write the test type library to `file`. */
export function buildTestTypeLib(file: string, ids: TestTypeLibIds): void {
  using stdole = loadTypeLib("stdole2.tlb");
  const typeInfoOf = (iid: GUID) => {
    const out = new ctypes.c_void_p();
    stdole.GetTypeInfoOfGuid(iid, ctypes.byref(out));
    return COMPointer.fromAddress(out.value!, ITypeInfo);
  };
  using library = createTypeLib(file);
  library.SetName("NodeComtypesTest");
  library.SetGuid(ids.libid);
  library.SetVersion(1, 0);
  library.SetLcid(0);

  using counter = library.CreateTypeInfo("ICounter", TKIND_INTERFACE)!;
  counter.SetGuid(ids.counter);
  counter.SetTypeFlags(TYPEFLAG_FOLEAUTOMATION);
  using unknownInfo = typeInfoOf(IUnknown.iid);
  counter.AddImplType(0, counter.AddRefTypeInfo(unknownInfo));
  addFunctions(
    counter,
    [
      { name: "Add", invkind: INVOKE_FUNC, params: [{ name: "a", vt: VT_I4, flags: PARAMFLAG_FIN }, { name: "b", vt: VT_I4, flags: PARAMFLAG_FIN }, outLong("sum")] },
      { name: "Name", invkind: INVOKE_PROPERTYGET, params: [outBSTR("name")] },
    ],
    3,
  );
  counter.LayOut();

  using greeter = library.CreateTypeInfo("IGreeter", TKIND_INTERFACE)!;
  greeter.SetGuid(ids.greeter);
  greeter.SetTypeFlags(TYPEFLAG_FDUAL | TYPEFLAG_FOLEAUTOMATION);
  using dispatchInfo = typeInfoOf(IDispatch.iid);
  greeter.AddImplType(0, greeter.AddRefTypeInfo(dispatchInfo));
  addFunctions(
    greeter,
    [
      { name: "Greet", invkind: INVOKE_FUNC, params: [{ name: "name", vt: VT_BSTR, flags: PARAMFLAG_FIN }, outBSTR("greeting")] },
      { name: "Count", invkind: INVOKE_PROPERTYGET, params: [outLong("count")] },
    ],
    7,
  );
  greeter.LayOut();

  using coclass = library.CreateTypeInfo("Greeter", TKIND_COCLASS)!;
  coclass.SetGuid(ids.clsid);
  coclass.SetTypeFlags(TYPEFLAG_FCANCREATE);
  using greeterInfo = greeter.QueryInterface(ITypeInfo);
  using counterInfo = counter.QueryInterface(ITypeInfo);
  coclass.AddImplType(0, coclass.AddRefTypeInfo(greeterInfo));
  coclass.SetImplTypeFlags(0, IMPLTYPEFLAG_FDEFAULT);
  coclass.AddImplType(1, coclass.AddRefTypeInfo(counterInfo));
  coclass.LayOut();

  library.SaveAllChanges();
}
