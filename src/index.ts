export { GUID, GUID_NULL, toGUID } from "./guid.ts";
export type { GUIDLike } from "./guid.ts";
export * from "./hresult.ts";
export { COMError, checkHRESULT, hresultText, systemMessage } from "./errors.ts";
export type { COMErrorDetails, COMErrorOptions } from "./errors.ts";
export { takeErrorInfo } from "./errorinfo.ts";
export { stats, finalize } from "./executor/executor.ts";
export { run } from "./executor/operations.ts";
export type { RunOptions } from "./executor/operations.ts";
export type { ExecutorStats } from "./executor/executor.ts";
export { COMInterface, STDMETHOD, HRESULT, findInterface, derivesFrom } from "./interface.ts";
export type {
  Asynchronous,
  DispatchCall,
  DispatchCallAsync,
  InterfaceDef,
  ComInterfaceDef,
  InterfacePointer,
  InterfaceSpec,
  InterfaceMethods,
  MethodSpec,
  MethodKind,
  MethodMarshal,
  MethodMarshalAsync,
  MethodFunction,
  MethodResult,
  MethodsOf,
  NativeRunner,
  PropertiesOf,
  PointerClass,
  RawFunction,
  SyncInterfaceMethods,
  SyncMethodsOf,
  SyncPointer,
  Synchronous,
} from "./interface.ts";
export { COMMETHOD, BSTR, LPOLESTR, REFGUID, REFIID, REFCLSID, SAFEARRAY, SELF, DATE, CURRENCY, lazy, dispid, helpstring, classifyParamType, resolveParamInterface } from "./commethod.ts";
export type {
  ComParamType,
  ParamFlag,
  ParamSpec,
  ParamInfo,
  IdlFlag,
  DispIdFlag,
  HelpStringFlag,
  InValue,
  OutValue,
  ArgumentsOf,
  OutputsOf,
  CommethodResult,
  CommethodFunction,
  CommethodSpec,
  ComObject,
  BSTRType,
  LPOLESTRType,
  RefGuidType,
  SafeArrayType,
  SelfMarker,
  LazyInterface,
  DateType,
  CurrencyType,
  ParamKind,
} from "./commethod.ts";
export { COMPointer } from "./pointer.ts";
export type { FromAddressOptions } from "./pointer.ts";
export * from "./constants.ts";
export { CoInitializeEx, CoInitialize, CoUninitialize, apartment, getApartmentType, getLivePointerCount, CoCreateInstance, CoCreateInstanceEx, CoGetClassObject, CoTaskMemAlloc, CoTaskMemRealloc, CoTaskMemFree } from "./runtime.ts";
export type { Apartment, ApartmentType, CreateInstanceOptions, CreateInstanceExOptions } from "./runtime.ts";
export * from "./interfaces/index.ts";
export * from "./automation/index.ts";
export * from "./client/index.ts";
export * from "./server/index.ts";
export { pumpMessages, PumpEvents, startMessagePump, waitFor } from "./pump.ts";
export { DISPMETHOD, DISPPROPERTY } from "./dispmethod.ts";
export type { DispSpec } from "./dispmethod.ts";
export * from "./typelib/index.ts";
export { marshalInterface, unmarshalInterface, unmarshalDispatch, releaseMarshaled, GlobalInterfaceTable, IGlobalInterfaceTable, CLSID_StdGlobalInterfaceTable } from "./threads.ts";
export type { MarshaledInterface } from "./threads.ts";
export type { MessagePump } from "./pump.ts";

export const version = "0.2.0";
