export { COMObject, VariantEnumerator, ConnectionPoint, ClassFactory, getLiveObjectCount, queryObject } from "./comobject.ts";
export type { Threading } from "./comobject.ts";
export { DispatchError, ReturnHRESULT, defaultIdsOfNames, defaultInvoke } from "./automation.ts";
export type { InvokeArguments } from "./automation.ts";
export { setErrorInfo } from "./errorinfo.ts";
export type { ErrorInfoFields } from "./errorinfo.ts";
export { registerClassObject, registerServer, unregisterServer, runLocalServer, UseCommandLine, REGCLS_SINGLEUSE, REGCLS_MULTIPLEUSE, REGCLS_MULTI_SEPARATE, REGCLS_SUSPENDED } from "./register.ts";
export type { ClassRegistration, ServerRegistration, UnregisterServerOptions, UseCommandLineOptions, LocalServerOptions } from "./register.ts";
