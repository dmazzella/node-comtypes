import { GUID } from "../guid.ts";

/** Shared well-known COM IIDs / CLSIDs used internally across modules. */
export const IID_IUnknown = new GUID("{00000000-0000-0000-C000-000000000046}");
export const IID_IDispatch = new GUID("{00020400-0000-0000-C000-000000000046}");
export const IID_ISupportErrorInfo = new GUID("{DF0B3D60-548F-101B-8E65-08002B2BD119}");
export const IID_IAgileObject = new GUID("{94EA2B94-E9CC-49E0-C0FF-EE64CA8F5B90}");
export const IID_IMarshal = new GUID("{00000003-0000-0000-C000-000000000046}");
export const IID_IClientSecurity = new GUID("{0000013D-0000-0000-C000-000000000046}");
export const IID_IGlobalInterfaceTable = new GUID("{00000146-0000-0000-C000-000000000046}");
export const CLSID_StdGlobalInterfaceTable = new GUID("{00000323-0000-0000-C000-000000000046}");
export const CLSID_InProcFreeMarshaler = new GUID("{0000033A-0000-0000-C000-000000000046}");
