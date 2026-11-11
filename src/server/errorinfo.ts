import * as ctypes from "node-ctypes";
import { checkHRESULT } from "../errors.ts";
import type { GUID } from "../guid.ts";
import { ICreateErrorInfo, IErrorInfo } from "../interfaces/errorinfo.ts";
import { oleaut32 } from "../low-level/native.ts";
import { COMPointer, sync } from "../pointer.ts";

export interface ErrorInfoFields {
  readonly description?: string | null;
  readonly source?: string | null;
  readonly guid?: GUID | null;
  readonly helpFile?: string | null;
  readonly helpContext?: number;
}

/**
 * Set the thread's error object (`CreateErrorInfo` + `SetErrorInfo`), read
 * by callers through `GetErrorInfo`. Python equivalent: `comtypes.ReportError`.
 */
export function setErrorInfo(fields: ErrorInfoFields): void {
  const out = new ctypes.c_void_p();
  checkHRESULT(oleaut32().CreateErrorInfo(ctypes.byref(out)), "CreateErrorInfo");
  using create = COMPointer.fromAddress(out.value!, ICreateErrorInfo);
  if (fields.description) sync(create).SetDescription(fields.description);
  if (fields.source) sync(create).SetSource(fields.source);
  if (fields.guid) sync(create).SetGUID(fields.guid);
  if (fields.helpFile) sync(create).SetHelpFile(fields.helpFile);
  if (fields.helpContext) sync(create).SetHelpContext(fields.helpContext);
  using info = create.QueryInterface(IErrorInfo);
  checkHRESULT(oleaut32().SetErrorInfo(0, info.address), "SetErrorInfo");
}
