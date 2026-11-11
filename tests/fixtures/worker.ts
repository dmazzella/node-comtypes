// Worker used by tests/src/threads.test.ts: calls a COM object of the main
// thread through a marshaled token or the Global Interface Table.
import { parentPort, threadId, workerData } from "node:worker_threads";
import { COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED, GlobalInterfaceTable, apartment, getApartmentType, unmarshalDispatch } from "@node-ctypes/comtypes";
import type { DispatchObject, MarshaledInterface } from "@node-ctypes/comtypes";

const data = workerData as { readonly token?: MarshaledInterface; readonly cookie?: number; readonly mta?: boolean };
using _ = apartment(data.mta ? COINIT_MULTITHREADED : COINIT_APARTMENTTHREADED);
let object: DispatchObject;
if (data.token) {
  object = unmarshalDispatch(data.token);
} else {
  using table = new GlobalInterfaceTable();
  object = table.getDispatch(data.cookie!);
}
try {
  const result = { apartment: getApartmentType()?.name, sum: await object.Add(20, 22), owner: object.ThreadId, worker: threadId, echo: await object.Echo(["a", 1]) };
  parentPort!.postMessage(result);
} finally {
  object[Symbol.dispose]();
}
