// LocalServer32 used by tests/src/server.test.ts. COM starts it as:
//   node localserver.ts {clsid} -Embedding
import { IDispatch } from "@node-ctypes/comtypes";
import { COMObject, runLocalServer } from "@node-ctypes/comtypes/server";

class Calculator extends COMObject {
  static _com_interfaces_ = [IDispatch];
  static _reg_clsid_ = process.argv[2];
  Name = "calculator";
  get Pid(): number {
    return process.pid;
  }
  Add(a: number, b: number): number {
    return a + b;
  }
  Echo(value: unknown): unknown {
    return value;
  }
  Fail(message: string): never {
    throw new Error(message);
  }
}

await runLocalServer([Calculator], { idleTimeout: 500, startupTimeout: 30_000 });
