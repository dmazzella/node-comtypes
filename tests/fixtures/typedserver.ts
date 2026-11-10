// LocalServer32 built from the type library of tests/fixtures/typelib.ts, used
// by tests/src/registration.test.ts:
//   node typedserver.ts <file.tlb> {clsid} [/regserver | /unregserver | -Embedding]
import { GetModule } from "@node-ctypes/comtypes";
import type { InterfaceDef } from "@node-ctypes/comtypes";
import { COMObject, UseCommandLine } from "@node-ctypes/comtypes/server";

const [file, clsid] = process.argv.slice(2);
const Test = GetModule(file);

class Greeter extends COMObject {
  static _com_interfaces_ = [Test.IGreeter as InterfaceDef, Test.ICounter as InterfaceDef];
  static _reg_clsid_ = clsid;
  static _reg_progid_ = "NodeComtypesTest.Greeter.1";
  static _reg_novers_progid_ = "NodeComtypesTest.Greeter";
  greeted = 0;
  Greet(name: string): string {
    this.greeted++;
    return `Hello, ${name} from ${process.pid}`;
  }
  get Count(): number {
    return this.greeted;
  }
  Add(a: number, b: number): number {
    return a + b;
  }
  get Name(): string {
    return `counter ${process.pid}`;
  }
}

await UseCommandLine([Greeter], { registration: { args: [file, clsid] }, server: { idleTimeout: 500, startupTimeout: 30_000 } });
