// A COM server written in JavaScript, registered for the current user and
// started by COM on demand (LocalServer32). Python equivalent: a comtypes
// server module ending with `UseCommandLine(...)`.
//
//   npm run build
//   node examples/localserver.ts /regserver      (register)
//   powershell -c "(New-Object -ComObject NodeComtypes.Example).Greet('PowerShell')"
//   node examples/localserver.ts /unregserver    (remove)
//
// COM runs `node examples/localserver.ts -Embedding` when a client creates the
// object; the process exits once its objects are released. Switches also
// accept "-" (-regserver), e.g. in Git Bash, which rewrites "/regserver" as a path.

import { IDispatch } from "@node-ctypes/comtypes";
import { COMObject, UseCommandLine } from "@node-ctypes/comtypes/server";

class Example extends COMObject {
  static _com_interfaces_ = [IDispatch];
  static _reg_clsid_ = "{7A3F5C1E-2B4D-4E8F-9A6B-1C2D3E4F5A6B}";
  static _reg_progid_ = "NodeComtypes.Example.1";
  static _reg_novers_progid_ = "NodeComtypes.Example";
  static _reg_desc_ = "node-comtypes example server";

  Greet(name: string): string {
    return `Hello, ${name}, from Node.js ${process.version} (pid ${process.pid})`;
  }

  Add(a: number, b: number): number {
    return a + b;
  }
}

await UseCommandLine([Example]);
