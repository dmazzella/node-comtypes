// Type libraries: GetModule at runtime, and generated modules for static types.
//
//   npm run build && node examples/typelib.ts
//
// For typed code, generate a module once:
//   npx node-comtypes generate scrrun.dll --out src/com
// then:
//   import * as Scripting from "./com/Scripting.ts";
//   using dict = CreateObject(Scripting.Dictionary);   // InterfacePointer<IDictionary>, fully typed

import { CreateObject, GetModule, apartment } from "@node-ctypes/comtypes";
import type { CoClass } from "@node-ctypes/comtypes";

using _ = apartment();

// Python: comtypes.client.GetModule("scrrun.dll")
const Scripting = GetModule("scrrun.dll");
console.log(`${Scripting.library.name} ${Scripting.library.major}.${Scripting.library.minor}`, String(Scripting.library.guid));
console.log("CompareMethod:", Scripting.CompareMethod);

// A coclass creates its default interface, called through the vtable.
using fso = CreateObject(Scripting.FileSystemObject as CoClass) as any;
console.log("Default interface:", fso.interface.name);
using temp = fso.GetSpecialFolder((Scripting.SpecialFolderConst as { TemporaryFolder: number }).TemporaryFolder);
console.log(`${temp.interface.name}.Path:`, temp.Path);
using drives = fso.Drives;
console.log("Drives:", drives.Count);
