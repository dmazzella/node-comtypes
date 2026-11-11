// COMMETHOD interfaces: out parameters are returned, SELF results are typed,
// several outputs come back as a tuple.
//
//   npm run build && node examples/stream.ts

import * as ctypes from "node-ctypes";
import { COMPointer, IStream, STATFLAG_NONAME, STREAM_SEEK_SET, apartment, checkHRESULT } from "@node-ctypes/comtypes";

using _ = apartment();

const ole32 = new ctypes.WinDLL("ole32.dll");
const CreateStreamOnHGlobal = ole32.func("CreateStreamOnHGlobal", ctypes.c_int32, [ctypes.c_void_p, ctypes.c_int32, ctypes.c_void_p]);
function createStream() {
  const out = new ctypes.c_void_p();
  checkHRESULT(CreateStreamOnHGlobal(null, 1, ctypes.byref(out)), "CreateStreamOnHGlobal");
  return COMPointer.fromAddress(out.value!, IStream);
}

using stream = createStream();

const data = Buffer.from("Hello from COMMETHOD");
console.log("written:", await stream.Write(data, data.length)); // [out] ULONG* -> number
console.log("position:", await stream.Seek(0n, STREAM_SEEK_SET)); // [out] ULARGE_INTEGER* -> bigint
console.log("size:", (await stream.Stat(STATFLAG_NONAME)).cbSize); // [out] STATSTG* -> STATSTG

// [out] IStream** (SELF) -> IStream pointer: same data, its own seek position.
using clone = (await stream.Clone())!;
await clone.Seek(11n, STREAM_SEEK_SET);
const tail = Buffer.alloc(9);
await clone.Read(tail, tail.length);
console.log("clone reads:", tail.toString());

using target = createStream();
const [read, written] = await stream.CopyTo(target, BigInt(data.length)); // two [out] -> tuple
console.log("copied:", read, written);
