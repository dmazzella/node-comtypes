// Asynchronous WMI events delivered on an STA using the low-latency
// MsgWaitForMultipleObjectsEx message pump (`startMessagePump({ mode: "wait" })`
// and `waitFor`).
//
//   npm run build
//   node examples/events.ts

import { CreateObject, GetEvents, GetObject, apartment, startMessagePump, waitFor } from "@node-ctypes/comtypes";

using _ = apartment();
using pump = startMessagePump({ mode: "wait", ref: false });
using wmi = await GetObject("winmgmts:");
using sink = await CreateObject("WbemScripting.SWbemSink");

const processes: Array<{ name: string; pid: number }> = [];
let completed = false;

using _connection = GetEvents(sink, {
  OnObjectReady(obj: { Name: string; ProcessId: number }) {
    processes.push({ name: obj.Name, pid: obj.ProcessId });
  },
  OnCompleted() {
    completed = true;
  },
});

const t0 = performance.now();
wmi.ExecQueryAsync(sink, "SELECT Name, ProcessId FROM Win32_Process WHERE Name = 'node.exe'");
const ok = await waitFor(() => completed, { timeout: 10_000 });
const elapsed = Math.round(performance.now() - t0);

console.log(`WMI async query completed=${ok} in ${elapsed} ms (${processes.length} node.exe process(es)):`);
for (const p of processes) {
  console.log(`  pid=${p.pid}  ${p.name}`);
}
void pump;
