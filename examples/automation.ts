// Automation: Scripting.Dictionary and FileSystemObject (typed from their type
// library, the default) and WMI (late bound). Methods return promises and run
// on the COM executor; properties are synchronous.
//
//   npm run build && node examples/automation.ts

import { CreateObject, GetObject, apartment, named, putProperty } from "@node-ctypes/comtypes";

using _ = apartment();

{
  using dict = await CreateObject("Scripting.Dictionary");
  await dict.Add("name", "@node-ctypes/comtypes");
  await dict.Add("created", new Date());
  await dict.Add("values", [1, 2.5, "three", true]);
  putProperty(dict, "Item", "name", "renamed"); // parameterized property put
  for await (const key of dict) console.log(`${key}:`, dict.Item(key)); // _NewEnum on the executor; Item is a property
  console.log("Keys():", await dict.Keys());
}

{
  using fso = await CreateObject("Scripting.FileSystemObject");
  using temp = await fso.GetSpecialFolder(2);
  const path = await fso.BuildPath(temp.Path, "node-comtypes.txt");
  using writer = await fso.CreateTextFile(path, named({ Overwrite: true }));
  await writer.WriteLine("written through ITextStream");
  await writer.Close();
  using reader = await fso.OpenTextFile(path);
  console.log(`${path}:`, (await reader.ReadAll()).trim());
  await reader.Close();
  await fso.DeleteFile(path);
}

{
  // WMI class properties (FreeSpace, Size, ...) are not in the type library: late binding.
  using wmi = await GetObject("winmgmts:", { dynamic: true });
  using disks = await wmi.ExecQuery("SELECT DeviceID, FreeSpace, Size FROM Win32_LogicalDisk WHERE DriveType = 3");
  // for await: the collection is enumerated on the COM executor.
  for await (const disk of disks) {
    using owned = disk;
    const free = Number(owned.FreeSpace) / 2 ** 30;
    const size = Number(owned.Size) / 2 ** 30;
    console.log(`${owned.DeviceID} ${free.toFixed(1)} GiB free of ${size.toFixed(1)} GiB`);
  }
}
