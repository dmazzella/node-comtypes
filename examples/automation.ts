// Automation: Scripting.Dictionary and FileSystemObject (typed from their type
// library, the default) and WMI (late bound).
//
//   npm run build && node examples/automation.ts

import { CreateObject, GetObject, apartment, named, putProperty } from "@node-ctypes/comtypes";

using _ = apartment();

{
  using dict = CreateObject("Scripting.Dictionary");
  dict.Add("name", "@node-ctypes/comtypes");
  dict.Add("created", new Date());
  dict.Add("values", [1, 2.5, "three", true]);
  putProperty(dict, "Item", "name", "renamed"); // parameterized property put
  for (const key of dict) console.log(`${key}:`, dict.Item(key));
  console.log("Keys():", dict.Keys());
}

{
  using fso = CreateObject("Scripting.FileSystemObject");
  using temp = fso.GetSpecialFolder(2);
  const path = fso.BuildPath(temp.Path, "node-comtypes.txt");
  using writer = fso.CreateTextFile(path, named({ Overwrite: true }));
  writer.WriteLine("written through ITextStream");
  writer.Close();
  using reader = fso.OpenTextFile(path);
  console.log(`${path}:`, reader.ReadAll().trim());
  reader.Close();
  fso.DeleteFile(path);
}

{
  // WMI class properties (FreeSpace, Size, ...) are not in the type library: late binding.
  using wmi = GetObject("winmgmts:", { dynamic: true });
  for (const disk of wmi.ExecQuery("SELECT DeviceID, FreeSpace, Size FROM Win32_LogicalDisk WHERE DriveType = 3")) {
    using owned = disk;
    const free = Number(owned.FreeSpace) / 2 ** 30;
    const size = Number(owned.Size) / 2 ** 30;
    console.log(`${owned.DeviceID} ${free.toFixed(1)} GiB free of ${size.toFixed(1)} GiB`);
  }
}
