// COM objects implemented in JavaScript: a vtable/dual interface, late
// binding, events, and native code holding the object.
//
//   npm run build && node examples/server.ts

import { BSTR, COMInterface, COMMETHOD, CreateObject, Dispatch, GetEvents, HRESULT, IDispatch, VARIANT, apartment, dispid } from "@node-ctypes/comtypes";
import { COMObject } from "@node-ctypes/comtypes/server";

const ICounter = COMInterface({
  name: "ICounter",
  iid: "{0F5A2C10-7C4B-4B8A-9E61-3D2B6E1C0A01}",
  extends: IDispatch,
  methods: [
    COMMETHOD([dispid(1)], HRESULT, "Increment", [["in", "optional"], VARIANT, "step"], [["out", "retval"], VARIANT, "value"]),
    COMMETHOD([dispid(2), "propget"], HRESULT, "Label", [["out", "retval"], BSTR, "label"]),
    COMMETHOD([dispid(2), "propput"], HRESULT, "Label", [["in"], BSTR, "label"]),
  ],
});

const ICounterEvents = COMInterface({
  name: "ICounterEvents",
  iid: "{0F5A2C10-7C4B-4B8A-9E61-3D2B6E1C0A02}",
  extends: IDispatch,
  methods: [COMMETHOD([dispid(1)], HRESULT, "Changed", [["in"], VARIANT, "value"])],
});

class Counter extends COMObject {
  static _com_interfaces_ = [ICounter];
  static _outgoing_interfaces_ = [ICounterEvents];
  Label = "counter";
  value = 0;

  Increment(step = 1) {
    this.value += step;
    this.fireEvent("Changed", this.value);
    return this.value;
  }
}

using _ = apartment();
const counter = new Counter();

// Events from the JavaScript source to a JavaScript sink, through connection points.
using connection = GetEvents(counter, { Changed: (value: number) => console.log("Changed ->", value) });

// Early binding: the vtable interface.
using pointer = counter.QueryInterface(ICounter);
pointer.Increment(5);
pointer.Label = "renamed";
console.log("Label:", pointer.Label);

// Late binding: IDispatch over the same object.
using dynamic = Dispatch(counter);
console.log("Increment():", dynamic.Increment());

// Native code holding the object: Scripting.Dictionary keeps a reference.
using dict = CreateObject("Scripting.Dictionary");
dict.Add("counter", counter);
console.log("references held by COM:", counter.refCount);
using stored = dict.Item("counter");
console.log("through the dictionary:", stored.Increment(10));
void connection;
