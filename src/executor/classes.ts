// Classes whose class objects this thread registered (`registerClassObject`):
// their objects are implemented in JavaScript and are created on this thread.
const classes = new Map<string, number>();

const key = (clsid: { toString(): string }) => clsid.toString().toUpperCase();

export function markLocalClass(clsid: { toString(): string }): void {
  classes.set(key(clsid), (classes.get(key(clsid)) ?? 0) + 1);
}

export function unmarkLocalClass(clsid: { toString(): string }): void {
  const count = (classes.get(key(clsid)) ?? 0) - 1;
  if (count > 0) classes.set(key(clsid), count);
  else classes.delete(key(clsid));
}

/** Whether this thread registered a class object for `clsid`. */
export function isLocalClass(clsid: { toString(): string }): boolean {
  return classes.has(key(clsid));
}
