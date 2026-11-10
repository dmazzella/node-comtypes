// OLE Automation DATE: days since 1899-12-30 00:00 as a double, in local
// wall-clock time (no time zone). Before the epoch the integer part counts
// days backwards while the fraction still counts the time forward.

const DAY = 86_400_000;
const EPOCH = Date.UTC(1899, 11, 30);
/** 0100-01-01 */
export const OADATE_MIN = -657434;
/** 9999-12-31 23:59:59.999 */
export const OADATE_MAX = 2958465.99999999;

/** Convert a JavaScript Date (local wall-clock time) to an OLE Automation DATE. */
export function dateToOADate(date: Date): number {
  const time = date.getTime();
  if (Number.isNaN(time)) throw new RangeError("Invalid Date");
  // Wall-clock time from the local components: getTimezoneOffset() is rounded
  // to minutes, which is wrong for historical (LMT) offsets.
  const wall = new Date(0);
  wall.setUTCFullYear(date.getFullYear(), date.getMonth(), date.getDate());
  wall.setUTCHours(date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds());
  let millis = wall.getTime() - EPOCH;
  if (millis < 0) {
    const fraction = millis % DAY;
    if (fraction !== 0) millis -= (DAY + fraction) * 2;
  }
  const value = millis / DAY;
  if (value < OADATE_MIN || value > OADATE_MAX) throw new RangeError("Date outside the OLE Automation DATE range");
  return value;
}

/** Convert an OLE Automation DATE to a JavaScript Date (local wall-clock time), to the millisecond. */
export function oaDateToDate(value: number): Date {
  if (!Number.isFinite(value) || value < OADATE_MIN || value > OADATE_MAX) throw new RangeError(`Invalid OLE Automation DATE: ${value}`);
  let millis = Math.trunc(value * DAY + (value >= 0 ? 0.5 : -0.5));
  if (millis < 0) millis -= (millis % DAY) * 2;
  const wall = new Date(EPOCH + millis);
  const date = new Date(2000, 0, 1);
  date.setFullYear(wall.getUTCFullYear(), wall.getUTCMonth(), wall.getUTCDate());
  date.setHours(wall.getUTCHours(), wall.getUTCMinutes(), wall.getUTCSeconds(), wall.getUTCMilliseconds());
  return date;
}
