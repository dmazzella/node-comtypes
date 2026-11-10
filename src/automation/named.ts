/** Named arguments for a call: `obj.Method(1, named({ Option: true }))`. */
export class NamedArguments {
  readonly values: Readonly<Record<string, unknown>>;
  constructor(values: Record<string, unknown>) {
    this.values = Object.freeze({ ...values });
  }
}

export function named(values: Record<string, unknown>): NamedArguments {
  return new NamedArguments(values);
}
