/**
 * RPR-STYLES-1 — raw-XML passthrough for the two *non*-`preserveOrder`
 * writers, `stylesWriter.ts` and `numberingWriter.ts`.
 *
 * `documentWriter.ts` already has this idea (`buildRawPassthroughPlaceholder`),
 * but it builds an ordered ARRAY of sibling nodes, so every placeholder there
 * can share one tag name and be told apart by a `data-id` attribute. These two
 * writers build plain objects whose KEYS are element names — two siblings
 * cannot share a key — so each placeholder gets a name of its own from the
 * counter below.
 *
 * Usage, once per part being written: `createRawPassthroughState()`, then
 * `spliceRawPassthrough` wherever captured fragments belong, then run the
 * finished string through `restoreRawPassthrough` before returning it. Forget
 * that last step and the placeholder tags reach the file, which is why both
 * writers take the state as an explicit argument rather than defaulting it —
 * an omitted state would otherwise silently drop the very content this exists
 * to preserve.
 */

export interface RawPassthroughState {
  readonly fragments: Map<string, string>
  next: number
}

export function createRawPassthroughState(): RawPassthroughState {
  return { fragments: new Map(), next: 0 }
}

/** Registers `xml` and returns the placeholder ELEMENT NAME to emit in its place. */
export function registerRawPassthrough(xml: string, state: RawPassthroughState): string {
  const name = `atlas-raw-${state.next}`
  state.next += 1
  state.fragments.set(name, xml)
  return name
}

/**
 * Swaps every placeholder back for its captured fragment.
 *
 * The replacement uses a FUNCTION rather than a string: a `$` in the captured
 * XML — legal in an attribute value, and routine in a `w:lvlText` or a field
 * instruction — would otherwise be read as `String.replace`'s `$&`/`$1`
 * substitution syntax and corrupt the output.
 */
export function restoreRawPassthrough(xml: string, state: RawPassthroughState): string {
  let out = xml
  for (const [name, fragment] of state.fragments) {
    out = out.replace(`<${name}/>`, () => fragment)
  }
  return out
}

/**
 * Returns `node` with each fragment spliced in as a placeholder key,
 * immediately ahead of the key named by its `before` — which is how the
 * captured child's position in the source `xsd:sequence` is restored. Falls
 * back to appending when `before` is `undefined` (the child was last in the
 * source) or names a key that isn't present (it was parsed, but whatever
 * drove it has since gone), matching `documentWriter.ts`'s
 * `insertRPrUnknownChildren`.
 *
 * Attribute keys (`@_`-prefixed) are emitted before any element key regardless
 * of where they sat, which is what `fast-xml-parser` does anyway — attributes
 * are not part of the child sequence.
 */
export function spliceRawPassthrough<T>(
  node: Readonly<Record<string, T>>,
  fragments: ReadonlyArray<{ readonly xml: string; readonly before?: string }> | undefined,
  state: RawPassthroughState,
): Record<string, T | string> {
  if (fragments === undefined || fragments.length === 0) {
    return { ...node }
  }

  const pending = fragments.map((fragment) => ({
    before: fragment.before,
    key: registerRawPassthrough(fragment.xml, state),
    placed: false,
  }))

  const out: Record<string, T | string> = {}
  for (const key of Object.keys(node)) {
    for (const entry of pending) {
      if (!entry.placed && entry.before === key) {
        out[entry.key] = ''
        entry.placed = true
      }
    }
    out[key] = node[key] as T
  }
  for (const entry of pending) {
    if (!entry.placed) out[entry.key] = ''
  }

  return out
}
