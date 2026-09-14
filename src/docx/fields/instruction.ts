/**
 * Atlas — DOCX field instruction parsing (DEFER-5 / DXS-20)
 *
 * A field instruction (`Field.instruction`, e.g. `DATE \@ "MMMM d, yyyy"` or
 * `REF _Ref123 \h`) is whitespace-separated tokens, with a double-quoted
 * span counting as one token (so `"MMMM d, yyyy"` — itself containing
 * spaces — survives as a single argument), and any token starting with `\`
 * introducing a "switch" that consumes the next token as its value UNLESS
 * that next token is itself another switch (a bare/flag switch like `\h`
 * takes no value). This is a pragmatic subset of Word's real field-code
 * grammar — real Word instructions can nest quotes and escape characters
 * in ways this doesn't attempt — sufficient for the switches the DEFER-5
 * evaluators actually consult (`\@`, `\h`, `\o`, `\c`, `\r`, `\l`, `\o`
 * (tooltip), `\t`, `\*`).
 */
export interface ParsedFieldInstruction {
  /** The field's type keyword, uppercased (e.g. `DATE`). Empty string for a blank instruction. */
  readonly keyword: string
  /** Non-switch tokens after the keyword, in order (e.g. a bookmark name for REF/PAGEREF, a sequence name for SEQ). */
  readonly arguments: ReadonlyArray<string>
  /** Switch name (without its leading `\`) -> its value, or `true` for a bare/flag switch that took no value. */
  readonly switches: ReadonlyMap<string, string | true>
}

export function parseFieldInstruction(instruction: string): ParsedFieldInstruction {
  const tokens = tokenizeInstruction(instruction)
  const keyword = (tokens[0] ?? '').toUpperCase()
  return { keyword, ...splitArgumentsAndSwitches(tokens.slice(1)) }
}

function splitArgumentsAndSwitches(
  tokens: ReadonlyArray<string>,
): { arguments: ReadonlyArray<string>; switches: ReadonlyMap<string, string | true> } {
  const args: string[] = []
  const switches = new Map<string, string | true>()
  let index = 0

  while (index < tokens.length) {
    const token = tokens[index]
    if (token === undefined) {
      break
    }

    if (isSwitchToken(token)) {
      const name = token.slice(1)
      const next = tokens[index + 1]
      if (next !== undefined && !isSwitchToken(next)) {
        switches.set(name, next)
        index += 2
      } else {
        switches.set(name, true)
        index += 1
      }
      continue
    }

    args.push(token)
    index += 1
  }

  return { arguments: args, switches }
}

function isSwitchToken(token: string): boolean {
  return token.startsWith('\\') && token.length > 1
}

/**
 * Splits `instruction` on whitespace, treating a `"..."` span as one token
 * (its surrounding quotes stripped) regardless of whitespace inside it.
 */
function tokenizeInstruction(instruction: string): ReadonlyArray<string> {
  const tokens: string[] = []
  const source = instruction.trim()
  let index = 0

  while (index < source.length) {
    while (index < source.length && isWhitespace(source[index])) {
      index += 1
    }
    if (index >= source.length) {
      break
    }

    if (source[index] === '"') {
      const closeIndex = source.indexOf('"', index + 1)
      if (closeIndex === -1) {
        tokens.push(source.slice(index + 1))
        break
      }
      tokens.push(source.slice(index + 1, closeIndex))
      index = closeIndex + 1
      continue
    }

    let end = index
    while (end < source.length && !isWhitespace(source[end])) {
      end += 1
    }
    tokens.push(source.slice(index, end))
    index = end
  }

  return tokens
}

function isWhitespace(char: string | undefined): boolean {
  return char === ' ' || char === '\t' || char === '\n' || char === '\r'
}
