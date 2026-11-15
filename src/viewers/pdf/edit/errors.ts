/** Errors the PDF writer raises for the UI to explain (PDF-EDIT-1/2). */

/** Thrown when text holds a character none of the available fonts can draw. */
export class UnsupportedCharacterError extends Error {
  readonly text: string

  constructor(text: string) {
    super(`None of the available fonts can draw some characters in "${text}".`)
    this.name = 'UnsupportedCharacterError'
    this.text = text
  }
}

/** Thrown when an edit refers to text that is no longer where it was (the file changed). */
export class StaleTextEditError extends Error {
  constructor() {
    super('The text being edited is no longer in the file.')
    this.name = 'StaleTextEditError'
  }
}
