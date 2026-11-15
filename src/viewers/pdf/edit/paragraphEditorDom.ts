/**
 * DOM side of the paragraph editor (PDF-EDIT-2): filling the contentEditable
 * box with styled spans, and reading it back — see ParagraphEditor.tsx.
 */

import type { CSSProperties } from 'react'

import { cssFontFamily } from './textLayout'
import type { SpanStyle, StyledSpan } from './content/paragraphs'

export type ParagraphEditorValue = {
  readonly spans: ReadonlyArray<StyledSpan>
  readonly styles: ReadonlyArray<SpanStyle>
}

const NO_BREAK_SPACE = String.fromCharCode(0xa0)

/** CSS colour for a fill the editor can show: plain gray/RGB/CMYK operators; anything else shows black. */
export function cssFill(style: SpanStyle): string {
  const color = style.fill.color ?? ''
  const numbers = color.trim().split(/\s+/).slice(0, -1).map(Number)
  const operator = color.trim().split(/\s+/).pop()
  const to255 = (value: number) => Math.round(Math.max(0, Math.min(1, value)) * 255)
  if (operator === 'g' && numbers.length === 1) return `rgb(${to255(numbers[0])}, ${to255(numbers[0])}, ${to255(numbers[0])})`
  if (operator === 'rg' && numbers.length === 3) return `rgb(${numbers.map(to255).join(', ')})`
  if (operator === 'k' && numbers.length === 4) {
    const [c, m, y, k] = numbers
    return `rgb(${to255((1 - c) * (1 - k))}, ${to255((1 - m) * (1 - k))}, ${to255((1 - y) * (1 - k))})`
  }
  return 'rgb(0, 0, 0)'
}

function spanCss(style: SpanStyle, scale: number): CSSProperties {
  return {
    fontFamily: cssFontFamily(style.font.family),
    fontWeight: style.font.bold ? 700 : 400,
    fontStyle: style.font.italic ? 'italic' : 'normal',
    fontSize: `${style.size * scale}px`,
    color: cssFill(style),
    letterSpacing: style.charSpacing ? `${style.charSpacing * scale}px` : undefined,
  }
}

export function fillEditor(root: HTMLElement, value: ParagraphEditorValue, scale: number): void {
  root.replaceChildren()
  for (const span of value.spans) {
    const style = value.styles[span.style]
    const element = document.createElement('span')
    element.dataset.style = String(span.style)
    Object.assign(element.style, spanCss(style, scale))
    element.textContent = span.text
    root.appendChild(element)
  }
}

/** Reads the editor's DOM back into spans; see the file comment for derived styles. */
export function readEditor(root: HTMLElement, base: ParagraphEditorValue): ParagraphEditorValue {
  const styles = [...base.styles]
  const derived = new Map<string, number>()
  const spans: StyledSpan[] = []
  const push = (text: string, style: number) => {
    if (text.length === 0) return
    const last = spans[spans.length - 1]
    if (last && last.style === style) spans[spans.length - 1] = { text: last.text + text, style }
    else spans.push({ text, style })
  }

  const styleFor = (node: Node): number => {
    const element = node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement
    const host = element?.closest<HTMLElement>('[data-style]')
    const baseIndex = host ? Number(host.dataset.style) : spans.length > 0 ? spans[spans.length - 1].style : 0
    const baseStyle = styles[Number.isFinite(baseIndex) ? baseIndex : 0] ?? styles[0]
    let bold = baseStyle.font.bold
    let italic = baseStyle.font.italic
    for (let el = element; el && el !== root; el = el.parentElement) {
      const tag = el.tagName
      if (tag === 'B' || tag === 'STRONG') bold = true
      if (tag === 'I' || tag === 'EM') italic = true
      if (el.style.fontWeight) bold = Number(el.style.fontWeight) >= 600 || el.style.fontWeight === 'bold'
      if (el.style.fontStyle) italic = el.style.fontStyle === 'italic'
      if (el === host) break
    }
    if (bold === baseStyle.font.bold && italic === baseStyle.font.italic) return styles.indexOf(baseStyle)
    const key = `${styles.indexOf(baseStyle)}|${bold}|${italic}`
    let index = derived.get(key)
    if (index === undefined) {
      index = styles.length
      styles.push({ ...baseStyle, fontKey: null, fontResource: null, font: { ...baseStyle.font, bold, italic } })
      derived.set(key, index)
    }
    return index
  }

  const walk = (node: Node, first: boolean) => {
    if (node.nodeType === Node.TEXT_NODE) {
      push((node.textContent ?? '').split(NO_BREAK_SPACE).join(' '), styleFor(node))
      return
    }
    if (!(node instanceof HTMLElement)) return
    if (node.tagName === 'BR') {
      push('\n', styleFor(node))
      return
    }
    // Chrome wraps lines in <div>s after Enter or a paste: each starts a new line.
    if ((node.tagName === 'DIV' || node.tagName === 'P') && !first && node !== root) push('\n', styleFor(node))
    node.childNodes.forEach((child, index) => walk(child, first && index === 0))
  }
  walk(root, true)
  // A trailing line break the browser keeps for the caret is not content.
  const last = spans[spans.length - 1]
  if (last && last.text.endsWith('\n')) spans[spans.length - 1] = { ...last, text: last.text.replace(/\n+$/, '') }
  return { spans: spans.filter((span) => span.text.length > 0), styles }
}

