/**
 * T11 (DAT-16) — TextViewer render/parse fixture coverage. Extracted onto
 * the shared VirtualizedPlainText component (T2); this locks in that the
 * refactor is behavior-preserving.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { render, screen } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'
import { TextViewer } from '../TextViewer'

const FIXTURES_DIR = path.resolve(process.cwd(), 'src/viewers/__fixtures__')

function readFixtureText(name: string): string {
  return readFileSync(path.join(FIXTURES_DIR, name), 'utf8')
}

describe('TextViewer — fixture render/parse (T11)', () => {
  it('renders a real .txt fixture line by line', () => {
    const file: LoadedFile = {
      kind: 'text',
      content: readFixtureText('sample.txt'),
      path: '/tmp/sample.txt',
      format: 'text',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <TextViewer file={file} />
      </ViewerProvider>,
    )

    expect(screen.getByText('Atlas text fixture')).toBeInTheDocument()
    expect(screen.getByText('Second line')).toBeInTheDocument()
  })

  it('decodes binary-kind content as UTF-8 text', () => {
    const file: LoadedFile = {
      kind: 'binary',
      content: new TextEncoder().encode('binary-sourced line\n').buffer as ArrayBuffer,
      path: '/tmp/binary.txt',
      format: 'text',
    }

    render(
      <ViewerProvider filePath={file.path}>
        <TextViewer file={file} />
      </ViewerProvider>,
    )

    expect(screen.getByText('binary-sourced line')).toBeInTheDocument()
  })

  it('carries its own scrollable root class (DAT-02/UX-04 parity)', () => {
    const file: LoadedFile = { kind: 'text', content: 'x\n', path: '/tmp/x.txt', format: 'text' }

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <TextViewer file={file} />
      </ViewerProvider>,
    )

    expect(container.querySelector('.text-viewer')).not.toBeNull()
  })
})
