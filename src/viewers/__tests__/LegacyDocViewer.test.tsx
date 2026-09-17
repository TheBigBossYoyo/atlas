import { render, waitFor } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import type { LoadedFile } from '../../formats/types'
import { buildMinimalDocBytes } from '../../legacy/__tests__/fixtures'
import { ViewerProvider } from '../shared/ViewerContext'
import { LegacyDocViewer } from '../LegacyDocViewer'

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return new Uint8Array(bytes).buffer
}

function fixtureFile(bodyText: string): LoadedFile {
  return { kind: 'binary', content: toArrayBuffer(buildMinimalDocBytes(bodyText)), path: '/fixtures/sample.doc', format: 'doc' }
}

describe('LegacyDocViewer', () => {
  it('parses a real .doc end to end and renders its paragraph text', async () => {
    const file = fixtureFile('Hello legacy world\rSecond paragraph here\r')

    const { getByText } = render(
      <ViewerProvider filePath={file.path}>
        <LegacyDocViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(getByText('Hello legacy world')).toBeInTheDocument()
    })
    expect(getByText('Second paragraph here')).toBeInTheDocument()
  })

  it('shows the legacy-format read-only banner naming the modern format', async () => {
    const file = fixtureFile('Text\r')

    const { getByRole } = render(
      <ViewerProvider filePath={file.path}>
        <LegacyDocViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(getByRole('note')).toHaveTextContent(/Word 97-2003/)
    })
    expect(getByRole('note')).toHaveTextContent(/\.docx/)
  })

  it('shows a friendly error instead of crashing on a file that is not a valid .doc', async () => {
    const file: LoadedFile = { kind: 'binary', content: new ArrayBuffer(3), path: '/fixtures/bad.doc', format: 'doc' }

    const { container } = render(
      <ViewerProvider filePath={file.path}>
        <LegacyDocViewer file={file} />
      </ViewerProvider>,
    )

    await waitFor(() => {
      expect(container.querySelector('.legacy-doc-viewer--error')).not.toBeNull()
    })
  })
})
