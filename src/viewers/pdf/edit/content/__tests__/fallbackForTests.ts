/** The bundled fallback fonts read from `public/fonts`, for tests running in Node. */
import fs from 'node:fs'
import path from 'node:path'
import fontkit from '@pdf-lib/fontkit'

import type { FallbackFontSource } from '../rewriteText'
import { fallbackFontFile } from '../textEdits'

const FONT_DIR = path.resolve(__dirname, '..', '..', '..', '..', '..', '..', 'public', 'fonts')

export const fallbackForTests: FallbackFontSource = {
  fontkit,
  load: async (style) => new Uint8Array(fs.readFileSync(path.join(FONT_DIR, fallbackFontFile(style)))),
}
