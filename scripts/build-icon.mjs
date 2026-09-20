// Atlas icon generator: SVG -> multi-size PNGs -> ICO + master PNG
// Uses sharp (rasterize SVG) + png-to-ico (pack ICO).
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import sharp from 'sharp';
import pngToIco from 'png-to-ico';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, '..');
const SRC_SVG = resolve(ROOT, 'assets', 'atlas.svg');
const BUILD_DIR = resolve(ROOT, 'build');
const ICO_OUT = resolve(BUILD_DIR, 'icon.ico');
const PNG_OUT = resolve(BUILD_DIR, 'icon.png');

const ICO_SIZES = [16, 24, 32, 48, 64, 128, 256];
const PNG_MASTER = 1024;

/**
 * @param {Buffer} svgBuffer
 * @param {number} size
 * @returns {Promise<Buffer>}
 */
async function rasterize(svgBuffer, size) {
  return sharp(svgBuffer, { density: 384 })
    .resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toBuffer();
}

async function main() {
  await mkdir(BUILD_DIR, { recursive: true });
  const svg = await readFile(SRC_SVG);

  const masterPng = await rasterize(svg, PNG_MASTER);
  await writeFile(PNG_OUT, masterPng);
  console.log(`wrote ${PNG_OUT} (${masterPng.length} bytes)`);

  const icoPngs = await Promise.all(ICO_SIZES.map((s) => rasterize(svg, s)));
  const ico = await pngToIco(icoPngs);
  await writeFile(ICO_OUT, ico);
  console.log(`wrote ${ICO_OUT} (${ico.length} bytes, sizes ${ICO_SIZES.join(',')})`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
