'use strict';

/**
 * USR-11 — lists installed font family names for the DOCX font picker.
 *
 * Reads the per-machine and per-user Windows font registrations with `reg
 * query` (fixed arguments, no user input, no shell) and reduces entries such
 * as "Arial Bold Italic (TrueType)" or "Segoe UI Semibold & Segoe UI Light
 * (TrueType)" to family names. The result is cached for the process lifetime.
 */

const { execFile } = require('node:child_process');

const REGISTRY_KEYS = [
  'HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts',
  'HKCU\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts',
];

const STYLE_WORDS = new Set([
  'regular', 'normal', 'book', 'roman', 'bold', 'italic', 'oblique', 'light', 'extralight', 'ultralight',
  'semilight', 'thin', 'hairline', 'medium', 'semibold', 'demibold', 'demi', 'extrabold', 'ultrabold', 'heavy',
  'black', 'condensed', 'semicondensed', 'extracondensed', 'narrow', 'expanded', 'extended', 'variable',
]);

/**
 * "Arial Bold Italic (TrueType)" -> ["Arial"]; "A & B (TrueType)" -> ["A", "B"].
 * @param {string} name
 * @returns {string[]}
 */
function familiesFromRegistryName(name) {
  const withoutTech = name.replace(/\s*\((TrueType|OpenType|All res|VGA res|[^)]*res[^)]*)\)\s*$/i, '').trim();
  return withoutTech
    .split('&')
    .map((part) => {
      const words = part.trim().split(/\s+/);
      while (words.length > 1 && STYLE_WORDS.has(words[words.length - 1].toLowerCase())) {
        words.pop();
      }
      return words.join(' ');
    })
    // Raster/bitmap registrations look like "Courier 10,12,15": not usable families.
    .filter((family) => family.length > 1 && !/^\d/.test(family) && !/,/.test(family));
}

/**
 * @param {string} stdout
 * @returns {string[]}
 */
function parseRegQueryOutput(stdout) {
  const families = [];
  for (const line of stdout.split(/\r?\n/)) {
    const match = /^\s{2,}(.+?)\s{2,}REG_(?:SZ|EXPAND_SZ)\s{2,}/.exec(line);
    if (match) {
      families.push(...familiesFromRegistryName(match[1]));
    }
  }
  return families;
}

/**
 * @param {string} key
 * @returns {Promise<string[]>}
 */
function queryKey(key) {
  return new Promise((resolve) => {
    execFile('reg', ['query', key], { windowsHide: true, timeout: 5000, maxBuffer: 8 * 1024 * 1024 }, (error, stdout) => {
      resolve(error ? [] : parseRegQueryOutput(String(stdout)));
    });
  });
}

/** @type {Promise<string[]> | null} */
let cached = null;

/** @returns {Promise<string[]>} */
function listSystemFontFamilies() {
  if (process.platform !== 'win32') {
    return Promise.resolve([]);
  }
  if (cached === null) {
    cached = Promise.all(REGISTRY_KEYS.map(queryKey)).then((lists) =>
      Array.from(new Set(lists.flat())).sort((a, b) => a.localeCompare(b)),
    );
  }
  return cached;
}

module.exports = { listSystemFontFamilies, familiesFromRegistryName, parseRegQueryOutput };
