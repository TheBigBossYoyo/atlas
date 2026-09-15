export declare class PrintToPdfError extends Error {
  constructor(message: string);
}

export const MAX_HTML_BYTES: number;

export function printHtmlToPdfBuffer(html: string): Promise<Buffer>;
