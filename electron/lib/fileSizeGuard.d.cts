export const DEFAULT_MAX_BYTES: number;

export declare class FileTooLargeError extends Error {
  constructor(message: string);
}

export function assertSizeAllowed(sizeBytes: number, maxBytes?: number): void;
export function assertFileSizeAllowed(filePath: string, maxBytes?: number): void;
