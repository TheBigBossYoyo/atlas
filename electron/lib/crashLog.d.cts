export const MAX_LOG_BYTES: number;
export const LOG_FILE_NAME: string;
export function getLogFilePath(logDir: string): string;
export function formatLogLine(level: string, message: string, error?: unknown): string;
export function appendLogLine(logDir: string, line: string): void;
export function logToFile(logDir: string, level: string, message: string, error?: unknown): void;
