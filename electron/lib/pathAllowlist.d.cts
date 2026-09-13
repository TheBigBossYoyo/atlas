export interface PathAllowlist {
  add(rawPath: unknown): string | null;
  has(rawPath: unknown): boolean;
  remove(rawPath: unknown): void;
  clear(): void;
  size(): number;
}

export function createPathAllowlist(): PathAllowlist;
export function normalizePath(rawPath: unknown): string | null;
