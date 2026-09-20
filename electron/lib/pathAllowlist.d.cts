export interface PathAllowlist {
  /** READ-tier registration only — see pathAllowlist.cjs's header (SEC-1). */
  add(rawPath: unknown): string | null;
  /** Full trust: READ + WRITE tier — see pathAllowlist.cjs's header (SEC-1). */
  trust(rawPath: unknown): string | null;
  has(rawPath: unknown): boolean;
  /** Whether `rawPath` may be silently overwritten (WRITE tier). */
  isWriteEligible(rawPath: unknown): boolean;
  remove(rawPath: unknown): void;
  clear(): void;
  size(): number;
}

export function createPathAllowlist(): PathAllowlist;
export function normalizePath(rawPath: unknown): string | null;
