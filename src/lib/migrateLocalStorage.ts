const FLAG = 'atlas-migration-v1'
const OLD_PREFIX = 'md-reader-'
const NEW_PREFIX = 'atlas-'

export function runMigration(): void {
  try {
    if (typeof localStorage === 'undefined') return
    if (localStorage.getItem(FLAG) === '1') return
    const keys: string[] = []
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i)
      if (k && k.startsWith(OLD_PREFIX)) keys.push(k)
    }
    for (const oldKey of keys) {
      const newKey = NEW_PREFIX + oldKey.slice(OLD_PREFIX.length)
      const value = localStorage.getItem(oldKey)
      if (value !== null && localStorage.getItem(newKey) === null) {
        localStorage.setItem(newKey, value)
      }
      localStorage.removeItem(oldKey)
    }
    localStorage.setItem(FLAG, '1')
  } catch {
    // localStorage unavailable / quota — fail silent, migration retried next boot
  }
}
