/**
 * SHEETFN-2 — Excel's serial-date numbering, in one place.
 *
 * A spreadsheet date is a NUMBER: whole days since the fictitious 1899-12-30
 * "day zero" (one day before 1900-01-01, which is itself day 1). That epoch
 * trick is what reproduces Excel's well-known "1900 was a leap year" bug for
 * every real date on or after 1900-03-01 without special-casing it, because the
 * phantom 1900-02-29 falls before that range.
 *
 * This existed as a private helper inside `xlsxPassthrough.ts`, used when the
 * user TYPES an ISO date into a cell. The formula evaluator needs the same
 * conversion, in both directions, for `TODAY`/`DATE`/`YEAR`/... — and two
 * copies of an epoch is precisely the kind of thing that drifts by a day and
 * then disagrees about what a cell means. So the passthrough now uses this too.
 *
 * Everything here works in UTC. A spreadsheet date has no timezone: 2026-10-02
 * is that calendar day, not an instant, and going through local time would
 * shift it by a day for anyone east or west of UTC depending on the hour.
 */

/**
 * Excel has TWO epochs, because of its 1900 leap-year bug.
 *
 * Excel believes 1900-02-29 existed. It did not. So:
 *   - serials 1..59 are 1900-01-01..1900-02-28, which is 1899-12-31 + n;
 *   - serial 60 is the phantom 1900-02-29, which is no real date at all;
 *   - serials 61+ are 1900-03-01 onward, which is 1899-12-30 + n — the extra
 *     day absorbing the date that never happened.
 *
 * Using only the 1899-12-30 epoch (as this code first did, inherited from the
 * save path where it is right for every date anyone types) puts every date
 * before March 1900 one day late. That is a narrow window, but being quietly
 * one day wrong is worse than refusing, so both epochs are modelled and the
 * phantom day is rejected.
 */
const EPOCH_PRE_MARCH_1900 = Date.UTC(1899, 11, 31)
const EPOCH_UTC_MS = Date.UTC(1899, 11, 30)
/** 1900-03-01, the first date the main epoch is correct for. */
const FIRST_MAIN_EPOCH_UTC = Date.UTC(1900, 2, 1)
/** Serial 60 — Excel's 1900-02-29, which never existed. */
const PHANTOM_LEAP_SERIAL = 60
const MS_PER_DAY = 86_400_000

/** An ISO calendar date, as a spreadsheet means it: a day, with no timezone. */
export type CalendarDate = {
  readonly year: number
  readonly month: number
  readonly day: number
}

export const ISO_DATE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})$/

/**
 * The serial number for a calendar date, or `null` when it is not a real date.
 *
 * `Date.UTC` NORMALIZES an impossible date (2024-02-30 becomes 2024-03-01)
 * rather than rejecting it, so the round-trip check is what actually catches
 * one — without it, `DATE(2024,2,30)` would silently answer for a different
 * day than the one asked about.
 */
export function serialFromDate(date: CalendarDate): number | null {
  const { year, month, day } = date
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null
  const asUtc = Date.UTC(year, month - 1, day)
  if (!Number.isFinite(asUtc)) return null
  const roundTrip = new Date(asUtc)
  if (
    roundTrip.getUTCFullYear() !== year ||
    roundTrip.getUTCMonth() !== month - 1 ||
    roundTrip.getUTCDate() !== day
  ) {
    return null
  }
  const epoch = asUtc < FIRST_MAIN_EPOCH_UTC ? EPOCH_PRE_MARCH_1900 : EPOCH_UTC_MS
  const serial = Math.round((asUtc - epoch) / MS_PER_DAY)
  return serial < 1 ? null : serial
}

/** The serial for an `YYYY-MM-DD` string, or `null` for anything else. */
export function serialFromIsoDate(text: string): number | null {
  const match = ISO_DATE_PATTERN.exec(text.trim())
  if (!match) return null
  return serialFromDate({ year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) })
}

/**
 * The calendar date a serial means, or `null` for a serial outside the range a
 * spreadsheet date can represent.
 *
 * The fractional part is the time of day and is discarded here — `Math.floor`,
 * not rounding, so 2026-10-02 18:00 is still 2026-10-02 rather than the 3rd.
 * Serial 0 is rejected along with everything below it: Excel shows it as the
 * fictitious 1900-01-00, which is not a date anything should answer with.
 */
export function dateFromSerial(serial: number): CalendarDate | null {
  if (!Number.isFinite(serial) || serial < 1) return null
  const whole = Math.floor(serial)
  // Serial 60 is Excel's 1900-02-29, a day that never existed. There is no
  // honest calendar date to return for it.
  if (whole === PHANTOM_LEAP_SERIAL) return null
  const epoch = whole < PHANTOM_LEAP_SERIAL ? EPOCH_PRE_MARCH_1900 : EPOCH_UTC_MS
  const date = new Date(epoch + whole * MS_PER_DAY)
  if (Number.isNaN(date.getTime())) return null
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() }
}

/** The day of the week for a serial, 0 = Sunday, or `null` for a serial that is not a date. */
export function weekdayFromSerial(serial: number): number | null {
  const date = dateFromSerial(serial)
  if (!date) return null
  // Derived from the real calendar rather than from `serial % 7`: the two
  // epochs mean the modulo answer is off by one for part of the range, and a
  // weekday that is quietly wrong is hard to notice and easy to rely on.
  return new Date(Date.UTC(date.year, date.month - 1, date.day)).getUTCDay()
}

/** `YYYY-MM-DD` for a serial, or `null` when it is not a representable date. */
export function isoDateFromSerial(serial: number): string | null {
  const date = dateFromSerial(serial)
  if (!date) return null
  const pad = (n: number): string => String(n).padStart(2, '0')
  return `${String(date.year).padStart(4, '0')}-${pad(date.month)}-${pad(date.day)}`
}

/** `YYYY-MM-DD hh:mm` for a serial whose fraction carries a time of day. */
export function isoDateTimeFromSerial(serial: number): string | null {
  const day = isoDateFromSerial(serial)
  if (!day) return null
  // The fraction of a day, turned into minutes. Rounded, because a time built
  // from a float otherwise lands on 16:59 instead of 17:00.
  const minutesInDay = Math.round((serial - Math.floor(serial)) * 24 * 60)
  const pad = (n: number): string => String(n).padStart(2, '0')
  // A fraction that rounds up to a whole day belongs to the NEXT day.
  if (minutesInDay >= 24 * 60) {
    const next = isoDateFromSerial(Math.floor(serial) + 1)
    return next === null ? null : `${next} 00:00`
  }
  return `${day} ${pad(Math.floor(minutesInDay / 60))}:${pad(minutesInDay % 60)}`
}

/** Today's serial, by the local calendar — "today" is the day the user is living in, not the UTC one. */
export function todaySerial(now: Date = new Date()): number {
  const serial = serialFromDate({ year: now.getFullYear(), month: now.getMonth() + 1, day: now.getDate() })
  // A local date always converts, so this fallback is unreachable in practice.
  return serial ?? 0
}

/** Now, as a serial with a time-of-day fraction, by the local clock. */
export function nowSerial(now: Date = new Date()): number {
  const minutes = now.getHours() * 60 + now.getMinutes()
  return todaySerial(now) + minutes / (24 * 60)
}
