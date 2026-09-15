/**
 * Atlas — DATE/TIME field `\@` picture formatting (DEFER-5 / DXS-20)
 *
 * Word's date-time field switch (`\@ "picture"`) uses case-sensitive
 * tokens borrowed from the same family as VBA's `Format$` date pictures —
 * notably `M`/`MM`/`MMM`/`MMMM` for month vs. `m`/`mm` for MINUTES (not
 * month), which is why this is a case-SENSITIVE token scan rather than
 * needing the "is this m near an h?" context-sensitivity some other
 * date-format systems require. A `'literal text'` span (single-quoted)
 * passes through unchanged, letting a picture embed fixed words.
 */
const MONTHS_LONG: ReadonlyArray<string> = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
]
const MONTHS_SHORT: ReadonlyArray<string> = MONTHS_LONG.map((month) => month.slice(0, 3))
const DAYS_LONG: ReadonlyArray<string> = [
  'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday',
]
const DAYS_SHORT: ReadonlyArray<string> = DAYS_LONG.map((day) => day.slice(0, 3))

/** Longest-token-first so e.g. `MMMM` isn't matched as `MM` + `MM`. */
const PICTURE_TOKEN_RE =
  /'[^']*'|dddd|ddd|dd|d|MMMM|MMM|MM|M|yyyy|yy|HH|H|hh|h|mm|m|ss|s|AM\/PM|am\/pm|A\/P|a\/p/g

export const DEFAULT_DATE_PICTURE = 'M/d/yyyy'
export const DEFAULT_TIME_PICTURE = 'h:mm am/pm'

export function formatDatePicture(date: Date, picture: string): string {
  return picture.replace(PICTURE_TOKEN_RE, (token) => formatToken(token, date))
}

function formatToken(token: string, date: Date): string {
  if (token.startsWith("'")) {
    return token.slice(1, -1)
  }

  switch (token) {
    case 'dddd':
      return DAYS_LONG[date.getDay()] ?? ''
    case 'ddd':
      return DAYS_SHORT[date.getDay()] ?? ''
    case 'dd':
      return pad2(date.getDate())
    case 'd':
      return String(date.getDate())
    case 'MMMM':
      return MONTHS_LONG[date.getMonth()] ?? ''
    case 'MMM':
      return MONTHS_SHORT[date.getMonth()] ?? ''
    case 'MM':
      return pad2(date.getMonth() + 1)
    case 'M':
      return String(date.getMonth() + 1)
    case 'yyyy':
      return String(date.getFullYear())
    case 'yy':
      return pad2(date.getFullYear() % 100)
    case 'HH':
      return pad2(date.getHours())
    case 'H':
      return String(date.getHours())
    case 'hh':
      return pad2(to12Hour(date.getHours()))
    case 'h':
      return String(to12Hour(date.getHours()))
    case 'mm':
      return pad2(date.getMinutes())
    case 'm':
      return String(date.getMinutes())
    case 'ss':
      return pad2(date.getSeconds())
    case 's':
      return String(date.getSeconds())
    case 'AM/PM':
      return date.getHours() < 12 ? 'AM' : 'PM'
    case 'am/pm':
      return date.getHours() < 12 ? 'am' : 'pm'
    case 'A/P':
      return date.getHours() < 12 ? 'A' : 'P'
    case 'a/p':
      return date.getHours() < 12 ? 'a' : 'p'
    default:
      return token
  }
}

function pad2(value: number): string {
  return String(value).padStart(2, '0')
}

function to12Hour(hour24: number): number {
  const remainder = hour24 % 12
  return remainder === 0 ? 12 : remainder
}
