import { describe, expect, it } from 'vitest'

import { DEFAULT_DATE_PICTURE, DEFAULT_TIME_PICTURE, formatDatePicture } from '../dateFormat'

// A fixed, unambiguous instant: Wednesday, January 7, 2026, 09:05:03 local.
const SAMPLE_DATE = new Date(2026, 0, 7, 9, 5, 3)
// Thursday, December 31, 2026, 15:45:00 local (afternoon, for AM/PM + hh checks).
const AFTERNOON_DATE = new Date(2026, 11, 31, 15, 45, 0)

describe('formatDatePicture', () => {
  it('formats the default DATE picture (M/d/yyyy)', () => {
    expect(formatDatePicture(SAMPLE_DATE, DEFAULT_DATE_PICTURE)).toBe('1/7/2026')
  })

  it('formats the default TIME picture (h:mm am/pm)', () => {
    expect(formatDatePicture(SAMPLE_DATE, DEFAULT_TIME_PICTURE)).toBe('9:05 am')
  })

  it('formats a two-digit month/day/year picture', () => {
    expect(formatDatePicture(SAMPLE_DATE, 'MM/dd/yy')).toBe('01/07/26')
  })

  it('formats long month and day-of-week names', () => {
    expect(formatDatePicture(SAMPLE_DATE, 'dddd, MMMM d, yyyy')).toBe('Wednesday, January 7, 2026')
  })

  it('formats short month and day-of-week names', () => {
    expect(formatDatePicture(SAMPLE_DATE, 'ddd, MMM d')).toBe('Wed, Jan 7')
  })

  it('distinguishes uppercase M (month) from lowercase m (minutes)', () => {
    expect(formatDatePicture(SAMPLE_DATE, 'M/d/yyyy h:mm')).toBe('1/7/2026 9:05')
  })

  it('formats 24-hour time with HH', () => {
    expect(formatDatePicture(AFTERNOON_DATE, 'HH:mm:ss')).toBe('15:45:00')
  })

  it('formats 12-hour time with hh and AM/PM, converting hour 0 to 12', () => {
    const noon = new Date(2026, 0, 1, 0, 0, 0)
    expect(formatDatePicture(noon, 'hh:mm AM/PM')).toBe('12:00 AM')
    expect(formatDatePicture(AFTERNOON_DATE, 'hh:mm AM/PM')).toBe('03:45 PM')
  })

  it('formats the single-letter am/pm variants (a/p, A/P)', () => {
    expect(formatDatePicture(SAMPLE_DATE, 'h:mm a/p')).toBe('9:05 a')
    expect(formatDatePicture(AFTERNOON_DATE, 'h:mm A/P')).toBe('3:45 P')
  })

  it('passes single-quoted literal text through unchanged', () => {
    expect(formatDatePicture(SAMPLE_DATE, "'Week of' M/d")).toBe('Week of 1/7')
  })

  it('leaves punctuation and unrecognized characters as-is', () => {
    expect(formatDatePicture(SAMPLE_DATE, 'yyyy-MM-dd')).toBe('2026-01-07')
  })
})
