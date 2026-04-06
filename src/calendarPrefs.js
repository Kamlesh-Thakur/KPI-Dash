/**
 * Persist English vs Nepali (Bikram Sambat) calendar display preference.
 * Internal filters stay Gregorian (YYYY-MM-DD); this only affects UI.
 */

const COOKIE_NAME = 'kpi_calendar_system';
const MAX_AGE_SEC = 365 * 24 * 60 * 60;

/** @returns {'english' | 'nepali'} */
export function getCalendarSystem() {
  const v = readCookie();
  if (v === 'nepali' || v === 'english') return v;
  return 'english';
}

function readCookie() {
  if (typeof document === 'undefined') return null;
  const m = document.cookie.match(new RegExp(`(?:^|; )${COOKIE_NAME}=([^;]*)`));
  return m ? decodeURIComponent(m[1]) : null;
}

/** @param {'english' | 'nepali'} system */
export function setCalendarSystem(system) {
  if (system !== 'english' && system !== 'nepali') return;
  document.cookie = `${COOKIE_NAME}=${encodeURIComponent(system)}; path=/; max-age=${MAX_AGE_SEC}; SameSite=Lax`;
}
