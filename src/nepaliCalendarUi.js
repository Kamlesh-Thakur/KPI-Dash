/**
 * Bikram Sambat month grid helpers (Gregorian YYYY-MM-DD still used for filters).
 */
import NepaliDate from 'nepali-date-converter';

export function getDaysInBsMonth(bsYear, bsMonthIndex) {
  for (let d = 32; d >= 28; d--) {
    const nd = new NepaliDate(bsYear, bsMonthIndex, d);
    if (nd.getYear() === bsYear && nd.getMonth() === bsMonthIndex) return d;
  }
  return 30;
}

/**
 * @returns {{ ad: Date, muted: boolean }[]}
 */
export function buildBsMonthCells(bsYear, bsMonthIndex) {
  const dim = getDaysInBsMonth(bsYear, bsMonthIndex);
  const first = new NepaliDate(bsYear, bsMonthIndex, 1);
  const startDow = first.toJsDate().getDay();

  const prevYm =
    bsMonthIndex === 0
      ? { y: bsYear - 1, m: 11 }
      : { y: bsYear, m: bsMonthIndex - 1 };
  const prevDim = getDaysInBsMonth(prevYm.y, prevYm.m);

  const cells = [];
  for (let i = 0; i < startDow; i++) {
    const day = prevDim - startDow + i + 1;
    const nd = new NepaliDate(prevYm.y, prevYm.m, day);
    cells.push({ ad: nd.toJsDate(), muted: true });
  }
  for (let d = 1; d <= dim; d++) {
    const nd = new NepaliDate(bsYear, bsMonthIndex, d);
    cells.push({ ad: nd.toJsDate(), muted: false });
  }
  const nextYm =
    bsMonthIndex === 11
      ? { y: bsYear + 1, m: 0 }
      : { y: bsYear, m: bsMonthIndex + 1 };
  let nextDay = 1;
  while (cells.length < 42) {
    const nd = new NepaliDate(nextYm.y, nextYm.m, nextDay);
    cells.push({ ad: nd.toJsDate(), muted: true });
    nextDay += 1;
  }
  return cells;
}

export function bsMonthTitle(bsYear, bsMonthIndex) {
  return new NepaliDate(bsYear, bsMonthIndex, 1).format('MMMM YYYY');
}

export function gregorianYmToBs(ymStr) {
  const m = /^(\d{4})-(\d{2})$/.exec(String(ymStr || '').trim());
  if (!m) return null;
  const ad = new Date(Number(m[1]), Number(m[2]) - 1, 1);
  const nd = NepaliDate.fromAD(ad);
  return { bsYear: nd.getYear(), bsMonth: nd.getMonth() };
}

/** BS year + month index 0–11 (Baisakh=00 … Chaitra=11). Not the same as Gregorian YYYY-MM. */
export function formatBsMonthCode(bsYear, bsMonthIndex) {
  return `${bsYear}-${String(bsMonthIndex).padStart(2, '0')}`;
}

export function parseBsMonthCode(str) {
  const m = /^(\d{4})-(\d{2})$/.exec(String(str || '').trim());
  if (!m) return null;
  const bsYear = Number(m[1]);
  const bsMonthIndex = Number(m[2]);
  if (Number.isNaN(bsYear) || Number.isNaN(bsMonthIndex) || bsMonthIndex < 0 || bsMonthIndex > 11) {
    return null;
  }
  return { bsYear, bsMonthIndex };
}

/** Keep hidden BS month in sync when only Gregorian YYYY-MM is known (English picker / init). */
export function syncBsMonthHiddenFromGregorianYm(monthYYYYMM) {
  const match = /^(\d{4})-(\d{2})$/.exec(String(monthYYYYMM || '').trim());
  const el = document.getElementById('filter-bs-month');
  if (!match || !el) return;
  const ad = new Date(Number(match[1]), Number(match[2]) - 1, 1);
  const nd = NepaliDate.fromAD(ad);
  el.value = formatBsMonthCode(nd.getYear(), nd.getMonth());
}
