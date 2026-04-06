/**
 * User-facing dates: English (Gregorian) vs Nepali (Bikram Sambat) from calendar preference.
 */
import NepaliDate from 'nepali-date-converter';
import { getCalendarSystem } from './calendarPrefs.js';

function categoryStringToLocalDate(v) {
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v.trim())) {
    const y = Number(v.slice(0, 4));
    const mo = Number(v.slice(5, 7));
    const day = Number(v.slice(8, 10));
    return new Date(y, mo - 1, day);
  }
  const d = new Date(v);
  return Number.isNaN(d.getTime()) ? null : d;
}

function nepaliFromAd(d) {
  if (typeof NepaliDate.fromAD === 'function') {
    return NepaliDate.fromAD(d);
  }
  return new NepaliDate(d);
}

/** Never throw — chart/grid formatters must not break renderAll */
function safeNepaliFormat(d, pattern, lang = 'np') {
  try {
    return nepaliFromAd(d).format(pattern, lang);
  } catch (e) {
    console.warn('[dateDisplay] Nepali format fallback', pattern, e);
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }
}

function englishShort(d) {
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

function englishIso(d) {
  return d.toISOString().slice(0, 10);
}

/** Grids / full cell dates */
export function formatDisplayDate(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '—';
  try {
    if (getCalendarSystem() === 'nepali') {
      return safeNepaliFormat(d, 'DD MMMM YYYY', 'np');
    }
    return englishIso(d);
  } catch (e) {
    console.warn('[dateDisplay] formatDisplayDate', e);
    return englishIso(d);
  }
}

export function formatDisplayDateShort(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '—';
  try {
    if (getCalendarSystem() === 'nepali') {
      return safeNepaliFormat(d, 'DD MMM, YYYY', 'np');
    }
    return englishShort(d);
  } catch (e) {
    return englishShort(d);
  }
}

/** Compact chart x-axis */
export function formatDisplayDateAxis(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  try {
    if (getCalendarSystem() === 'nepali') {
      return safeNepaliFormat(d, 'DD MMM', 'np');
    }
    return englishShort(d);
  } catch (e) {
    return englishShort(d);
  }
}

export function formatDisplayDateAxisFromCategory(value) {
  try {
    const d = categoryStringToLocalDate(value);
    if (!d || isNaN(d.getTime())) return String(value ?? '');
    return formatDisplayDateAxis(d);
  } catch (e) {
    return String(value ?? '');
  }
}

export function formatDisplayDateRange(start, end) {
  if (!start || !end || !(start instanceof Date) || !(end instanceof Date)) return '—';
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return '—';
  try {
    if (getCalendarSystem() === 'nepali') {
      const a = safeNepaliFormat(start, 'DD MMM YYYY', 'np');
      const b = safeNepaliFormat(end, 'DD MMM YYYY', 'np');
      if (start.getTime() === end.getTime()) return a;
      return `${a} – ${b}`;
    }
    const opts = { month: 'short', day: 'numeric', year: 'numeric' };
    const s = start.toLocaleDateString('en-US', opts);
    const e = end.toLocaleDateString('en-US', opts);
    if (start.getTime() === end.getTime()) return s;
    return `${s} - ${e}`;
  } catch (e) {
    console.warn('[dateDisplay] formatDisplayDateRange', e);
    return '—';
  }
}

/** ECharts axis tooltip for time-series with YYYY-MM-DD categories */
export function buildAxisTooltipHtmlWithDates(params) {
  try {
    if (!params?.length) return '';
    const raw = params[0].axisValue;
    const d = categoryStringToLocalDate(raw);
    let header;
    if (d && !isNaN(d.getTime())) {
      if (getCalendarSystem() === 'nepali') {
        header = safeNepaliFormat(d, 'ddd DD, MMMM YYYY', 'np');
      } else {
        header = d.toLocaleDateString('en-US', {
          weekday: 'short',
          month: 'short',
          day: 'numeric',
          year: 'numeric'
        });
      }
    } else {
      header = String(raw ?? '');
    }
    let html = `${header}<br/>`;
    params.forEach((p) => {
      const name = p.seriesName != null && p.seriesName !== '' ? p.seriesName : '—';
      const marker = p.marker || '';
      html += `${marker}${name}: ${p.value}<br/>`;
    });
    return html.trim();
  } catch (e) {
    console.warn('[dateDisplay] buildAxisTooltipHtmlWithDates', e);
    return '';
  }
}
