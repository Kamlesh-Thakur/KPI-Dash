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

/**
 * Latin Bikram Sambat labels (Chaitra, Falgun, …) — same idea as the calendar picker, not Devanagari.
 * Second arg to NepaliDate.format: 'en' = English month names, 'np' = Nepali script.
 */
function safeNepaliFormat(d, pattern, lang = 'en') {
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
      return safeNepaliFormat(d, 'DD MMMM YYYY', 'en');
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
      return safeNepaliFormat(d, 'DD MMM, YYYY', 'en');
    }
    return englishShort(d);
  } catch (e) {
    return englishShort(d);
  }
}

/** Chart x-axis / single-point display: full month name + day */
export function formatDisplayDateAxis(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
  try {
    if (getCalendarSystem() === 'nepali') {
      return safeNepaliFormat(d, 'MMMM DD', 'en');
    }
    return d.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
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

/**
 * Time-series category axis: full month name + day on every tick (BS: MMMM DD, AD: long month).
 * `index` / `categories` kept for call-site compatibility with ECharts formatters.
 */
export function formatCategoryAxisDateLabel(value, index, categories) {
  try {
    if (!Array.isArray(categories) || categories.length === 0) {
      return formatDisplayDateAxisFromCategory(value);
    }
    const d = categoryStringToLocalDate(value);
    if (!d || isNaN(d.getTime())) return String(value ?? '');
    if (getCalendarSystem() === 'nepali') {
      return safeNepaliFormat(d, 'MMMM DD', 'en');
    }
    return d.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
  } catch (e) {
    return formatDisplayDateAxisFromCategory(value);
  }
}

export function formatDisplayDateRange(start, end) {
  if (!start || !end || !(start instanceof Date) || !(end instanceof Date)) return '—';
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return '—';
  try {
    if (getCalendarSystem() === 'nepali') {
      const a = safeNepaliFormat(start, 'DD MMM YYYY', 'en');
      const b = safeNepaliFormat(end, 'DD MMM YYYY', 'en');
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
        header = safeNepaliFormat(d, 'ddd DD, MMMM YYYY', 'en');
      } else {
        header = d.toLocaleDateString('en-US', {
          weekday: 'short',
          month: 'long',
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
