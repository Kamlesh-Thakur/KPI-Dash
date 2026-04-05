/**
 * Shared themed calendar popover: week, single day, month, and date range.
 */

const CAL_ICON_SVG = `<svg class="picker-trigger-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><rect x="3" y="4" width="18" height="18" rx="2" stroke="currentColor" stroke-width="2"/><path d="M16 2v4M8 2v4M3 10h18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>`;

const CHEVRON_SVG = `<svg class="picker-trigger-chevron" width="12" height="12" viewBox="0 0 10 6" aria-hidden="true"><path d="M1 1l4 4 4-4" stroke="currentColor" fill="none" stroke-width="1.5"/></svg>`;

function getSundayOfWeek(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() - d.getDay());
  return d;
}

function formatYmdLocal(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function parseYmdLocal(ymd) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || '').trim());
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const day = Number(m[3]);
  const dt = new Date(y, mo, day);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo || dt.getDate() !== day) return null;
  return dt;
}

function dateInSelectedWeek(d, weekSunday) {
  const t = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const s = new Date(weekSunday.getFullYear(), weekSunday.getMonth(), weekSunday.getDate()).getTime();
  return t >= s && t <= s + 6 * 86400000;
}

function sameDay(a, b) {
  return (
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate()
  );
}

function dateInRangeInclusive(d, from, to) {
  const t = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const a = new Date(from.getFullYear(), from.getMonth(), from.getDate()).getTime();
  const b = new Date(to.getFullYear(), to.getMonth(), to.getDate()).getTime();
  return t >= Math.min(a, b) && t <= Math.max(a, b);
}

export function createCalendarPicker({
  setFilter,
  formatYmdLocal: fmt = formatYmdLocal,
  parseYmdLocal: parse = parseYmdLocal,
  getSundayOfWeek: gsw = getSundayOfWeek,
  renderAll
}) {
  let calUIMode = 'week';
  let calViewMonth = null;
  let calViewYear = null;
  let calPopoverOpen = false;
  let activeTrigger = null;
  let rangePickStart = null;

  function positionPopover() {
    const pop = document.getElementById('shared-calendar-popover');
    if (!pop || !activeTrigger) return;
    const r = activeTrigger.getBoundingClientRect();
    const pad = 8;
    let left = r.left;
    let top = r.bottom + 6;
    const pw = pop.offsetWidth || 320;
    const ph = pop.offsetHeight || 360;
    if (left + pw > window.innerWidth - pad) left = Math.max(pad, window.innerWidth - pw - pad);
    if (top + ph > window.innerHeight - pad) top = Math.max(pad, r.top - ph - 6);
    pop.style.left = `${left}px`;
    pop.style.top = `${top}px`;
  }

  function updateWeekTrigger() {
    const el = document.getElementById('filter-date-week');
    const textEl = document.getElementById('week-picker-trigger-text');
    if (!el || !textEl) return;
    const sun = parse(el.value) || gsw(new Date());
    const sat = new Date(sun.getFullYear(), sun.getMonth(), sun.getDate() + 6);
    const opts = { month: 'short', day: 'numeric', year: 'numeric' };
    textEl.textContent = `${sun.toLocaleDateString('en-US', opts)} – ${sat.toLocaleDateString('en-US', opts)}`;
  }

  function updateDailyTrigger() {
    const el = document.getElementById('filter-date-anchor');
    const textEl = document.getElementById('daily-picker-trigger-text');
    if (!el || !textEl) return;
    const d = parse(el.value) || new Date();
    textEl.textContent = d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }

  function updateMonthTrigger() {
    const el = document.getElementById('filter-date-month');
    const textEl = document.getElementById('month-picker-trigger-text');
    if (!el || !textEl) return;
    const m = /^(\d{4})-(\d{2})$/.exec(el.value || '');
    if (!m) {
      const t = new Date();
      textEl.textContent = t.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
      return;
    }
    const dt = new Date(Number(m[1]), Number(m[2]) - 1, 1);
    textEl.textContent = dt.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  }

  function updateRangeTrigger() {
    const fromEl = document.getElementById('filter-date-from');
    const toEl = document.getElementById('filter-date-to');
    const textEl = document.getElementById('range-picker-trigger-text');
    if (!fromEl || !toEl || !textEl) return;
    const a = parse(fromEl.value);
    const b = parse(toEl.value);
    const opts = { month: 'short', day: 'numeric', year: 'numeric' };
    if (a && b) {
      textEl.textContent = `${a.toLocaleDateString('en-US', opts)} – ${b.toLocaleDateString('en-US', opts)}`;
    } else if (a) {
      textEl.textContent = `${a.toLocaleDateString('en-US', opts)} – …`;
    } else {
      textEl.textContent = 'Select date range';
    }
  }

  function clearWeekHover(grid) {
    grid.querySelectorAll('.week-cal-day--hover-week').forEach((n) => n.classList.remove('week-cal-day--hover-week'));
  }

  function applyWeekHover(grid, hoverSunday) {
    clearWeekHover(grid);
    if (!hoverSunday) return;
    grid.querySelectorAll('.week-cal-day[data-ymd]').forEach((btn) => {
      const d = parse(btn.dataset.ymd);
      if (d && dateInSelectedWeek(d, hoverSunday)) btn.classList.add('week-cal-day--hover-week');
    });
  }

  function renderDayGrid(grid, viewY, viewM, opts) {
    const {
      mode,
      selectedSunday,
      selectedDay,
      rangeFrom,
      rangeTo,
      rangeTempStart,
      onDayClick,
      showWeekHover
    } = opts;

    const gridStart = gsw(new Date(viewY, viewM, 1));
    grid.replaceChildren();

    for (let i = 0; i < 42; i++) {
      const d = new Date(gridStart.getFullYear(), gridStart.getMonth(), gridStart.getDate() + i);
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'week-cal-day';
      btn.dataset.ymd = fmt(d);
      if (d.getMonth() !== viewM) btn.classList.add('week-cal-day--muted');

      if (mode === 'week' && selectedSunday && dateInSelectedWeek(d, selectedSunday)) {
        btn.classList.add('week-cal-day--in-week');
      } else if (mode === 'daily' && selectedDay && sameDay(d, selectedDay)) {
        btn.classList.add('week-cal-day--selected-day');
      } else if (mode === 'range' && rangeFrom && rangeTo) {
        if (dateInRangeInclusive(d, rangeFrom, rangeTo)) btn.classList.add('week-cal-day--in-range');
      } else if (mode === 'range' && rangeTempStart && sameDay(d, rangeTempStart)) {
        btn.classList.add('week-cal-day--range-start');
      }

      btn.textContent = String(d.getDate());
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        onDayClick(d);
      });

      grid.appendChild(btn);
    }

    if (showWeekHover) {
      grid.onmouseover = (e) => {
        const btn = e.target.closest('.week-cal-day[data-ymd]');
        if (!btn) return;
        const d = parse(btn.dataset.ymd);
        if (d) applyWeekHover(grid, gsw(d));
      };
      grid.onmouseleave = () => clearWeekHover(grid);
    } else {
      grid.onmouseover = null;
      grid.onmouseleave = null;
    }
  }

  function renderMonthGrid(body, year, selectedMonth0) {
    body.replaceChildren();
    body.className = 'month-cal-grid';
    const names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    names.forEach((name, idx) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'month-cal-month';
      if (idx === selectedMonth0) btn.classList.add('month-cal-month--selected');
      btn.textContent = name;
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const monthEl = document.getElementById('filter-date-month');
        const ym = `${year}-${String(idx + 1).padStart(2, '0')}`;
        if (monthEl) monthEl.value = ym;
        setFilter('dateAnchor', `${ym}-01`);
        updateMonthTrigger();
        closeCalPopover();
        renderAll();
      });
      body.appendChild(btn);
    });
  }

  function renderSharedCalendar() {
    const pop = document.getElementById('shared-calendar-popover');
    const titleEl = document.getElementById('cal-pop-title');
    const weekdaysRow = document.getElementById('cal-weekdays-row');
    const body = document.getElementById('cal-pop-body');
    const todayBtn = document.getElementById('cal-footer-today');
    if (!pop || !titleEl || !body) return;

    if (calUIMode === 'month') {
      if (calViewYear == null) {
        const m = document.getElementById('filter-date-month')?.value;
        const match = /^(\d{4})-(\d{2})$/.exec(m || '');
        calViewYear = match ? Number(match[1]) : new Date().getFullYear();
      }
      titleEl.textContent = String(calViewYear);
      if (weekdaysRow) {
        weekdaysRow.hidden = true;
        weekdaysRow.classList.add('week-cal-weekdays--hidden');
      }
      const m = document.getElementById('filter-date-month')?.value;
      const match = /^(\d{4})-(\d{2})$/.exec(m || '');
      const selM = match ? Number(match[2]) - 1 : new Date().getMonth();
      renderMonthGrid(body, calViewYear, selM);
      if (todayBtn) {
        todayBtn.textContent = 'This month';
        todayBtn.onclick = (e) => {
          e.stopPropagation();
          const t = new Date();
          document.getElementById('filter-date-month').value = `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}`;
          setFilter('dateAnchor', fmt(new Date(t.getFullYear(), t.getMonth(), 1)));
          updateMonthTrigger();
          closeCalPopover();
          renderAll();
        };
      }
      return;
    }

    if (weekdaysRow) {
      weekdaysRow.hidden = false;
      weekdaysRow.classList.remove('week-cal-weekdays--hidden');
    }
    body.className = 'week-cal-grid';

    if (!calViewMonth) {
      const n = new Date();
      calViewMonth = new Date(n.getFullYear(), n.getMonth(), 1);
    }
    const y = calViewMonth.getFullYear();
    const m = calViewMonth.getMonth();
    titleEl.textContent = new Date(y, m, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

    if (calUIMode === 'week') {
      const hidden = document.getElementById('filter-date-week');
      const anchor = parse(hidden?.value) || gsw(new Date());
      const selSun = gsw(anchor);
      renderDayGrid(body, y, m, {
        mode: 'week',
        selectedSunday: selSun,
        onDayClick: (d) => {
          const sun = gsw(d);
          const el = document.getElementById('filter-date-week');
          const v = fmt(sun);
          if (el) el.value = v;
          setFilter('dateAnchor', v);
          updateWeekTrigger();
          closeCalPopover();
          renderAll();
        },
        showWeekHover: true
      });
      if (todayBtn) {
        todayBtn.textContent = 'This week';
        todayBtn.onclick = (e) => {
          e.stopPropagation();
          const sun = gsw(new Date());
          document.getElementById('filter-date-week').value = fmt(sun);
          setFilter('dateAnchor', fmt(sun));
          calViewMonth = new Date(sun.getFullYear(), sun.getMonth(), 1);
          updateWeekTrigger();
          closeCalPopover();
          renderAll();
        };
      }
      return;
    }

    if (calUIMode === 'daily') {
      const hidden = document.getElementById('filter-date-anchor');
      const anchor = parse(hidden?.value) || new Date();
      renderDayGrid(body, y, m, {
        mode: 'daily',
        selectedDay: anchor,
        onDayClick: (d) => {
          const v = fmt(d);
          document.getElementById('filter-date-anchor').value = v;
          setFilter('dateAnchor', v);
          updateDailyTrigger();
          closeCalPopover();
          renderAll();
        },
        showWeekHover: false
      });
      if (todayBtn) {
        todayBtn.textContent = 'Today';
        todayBtn.onclick = (e) => {
          e.stopPropagation();
          const t = new Date();
          const v = fmt(t);
          document.getElementById('filter-date-anchor').value = v;
          setFilter('dateAnchor', v);
          calViewMonth = new Date(t.getFullYear(), t.getMonth(), 1);
          updateDailyTrigger();
          closeCalPopover();
          renderAll();
        };
      }
      return;
    }

    if (calUIMode === 'range') {
      const fromEl = document.getElementById('filter-date-from');
      const toEl = document.getElementById('filter-date-to');
      const from = parse(fromEl?.value);
      const to = parse(toEl?.value);
      renderDayGrid(body, y, m, {
        mode: 'range',
        rangeFrom: from,
        rangeTo: to,
        rangeTempStart: rangePickStart,
        onDayClick: (d) => {
          if (!rangePickStart) {
            rangePickStart = d;
            renderSharedCalendar();
            return;
          }
          let a = rangePickStart;
          let b = d;
          if (a > b) [a, b] = [b, a];
          fromEl.value = fmt(a);
          toEl.value = fmt(b);
          setFilter('dateFrom', fromEl.value);
          setFilter('dateTo', toEl.value);
          rangePickStart = null;
          updateRangeTrigger();
          closeCalPopover();
          renderAll();
        },
        showWeekHover: true
      });
      if (todayBtn) {
        todayBtn.textContent = 'Today';
        todayBtn.onclick = (e) => {
          e.stopPropagation();
          const t = new Date();
          const v = fmt(t);
          fromEl.value = v;
          toEl.value = v;
          setFilter('dateFrom', v);
          setFilter('dateTo', v);
          rangePickStart = null;
          updateRangeTrigger();
          closeCalPopover();
          renderAll();
        };
      }
    }
  }

  function openCalPopover(mode, triggerEl) {
    closeCalPopover();
    calUIMode = mode;
    activeTrigger = triggerEl;
    rangePickStart = null;
    const pop = document.getElementById('shared-calendar-popover');
    if (!pop) return;

    if (mode === 'month') {
      calViewYear = null;
    } else {
      let ref = new Date();
      if (mode === 'week') {
        const hid = document.getElementById('filter-date-week');
        ref = parse(hid?.value) || gsw(new Date());
      } else if (mode === 'daily') {
        ref = parse(document.getElementById('filter-date-anchor')?.value) || new Date();
      } else if (mode === 'range') {
        ref = parse(document.getElementById('filter-date-from')?.value) || new Date();
      }
      calViewMonth = new Date(ref.getFullYear(), ref.getMonth(), 1);
    }

    calPopoverOpen = true;
    pop.hidden = false;
    triggerEl.setAttribute('aria-expanded', 'true');
    renderSharedCalendar();
    requestAnimationFrame(() => {
      positionPopover();
      window.addEventListener('scroll', positionPopover, true);
      window.addEventListener('resize', positionPopover);
    });
  }

  function closeCalPopover() {
    const pop = document.getElementById('shared-calendar-popover');
    window.removeEventListener('scroll', positionPopover, true);
    window.removeEventListener('resize', positionPopover);
    calPopoverOpen = false;
    rangePickStart = null;
    if (activeTrigger) {
      activeTrigger.setAttribute('aria-expanded', 'false');
      activeTrigger = null;
    }
    if (pop) pop.hidden = true;
  }

  function initCalendarPickerInner(today) {
    const pop = document.getElementById('shared-calendar-popover');
    if (!pop) return;

    document.getElementById('cal-prev')?.addEventListener('click', (e) => {
      e.stopPropagation();
      if (calUIMode === 'month') {
        calViewYear = (calViewYear ?? new Date().getFullYear()) - 1;
      } else {
        if (!calViewMonth) {
          const n = new Date();
          calViewMonth = new Date(n.getFullYear(), n.getMonth(), 1);
        }
        calViewMonth.setMonth(calViewMonth.getMonth() - 1);
      }
      renderSharedCalendar();
      requestAnimationFrame(positionPopover);
    });
    document.getElementById('cal-next')?.addEventListener('click', (e) => {
      e.stopPropagation();
      if (calUIMode === 'month') {
        calViewYear = (calViewYear ?? new Date().getFullYear()) + 1;
      } else {
        if (!calViewMonth) {
          const n = new Date();
          calViewMonth = new Date(n.getFullYear(), n.getMonth(), 1);
        }
        calViewMonth.setMonth(calViewMonth.getMonth() + 1);
      }
      renderSharedCalendar();
      requestAnimationFrame(positionPopover);
    });

    document.getElementById('cal-close')?.addEventListener('click', (e) => {
      e.stopPropagation();
      closeCalPopover();
    });

    pop.addEventListener('click', (e) => e.stopPropagation());

    const triggers = [
      ['week-picker-trigger', 'week'],
      ['daily-picker-trigger', 'daily'],
      ['month-picker-trigger', 'month'],
      ['range-picker-trigger', 'range']
    ];

    triggers.forEach(([id, mode]) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        if (calPopoverOpen && activeTrigger === el) {
          closeCalPopover();
        } else {
          openCalPopover(mode, el);
        }
      });
    });

    document.addEventListener(
      'pointerdown',
      (e) => {
        if (!calPopoverOpen) return;
        if (pop.contains(e.target)) return;
        if (triggers.some(([id]) => document.getElementById(id)?.contains(e.target))) return;
        closeCalPopover();
      },
      true
    );

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && calPopoverOpen) closeCalPopover();
    });
  }

  function syncCalTriggers(opts) {
    const { showWeek, showDaily, showMonth, showRange } = opts;
    if (showWeek) {
      updateWeekTrigger();
      if (calPopoverOpen && calUIMode === 'week') renderSharedCalendar();
    }
    if (showDaily) {
      updateDailyTrigger();
      if (calPopoverOpen && calUIMode === 'daily') renderSharedCalendar();
    }
    if (showMonth) {
      updateMonthTrigger();
      if (calPopoverOpen && calUIMode === 'month') renderSharedCalendar();
    }
    if (showRange) {
      updateRangeTrigger();
      if (calPopoverOpen && calUIMode === 'range') renderSharedCalendar();
    }
  }

  return {
    initCalendarPickerInner,
    syncCalTriggers,
    updateWeekTrigger,
    updateDailyTrigger,
    updateMonthTrigger,
    updateRangeTrigger,
    closeCalPopover,
    renderSharedCalendar,
    CAL_ICON_SVG,
    CHEVRON_SVG
  };
}
