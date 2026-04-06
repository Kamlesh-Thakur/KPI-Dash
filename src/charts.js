/**
 * Charts Module — all ECharts visualizations
 */
import * as echarts from 'echarts';
import {
  getFilteredRawData,
  getFilteredIncidentData,
  getBranchEfficiencyData,
  getTeamPerformanceData,
  excelDateToJS,
  formatDuration,
  formatNumber
} from './dataStore.js';
import { formatCategoryAxisDateLabel, buildAxisTooltipHtmlWithDates } from './dateDisplay.js';
import {
  chartAxisLabelCat,
  chartAxisLabelValue,
  chartAxisLine,
  chartSplitLine,
  chartTooltipTheme,
  chartHeatmapCellTextColor,
  chartRadarSplitAreaColors,
  chartRadarLineColor
} from './chartTheme.js';

const chartInstances = {};

/** Shown in top KPI cards; omitted from Task KPI Matrix heatmap. */
const TASK_KPI_MATRIX_EXCLUDED_KEYS = new Set([
  'closedWithinSla',
  'closedSameDay',
  'closedWithin24h',
  'kpiAfterExclusion'
]);

const incidentCategoryShowAll = {};

/** Per chart container so Dashboard and Branch KPI toggles stay independent. */
const branchTasksVolumeStateById = {};

function getBranchVolumeState(containerId) {
  if (!branchTasksVolumeStateById[containerId]) {
    branchTasksVolumeStateById[containerId] = {
      showAllBranches: false,
      consideredOnly: false
    };
  }
  return branchTasksVolumeStateById[containerId];
}

function isRowConsideredForKpi(r) {
  return (r['Exceptions'] || '').toString().toLowerCase().includes('consider');
}

/** First word on one line, remainder on the next (e.g. Fiber Support → Fiber\\nSupport). */
function formatTaskTypeAxisLabelTwoLines(name) {
  if (name == null || name === '') return '';
  const s = String(name).trim();
  const i = s.indexOf(' ');
  if (i === -1) return s;
  return `${s.slice(0, i)}\n${s.slice(i + 1).trim()}`;
}

const COLORS = {
  blue: '#6384ff',
  purple: '#a78bfa',
  cyan: '#22d3ee',
  green: '#34d399',
  amber: '#fbbf24',
  red: '#f87171',
  pink: '#f472b6',
  indigo: '#818cf8',
  teal: '#2dd4bf',
  orange: '#fb923c'
};

const PALETTE = [COLORS.blue, COLORS.purple, COLORS.cyan, COLORS.green, COLORS.amber, COLORS.red, COLORS.pink, COLORS.indigo, COLORS.teal, COLORS.orange];

function getOrCreate(containerId) {
  const el = document.getElementById(containerId);
  if (!el) return null;

  // Clear loader
  el.innerHTML = '';

  // Add title + chart body
  const titleEl = document.createElement('div');
  titleEl.className = 'chart-title';
  el.appendChild(titleEl);

  const bodyEl = document.createElement('div');
  bodyEl.className = 'chart-body';
  el.appendChild(bodyEl);

  // Dispose existing instance
  if (chartInstances[containerId]) {
    chartInstances[containerId].dispose();
  }

  const chart = echarts.init(bodyEl, null, { renderer: 'canvas' });
  chartInstances[containerId] = chart;

  return { chart, titleEl };
}

function setTitle(titleEl, text, dotColor = 'blue') {
  titleEl.innerHTML = `<span class="dot ${dotColor}"></span>${text}`;
}

/** Resize all charts */
export function resizeCharts() {
  Object.values(chartInstances).forEach(c => c && c.resize());
}

// ==========================================
// DASHBOARD CHARTS
// ==========================================

export function renderTasksTrend(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Tasks & Incidents Over Time', 'blue');

  const taskData = getFilteredRawData();
  const incidentData = getFilteredIncidentData();

  // Group tasks by completed date
  const taskDateMap = {};
  taskData.forEach(row => {
    const d = excelDateToJS(row['Completed date']);
    if (!d) return;
    const key = d.toISOString().slice(0, 10);
    taskDateMap[key] = (taskDateMap[key] || 0) + 1;
  });

  // Group incidents by completed date
  const incidentDateMap = {};
  incidentData.forEach(row => {
    const d = excelDateToJS(row['Completed date']);
    if (!d) return;
    const key = d.toISOString().slice(0, 10);
    incidentDateMap[key] = (incidentDateMap[key] || 0) + 1;
  });

  // Build a shared timeline across tasks and incidents
  const allDates = [...new Set([
    ...Object.keys(taskDateMap),
    ...Object.keys(incidentDateMap)
  ])].sort((a, b) => a.localeCompare(b));

  const taskCounts = allDates.map(date => taskDateMap[date] || 0);
  const incidentCounts = allDates.map(date => incidentDateMap[date] || 0);

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      ...chartTooltipTheme(13),
      formatter: (params) => buildAxisTooltipHtmlWithDates(params)
    },
    legend: {
      data: ['Tasks', 'Incidents'],
      textStyle: { color: chartAxisLabelCat(), fontSize: 12 },
      top: 0
    },
    grid: { left: 52, right: 20, top: 24, bottom: 56 },
    xAxis: {
      type: 'category',
      data: allDates,
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12, rotate: 30,
        formatter: (v, idx) => formatCategoryAxisDateLabel(v, idx, allDates)
      },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    series: [{
      name: 'Tasks',
      type: 'line',
      data: taskCounts,
      smooth: true,
      symbol: 'none',
      lineStyle: { width: 2.5, color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
        { offset: 0, color: COLORS.blue },
        { offset: 1, color: COLORS.purple }
      ]) },
      areaStyle: {
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: 'rgba(99,132,255,0.3)' },
          { offset: 1, color: 'rgba(99,132,255,0)' }
        ])
      }
    }, {
      name: 'Incidents',
      type: 'line',
      data: incidentCounts,
      smooth: true,
      symbol: 'none',
      lineStyle: { width: 2.5, color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
        { offset: 0, color: COLORS.red },
        { offset: 1, color: COLORS.pink }
      ]) },
      areaStyle: {
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: 'rgba(248,113,113,0.2)' },
          { offset: 1, color: 'rgba(248,113,113,0)' }
        ])
      }
    }],
    animationDuration: 1200,
    animationEasing: 'cubicOut'
  });
}

export function renderTaskTypePie(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Task KPI Matrix', 'purple');

  const { rows, metrics: allMetrics } = buildTaskKPIRows();
  const metrics = allMetrics.filter((m) => !TASK_KPI_MATRIX_EXCLUDED_KEYS.has(m.key));
  const normalizedData = [];
  metrics.forEach((metric, metricIndex) => {
    const values = rows.map(r => Number(r[metric.key]) || 0);
    const min = Math.min(...values, 0);
    const max = Math.max(...values, 1);
    const range = max - min || 1;
    rows.forEach((row, rowIndex) => {
      const rawValue = Number(row[metric.key]) || 0;
      const normalized = (rawValue - min) / range;
      normalizedData.push([metricIndex, rowIndex, normalized, rawValue, metric.kind]);
    });
  });

  const formatValue = (rawValue, kind) => {
    if (kind === 'pct') return `${(rawValue * 100).toFixed(1)}%`;
    if (kind === 'number' && rawValue >= 1000) return rawValue.toLocaleString('en-US');
    return Number(rawValue.toFixed(2)).toString();
  };

  chart.setOption({
    tooltip: {
      position: 'top',
      ...chartTooltipTheme(13),
      formatter: (p) => {
        const row = rows[p.data[1]];
        const metric = metrics[p.data[0]];
        return `${row.task}<br/>${metric.label}: ${formatValue(p.data[3], p.data[4])}`;
      }
    },
    grid: {
      left: 16,
      right: 16,
      top: 32,
      bottom: 8,
      containLabel: true
    },
    xAxis: {
      type: 'category',
      data: metrics.map(m => m.label),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: {
        color: chartAxisLabelCat(),
        fontSize: 12,
        rotate: 0,
        interval: 0,
        hideOverlap: false,
        margin: 12
      },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'category',
      data: rows.map(r => r.task),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: {
        color: chartAxisLabelCat(),
        fontSize: 12,
        lineHeight: 16
      },
      axisTick: { show: false }
    },
    visualMap: {
      min: 0,
      max: 1,
      dimension: 2,
      orient: 'horizontal',
      left: 'center',
      top: 0,
      calculable: false,
      inRange: {
        color: ['rgba(99,132,255,0.12)', 'rgba(99,132,255,0.95)']
      },
      textStyle: { color: chartAxisLabelCat(), fontSize: 12 }
    },
    series: [{
      type: 'heatmap',
      data: normalizedData,
      label: {
        show: true,
        color: chartHeatmapCellTextColor(),
        fontSize: 12,
        fontWeight: 600,
        formatter: (p) => formatValue(p.data[3], p.data[4])
      },
      itemStyle: {
        borderColor: 'rgba(255,255,255,0.08)',
        borderWidth: 1,
        borderRadius: 4
      },
      emphasis: {
        itemStyle: {
          borderColor: COLORS.cyan,
          borderWidth: 1.5
        }
      }
    }],
    animationDuration: 900
  });
}

function buildTaskKPIRows() {
  const data = getFilteredRawData();
  const dayKeys = new Set();
  const taskMap = {};

  data.forEach((r) => {
    const task = r['Task Type'] || 'Unknown';
    if (!taskMap[task]) {
      taskMap[task] = {
        total: 0,
        after5: 0,
        delayedBefore5: 0,
        after7: 0,
        delayedBefore7: 0,
        closedWithinSla: 0,
        closedSameDay: 0,
        closedWithin24h: 0,
        considered: 0
      };
    }

    const m = taskMap[task];
    m.total += 1;

    const completed = excelDateToJS(r['Completed date'] ?? r['Task Completed']);
    if (completed instanceof Date && !Number.isNaN(completed.getTime())) {
      dayKeys.add(completed.toISOString().slice(0, 10));
    }

    const assignedHour = Number(r['Task Assigned hour']);
    if (!Number.isNaN(assignedHour) && assignedHour >= 17) m.after5 += 1;
    if (!Number.isNaN(assignedHour) && assignedHour >= 19) m.after7 += 1;

    const exceptions = (r['Exceptions'] || '').toString().toLowerCase();
    const isDelayed = exceptions.includes('delay');
    const isConsidered = exceptions.includes('consider');
    if (isDelayed && !Number.isNaN(assignedHour) && assignedHour < 17) m.delayedBefore5 += 1;
    if (isDelayed && !Number.isNaN(assignedHour) && assignedHour < 19) m.delayedBefore7 += 1;
    if (isConsidered) m.considered += 1;

    const durationDays = Number(r['Duration']);
    const durationHours = Number.isNaN(durationDays) ? null : durationDays * 24;
    const slaHours = task.toLowerCase().includes('incident') ? 4 : 8;
    if (durationHours != null && durationHours <= slaHours) m.closedWithinSla += 1;
    if (r['Same day Closure'] === true) m.closedSameDay += 1;
    if (durationHours != null && durationHours <= 24) m.closedWithin24h += 1;
  });

  const totalDays = Math.max(dayKeys.size, 1);
  const rows = Object.entries(taskMap)
    .map(([task, m]) => {
      const total = m.total || 1;
      return {
        task,
        totalTickets: m.total,
        assignedAfter5: m.after5 / total,
        delayedBefore5: m.delayedBefore5 / total,
        assignedAfter7: m.after7 / total,
        delayedBefore7: m.delayedBefore7 / total,
        closedWithinSla: m.closedWithinSla / total,
        closedSameDay: m.closedSameDay / total,
        closedWithin24h: m.closedWithin24h / total,
        kpiAfterExclusion: m.considered / total
      };
    })
    .sort((a, b) => b.totalTickets - a.totalTickets);

  const metrics = [
    { key: 'totalTickets', label: 'Total Tickets', kind: 'number' },
    { key: 'assignedAfter5', label: 'Assigned > 5 PM', kind: 'pct' },
    { key: 'delayedBefore5', label: 'Delayed (< 5 PM)', kind: 'pct' },
    { key: 'assignedAfter7', label: 'Assigned > 7 PM', kind: 'pct' },
    { key: 'delayedBefore7', label: 'Delayed (< 7 PM)', kind: 'pct' },
    { key: 'closedWithinSla', label: 'Closed in 4/8h', kind: 'pct' },
    { key: 'closedSameDay', label: 'Same Day', kind: 'pct' },
    { key: 'closedWithin24h', label: 'Within 24h', kind: 'pct' },
    { key: 'kpiAfterExclusion', label: 'KPI After Exclusion', kind: 'pct' }
  ];

  return { rows, metrics };
}

export function renderTaskTypeResolutionTargets(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Task Type Resolution Performance (incl. Incident)', 'green');

  const stats = getTaskTypeResolutionStats();
  const labels = stats.map(s => s.name);

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      ...chartTooltipTheme(12)
    },
    legend: {
      data: ['Total', 'Resolved <= 4h', 'Resolved <= 24h', 'Same Day Closed', 'Target Attainment %'],
      textStyle: { color: chartAxisLabelCat(), fontSize: 11 },
      top: 2,
      itemGap: 8
    },
    grid: { left: 60, right: 50, top: 42, bottom: 44 },
    xAxis: {
      type: 'category',
      data: labels,
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: {
        color: chartAxisLabelCat(),
        fontSize: 12,
        rotate: 0,
        interval: 0,
        lineHeight: 15,
        formatter: (value) => formatTaskTypeAxisLabelTwoLines(value)
      },
      axisTick: { show: false }
    },
    yAxis: [
      {
        type: 'value',
        name: 'Count',
        nameTextStyle: { color: chartAxisLabelValue(), fontSize: 12 },
        splitLine: { lineStyle: { color: chartSplitLine() } },
        axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
      },
      {
        type: 'value',
        name: 'Target %',
        min: 0,
        max: 100,
        splitLine: { show: false },
        axisLabel: { color: chartAxisLabelValue(), fontSize: 12, formatter: '{value}%' }
      }
    ],
    series: [
      {
        name: 'Total',
        type: 'bar',
        barMaxWidth: 18,
        data: stats.map(s => s.total),
        itemStyle: { color: COLORS.blue, borderRadius: [4, 4, 0, 0] }
      },
      {
        name: 'Resolved <= 4h',
        type: 'bar',
        barMaxWidth: 18,
        data: stats.map(s => s.within4h),
        itemStyle: { color: COLORS.amber, borderRadius: [4, 4, 0, 0] }
      },
      {
        name: 'Resolved <= 24h',
        type: 'bar',
        barMaxWidth: 18,
        data: stats.map(s => s.within24h),
        itemStyle: { color: COLORS.cyan, borderRadius: [4, 4, 0, 0] }
      },
      {
        name: 'Same Day Closed',
        type: 'bar',
        barMaxWidth: 18,
        data: stats.map(s => s.sameDay),
        itemStyle: { color: COLORS.purple, borderRadius: [4, 4, 0, 0] }
      },
      {
        name: 'Target Attainment %',
        type: 'line',
        yAxisIndex: 1,
        smooth: true,
        symbol: 'circle',
        symbolSize: 6,
        data: stats.map(s => s.targetAttainmentPct),
        lineStyle: { color: COLORS.green, width: 2 },
        itemStyle: { color: COLORS.green }
      }
    ],
    animationDuration: 1100
  });
}

function getTaskTypeResolutionStats() {
  const tasks = getFilteredRawData();
  const incidents = getFilteredIncidentData();
  const byType = {};

  const getOrInit = (name) => {
    if (!byType[name]) {
      byType[name] = { name, total: 0, within4h: 0, within24h: 0, sameDay: 0, targetAttainmentPct: 0 };
    }
    return byType[name];
  };

  tasks.forEach((r) => {
    const name = r['Task Type'] || 'Unknown';
    const rec = getOrInit(name);
    rec.total += 1;
    const duration = parseFloat(r['Duration']);
    if (!Number.isNaN(duration)) {
      const hours = duration * 24;
      if (hours <= 4) rec.within4h += 1;
      if (hours <= 24) rec.within24h += 1;
    }
    if (r['Same day Closure'] === true) rec.sameDay += 1;
  });

  incidents.forEach((r) => {
    const rec = getOrInit('Incident');
    rec.total += 1;
    const duration = parseFloat(r['Duration']);
    if (!Number.isNaN(duration)) {
      const hours = duration * 24;
      if (hours <= 4) rec.within4h += 1;
      if (hours <= 24) rec.within24h += 1;
    }
    if (r['Same day Closure'] === true) rec.sameDay += 1;
  });

  const stats = Object.values(byType)
    .map((s) => {
      const total = s.total || 1;
      const use4hTarget = s.name.toLowerCase().includes('fiber support') || s.name.toLowerCase() === 'incident';
      const targetCount = use4hTarget ? s.within4h : s.within24h;
      return {
        ...s,
        targetAttainmentPct: Math.round((targetCount / total) * 100)
      };
    })
    .sort((a, b) => b.total - a.total);

  return stats;
}

export function renderRepeatedSupportWindow(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Repeated Support Volume by Time Window', 'amber');

  const data = getFilteredRawData();
  const buckets = [
    { label: 'Within 1 day', key: 'Repeated Fiber Support (yes/No), within 1 days' },
    { label: 'Within 3 days', key: 'Repeated Fiber Support (yes/No), within 3 days' },
    { label: 'Within 7 days', key: 'Repeated Fiber Support (yes/No), within 7 days' },
    { label: 'Within 30 days', key: 'Repeated Fiber Support (yes/No), within 30 days' }
  ];
  const counts = buckets.map((b) => data.reduce((sum, row) => sum + (isPositiveFlag(row[b.key]) ? 1 : 0), 0));

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      ...chartTooltipTheme(12)
    },
    grid: { left: 60, right: 20, top: 18, bottom: 40 },
    xAxis: {
      type: 'category',
      data: buckets.map(b => b.label),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    series: [{
      name: 'Repeated Support Count',
      type: 'bar',
      data: counts,
      barMaxWidth: 48,
      itemStyle: {
        borderRadius: [6, 6, 0, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: COLORS.amber },
          { offset: 1, color: COLORS.orange }
        ])
      },
      label: {
        show: true,
        position: 'top',
        color: chartAxisLabelCat(),
        fontSize: 12
      }
    }],
    animationDuration: 900
  });
}

export function renderImmediateSupportWindow(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Immediate Support Volume by Time Window', 'cyan');

  const data = getFilteredRawData();
  const buckets = [
    { label: 'Within 1 day', key: 'Immediate Fiber support (yes/No), within 1 days' },
    { label: 'Within 3 days', key: 'Immediate Fiber support (yes/No), within 3 days' },
    { label: 'Within 7 days', key: 'Immediate fiber support (yes/No), within 7 days' },
    { label: 'Within 30 days', key: 'Immediate fiber support (yes/No), within 30 days' }
  ];
  const counts = buckets.map((b) => data.reduce((sum, row) => sum + (isPositiveFlag(row[b.key]) ? 1 : 0), 0));

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      ...chartTooltipTheme(12)
    },
    grid: { left: 60, right: 20, top: 18, bottom: 40 },
    xAxis: {
      type: 'category',
      data: buckets.map(b => b.label),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    series: [{
      name: 'Immediate Support Count',
      type: 'bar',
      data: counts,
      barMaxWidth: 48,
      itemStyle: {
        borderRadius: [6, 6, 0, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: COLORS.cyan },
          { offset: 1, color: COLORS.blue }
        ])
      },
      label: {
        show: true,
        position: 'top',
        color: chartAxisLabelCat(),
        fontSize: 12
      }
    }],
    animationDuration: 900
  });
}

export function renderRepeatedSupportByBranch(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Repeated Support by Branch (Top 12)', 'amber');

  const { branches, seriesByWindow } = buildSupportByBranch('repeated');
  renderSupportBranchChart(chart, branches, seriesByWindow);
}

export function renderImmediateSupportByBranch(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Immediate Support by Branch (Top 12)', 'cyan');

  const { branches, seriesByWindow } = buildSupportByBranch('immediate');
  renderSupportBranchChart(chart, branches, seriesByWindow);
}

function renderSupportBranchChart(chart, branches, seriesByWindow) {
  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      ...chartTooltipTheme(12)
    },
    legend: {
      data: ['Within 1 day', 'Within 3 days', 'Within 7 days', 'Within 30 days'],
      textStyle: { color: chartAxisLabelCat(), fontSize: 12 },
      top: 0
    },
    grid: { left: 95, right: 20, top: 34, bottom: 24 },
    xAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    yAxis: {
      type: 'category',
      data: branches.reverse(),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12 },
      axisTick: { show: false }
    },
    series: [
      { name: 'Within 1 day', type: 'bar', stack: 'total', data: seriesByWindow[0].slice().reverse(), itemStyle: { color: COLORS.green } },
      { name: 'Within 3 days', type: 'bar', stack: 'total', data: seriesByWindow[1].slice().reverse(), itemStyle: { color: COLORS.cyan } },
      { name: 'Within 7 days', type: 'bar', stack: 'total', data: seriesByWindow[2].slice().reverse(), itemStyle: { color: COLORS.blue } },
      { name: 'Within 30 days', type: 'bar', stack: 'total', data: seriesByWindow[3].slice().reverse(), itemStyle: { color: COLORS.amber } }
    ],
    animationDuration: 900
  });
}

function buildSupportByBranch(kind) {
  const data = getFilteredRawData();
  const keys = kind === 'repeated'
    ? [
      'Repeated Fiber Support (yes/No), within 1 days',
      'Repeated Fiber Support (yes/No), within 3 days',
      'Repeated Fiber Support (yes/No), within 7 days',
      'Repeated Fiber Support (yes/No), within 30 days'
    ]
    : [
      'Immediate Fiber support (yes/No), within 1 days',
      'Immediate Fiber support (yes/No), within 3 days',
      'Immediate fiber support (yes/No), within 7 days',
      'Immediate fiber support (yes/No), within 30 days'
    ];

  const branchMap = {};
  data.forEach((row) => {
    const branch = row['Branch'] || 'N/A';
    if (!branchMap[branch]) {
      branchMap[branch] = [0, 0, 0, 0];
    }
    keys.forEach((key, idx) => {
      if (isPositiveFlag(row[key])) branchMap[branch][idx] += 1;
    });
  });

  const sorted = Object.entries(branchMap)
    .map(([branch, values]) => ({ branch, values, total: values.reduce((a, b) => a + b, 0) }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 12);

  return {
    branches: sorted.map(s => s.branch),
    seriesByWindow: [0, 1, 2, 3].map(i => sorted.map(s => s.values[i]))
  };
}

export function renderTeamTopScores(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Top Team Scores (FS & New)', 'blue');

  const { fsNew } = getTeamPerformanceData();
  const top = (fsNew || [])
    .slice()
    .sort((a, b) => b.score - a.score)
    .slice(0, 12);

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      ...chartTooltipTheme(12),
      formatter: (p) => {
        const d = top[p[0].dataIndex];
        return `${d.agent}<br/>Branch: ${d.branch}<br/>Score: ${(d.score * 100).toFixed(1)}%`;
      }
    },
    grid: { left: 180, right: 20, top: 16, bottom: 24 },
    xAxis: {
      type: 'value',
      min: 0,
      max: 1,
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), formatter: (v) => `${Math.round(v * 100)}%` }
    },
    yAxis: {
      type: 'category',
      data: top.map(r => r.agent.split('[')[0].trim()).reverse(),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12, width: 260, overflow: 'truncate' },
      axisTick: { show: false }
    },
    series: [{
      name: 'Obtained Score',
      type: 'bar',
      data: top.map(r => r.score).reverse(),
      barMaxWidth: 16,
      itemStyle: {
        borderRadius: [0, 6, 6, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
          { offset: 0, color: COLORS.blue },
          { offset: 1, color: COLORS.purple }
        ])
      }
    }],
    animationDuration: 900
  });
}

export function renderTeamPerformanceByBranch(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Branch Team Performance (Avg Score)', 'green');

  const { fsNew } = getTeamPerformanceData();
  const byBranch = {};
  (fsNew || []).forEach((r) => {
    if (!byBranch[r.branch]) byBranch[r.branch] = { total: 0, count: 0 };
    byBranch[r.branch].total += r.score;
    byBranch[r.branch].count += 1;
  });
  const rows = Object.entries(byBranch)
    .map(([branch, v]) => ({ branch, avg: v.count ? v.total / v.count : 0, teams: v.count }))
    .sort((a, b) => b.avg - a.avg)
    .slice(0, 12);

  chart.setOption({
    tooltip: { trigger: 'axis', ...chartTooltipTheme(12) },
    grid: { left: 50, right: 20, top: 20, bottom: 45 },
    xAxis: {
      type: 'category',
      data: rows.map(r => r.branch),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12, rotate: 30 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      min: 0,
      max: 1,
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), formatter: (v) => `${Math.round(v * 100)}%` }
    },
    series: [{
      name: 'Avg Score',
      type: 'bar',
      data: rows.map(r => r.avg),
      barMaxWidth: 24,
      itemStyle: { color: COLORS.green, borderRadius: [4, 4, 0, 0] }
    }],
    animationDuration: 900
  });
}

export function renderTeamOpsHealth(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Operations Health (RS & ARE)', 'amber');

  const { rsAre } = getTeamPerformanceData();
  const items = (rsAre || []).slice(0, 20);
  const avgWorkingDays = avg(items.map(r => r.workingDaysPct));
  const avgTaskHandledPct = avg(items.map(r => r.taskHandledPct));
  const avgTaskHandled = avg(items.map(r => r.avgTasksHandled));

  chart.setOption({
    tooltip: { trigger: 'item', ...chartTooltipTheme(12) },
    grid: { left: 40, right: 20, top: 20, bottom: 25 },
    xAxis: { type: 'category', data: ['Working Days %', 'Task Handled %', 'Avg Tasks/Day'], axisLine: { lineStyle: { color: chartAxisLine() } }, axisLabel: { color: chartAxisLabelCat() }, axisTick: { show: false } },
    yAxis: { type: 'value', splitLine: { lineStyle: { color: chartSplitLine() } }, axisLabel: { color: chartAxisLabelValue() } },
    series: [{
      type: 'bar',
      data: [avgWorkingDays * 100, avgTaskHandledPct * 100, avgTaskHandled],
      barMaxWidth: 46,
      itemStyle: {
        color: (p) => [COLORS.amber, COLORS.orange, COLORS.blue][p.dataIndex],
        borderRadius: [6, 6, 0, 0]
      },
      label: { show: true, position: 'top', color: chartAxisLabelCat(), formatter: (p) => (p.dataIndex < 2 ? `${p.value.toFixed(1)}%` : `${p.value.toFixed(1)}`) }
    }],
    animationDuration: 900
  });
}

function avg(values) {
  const clean = values.filter(v => typeof v === 'number' && !Number.isNaN(v));
  if (!clean.length) return 0;
  return clean.reduce((a, b) => a + b, 0) / clean.length;
}

function isPositiveFlag(value) {
  if (value === true || value === 1) return true;
  const text = (value ?? '').toString().trim().toLowerCase();
  return text === 'yes' || text === 'y' || text === 'true' || text === '1';
}

export function renderRegionBar(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Tasks by Region', 'green');

  const data = getFilteredRawData();
  const regionMap = {};
  data.forEach(r => {
    const reg = r['Region'] || 'N/A';
    regionMap[reg] = (regionMap[reg] || 0) + 1;
  });

  const sorted = Object.entries(regionMap).sort((a, b) => b[1] - a[1]);
  const regions = sorted.map(s => s[0]);
  const counts = sorted.map(s => s[1]);

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      ...chartTooltipTheme(12)
    },
    grid: { left: 100, right: 30, top: 10, bottom: 20 },
    xAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    yAxis: {
      type: 'category',
      data: regions.reverse(),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12 },
      axisTick: { show: false }
    },
    series: [{
      type: 'bar',
      data: counts.reverse(),
      barWidth: 18,
      itemStyle: {
        borderRadius: [0, 6, 6, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
          { offset: 0, color: COLORS.green },
          { offset: 1, color: COLORS.cyan }
        ])
      }
    }],
    animationDuration: 1000
  });
}

export function renderPriorityDist(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Task Priority Distribution', 'amber');

  const data = getFilteredRawData();
  const prioMap = {};
  data.forEach(r => {
    const p = `P${r['Task Priority']}`;
    prioMap[p] = (prioMap[p] || 0) + 1;
  });

  const sorted = Object.entries(prioMap).sort((a, b) => a[0].localeCompare(b[0]));
  const labels = sorted.map(s => s[0]);
  const values = sorted.map(s => s[1]);
  const colors = [COLORS.red, COLORS.amber, COLORS.blue, COLORS.green, COLORS.purple];

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      ...chartTooltipTheme(12)
    },
    grid: { left: 50, right: 20, top: 10, bottom: 30 },
    xAxis: {
      type: 'category',
      data: labels,
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    series: [{
      type: 'bar',
      data: values.map((v, i) => ({
        value: v,
        itemStyle: {
          color: colors[i % colors.length],
          borderRadius: [6, 6, 0, 0]
        }
      })),
      barWidth: 36
    }],
    animationDuration: 1000
  });
}

export function renderSLAGauge(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Same-Day Closure Rate', 'cyan');

  const data = getFilteredRawData();
  const total = data.length;
  const sameDay = data.filter(r => r['Same day Closure'] === true).length;
  const rate = total > 0 ? Math.round((sameDay / total) * 100) : 0;

  chart.setOption({
    series: [{
      type: 'gauge',
      startAngle: 200,
      endAngle: -20,
      min: 0,
      max: 100,
      center: ['50%', '55%'],
      radius: '85%',
      progress: {
        show: true,
        width: 18,
        itemStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
            { offset: 0, color: COLORS.cyan },
            { offset: 1, color: COLORS.green }
          ])
        }
      },
      axisLine: {
        lineStyle: { width: 18, color: [[1, 'rgba(255,255,255,0.06)']] }
      },
      axisTick: { show: false },
      splitLine: { show: false },
      axisLabel: { show: false },
      anchor: { show: false },
      pointer: { show: false },
      title: {
        offsetCenter: [0, '20%'],
        fontSize: 14,
        color: chartAxisLabelCat(),
        fontWeight: 500
      },
      detail: {
        valueAnimation: true,
        fontSize: 36,
        fontWeight: 800,
        offsetCenter: [0, '-10%'],
        formatter: '{value}%',
        color: COLORS.cyan
      },
      data: [{ value: rate, name: 'Same Day' }]
    }],
    animationDuration: 1500,
    animationEasing: 'cubicOut'
  });
}

// ==========================================
// INCIDENT CHARTS
// ==========================================

export function renderIncidentTrend(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Incidents Over Time', 'red');

  const data = getFilteredIncidentData();
  const dateMap = {};
  data.forEach(row => {
    const d = excelDateToJS(row['Completed date']);
    if (!d) return;
    const key = d.toISOString().slice(0, 10);
    dateMap[key] = (dateMap[key] || 0) + 1;
  });

  const sorted = Object.entries(dateMap).sort((a, b) => a[0].localeCompare(b[0]));
  const incidentCategories = sorted.map(s => s[0]);

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      ...chartTooltipTheme(13),
      formatter: (params) => buildAxisTooltipHtmlWithDates(params)
    },
    grid: { left: 52, right: 20, top: 24, bottom: 56 },
    xAxis: {
      type: 'category',
      data: incidentCategories,
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12, rotate: 30,
        formatter: (v, idx) => formatCategoryAxisDateLabel(v, idx, incidentCategories)
      },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    series: [{
      type: 'bar',
      data: sorted.map(s => s[1]),
      itemStyle: {
        borderRadius: [4, 4, 0, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: COLORS.red },
          { offset: 1, color: COLORS.pink }
        ])
      },
      barMaxWidth: 20
    }],
    animationDuration: 1000
  });
}

export function renderIncidentCategory(containerId) {
  const el = document.getElementById(containerId);
  if (!el) return;
  if (incidentCategoryShowAll[containerId] === undefined) incidentCategoryShowAll[containerId] = false;
  const showAll = incidentCategoryShowAll[containerId];

  // Drop inline sizes from switching "All categories" → "Top 10" (min-height alone may not shrink).
  el.style.removeProperty('min-height');
  el.style.removeProperty('height');

  el.innerHTML = '';
  const header = document.createElement('div');
  header.className = 'chart-card-header';
  header.innerHTML = `
    <div class="chart-title"><span class="dot amber"></span>Incident Categories</div>
    <div class="chart-card-toolbar">
      <button type="button" class="chart-toggle-btn ${!showAll ? 'active' : ''}" data-mode="top">Top 10</button>
      <button type="button" class="chart-toggle-btn ${showAll ? 'active' : ''}" data-mode="all">All categories</button>
    </div>
  `;
  const bodyEl = document.createElement('div');
  bodyEl.className = 'chart-body';
  el.appendChild(header);
  el.appendChild(bodyEl);

  header.querySelectorAll('.chart-toggle-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      incidentCategoryShowAll[containerId] = btn.dataset.mode === 'all';
      renderIncidentCategory(containerId);
    });
  });

  const data = getFilteredIncidentData();
  const catMap = {};
  data.forEach((r) => {
    const c = r['Category'] || 'Unknown';
    catMap[c] = (catMap[c] || 0) + 1;
  });

  const sorted = Object.entries(catMap).sort((a, b) => b[1] - a[1]);
  const display = showAll ? sorted : sorted.slice(0, 10);
  const names = display.map((s) => s[0]);
  const counts = display.map((s) => s[1]);
  const n = Math.max(names.length, 1);
  const barH = Math.min(26, Math.max(14, 340 / n));
  const bodyMinH = Math.max(280, n * barH + 72);
  // Fixed height overrides global .chart-body { height: calc(100% - 40px) }, which can block shrinking.
  bodyEl.style.minHeight = `${bodyMinH}px`;
  bodyEl.style.height = `${bodyMinH}px`;
  // Header row + card vertical padding (~20px × 2) — keep card from staying tall after "All".
  el.style.minHeight = `${bodyMinH + 100}px`;

  if (chartInstances[containerId]) {
    chartInstances[containerId].dispose();
    delete chartInstances[containerId];
  }
  const chart = echarts.init(bodyEl, null, { renderer: 'canvas' });
  chartInstances[containerId] = chart;

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      ...chartTooltipTheme(12)
    },
    grid: { left: 180, right: 30, top: 10, bottom: 20 },
    xAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    yAxis: {
      type: 'category',
      data: [...names].reverse(),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12, width: 160, overflow: 'truncate' },
      axisTick: { show: false }
    },
    series: [{
      type: 'bar',
      data: [...counts].reverse(),
      barMaxWidth: 22,
      itemStyle: {
        borderRadius: [0, 6, 6, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
          { offset: 0, color: COLORS.amber },
          { offset: 1, color: COLORS.orange }
        ])
      }
    }],
    animationDuration: 1000
  });
  chart.resize();
  requestAnimationFrame(() => {
    chart.resize();
  });
}

export function renderIncidentComplexity(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Incident Complexity', 'purple');

  const data = getFilteredIncidentData();
  const cMap = {};
  data.forEach(r => {
    const c = r['Complexity'] || 'Unknown';
    cMap[c] = (cMap[c] || 0) + 1;
  });

  const pieData = Object.entries(cMap).map(([name, value]) => ({ name, value }));

  chart.setOption({
    tooltip: {
      trigger: 'item',
      ...chartTooltipTheme(12),
      formatter: '{b}: {c} ({d}%)'
    },
    color: [COLORS.green, COLORS.amber, COLORS.red, COLORS.purple, COLORS.blue],
    series: [{
      type: 'pie',
      radius: ['40%', '70%'],
      center: ['50%', '50%'],
      itemStyle: { borderRadius: 6, borderColor: 'rgba(10,14,26,0.8)', borderWidth: 2 },
      label: {
        color: chartAxisLabelCat(),
        fontSize: 11,
        formatter: '{b}\n{d}%'
      },
      data: pieData
    }],
    animationDuration: 1000
  });
}

export function renderIncidentClosure(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Same-Day Closure vs Delayed', 'green');

  const data = getFilteredIncidentData();
  const sameDay = data.filter(r => r['Same day Closure'] === true).length;
  const delayed = data.filter(r => {
    const exc = (r['Exceptions'] || '').toLowerCase();
    return exc.includes('delayed');
  }).length;
  const other = data.length - sameDay - delayed;

  chart.setOption({
    tooltip: {
      trigger: 'item',
      ...chartTooltipTheme(12),
      formatter: '{b}: {c} ({d}%)'
    },
    color: [COLORS.green, COLORS.red, COLORS.blue],
    series: [{
      type: 'pie',
      radius: ['40%', '70%'],
      center: ['50%', '50%'],
      itemStyle: { borderRadius: 6, borderColor: 'rgba(10,14,26,0.8)', borderWidth: 2 },
      label: {
        color: chartAxisLabelCat(),
        fontSize: 11,
        formatter: '{b}\n{c}'
      },
      data: [
        { name: 'Same Day', value: sameDay },
        { name: 'Delayed', value: delayed },
        { name: 'Other', value: other > 0 ? other : 0 }
      ]
    }],
    animationDuration: 1000
  });
}

// ==========================================
// BRANCH KPI CHARTS
// ==========================================

export function renderBranchPerformance(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Branch Performance — Task & Closure Counts', 'blue');

  const sorted = getBranchClosureStats();

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      ...chartTooltipTheme(12),
      axisPointer: { type: 'shadow' }
    },
    legend: {
      data: ['Tasks', 'Same Day (count)', 'Closed <= 4h', 'Closed <= 24h'],
      textStyle: { color: chartAxisLabelCat(), fontSize: 11 },
      top: 0
    },
    grid: { left: 70, right: 30, top: 35, bottom: 40 },
    xAxis: {
      type: 'category',
      data: sorted.map(s => s.name),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12, rotate: 40 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      name: 'Tasks',
      nameTextStyle: { color: chartAxisLabelValue(), fontSize: 12 },
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    series: [
      {
        name: 'Tasks',
        type: 'bar',
        data: sorted.map(s => s.total),
        barWidth: 18,
        itemStyle: {
          borderRadius: [4, 4, 0, 0],
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: COLORS.blue },
            { offset: 1, color: 'rgba(99,132,255,0.3)' }
          ])
        }
      },
      {
        name: 'Same Day (count)',
        type: 'line',
        data: sorted.map(s => s.sameDay),
        smooth: true,
        symbol: 'circle',
        symbolSize: 6,
        lineStyle: { color: COLORS.purple, width: 2 },
        itemStyle: { color: COLORS.purple }
      },
      {
        name: 'Closed <= 4h',
        type: 'line',
        data: sorted.map(s => s.within4h),
        smooth: true,
        symbol: 'circle',
        symbolSize: 5,
        lineStyle: { color: COLORS.amber, width: 2 },
        itemStyle: { color: COLORS.amber }
      },
      {
        name: 'Closed <= 24h',
        type: 'line',
        data: sorted.map(s => s.within24h),
        smooth: true,
        symbol: 'circle',
        symbolSize: 5,
        lineStyle: { color: COLORS.red, width: 2 },
        itemStyle: { color: COLORS.red }
      }
    ],
    animationDuration: 1200
  });
}

export function renderBranchClosureRates(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Branch Closure Rates (%) — Same Day, 4h, 24h', 'green');

  const sorted = getBranchClosureStats();

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      ...chartTooltipTheme(12)
    },
    legend: {
      data: ['Same Day %', 'Closed <= 4h %', 'Closed <= 24h %'],
      textStyle: { color: chartAxisLabelCat(), fontSize: 11 },
      top: 0
    },
    grid: { left: 70, right: 30, top: 35, bottom: 40 },
    xAxis: {
      type: 'category',
      data: sorted.map(s => s.name),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12, rotate: 40 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      min: 0,
      max: 100,
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12, formatter: '{value}%' }
    },
    series: [
      {
        name: 'Same Day %',
        type: 'line',
        data: sorted.map(s => s.rate),
        smooth: true,
        symbol: 'circle',
        symbolSize: 6,
        lineStyle: { color: COLORS.purple, width: 2 },
        itemStyle: { color: COLORS.purple }
      },
      {
        name: 'Closed <= 4h %',
        type: 'line',
        data: sorted.map(s => s.within4hRate),
        smooth: true,
        symbol: 'circle',
        symbolSize: 5,
        lineStyle: { color: COLORS.amber, width: 2 },
        itemStyle: { color: COLORS.amber }
      },
      {
        name: 'Closed <= 24h %',
        type: 'line',
        data: sorted.map(s => s.within24hRate),
        smooth: true,
        symbol: 'circle',
        symbolSize: 5,
        lineStyle: { color: COLORS.red, width: 2 },
        itemStyle: { color: COLORS.red }
      }
    ],
    animationDuration: 1200
  });
}

export function renderBranchEfficiencyFromSheet(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Branch Efficiency (from Branch Effi.)', 'green');

  const data = getBranchEfficiencyData()
    .filter(r => r.efficiencyPerWorkingDay != null)
    .sort((a, b) => (b.efficiencyPerWorkingDay - a.efficiencyPerWorkingDay))
    .slice(0, 15);
  const maxEfficiencyValue = Math.max(
    1,
    ...data.map(d => Number(d.efficiency) || 0),
    ...data.map(d => Number(d.efficiencyPerWorkingDay) || 0)
  );
  const yAxisMax = Math.min(1.5, Math.max(1.05, maxEfficiencyValue * 1.08));

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      ...chartTooltipTheme(12),
      formatter: (p) => {
        const first = p[0];
        const second = p[1];
        return `${first.axisValue}<br/>Efficiency: ${(first.value * 100).toFixed(1)}%<br/>Efficiency (working day): ${(second.value * 100).toFixed(1)}%`;
      }
    },
    legend: {
      data: ['Efficiency', 'Efficiency (Working Day)'],
      textStyle: { color: chartAxisLabelCat(), fontSize: 12 },
      top: 0
    },
    grid: { left: 70, right: 30, top: 52, bottom: 64 },
    xAxis: {
      type: 'category',
      data: data.map(d => d.branch),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12, rotate: 28 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      min: 0,
      max: yAxisMax,
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12, formatter: (v) => `${Math.round(v * 100)}%` }
    },
    series: [
      {
        name: 'Efficiency',
        type: 'bar',
        data: data.map(d => d.efficiency),
        barMaxWidth: 24,
        itemStyle: { color: COLORS.green, borderRadius: [4, 4, 0, 0] }
      },
      {
        name: 'Efficiency (Working Day)',
        type: 'line',
        data: data.map(d => d.efficiencyPerWorkingDay),
        smooth: true,
        symbol: 'circle',
        symbolSize: 5,
        lineStyle: { color: COLORS.cyan, width: 2 },
        itemStyle: { color: COLORS.cyan }
      }
    ],
    animationDuration: 900
  });
}

export function renderBranchWorkloadFromSheet(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Branch Workload (from Branch Effi.)', 'amber');

  const base = getBranchEfficiencyData()
    .filter(r => r.workload != null)
    .sort((a, b) => (b.workload - a.workload));
  const grandTotal = base.reduce((sum, r) => sum + (Number(r.workload) || 0), 0);
  const data = [
    { branch: 'Grand Total', workload: grandTotal },
    ...base.slice(0, 14)
  ];

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      ...chartTooltipTheme(12)
    },
    grid: { left: 100, right: 20, top: 16, bottom: 24 },
    xAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    yAxis: {
      type: 'category',
      data: data.map(d => d.branch).reverse(),
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12 },
      axisTick: { show: false }
    },
    series: [{
      name: 'Total Workload',
      type: 'bar',
      data: data.map(d => d.workload).reverse(),
      barMaxWidth: 18,
      itemStyle: {
        borderRadius: [0, 6, 6, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
          { offset: 0, color: COLORS.amber },
          { offset: 1, color: COLORS.orange }
        ])
      }
    }],
    animationDuration: 900
  });
}

function getBranchClosureStats() {
  const data = getFilteredRawData();
  const branchMap = {};
  data.forEach(r => {
    const b = r['Branch'] || 'N/A';
    if (!branchMap[b]) branchMap[b] = { total: 0, sameDay: 0, within4h: 0, within24h: 0 };
    branchMap[b].total++;
    if (r['Same day Closure'] === true) branchMap[b].sameDay++;
    const durationValue = parseFloat(r['Duration']);
    if (!Number.isNaN(durationValue)) {
      const durHours = durationValue * 24;
      if (durHours <= 4) branchMap[b].within4h++;
      if (durHours <= 24) branchMap[b].within24h++;
    }
  });

  return Object.entries(branchMap)
    .map(([name, v]) => ({
      name,
      total: v.total,
      sameDay: v.sameDay,
      within4h: v.within4h,
      within24h: v.within24h,
      rate: v.total > 0 ? Math.round((v.sameDay / v.total) * 100) : 0,
      within4hRate: v.total > 0 ? Math.round((v.within4h / v.total) * 100) : 0,
      within24hRate: v.total > 0 ? Math.round((v.within24h / v.total) * 100) : 0
    }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 20);
}

export function renderDivisionCompare(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Division Comparison', 'purple');

  const data = getFilteredRawData();
  const divMap = {};
  data.forEach(r => {
    const d = r['Division'] || 'N/A';
    if (!divMap[d]) divMap[d] = { total: 0, sameDay: 0 };
    divMap[d].total++;
    if (r['Same day Closure'] === true) divMap[d].sameDay++;
  });

  const entries = Object.entries(divMap);
  const indicator = entries.map(([name]) => ({ name, max: Math.max(...entries.map(e => e[1].total)) * 1.2 }));

  chart.setOption({
    tooltip: {
      ...chartTooltipTheme(12)
    },
    color: [COLORS.blue, COLORS.green],
    legend: {
      data: ['Total Tasks', 'Same Day Closed'],
      textStyle: { color: chartAxisLabelCat(), fontSize: 11 },
      bottom: 0
    },
    radar: {
      indicator,
      shape: 'polygon',
      axisName: { color: chartAxisLabelCat(), fontSize: 11 },
      splitArea: { areaStyle: { color: chartRadarSplitAreaColors() } },
      splitLine: { lineStyle: { color: chartRadarLineColor() } },
      axisLine: { lineStyle: { color: chartRadarLineColor() } }
    },
    series: [{
      type: 'radar',
      data: [
        { name: 'Total Tasks', value: entries.map(e => e[1].total), areaStyle: { color: 'rgba(99,132,255,0.15)' } },
        { name: 'Same Day Closed', value: entries.map(e => e[1].sameDay), areaStyle: { color: 'rgba(52,211,153,0.15)' } }
      ]
    }],
    animationDuration: 1000
  });
}

export function renderSameDayClosure(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Weekday Open/New/Same-Day/Open Next Day', 'amber');

  const data = getFilteredRawData();
  const assignedByDate = {};
  data.forEach((r) => {
    const assignedDate = excelDateToJS(r['Assigned date'] ?? r['Task Assigned'] ?? r['Task Created']);
    if (!assignedDate || Number.isNaN(assignedDate.getTime())) return;
    const key = assignedDate.toISOString().slice(0, 10);
    if (!assignedByDate[key]) assignedByDate[key] = { assigned: 0, sameDay: 0 };
    assignedByDate[key].assigned += 1;
    if (r['Same day Closure'] === true) assignedByDate[key].sameDay += 1;
  });

  const dates = Object.keys(assignedByDate).sort((a, b) => a.localeCompare(b));
  const dayOrder = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const weekdayTotals = dayOrder.map(() => ({ assigned: 0, sameDay: 0 }));

  dates.forEach((dateKey) => {
    const weekdayIdx = new Date(dateKey).getDay();
    weekdayTotals[weekdayIdx].assigned += assignedByDate[dateKey].assigned;
    weekdayTotals[weekdayIdx].sameDay += assignedByDate[dateKey].sameDay;
  });

  const weekdayFlow = dayOrder.map(() => ({
    opening: 0,
    assigned: 0,
    sameDay: 0,
    nextOpen: 0
  }));
  let carryOpen = 0; // Sunday Open(start) must be 0
  dayOrder.forEach((_, idx) => {
    const assigned = weekdayTotals[idx].assigned;
    const sameDayClosed = weekdayTotals[idx].sameDay;
    const nextOpen = Math.max(carryOpen + assigned - sameDayClosed, 0);
    weekdayFlow[idx].opening = carryOpen;
    weekdayFlow[idx].assigned = assigned;
    weekdayFlow[idx].sameDay = sameDayClosed;
    weekdayFlow[idx].nextOpen = nextOpen;
    carryOpen = nextOpen;
  });

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      ...chartTooltipTheme(12),
      axisPointer: { type: 'shadow' }
    },
    legend: {
      data: ['Open (start)', 'New', 'Same Day Closure', 'Open (next day)'],
      textStyle: { color: chartAxisLabelCat(), fontSize: 11 },
      top: 0
    },
    grid: { left: 50, right: 20, top: 35, bottom: 48 },
    xAxis: {
      type: 'category',
      data: dayOrder,
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: chartSplitLine() } },
      axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
    },
    series: [
      {
        name: 'Open (start)',
        type: 'bar',
        data: weekdayFlow.map(d => d.opening),
        barWidth: 14,
        itemStyle: { borderRadius: [4, 4, 0, 0], color: COLORS.purple }
      },
      {
        name: 'New',
        type: 'bar',
        data: weekdayFlow.map(d => d.assigned),
        barWidth: 14,
        itemStyle: { borderRadius: [4, 4, 0, 0], color: COLORS.blue }
      },
      {
        name: 'Same Day Closure',
        type: 'bar',
        data: weekdayFlow.map(d => d.sameDay),
        barWidth: 14,
        itemStyle: { borderRadius: [4, 4, 0, 0], color: COLORS.green }
      },
      {
        name: 'Open (next day)',
        type: 'bar',
        data: weekdayFlow.map(d => d.nextOpen),
        barWidth: 14,
        itemStyle: { borderRadius: [4, 4, 0, 0], color: COLORS.amber }
      }
    ],
    animationDuration: 1000
  });
}

export function renderBranchTasksVolume(containerId) {
  const el = document.getElementById(containerId);
  if (!el) return;
  const state = getBranchVolumeState(containerId);
  const { showAllBranches, consideredOnly } = state;

  el.style.removeProperty('min-height');
  el.style.removeProperty('height');

  el.innerHTML = '';
  const header = document.createElement('div');
  header.className = 'chart-card-header';
  header.innerHTML = `
    <div class="chart-title"><span class="dot cyan"></span>Branch Task Volume & Avg Duration</div>
    <div class="chart-card-toolbar chart-card-toolbar--branch-controls">
      <div class="chart-toolbar-category" role="group" aria-label="Branch list scope">
        <span class="chart-toolbar-label">Branches</span>
        <div class="chart-toolbar-category-buttons">
          <button type="button" class="chart-toggle-btn ${!showAllBranches ? 'active' : ''}" data-branches="top">Top 15</button>
          <button type="button" class="chart-toggle-btn ${showAllBranches ? 'active' : ''}" data-branches="all">All</button>
        </div>
      </div>
      <div class="chart-toolbar-category" role="group" aria-label="Task filter scope">
        <span class="chart-toolbar-label">Scope</span>
        <div class="chart-toolbar-category-buttons">
          <button type="button" class="chart-toggle-btn ${!consideredOnly ? 'active' : ''}" data-scope="all">All tasks</button>
          <button type="button" class="chart-toggle-btn ${consideredOnly ? 'active' : ''}" data-scope="considered">Considered</button>
        </div>
      </div>
    </div>
  `;
  const bodyEl = document.createElement('div');
  bodyEl.className = 'chart-body';
  el.appendChild(header);
  el.appendChild(bodyEl);

  header.querySelectorAll('[data-branches]').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.showAllBranches = btn.dataset.branches === 'all';
      renderBranchTasksVolume(containerId);
    });
  });
  header.querySelectorAll('[data-scope]').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.consideredOnly = btn.dataset.scope === 'considered';
      renderBranchTasksVolume(containerId);
    });
  });

  let rows = getFilteredRawData();
  if (consideredOnly) rows = rows.filter(isRowConsideredForKpi);

  const branchMap = {};
  rows.forEach((r) => {
    const b = r['Branch'] || 'N/A';
    if (!branchMap[b]) branchMap[b] = { count: 0, durSum: 0 };
    branchMap[b].count += 1;
    branchMap[b].durSum += parseFloat(r['Duration']) || 0;
  });

  let list = Object.entries(branchMap)
    .map(([name, v]) => ({
      name,
      total: v.count,
      avgDuration: v.count ? v.durSum / v.count : 0
    }))
    .sort((a, b) => b.total - a.total);

  if (!showAllBranches) list = list.slice(0, 15);

  const names = list.map((x) => x.name);
  const volumes = list.map((x) => x.total);
  const avgDays = list.map((x) => x.avgDuration);
  const maxAvg = Math.max(...avgDays, 0.01) * 1.15;
  const n = Math.max(names.length, 1);
  const rowH = Math.min(24, Math.max(16, 360 / n));
  const bodyMinH = list.length === 0
    ? 280
    : Math.max(320, n * rowH + 120);
  bodyEl.style.minHeight = `${bodyMinH}px`;
  bodyEl.style.height = `${bodyMinH}px`;
  el.style.minHeight = `${bodyMinH + 110}px`;

  if (chartInstances[containerId]) {
    chartInstances[containerId].dispose();
    delete chartInstances[containerId];
  }
  const chart = echarts.init(bodyEl, null, { renderer: 'canvas' });
  chartInstances[containerId] = chart;

  if (list.length === 0) {
    chart.setOption({
      title: {
        text: consideredOnly ? 'No Considered tasks in this period' : 'No task data for branches',
        left: 'center',
        top: 'center',
        textStyle: { color: chartAxisLabelCat(), fontSize: 14, fontWeight: 500 }
      },
      xAxis: { show: false },
      yAxis: { show: false },
      series: []
    });
    chart.resize();
    requestAnimationFrame(() => {
      chart.resize();
    });
    return;
  }

  const namesRev = [...names].reverse();
  const volRev = [...volumes].reverse();
  const avgRev = [...avgDays].reverse();

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      ...chartTooltipTheme(12),
      formatter: (params) => {
        if (!params?.length) return '';
        const branch = params[0].name;
        const row = list.find((l) => l.name === branch);
        if (!row) return '';
        const scope = consideredOnly ? ' (Considered)' : '';
        return `${branch}${scope}<br/>Tasks: ${formatNumber(row.total)}<br/>Avg duration: ${formatDuration(row.avgDuration)}`;
      }
    },
    legend: {
      data: ['Task volume', 'Avg duration'],
      textStyle: { color: chartAxisLabelCat(), fontSize: 11 },
      left: 'center',
      bottom: 4,
      itemGap: 20
    },
    /* Legend below plot so it does not cover the top “Avg days” x-axis */
    grid: { left: '14%', right: '10%', top: 32, bottom: 52 },
    xAxis: [
      {
        type: 'value',
        name: 'Tasks',
        position: 'bottom',
        nameTextStyle: { color: chartAxisLabelValue(), fontSize: 12 },
        splitLine: { lineStyle: { color: chartSplitLine() } },
        axisLabel: { color: chartAxisLabelValue(), fontSize: 12 }
      },
      {
        type: 'value',
        name: 'Avg days',
        position: 'top',
        nameTextStyle: { color: chartAxisLabelValue(), fontSize: 12 },
        max: maxAvg,
        splitLine: { show: false },
        axisLabel: {
          color: chartAxisLabelValue(),
          fontSize: 12,
          formatter: (v) => Number(v).toFixed(1)
        }
      }
    ],
    yAxis: {
      type: 'category',
      data: namesRev,
      axisLine: { lineStyle: { color: chartAxisLine() } },
      axisLabel: { color: chartAxisLabelCat(), fontSize: 12, width: 130, overflow: 'truncate' },
      axisTick: { show: false }
    },
    series: [
      {
        name: 'Task volume',
        type: 'bar',
        xAxisIndex: 0,
        data: volRev,
        barMaxWidth: 18,
        itemStyle: {
          borderRadius: [0, 6, 6, 0],
          color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
            { offset: 0, color: COLORS.cyan },
            { offset: 1, color: COLORS.blue }
          ])
        }
      },
      {
        name: 'Avg duration',
        type: 'line',
        xAxisIndex: 1,
        yAxisIndex: 0,
        data: avgRev,
        symbol: 'circle',
        symbolSize: 8,
        lineStyle: { width: 2.5, color: COLORS.amber },
        itemStyle: { color: COLORS.amber },
        emphasis: { focus: 'series' }
      }
    ],
    animationDuration: 900
  });
  chart.resize();
  requestAnimationFrame(() => {
    chart.resize();
  });
}
