/**
 * Charts Module — all ECharts visualizations
 */
import * as echarts from 'echarts';
import {
  getFilteredRawData,
  getFilteredIncidentData,
  excelDateToJS,
  formatDateShort
} from './dataStore.js';

const chartInstances = {};

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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 }
    },
    legend: {
      data: ['Tasks', 'Incidents'],
      textStyle: { color: '#94a3b8', fontSize: 11 },
      top: 0
    },
    grid: { left: 50, right: 20, top: 20, bottom: 40 },
    xAxis: {
      type: 'category',
      data: allDates,
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#64748b', fontSize: 10, rotate: 30,
        formatter: v => { const d = new Date(v); return formatDateShort(d); }
      },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', fontSize: 10 }
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

  const { rows, metrics } = buildTaskKPIRows();
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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      formatter: (p) => {
        const row = rows[p.data[1]];
        const metric = metrics[p.data[0]];
        return `${row.task}<br/>${metric.label}: ${formatValue(p.data[3], p.data[4])}`;
      }
    },
    grid: { left:100,right: 20, top: 50, bottom: 58 },
    xAxis: {
      type: 'category',
      data: metrics.map(m => m.label),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 10, rotate: 25 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'category',
      data: rows.map(r => r.task),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 10, width: 130, overflow: 'truncate' },
      axisTick: { show: false }
    },
    visualMap: {
      min: 0,
      max: 1,
      dimension: 2,
      orient: 'horizontal',
      left: 'center',
      top: 16,
      calculable: false,
      inRange: {
        color: ['rgba(99,132,255,0.12)', 'rgba(99,132,255,0.95)']
      },
      textStyle: { color: '#94a3b8' }
    },
    series: [{
      type: 'heatmap',
      data: normalizedData,
      label: {
        show: true,
        color: document.body.classList.contains('theme-light') ? '#0f172a' : '#f8fafc',
        fontSize: 9,
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

export function renderTaskKPIBars(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Task KPI Comparison (Alternative View)', 'cyan');

  const { rows, metrics } = buildTaskKPIRows();
  const percentMetrics = metrics.filter(m => m.kind === 'pct');

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      formatter: (params) => {
        if (!params?.length) return '';
        const metricLabel = params[0].axisValue;
        const lines = params
          .filter(p => p.value > 0)
          .map(p => `${p.marker} ${p.seriesName}: ${Number(p.value).toFixed(1)}%`);
        return `${metricLabel}<br/>${lines.join('<br/>')}`;
      }
    },
    legend: {
      type: 'scroll',
      top: 0,
      textStyle: { color: '#94a3b8', fontSize: 10 },
      pageTextStyle: { color: '#94a3b8' }
    },
    grid: { left: 60, right: 20, top: 44, bottom: 52 },
    xAxis: {
      type: 'category',
      data: percentMetrics.map(m => m.label),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 10, rotate: 18 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      min: 0,
      max: 1,
      splitNumber: 5,
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', formatter: (v) => `${Math.round(v * 100)}%` }
    },
    color: PALETTE,
    series: rows.map((row) => ({
      name: row.task,
      type: 'bar',
      barMaxWidth: 20,
      data: percentMetrics.map(m => row[m.key] || 0),
      emphasis: {
        focus: 'series'
      }
    })),
    animationDuration: 1000
  });
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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 }
    },
    grid: { left: 100, right: 30, top: 10, bottom: 20 },
    xAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', fontSize: 10 }
    },
    yAxis: {
      type: 'category',
      data: regions.reverse(),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 11 },
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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 }
    },
    grid: { left: 50, right: 20, top: 10, bottom: 30 },
    xAxis: {
      type: 'category',
      data: labels,
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 11 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', fontSize: 10 }
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
        color: '#94a3b8',
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

export function renderBranchHeatmap(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Branch Task Volume & Avg Duration', 'blue');

  const data = getFilteredRawData();
  const branchMap = {};
  data.forEach(r => {
    const b = r['Branch'] || 'N/A';
    if (!branchMap[b]) branchMap[b] = { count: 0, totalDuration: 0 };
    branchMap[b].count++;
    const dur = parseFloat(r['Duration']);
    if (!isNaN(dur)) branchMap[b].totalDuration += dur;
  });

  const sorted = Object.entries(branchMap)
    .map(([name, v]) => ({ name, count: v.count, avgDur: v.count > 0 ? (v.totalDuration / v.count) : 0 }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 25);

  const branches = sorted.map(s => s.name);
  const counts = sorted.map(s => s.count);
  const durations = sorted.map(s => Math.round(s.avgDur * 24 * 10) / 10);

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      axisPointer: { type: 'shadow' }
    },
    legend: {
      data: ['Task Count', 'Avg Duration (hrs)'],
      textStyle: { color: '#94a3b8', fontSize: 11 },
      top: 0
    },
    grid: { left: 100, right: 60, top: 35, bottom: 20 },
    xAxis: [
      {
        type: 'value',
        position: 'bottom',
        splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
        axisLabel: { color: '#64748b', fontSize: 10 }
      },
      {
        type: 'value',
        position: 'top',
        splitLine: { show: false },
        axisLabel: { color: '#64748b', fontSize: 10, formatter: '{value}h' }
      }
    ],
    yAxis: {
      type: 'category',
      data: branches.reverse(),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 10 },
      axisTick: { show: false }
    },
    series: [
      {
        name: 'Task Count',
        type: 'bar',
        data: counts.reverse(),
        barWidth: 12,
        itemStyle: {
          borderRadius: [0, 4, 4, 0],
          color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
            { offset: 0, color: COLORS.blue },
            { offset: 1, color: COLORS.purple }
          ])
        }
      },
      {
        name: 'Avg Duration (hrs)',
        type: 'bar',
        xAxisIndex: 1,
        data: durations.reverse(),
        barWidth: 12,
        itemStyle: {
          borderRadius: [0, 4, 4, 0],
          color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
            { offset: 0, color: COLORS.amber },
            { offset: 1, color: COLORS.orange }
          ])
        }
      }
    ],
    animationDuration: 1200
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

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 }
    },
    grid: { left: 50, right: 20, top: 20, bottom: 40 },
    xAxis: {
      type: 'category',
      data: sorted.map(s => s[0]),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#64748b', fontSize: 10, rotate: 30,
        formatter: v => formatDateShort(new Date(v))
      },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', fontSize: 10 }
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
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Incident Categories', 'amber');

  const data = getFilteredIncidentData();
  const catMap = {};
  data.forEach(r => {
    const c = r['Category'] || 'Unknown';
    catMap[c] = (catMap[c] || 0) + 1;
  });

  const sorted = Object.entries(catMap).sort((a, b) => b[1] - a[1]).slice(0, 10);

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 }
    },
    grid: { left: 180, right: 30, top: 10, bottom: 20 },
    xAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', fontSize: 10 }
    },
    yAxis: {
      type: 'category',
      data: sorted.map(s => s[0]).reverse(),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 10, width: 160, overflow: 'truncate' },
      axisTick: { show: false }
    },
    series: [{
      type: 'bar',
      data: sorted.map(s => s[1]).reverse(),
      barWidth: 16,
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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      formatter: '{b}: {c} ({d}%)'
    },
    color: [COLORS.green, COLORS.amber, COLORS.red, COLORS.purple, COLORS.blue],
    series: [{
      type: 'pie',
      radius: ['40%', '70%'],
      center: ['50%', '50%'],
      itemStyle: { borderRadius: 6, borderColor: 'rgba(10,14,26,0.8)', borderWidth: 2 },
      label: {
        color: '#94a3b8',
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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      formatter: '{b}: {c} ({d}%)'
    },
    color: [COLORS.green, COLORS.red, COLORS.blue],
    series: [{
      type: 'pie',
      radius: ['40%', '70%'],
      center: ['50%', '50%'],
      itemStyle: { borderRadius: 6, borderColor: 'rgba(10,14,26,0.8)', borderWidth: 2 },
      label: {
        color: '#94a3b8',
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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      axisPointer: { type: 'shadow' }
    },
    legend: {
      data: ['Tasks', 'Same Day (count)', 'Closed <= 4h', 'Closed <= 24h'],
      textStyle: { color: '#94a3b8', fontSize: 11 },
      top: 0
    },
    grid: { left: 70, right: 30, top: 35, bottom: 40 },
    xAxis: {
      type: 'category',
      data: sorted.map(s => s.name),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 10, rotate: 40 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      name: 'Tasks',
      nameTextStyle: { color: '#64748b', fontSize: 10 },
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', fontSize: 10 }
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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 }
    },
    legend: {
      data: ['Same Day %', 'Closed <= 4h %', 'Closed <= 24h %'],
      textStyle: { color: '#94a3b8', fontSize: 11 },
      top: 0
    },
    grid: { left: 70, right: 30, top: 35, bottom: 40 },
    xAxis: {
      type: 'category',
      data: sorted.map(s => s.name),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 10, rotate: 40 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      min: 0,
      max: 100,
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', fontSize: 10, formatter: '{value}%' }
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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 }
    },
    color: [COLORS.blue, COLORS.green],
    legend: {
      data: ['Total Tasks', 'Same Day Closed'],
      textStyle: { color: '#94a3b8', fontSize: 11 },
      bottom: 0
    },
    radar: {
      indicator,
      shape: 'polygon',
      axisName: { color: '#94a3b8', fontSize: 11 },
      splitArea: { areaStyle: { color: ['rgba(255,255,255,0.02)', 'rgba(255,255,255,0.04)'] } },
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } },
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.06)' } }
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
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      axisPointer: { type: 'shadow' }
    },
    legend: {
      data: ['Open (start)', 'New', 'Same Day Closure', 'Open (next day)'],
      textStyle: { color: '#94a3b8', fontSize: 11 },
      top: 0
    },
    grid: { left: 50, right: 20, top: 35, bottom: 48 },
    xAxis: {
      type: 'category',
      data: dayOrder,
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 11 },
      axisTick: { show: false }
    },
    yAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', fontSize: 10 }
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
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Task Volume by Branch & Type (Top 15)', 'cyan');

  const data = getFilteredRawData();

  // Get task types
  const taskTypes = new Set();
  data.forEach(r => { if (r['Task Type']) taskTypes.add(r['Task Type']); });
  const types = [...taskTypes];

  // Group by branch
  const branchMap = {};
  data.forEach(r => {
    const b = r['Branch'] || 'N/A';
    const t = r['Task Type'] || 'Unknown';
    if (!branchMap[b]) branchMap[b] = {};
    branchMap[b][t] = (branchMap[b][t] || 0) + 1;
  });

  // Sort by total
  const sorted = Object.entries(branchMap)
    .map(([name, types_]) => ({ name, types: types_, total: Object.values(types_).reduce((a, b) => a + b, 0) }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 15);

  const branches = sorted.map(s => s.name);

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      axisPointer: { type: 'shadow' }
    },
    legend: {
      data: types,
      type: 'scroll',
      textStyle: { color: '#94a3b8', fontSize: 10 },
      pageTextStyle: { color: '#94a3b8' },
      top: 0
    },
    grid: { left: 100, right: 20, top: 40, bottom: 20 },
    color: PALETTE,
    xAxis: {
      type: 'value',
      splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
      axisLabel: { color: '#64748b', fontSize: 10 }
    },
    yAxis: {
      type: 'category',
      data: branches.reverse(),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 10 },
      axisTick: { show: false }
    },
    series: types.map(type => ({
      name: type,
      type: 'bar',
      stack: 'total',
      data: branches.map(b => {
        const entry = sorted.find(s => s.name === b);
        return entry ? (entry.types[type] || 0) : 0;
      }),
      barWidth: 16,
      itemStyle: { borderRadius: 0 }
    })),
    animationDuration: 1200
  });
}
