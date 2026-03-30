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
  setTitle(titleEl, 'Tasks Over Time', 'blue');

  const data = getFilteredRawData();

  // Group by completed date
  const dateMap = {};
  data.forEach(row => {
    const d = excelDateToJS(row['Completed date']);
    if (!d) return;
    const key = d.toISOString().slice(0, 10);
    dateMap[key] = (dateMap[key] || 0) + 1;
  });

  const sorted = Object.entries(dateMap).sort((a, b) => a[0].localeCompare(b[0]));
  const dates = sorted.map(s => s[0]);
  const counts = sorted.map(s => s[1]);

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
      data: dates,
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
      type: 'line',
      data: counts,
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
    }],
    animationDuration: 1200,
    animationEasing: 'cubicOut'
  });
}

export function renderTaskTypePie(containerId) {
  const { chart, titleEl } = getOrCreate(containerId) || {};
  if (!chart) return;
  setTitle(titleEl, 'Tasks by Type', 'purple');

  const data = getFilteredRawData();
  const typeMap = {};
  data.forEach(r => {
    const t = r['Task Type'] || 'Unknown';
    typeMap[t] = (typeMap[t] || 0) + 1;
  });

  const pieData = Object.entries(typeMap)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);

  chart.setOption({
    tooltip: {
      trigger: 'item',
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      formatter: '{b}: {c} ({d}%)'
    },
    legend: {
      type: 'scroll',
      orient: 'vertical',
      right: 10,
      top: 'middle',
      textStyle: { color: '#94a3b8', fontSize: 11 },
      pageTextStyle: { color: '#94a3b8' }
    },
    color: PALETTE,
    series: [{
      type: 'pie',
      radius: ['45%', '72%'],
      center: ['35%', '50%'],
      avoidLabelOverlap: true,
      itemStyle: { borderRadius: 6, borderColor: 'rgba(10,14,26,0.8)', borderWidth: 2 },
      label: { show: false },
      emphasis: {
        label: { show: true, fontSize: 13, fontWeight: 600, color: '#f1f5f9' },
        itemStyle: { shadowBlur: 20, shadowColor: 'rgba(0,0,0,0.3)' }
      },
      data: pieData
    }],
    animationDuration: 1000,
    animationEasing: 'cubicOut'
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
  setTitle(titleEl, 'Branch Performance — Tasks & Same-Day Closure %', 'blue');

  const data = getFilteredRawData();
  const branchMap = {};
  data.forEach(r => {
    const b = r['Branch'] || 'N/A';
    if (!branchMap[b]) branchMap[b] = { total: 0, sameDay: 0 };
    branchMap[b].total++;
    if (r['Same day Closure'] === true) branchMap[b].sameDay++;
  });

  const sorted = Object.entries(branchMap)
    .map(([name, v]) => ({ name, total: v.total, rate: v.total > 0 ? Math.round((v.sameDay / v.total) * 100) : 0 }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 20);

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      backgroundColor: 'rgba(17,24,39,0.95)',
      borderColor: 'rgba(255,255,255,0.08)',
      textStyle: { color: '#f1f5f9', fontSize: 12 },
      axisPointer: { type: 'shadow' }
    },
    legend: {
      data: ['Tasks', 'Same Day %'],
      textStyle: { color: '#94a3b8', fontSize: 11 },
      top: 0
    },
    grid: { left: 90, right: 50, top: 35, bottom: 40 },
    xAxis: {
      type: 'category',
      data: sorted.map(s => s.name),
      axisLine: { lineStyle: { color: 'rgba(255,255,255,0.08)' } },
      axisLabel: { color: '#94a3b8', fontSize: 10, rotate: 40 },
      axisTick: { show: false }
    },
    yAxis: [
      {
        type: 'value',
        name: 'Tasks',
        nameTextStyle: { color: '#64748b', fontSize: 10 },
        splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } },
        axisLabel: { color: '#64748b', fontSize: 10 }
      },
      {
        type: 'value',
        name: '%',
        max: 100,
        nameTextStyle: { color: '#64748b', fontSize: 10 },
        splitLine: { show: false },
        axisLabel: { color: '#64748b', fontSize: 10, formatter: '{value}%' }
      }
    ],
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
        name: 'Same Day %',
        type: 'line',
        yAxisIndex: 1,
        data: sorted.map(s => s.rate),
        smooth: true,
        symbol: 'circle',
        symbolSize: 6,
        lineStyle: { color: COLORS.green, width: 2 },
        itemStyle: { color: COLORS.green }
      }
    ],
    animationDuration: 1200
  });
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
  setTitle(titleEl, 'Closure by Day of Week', 'amber');

  const data = getFilteredRawData();
  const dayMap = {};
  const dayOrder = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  dayOrder.forEach(d => { dayMap[d] = { total: 0, sameDay: 0 }; });

  data.forEach(r => {
    const day = r['Completed Day'];
    if (day && dayMap[day]) {
      dayMap[day].total++;
      if (r['Same day Closure'] === true) dayMap[day].sameDay++;
    }
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
      data: ['Total', 'Same Day'],
      textStyle: { color: '#94a3b8', fontSize: 11 },
      top: 0
    },
    grid: { left: 50, right: 20, top: 35, bottom: 30 },
    xAxis: {
      type: 'category',
      data: dayOrder.map(d => d.slice(0, 3)),
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
        name: 'Total',
        type: 'bar',
        data: dayOrder.map(d => dayMap[d].total),
        barWidth: 20,
        itemStyle: { borderRadius: [4, 4, 0, 0], color: COLORS.blue }
      },
      {
        name: 'Same Day',
        type: 'bar',
        data: dayOrder.map(d => dayMap[d].sameDay),
        barWidth: 20,
        itemStyle: { borderRadius: [4, 4, 0, 0], color: COLORS.green }
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
