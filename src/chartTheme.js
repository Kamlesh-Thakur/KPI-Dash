/**
 * Theme-aware colors for ECharts (body.theme-light vs dark).
 * Aligns with CSS tokens: --text-secondary, --text-muted, borders.
 */

export function chartIsLight() {
  return typeof document !== 'undefined' && document.body.classList.contains('theme-light');
}

/** Category / legend / secondary labels */
export function chartAxisLabelCat() {
  return chartIsLight() ? '#334155' : '#94a3b8';
}

/** Value-axis tick labels */
export function chartAxisLabelValue() {
  return chartIsLight() ? '#475569' : '#64748b';
}

export function chartAxisLine() {
  return chartIsLight() ? 'rgba(15, 23, 42, 0.12)' : 'rgba(255, 255, 255, 0.08)';
}

export function chartSplitLine() {
  return chartIsLight() ? 'rgba(15, 23, 42, 0.09)' : 'rgba(255, 255, 255, 0.05)';
}

/**
 * Tooltip panel styling (spread into tooltip: { ... }).
 * @param {number} fontSize
 */
export function chartTooltipTheme(fontSize = 12) {
  if (chartIsLight()) {
    return {
      backgroundColor: 'rgba(255, 255, 255, 0.97)',
      borderColor: 'rgba(15, 23, 42, 0.14)',
      textStyle: { color: '#0f172a', fontSize }
    };
  }
  return {
    backgroundColor: 'rgba(17, 24, 39, 0.96)',
    borderColor: 'rgba(255, 255, 255, 0.08)',
    textStyle: { color: '#f1f5f9', fontSize }
  };
}

export function chartHeatmapCellTextColor() {
  return chartIsLight() ? '#0f172a' : '#f8fafc';
}

/** Inline HTML in ECharts tooltip formatters (textStyle alone may not apply to HTML). */
export function chartTooltipHtmlColor() {
  return chartIsLight() ? '#0f172a' : '#f1f5f9';
}

export function chartRadarSplitAreaColors() {
  return chartIsLight()
    ? ['rgba(15, 23, 42, 0.04)', 'rgba(15, 23, 42, 0.08)']
    : ['rgba(255, 255, 255, 0.02)', 'rgba(255, 255, 255, 0.04)'];
}

export function chartRadarLineColor() {
  return chartIsLight() ? 'rgba(15, 23, 42, 0.12)' : 'rgba(255, 255, 255, 0.06)';
}
