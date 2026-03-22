import type { ScaleAnalytics, QuestionInfo } from '@/types/survey';
import type { ChartSettings } from '@/types/chartSettings';
import { getColorArray, hexToRgba, darkenHex, lightenHex } from '@/utils/colorUtils';
import { useTemplateStore } from '@/store/templateStore';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];

/** Convert a mean value (1-10) to a categorical axis position.
 *  SCALE_ORDER: '10'=0, '9'=1, ... '1'=9, 'N/A'=10
 *  So position = 10 - mean */
function meanToAxisPos(mean: number): number {
  return Math.max(0, Math.min(10, 10 - mean));
}

function getChartHeightPx(settings: ChartSettings): number {
  if (settings.chartHeight === 'compact') return 280;
  if (settings.chartHeight === 'tall') return 520;
  return 380;
}

function getTitle(question: QuestionInfo, fontFamily: string) {
  const text = `${question.questionKey || ''} - ${question.questionText.slice(0, 60)}${question.questionText.length > 60 ? '...' : ''}`;
  return { text, font: { size: 14, color: '#1E293B', family: fontFamily } };
}

function getSubtitle(analytics: ScaleAnalytics, fontFamily: string) {
  return {
    x: 0.5, y: 1.12, xref: 'paper' as const, yref: 'paper' as const,
    text: `Media: ${analytics.mean.toFixed(2)} | Risposte: ${analytics.validResponses}/${analytics.totalRespondents}`,
    showarrow: false, font: { size: 11, color: '#64748B', family: fontFamily },
  };
}

/** Mean line for vertical bar charts — vertical line on the categorical X axis */
function meanLineVertical(analytics: ScaleAnalytics) {
  const pos = meanToAxisPos(analytics.mean);
  return {
    type: 'line' as const, xref: 'x' as const, yref: 'paper' as const,
    x0: pos, x1: pos, y0: 0, y1: 1,
    line: { color: '#EF4444', width: 2, dash: 'dash' as const },
  };
}

/** Mean line for horizontal bar charts — horizontal line on the categorical Y axis */
function meanLineHorizontal(analytics: ScaleAnalytics) {
  const pos = meanToAxisPos(analytics.mean);
  return {
    type: 'line' as const, xref: 'paper' as const, yref: 'y' as const,
    x0: 0, x1: 1, y0: pos, y1: pos,
    line: { color: '#EF4444', width: 2, dash: 'dash' as const },
  };
}

export function buildPlotlyConfig(
  analytics: ScaleAnalytics,
  question: QuestionInfo,
  settings: ChartSettings
): { data: any[]; layout: any } {
  const template = useTemplateStore.getState().getActiveTemplate();
  const fontFamily = template?.fontFamily || 'Arial';
  const values = SCALE_ORDER.map((key) => analytics.counts[key] || 0);
  const colors = getColorArray(settings);
  const height = getChartHeightPx(settings);
  const showLabels = settings.showDataLabels && settings.dataLabelPosition !== 'none';

  const baseLayout = {
    margin: { t: 80, r: 30, b: 50, l: 50 },
    paper_bgcolor: '#FFFFFF',
    plot_bgcolor: '#F8FAFC',
    autosize: true,
    height,
    showlegend: settings.showLegend,
  };

  switch (settings.chartType) {
    case 'bar_horizontal': {
      return {
        data: [{
          type: 'bar', orientation: 'h',
          y: SCALE_ORDER, x: values,
          marker: { color: colors, line: { width: settings.barBorderWidth, color: '#1E293B' } },
          text: showLabels ? values.map(String) : [],
          textposition: settings.dataLabelPosition === 'inside' ? 'inside' : 'outside',
        }],
        layout: {
          ...baseLayout,
          title: getTitle(question, fontFamily),
          annotations: [getSubtitle(analytics, fontFamily)],
          xaxis: { title: { text: 'Conteggio', font: { family: fontFamily } }, tickfont: { family: fontFamily }, showgrid: settings.showGridLines },
          yaxis: { title: { text: 'Valutazione', font: { family: fontFamily } }, tickfont: { family: fontFamily } },
          bargap: settings.barSpacing,
          shapes: settings.showMean ? [meanLineHorizontal(analytics)] : [],
        },
      };
    }

    case 'bar_3d': {
      const depth = settings.depthEffect;
      const shapes: any[] = [];
      // Simulated 3D via layout shapes - offset side and top faces
      // Note: Plotly shape coordinates in data space
      SCALE_ORDER.forEach((_, i) => {
        const v = values[i];
        if (v === 0) return;
        const x = i;
        const barWidth = 0.35;
        const dxPx = depth * 0.02;
        const dyPx = depth * 0.015;
        // Side face (right)
        shapes.push({
          type: 'path', path: `M ${x + barWidth},0 L ${x + barWidth + dxPx},${dyPx} L ${x + barWidth + dxPx},${v + dyPx} L ${x + barWidth},${v} Z`,
          fillcolor: darkenHex(typeof colors[i] === 'string' && colors[i].startsWith('#') ? colors[i] : '#3B82F6', 40),
          line: { width: 0.5, color: '#1E293B' },
        });
        // Top face
        shapes.push({
          type: 'path', path: `M ${x - barWidth},${v} L ${x + barWidth},${v} L ${x + barWidth + dxPx},${v + dyPx} L ${x - barWidth + dxPx},${v + dyPx} Z`,
          fillcolor: lightenHex(typeof colors[i] === 'string' && colors[i].startsWith('#') ? colors[i] : '#3B82F6', 30),
          line: { width: 0.5, color: '#1E293B' },
        });
      });

      return {
        data: [{
          type: 'bar',
          x: SCALE_ORDER, y: values,
          marker: { color: colors, line: { width: settings.barBorderWidth } },
          text: showLabels ? values.map(String) : [],
          textposition: settings.dataLabelPosition === 'inside' ? 'inside' : 'outside',
        }],
        layout: {
          ...baseLayout,
          title: getTitle(question, fontFamily),
          annotations: [getSubtitle(analytics, fontFamily)],
          xaxis: { title: { text: 'Valutazione', font: { family: fontFamily } }, tickfont: { family: fontFamily } },
          yaxis: { title: { text: 'Conteggio', font: { family: fontFamily } }, tickfont: { family: fontFamily }, showgrid: settings.showGridLines },
          bargap: settings.barSpacing,
          shapes: [...shapes, ...(settings.showMean ? [meanLineVertical(analytics)] : [])],
        },
      };
    }

    case 'pie': {
      const filteredIdx = SCALE_ORDER.map((s, i) => ({ s, v: values[i], c: colors[i] })).filter((d) => d.s !== 'N/A' && d.v > 0);
      return {
        data: [{
          type: 'pie',
          values: filteredIdx.map((d) => d.v),
          labels: filteredIdx.map((d) => d.s),
          marker: { colors: filteredIdx.map((d) => d.c) },
          textinfo: showLabels ? 'label+percent' : 'none',
          textposition: 'outside',
          hole: 0,
        }],
        layout: {
          ...baseLayout,
          title: getTitle(question, fontFamily),
          annotations: [getSubtitle(analytics, fontFamily)],
        },
      };
    }

    case 'donut': {
      const filteredIdx = SCALE_ORDER.map((s, i) => ({ s, v: values[i], c: colors[i] })).filter((d) => d.s !== 'N/A' && d.v > 0);
      return {
        data: [{
          type: 'pie',
          values: filteredIdx.map((d) => d.v),
          labels: filteredIdx.map((d) => d.s),
          marker: { colors: filteredIdx.map((d) => d.c) },
          textinfo: showLabels ? 'label+percent' : 'none',
          textposition: 'outside',
          hole: 0.45,
        }],
        layout: {
          ...baseLayout,
          title: getTitle(question, fontFamily),
          annotations: [
            getSubtitle(analytics, fontFamily),
            {
              text: `<b>${analytics.mean.toFixed(1)}</b><br>media`,
              x: 0.5, y: 0.5, xref: 'paper', yref: 'paper',
              showarrow: false, font: { size: 18, family: fontFamily },
            },
          ],
        },
      };
    }

    case 'radar': {
      const radarValues = values.slice(0, 10);
      const radarLabels = SCALE_ORDER.slice(0, 10);
      return {
        data: [{
          type: 'scatterpolar',
          r: [...radarValues, radarValues[0]], // close the polygon
          theta: [...radarLabels, radarLabels[0]],
          fill: 'toself',
          fillcolor: hexToRgba(settings.solidColor, 0.3),
          line: { color: settings.solidColor },
          name: question.questionKey || '',
        }],
        layout: {
          ...baseLayout,
          title: getTitle(question, fontFamily),
          annotations: [getSubtitle(analytics, fontFamily)],
          polar: { radialaxis: { range: [0, Math.max(...radarValues, 1)], showgrid: settings.showGridLines } },
        },
      };
    }

    case 'bubble': {
      return {
        data: [{
          type: 'scatter', mode: 'markers',
          x: SCALE_ORDER, y: values.map(() => 0),
          marker: {
            size: values.map((v) => Math.max(v * 8, 5)),
            color: colors, opacity: 0.8,
            line: { width: 1, color: '#fff' },
          },
          text: values.map((v, i) => `${SCALE_ORDER[i]}: ${v} risposte`),
          hoverinfo: 'text',
        }],
        layout: {
          ...baseLayout,
          title: getTitle(question, fontFamily),
          annotations: [getSubtitle(analytics, fontFamily)],
          xaxis: { title: { text: 'Valutazione', font: { family: fontFamily } }, tickfont: { family: fontFamily } },
          yaxis: { visible: false, range: [-1, 1] },
        },
      };
    }

    case 'waterfall': {
      return {
        data: [{
          type: 'waterfall',
          x: SCALE_ORDER, y: values,
          measure: values.map(() => 'relative'),
          connector: { line: { color: '#888', width: 1 } },
          increasing: { marker: { color: settings.gradientTo } },
          decreasing: { marker: { color: settings.gradientFrom } },
          text: showLabels ? values.map(String) : [],
          textposition: 'outside',
        }],
        layout: {
          ...baseLayout,
          title: getTitle(question, fontFamily),
          annotations: [getSubtitle(analytics, fontFamily)],
          xaxis: { title: { text: 'Valutazione', font: { family: fontFamily } }, tickfont: { family: fontFamily } },
          yaxis: { title: { text: 'Conteggio', font: { family: fontFamily } }, tickfont: { family: fontFamily }, showgrid: settings.showGridLines },
        },
      };
    }

    case 'gauge': {
      return {
        data: [{
          type: 'indicator',
          mode: 'gauge+number+delta',
          value: analytics.mean,
          delta: { reference: 7.5, increasing: { color: '#22C55E' } },
          gauge: {
            axis: { range: [0, 10], tickwidth: 1 },
            bar: { color: settings.solidColor },
            bgcolor: 'white',
            steps: [
              { range: [0, 4], color: '#FEE2E2' },
              { range: [4, 7], color: '#FEF9C3' },
              { range: [7, 10], color: '#DCFCE7' },
            ],
            threshold: {
              line: { color: 'red', width: 4 },
              thickness: 0.75,
              value: analytics.mean,
            },
          },
          title: { text: question.questionText.slice(0, 60), font: { family: fontFamily } },
        }],
        layout: {
          ...baseLayout,
          annotations: [getSubtitle(analytics, fontFamily)],
        },
      };
    }

    // bar_vertical (default)
    default: {
      return {
        data: [{
          type: 'bar',
          x: SCALE_ORDER, y: values,
          marker: { color: colors, line: { width: settings.barBorderWidth, color: template?.accentColor || '#1E40AF' } },
          text: showLabels ? values.map(String) : [],
          textposition: settings.dataLabelPosition === 'inside' ? 'inside' : 'outside',
        }],
        layout: {
          ...baseLayout,
          title: getTitle(question, fontFamily),
          annotations: [getSubtitle(analytics, fontFamily)],
          xaxis: { title: { text: 'Valutazione', font: { family: fontFamily } }, tickfont: { size: 11, family: fontFamily } },
          yaxis: { title: { text: 'Conteggio', font: { family: fontFamily } }, tickfont: { size: 11, family: fontFamily }, showgrid: settings.showGridLines },
          bargap: settings.barSpacing,
          shapes: settings.showMean ? [meanLineVertical(analytics)] : [],
        },
      };
    }
  }
}
