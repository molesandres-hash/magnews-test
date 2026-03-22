import type { QuestionInfo, ScaleAnalytics } from '@/types/survey';
import type { CompanyTemplate } from '@/types/companyTemplate';
import { getBlockDisplayName } from './analytics';
import { getMeanBarColor } from './templateColors';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];
const CHART_COLORS = [
  '#22C55E', '#22C55E', '#22C55E', '#84CC16',
  '#FDE047', '#F59E0B', '#F97316', '#EF4444',
  '#DC2626', '#B91C1C', '#94A3B8'
];

/**
 * Generate a horizontal bar chart of mean values per question in a block.
 * Returns base64 PNG (without prefix).
 */
export async function generateBlockMeanChartPNG(
  blockId: number | null,
  questions: QuestionInfo[],
  scaleAnalytics: Map<string, ScaleAnalytics>,
  template: CompanyTemplate | null,
  Plotly: any,
  container: HTMLDivElement
): Promise<string> {
  const sorted = [...questions].sort((a, b) => a.subId - b.subId);
  const labels = sorted.map(q => q.questionKey || q.questionText.slice(0, 20)).reverse();
  const means = sorted.map(q => {
    const a = scaleAnalytics.get(q.id);
    return a ? a.mean : 0;
  }).reverse();
  const colors = means.map(m => getMeanBarColor(m, template));
  const fontFamily = template?.fontFamily || 'Arial';
  const chartHeight = Math.max(300, questions.length * 35 + 100);

  const data = [{
    y: labels, x: means, type: 'bar', orientation: 'h',
    marker: { color: colors, line: { color: '#1E293B', width: 1 } },
    text: means.map(m => m.toFixed(2)), textposition: 'outside',
  }];

  const layout = {
    title: { text: getBlockDisplayName(blockId), font: { size: 16, family: fontFamily } },
    xaxis: { range: [0, 10], title: { text: 'Media', font: { family: fontFamily } }, tickfont: { family: fontFamily } },
    yaxis: { tickfont: { size: 10, family: fontFamily } },
    margin: { t: 50, r: 60, b: 50, l: 80 },
    paper_bgcolor: '#FFFFFF', plot_bgcolor: '#F8FAFC',
    width: 900, height: chartHeight,
  };

  await Plotly.default.newPlot(container, data, layout as any, { displayModeBar: false });
  const imgData = await Plotly.default.toImage(container, { format: 'png', width: 900, height: chartHeight });
  return imgData.replace('data:image/png;base64,', '');
}

/**
 * Generate a grouped distribution chart for all scale questions in a block.
 * Returns base64 PNG (without prefix).
 */
export async function generateBlockDistributionChartPNG(
  blockId: number | null,
  questions: QuestionInfo[],
  scaleAnalytics: Map<string, ScaleAnalytics>,
  template: CompanyTemplate | null,
  Plotly: any,
  container: HTMLDivElement
): Promise<string> {
  const sorted = [...questions].sort((a, b) => a.subId - b.subId);
  const questionLabels = sorted.map(q => q.questionKey || q.questionText.slice(0, 15));
  const fontFamily = template?.fontFamily || 'Arial';

  const traces = SCALE_ORDER.map((scaleVal, idx) => ({
    name: scaleVal,
    x: questionLabels,
    y: sorted.map(q => {
      const a = scaleAnalytics.get(q.id);
      return a ? (a.counts[scaleVal] || 0) : 0;
    }),
    type: 'bar' as const,
    marker: { color: CHART_COLORS[idx] },
  }));

  const layout = {
    title: { text: `${getBlockDisplayName(blockId)} - Distribuzione`, font: { size: 14, family: fontFamily } },
    barmode: 'group',
    xaxis: { tickfont: { size: 9, family: fontFamily }, tickangle: -45 },
    yaxis: { title: { text: 'Conteggio', font: { family: fontFamily } } },
    margin: { t: 50, r: 30, b: 100, l: 50 },
    paper_bgcolor: '#FFFFFF', plot_bgcolor: '#F8FAFC',
    width: 900, height: 400,
    legend: { font: { size: 9, family: fontFamily }, orientation: 'h' as const, y: -0.3 },
  };

  await Plotly.default.newPlot(container, traces, layout as any, { displayModeBar: false });
  const imgData = await Plotly.default.toImage(container, { format: 'png', width: 900, height: 400 });
  return imgData.replace('data:image/png;base64,', '');
}
