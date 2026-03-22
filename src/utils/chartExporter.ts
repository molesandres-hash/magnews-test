import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import type { ParsedSurvey, QuestionInfo, ScaleAnalytics } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';
import { getShortQuestionText } from './headerNormalizer';
import { useTemplateStore } from '@/store/templateStore';
import { useChartSettingsStore } from '@/store/chartSettingsStore';
import { buildPlotlyConfig } from './chartBuilder';
import { getTemplateChartColors, getTemplatePlotBg, getMeanBarColor } from './templateColors';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];

export { getBlockDisplayName };

function getGradientColor(value: number | null): string {
  if (value === null) return '#94A3B8';
  if (value <= 3) return '#EF4444';
  if (value <= 5) return '#F59E0B';
  if (value <= 7) return '#FDE047';
  return '#22C55E';
}

function getTextColor(bgColor: string): string {
  if (bgColor === '#FDE047' || bgColor === '#22C55E') return '#1E293B';
  return '#FFFFFF';
}

export function getDistributionChartData(analytics: ScaleAnalytics, question: QuestionInfo) {
  const settings = useChartSettingsStore.getState().getEffectiveSettings(question.id);
  return buildPlotlyConfig(analytics, question, settings);
}

export function getBlockSummaryChartData(
  blockId: number | null, questions: QuestionInfo[], scaleAnalytics: Map<string, ScaleAnalytics>
) {
  const template = useTemplateStore.getState().getActiveTemplate();
  const fontFamily = template?.fontFamily || 'Arial';
  const blockQuestions = questions.filter(q => q.blockId === blockId && q.type === 'scale_1_10_na');
  const labels = blockQuestions.map(q => q.questionKey ? `${q.questionKey}` : getShortQuestionText(q.questionText, 20));
  const means = blockQuestions.map(q => { const a = scaleAnalytics.get(q.id); return a ? a.mean : 0; });
  const colors = means.map(m => getMeanBarColor(m, template));

  return {
    data: [{ x: labels, y: means, type: 'bar' as const, marker: { color: colors, line: { color: '#1E293B', width: 1 } },
      text: means.map(m => m.toFixed(2)), textposition: 'outside' as const }],
    layout: {
      title: { text: `${getBlockDisplayName(blockId)} - Medie per domanda`, font: { size: 18, color: '#1E293B', family: fontFamily } },
      xaxis: { title: { text: 'Domanda', font: { size: 12, family: fontFamily } }, tickfont: { size: 10, family: fontFamily }, tickangle: -45 },
      yaxis: { title: { text: 'Media', font: { size: 12, family: fontFamily } }, range: [0, 10], tickfont: { size: 12, family: fontFamily } },
      margin: { t: 60, r: 40, b: 120, l: 60 },
      paper_bgcolor: '#FFFFFF', plot_bgcolor: getTemplatePlotBg(template),
    },
  };
}

function createResponseMatrix(
  analytics: ScaleAnalytics, respondents: { id: string; displayName: string }[], width: number
): HTMLCanvasElement {
  const cellWidth = 120; const cellHeight = 32; const headerHeight = 40; const padding = 20;
  const respondentData = respondents.map(r => ({ name: r.displayName.split(' ')[0] || 'Anonimo', value: analytics.respondentValues[r.id] })).filter(r => r.value !== undefined);
  const numCols = Math.ceil(width / cellWidth);
  const numRows = Math.ceil(respondentData.length / numCols);
  const canvasWidth = width;
  const canvasHeight = headerHeight + (numRows * cellHeight) + padding * 2;
  const canvas = document.createElement('canvas');
  canvas.width = canvasWidth; canvas.height = canvasHeight;
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, canvasWidth, canvasHeight);
  ctx.fillStyle = '#1E293B'; ctx.font = 'bold 16px Arial'; ctx.textAlign = 'center';
  ctx.fillText('Matrice Risposte', canvasWidth / 2, headerHeight / 2 + 6);
  respondentData.forEach((r, idx) => {
    const col = idx % numCols; const row = Math.floor(idx / numCols);
    const x = padding + col * cellWidth; const y = headerHeight + row * cellHeight;
    const bgColor = getGradientColor(r.value);
    ctx.fillStyle = bgColor; ctx.fillRect(x, y, cellWidth - 4, cellHeight - 4);
    ctx.strokeStyle = '#CBD5E1'; ctx.lineWidth = 1; ctx.strokeRect(x, y, cellWidth - 4, cellHeight - 4);
    const textColor = getTextColor(bgColor);
    ctx.fillStyle = textColor; ctx.font = '12px Arial'; ctx.textAlign = 'left';
    const maxNameWidth = cellWidth - 50;
    let displayName = r.name;
    while (ctx.measureText(displayName).width > maxNameWidth && displayName.length > 3) displayName = displayName.slice(0, -1);
    if (displayName !== r.name) displayName += '…';
    ctx.fillText(displayName, x + 8, y + cellHeight / 2 + 4);
    ctx.font = 'bold 14px Arial'; ctx.textAlign = 'right';
    ctx.fillText(r.value !== null ? r.value.toString() : 'N/A', x + cellWidth - 12, y + cellHeight / 2 + 4);
  });
  return canvas;
}

async function combineChartWithMatrix(chartDataUrl: string, matrixCanvas: HTMLCanvasElement): Promise<string> {
  const chartImg = new Image(); chartImg.src = chartDataUrl;
  await new Promise<void>(resolve => { chartImg.onload = () => resolve(); });
  const totalWidth = chartImg.width; const totalHeight = chartImg.height + matrixCanvas.height + 20;
  const canvas = document.createElement('canvas'); canvas.width = totalWidth; canvas.height = totalHeight;
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, totalWidth, totalHeight);
  ctx.drawImage(chartImg, 0, 0);
  ctx.drawImage(matrixCanvas, (totalWidth - matrixCanvas.width) / 2, chartImg.height + 10);
  return canvas.toDataURL('image/png');
}

export async function generateChartsZip(
  survey: ParsedSurvey, onProgress?: (current: number, total: number) => void
): Promise<void> {
  const zip = new JSZip();
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  const Plotly = await import('plotly.js-dist-min');
  const container = document.createElement('div');
  container.style.position = 'absolute'; container.style.left = '-9999px';
  container.style.width = '1600px'; container.style.height = '900px';
  document.body.appendChild(container);

  try {
    let processed = 0;
    const total = scaleQuestions.length + new Set(scaleQuestions.map(q => q.blockId)).size;

    for (const question of scaleQuestions) {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (!analytics) continue;
      const { data, layout } = getDistributionChartData(analytics, question);
      await Plotly.default.newPlot(container, data, layout as any, { displayModeBar: false });
      const chartImageData = await Plotly.default.toImage(container, { format: 'png', width: 1600, height: 700 });
      const matrixCanvas = createResponseMatrix(analytics, survey.respondents, 1600);
      const combinedImage = await combineChartWithMatrix(chartImageData, matrixCanvas);
      const base64Data = combinedImage.replace('data:image/png;base64,', '');
      const fileName = question.questionKey ? `${question.questionKey.replace('.', '_')}.png` : `domanda_${processed + 1}.png`;
      zip.file(fileName, base64Data, { base64: true });
      processed++; onProgress?.(processed, total);
    }

    const grouped = groupQuestionsByBlock(scaleQuestions);
    for (const [blockId, questions] of grouped) {
      if (questions.length === 0) continue;
      const { data, layout } = getBlockSummaryChartData(blockId, questions, survey.scaleAnalytics);
      await Plotly.default.newPlot(container, data, layout as any, { displayModeBar: false });
      const imageData = await Plotly.default.toImage(container, { format: 'png', width: 1600, height: 900 });
      const base64Data = imageData.replace('data:image/png;base64,', '');
      const fileName = blockId !== null ? `blocco_${blockId}_medie.png` : 'blocco_senza_numero_medie.png';
      zip.file(fileName, base64Data, { base64: true });
      processed++; onProgress?.(processed, total);
    }

    const content = await zip.generateAsync({ type: 'blob' });
    saveAs(content, `Charts_${survey.metadata.fileName.replace('.csv', '')}.zip`);
  } finally {
    Plotly.default.purge(container);
    document.body.removeChild(container);
  }
}
