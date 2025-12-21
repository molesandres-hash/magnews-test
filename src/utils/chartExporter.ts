import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import type { ParsedSurvey, QuestionInfo, ScaleAnalytics } from '@/types/survey';
import { groupQuestionsByBlock, getBlockDisplayName } from './analytics';
import { getShortQuestionText } from './headerNormalizer';

const SCALE_ORDER = ['10', '9', '8', '7', '6', '5', '4', '3', '2', '1', 'N/A'];
const CHART_COLORS = [
  '#3B82F6', '#3B82F6', '#3B82F6', '#3B82F6', '#3B82F6',
  '#60A5FA', '#60A5FA', '#60A5FA', '#93C5FD', '#BFDBFE', '#94A3B8'
];

/**
 * Generate distribution chart data for a scale question
 */
export function getDistributionChartData(analytics: ScaleAnalytics, question: QuestionInfo) {
  const values = SCALE_ORDER.map(key => analytics.counts[key] || 0);
  
  return {
    data: [{
      x: SCALE_ORDER,
      y: values,
      type: 'bar' as const,
      marker: {
        color: CHART_COLORS,
        line: {
          color: '#1E40AF',
          width: 1,
        },
      },
      text: values.map(v => v.toString()),
      textposition: 'outside' as const,
    }],
    layout: {
      title: {
        text: `${question.questionKey || ''} - ${getShortQuestionText(question.questionText, 60)}`,
        font: { size: 16, color: '#1E293B' },
      },
      annotations: [{
        x: 0.5,
        y: 1.08,
        xref: 'paper' as const,
        yref: 'paper' as const,
        text: `Media: ${analytics.mean.toFixed(2)} | Risposte valide: ${analytics.validResponses}/${analytics.totalRespondents}`,
        showarrow: false,
        font: { size: 12, color: '#64748B' },
      }],
      xaxis: {
        title: { text: 'Valutazione', font: { size: 12 } },
        tickfont: { size: 12 },
      },
      yaxis: {
        title: { text: 'Conteggio', font: { size: 12 } },
        tickfont: { size: 12 },
      },
      margin: { t: 80, r: 40, b: 60, l: 60 },
      paper_bgcolor: '#FFFFFF',
      plot_bgcolor: '#F8FAFC',
    },
  };
}

/**
 * Generate block summary chart data
 */
export function getBlockSummaryChartData(
  blockId: number | null,
  questions: QuestionInfo[],
  scaleAnalytics: Map<string, ScaleAnalytics>
) {
  const blockQuestions = questions.filter(
    q => q.blockId === blockId && q.type === 'scale_1_10_na'
  );

  const labels = blockQuestions.map(q => 
    q.questionKey ? `${q.questionKey}` : getShortQuestionText(q.questionText, 20)
  );
  
  const means = blockQuestions.map(q => {
    const analytics = scaleAnalytics.get(q.id);
    return analytics ? analytics.mean : 0;
  });

  const colors = means.map(m => {
    if (m >= 8) return '#22C55E';
    if (m >= 6) return '#3B82F6';
    if (m >= 4) return '#F59E0B';
    return '#EF4444';
  });

  return {
    data: [{
      x: labels,
      y: means,
      type: 'bar' as const,
      marker: {
        color: colors,
        line: {
          color: '#1E293B',
          width: 1,
        },
      },
      text: means.map(m => m.toFixed(2)),
      textposition: 'outside' as const,
    }],
    layout: {
      title: {
        text: `${getBlockDisplayName(blockId)} - Medie per domanda`,
        font: { size: 18, color: '#1E293B' },
      },
      xaxis: {
        title: { text: 'Domanda', font: { size: 12 } },
        tickfont: { size: 10 },
        tickangle: -45,
      },
      yaxis: {
        title: { text: 'Media', font: { size: 12 } },
        range: [0, 10],
        tickfont: { size: 12 },
      },
      margin: { t: 60, r: 40, b: 120, l: 60 },
      paper_bgcolor: '#FFFFFF',
      plot_bgcolor: '#F8FAFC',
    },
  };
}

/**
 * Generate and download charts as ZIP
 */
export async function generateChartsZip(
  survey: ParsedSurvey,
  onProgress?: (current: number, total: number) => void
): Promise<void> {
  const zip = new JSZip();
  const scaleQuestions = survey.questions.filter(q => q.type === 'scale_1_10_na');
  
  // Dynamic import Plotly
  const Plotly = await import('plotly.js-dist-min');
  
  // Create a hidden div for rendering
  const container = document.createElement('div');
  container.style.position = 'absolute';
  container.style.left = '-9999px';
  container.style.width = '1600px';
  container.style.height = '900px';
  document.body.appendChild(container);

  try {
    let processed = 0;
    const total = scaleQuestions.length + new Set(scaleQuestions.map(q => q.blockId)).size;

    // Generate individual question charts
    for (const question of scaleQuestions) {
      const analytics = survey.scaleAnalytics.get(question.id);
      if (!analytics) continue;

      const { data, layout } = getDistributionChartData(analytics, question);
      
      await Plotly.default.newPlot(container, data, layout as any, { 
        displayModeBar: false 
      });
      
      const imageData = await Plotly.default.toImage(container, {
        format: 'png',
        width: 1600,
        height: 900,
      });

      // Convert base64 to blob
      const base64Data = imageData.replace('data:image/png;base64,', '');
      const fileName = question.questionKey 
        ? `${question.questionKey.replace('.', '_')}.png`
        : `domanda_${processed + 1}.png`;
      
      zip.file(fileName, base64Data, { base64: true });
      
      processed++;
      onProgress?.(processed, total);
    }

    // Generate block summary charts
    const grouped = groupQuestionsByBlock(scaleQuestions);
    for (const [blockId, questions] of grouped) {
      if (questions.length === 0) continue;

      const { data, layout } = getBlockSummaryChartData(blockId, questions, survey.scaleAnalytics);
      
      await Plotly.default.newPlot(container, data, layout as any, { 
        displayModeBar: false 
      });
      
      const imageData = await Plotly.default.toImage(container, {
        format: 'png',
        width: 1600,
        height: 900,
      });

      const base64Data = imageData.replace('data:image/png;base64,', '');
      const fileName = blockId !== null 
        ? `blocco_${blockId}_medie.png`
        : 'blocco_senza_numero_medie.png';
      
      zip.file(fileName, base64Data, { base64: true });
      
      processed++;
      onProgress?.(processed, total);
    }

    // Generate and download zip
    const content = await zip.generateAsync({ type: 'blob' });
    saveAs(content, `Charts_${survey.metadata.fileName.replace('.csv', '')}.zip`);

  } finally {
    // Cleanup
    Plotly.default.purge(container);
    document.body.removeChild(container);
  }
}
